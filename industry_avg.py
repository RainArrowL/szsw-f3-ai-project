#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
行业平均值计算模块
按东方财富行业分类/CSRC(证监会)行业分类，获取同行业大类其他公司的财务数据，计算行业平均值。

说明：财务数据API中返回的 INDUSTRY_NAME 和 INDUSTRY_CODE 为东方财富申万行业分类（大类层级）。
用户要求的"国标行业"在公开免费API中无直接映射，我们先用申万行业大类（通常与CSRC大类高度相关）
作为行业分组依据。未来可加入CSRC证监会行业分类映射。
"""

import logging
import time
import concurrent.futures
from collections import defaultdict
from typing import Dict, List, Optional, Tuple, Set

from cninfo_api import get_api, CninfoAPI

logger = logging.getLogger(__name__)

# 全A行业列表缓存 (避免每次重复请求)
_stock_list_cache: Optional[List[Dict]] = None

# 金融计算时忽略的元数据/代码字段
SKIP_FIELDS_FOR_AVG = {
    "SECUCODE", "SECURITY_CODE", "SECURITY_NAME_ABBR", "ORG_CODE",
    "SECURITY_TYPE_CODE", "TRADE_MARKET_CODE", "DATE_TYPE_CODE",
    "REPORT_TYPE_CODE", "DATA_STATE", "MARKET", "REPORT_DATE",
    "NOTICE_DATE", "INDUSTRY_CODE", "INDUSTRY_NAME",
}

# ==================== 行业分类获取 ====================

def get_industry_stock_list(target_ind_code: str = None, target_ind_name: str = None) -> List[Dict]:
    """
    从东方财富获取全A股及行业分类信息，筛选同行业股票

    使用 datacenter-web API（RPT_DMSK_FN_BALANCE），查询最新报告期的全A股票行业信息。
    结果会缓存，同一次运行只请求一次。

    参数：
        target_ind_code: 目标行业代码
        target_ind_name: 目标行业名称

    返回：
        同行业股票列表: [{"code": "600519", "name": "贵州茅台", "industry": "白酒Ⅱ"}, ...]
    """
    global _stock_list_cache

    api = get_api()

    # 使用缓存
    if _stock_list_cache is not None:
        all_stocks = _stock_list_cache
    else:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept": "application/json, text/plain, */*",
            "Referer": "https://data.eastmoney.com/",
        }

        url = "https://datacenter-web.eastmoney.com/api/data/v1/get"

        all_stocks = []
        page = 1
        max_pages = 30  # 安全限制

        while page <= max_pages:
            params = {
                "reportName": "RPT_DMSK_FN_BALANCE",
                "columns": "SECURITY_CODE,SECURITY_NAME_ABBR,INDUSTRY_CODE,INDUSTRY_NAME",
                "pageSize": 500,
                "pageNumber": page,
                "sortColumns": "SECURITY_CODE",
                "sortTypes": "1",
                "filter": "(REPORT_DATE='2023-12-31')",
                "source": "WEB",
                "client": "WEB",
            }

            try:
                resp = api.session.get(url, params=params, headers=headers, timeout=30)
                if resp.status_code != 200:
                    logger.warning(f"获取行业列表第{page}页 HTTP {resp.status_code}")
                    break
                data = resp.json()
                if not data.get("success") or not data.get("result"):
                    break

                rows = data["result"].get("data", [])
                if not rows:
                    break

                for row in rows:
                    all_stocks.append({
                        "code": row.get("SECURITY_CODE", ""),
                        "name": row.get("SECURITY_NAME_ABBR", ""),
                        "industry": row.get("INDUSTRY_NAME", ""),
                        "ind_code": str(row.get("INDUSTRY_CODE", "")),
                    })

                total = data["result"].get("count", 0)
                fetched = page * 500
                if fetched >= total:
                    break

                page += 1
                time.sleep(0.3)

            except Exception as e:
                logger.warning(f"获取行业列表第{page}页失败: {e}")
                break

        logger.info(f"获取到 {len(all_stocks)} 只A股行业信息")
        _stock_list_cache = all_stocks

    # 如果没传筛选条件，返回全部
    if not target_ind_code and not target_ind_name:
        return all_stocks

    # 筛选同行业股票
    same_industry = []
    for s in all_stocks:
        match = False
        if target_ind_name and s.get("industry") == target_ind_name:
            match = True
        if target_ind_code and s.get("ind_code") == str(target_ind_code):
            match = True
        if match:
            same_industry.append(s)

    logger.info(f"行业 [{target_ind_name or target_ind_code}] 共有 {len(same_industry)} 家公司")
    return same_industry


def get_company_industry(stock_code: str) -> Optional[Tuple[str, str]]:
    """
    获取某公司的行业分类

    返回: (industry_code, industry_name) 或 None
    """
    api = get_api()
    # 尝试从财务数据中获取
    try:
        records = api.fetch_from_eastmoney(stock_code, "balance", 2023, 2023)
        if records:
            r = records[0]
            ind_code = r.get("INDUSTRY_CODE", "")
            ind_name = r.get("INDUSTRY_NAME", "")
            if ind_code or ind_name:
                return (str(ind_code) if ind_code else "", ind_name if ind_name else "")
    except Exception:
        pass

    # 备用：从股票列表API获取
    all_stocks = get_industry_stock_list()
    for s in all_stocks:
        if s["code"] == stock_code:
            return (s["ind_code"], s["industry"])

    return None


# ==================== 行业平均值计算 ====================

def compute_industry_averages(
    stock_code: str,
    report_type: str,
    start_year: int,
    end_year: int,
    target_ind_name: str = None,
    target_ind_code: str = None,
    progress_callback=None,
) -> Dict[str, List[Dict]]:
    """
    计算指定行业的各年度平均值

    参数：
        stock_code: 目标股票代码（仅用于获取行业信息）
        report_type: 报表类型 (balance/income/cashflow)
        start_year: 开始年份
        end_year: 截止年份
        target_ind_name: 目标行业名称
        target_ind_code: 目标行业代码

    返回：
        各年度行业平均值数据，格式同 fetch_company_data
    """
    api = get_api()

    # 1. 获取目标公司的行业分类
    if not target_ind_name and not target_ind_code:
        industry = get_company_industry(stock_code)
        if industry:
            target_ind_code, target_ind_name = industry
        else:
            logger.warning(f"未找到 {stock_code} 的行业分类，跳过行业均值")
            return {}

    if not target_ind_name:
        return {}

    logger.info(f"计算行业 [{target_ind_name}] 的财务报表平均值...")

    # 2. 获取同行业股票列表
    same_industry = get_industry_stock_list(
        target_ind_code=target_ind_code,
        target_ind_name=target_ind_name,
    )

    if not same_industry:
        logger.warning(f"未找到同行业股票，跳过行业均值")
        return {}

    # 去掉目标公司自身
    peer_codes = [s["code"] for s in same_industry if s["code"] != stock_code]

    if not peer_codes:
        logger.warning(f"行业 [{target_ind_name}] 仅包含目标公司自身，无法计算均值")
        return {}

    logger.info(f"计算行业均值: {len(peer_codes)} 只同业股票")

    # 3. 批量获取同行业公司财务数据（限制50只，避免过于耗时）
    max_peers = 50
    if len(peer_codes) > max_peers:
        logger.info(f"同行业股票过多，取前 {max_peers} 只")
        peer_codes = peer_codes[:max_peers]

    # 按年度收集所有行业公司数据
    year_data: Dict[str, List[Dict]] = defaultdict(list)

    for idx, code in enumerate(peer_codes):
        if progress_callback:
            progress_callback(idx + 1, len(peer_codes), f"获取同业: {code}")

        try:
            records = api.fetch_from_eastmoney(code, report_type, start_year, end_year)
            for record in records:
                # 提取报告年份
                date_str = record.get("REPORT_DATE", "")
                year = date_str[:4] if len(date_str) >= 4 else ""
                if year and start_year <= int(year) <= end_year:
                    year_data[year].append(record)
            time.sleep(0.3)
        except Exception as e:
            logger.debug(f"获取 {code} 财务数据失败: {e}")

    # 4. 计算各年度平均值
    result = {}
    from cninfo_fin_data import FIELD_MAPS

    field_map = FIELD_MAPS.get(report_type, {})

    for year in sorted(year_data.keys()):
        records = year_data[year]

        if not records:
            continue

        # 筛选同一年度的年报（优先12月31日）
        annual_records = [
            r for r in records
            if "12-31" in str(r.get("REPORT_DATE", ""))
        ]
        if not annual_records:
            annual_records = records

        # 收集所有数值字段（排除元数据字段）
        numeric_fields = set()
        for r in annual_records:
            for k, v in r.items():
                if k in SKIP_FIELDS_FOR_AVG:
                    continue
                if isinstance(v, (int, float)):
                    numeric_fields.add(k)
                elif isinstance(v, str) and v.replace(".", "").replace("-", "").replace("e", "").replace("E", "").isdigit():
                    numeric_fields.add(k)

        # 计算平均值
        avg_record = {}
        for field in numeric_fields:
            values = []
            for r in annual_records:
                v = r.get(field)
                if v is None:
                    continue
                try:
                    fv = float(v)
                    values.append(fv)
                except (ValueError, TypeError):
                    pass
            if values:
                avg_record[field] = sum(values) / len(values)

        # 翻译字段
        translated = {}
        for k, v in avg_record.items():
            cn_key = field_map.get(k, k)
            translated[cn_key] = v

        result[year] = [translated]

    logger.info(f"行业 [{target_ind_name}] 均值计算完成: {len(result)} 个年度")
    return result


def compute_all_industry_averages(
    stock_code: str,
    ind_name: str,
    ind_code: str,
    start_year: int,
    end_year: int,
    progress_callback=None,
) -> Dict[str, Dict[str, List[Dict]]]:
    """
    计算所有三种报表的行业平均值

    返回格式：
    {
        "行业均值_资产负债表": {...},
        "行业均值_利润表": {...},
        "行业均值_现金流量表": {...},
    }
    """
    result = {}
    report_types = {
        "balance": "行业均值_资产负债表",
        "income": "行业均值_利润表",
        "cashflow": "行业均值_现金流量表",
    }

    for rt, label in report_types.items():
        logger.info(f"计算 {label}...")
        avg = compute_industry_averages(
            stock_code=stock_code,
            report_type=rt,
            start_year=start_year,
            end_year=end_year,
            target_ind_name=ind_name,
            target_ind_code=ind_code,
            progress_callback=progress_callback,
        )
        # 转换为按年度分组的格式
        result[label] = avg

    return result


# ==================== 统计计算 ====================

def compute_statistics(values: List[float]) -> Dict[str, float]:
    """计算平均值和中位数"""
    if not values:
        return {"avg": 0, "median": 0, "count": 0}

    n = len(values)
    sorted_vals = sorted(values)
    if n % 2 == 1:
        median = sorted_vals[n // 2]
    else:
        median = (sorted_vals[n // 2 - 1] + sorted_vals[n // 2]) / 2

    return {
        "avg": sum(values) / n,
        "median": median,
        "count": n,
    }