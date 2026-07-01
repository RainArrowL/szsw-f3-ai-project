#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
深证信数据服务平台 (webapi.cninfo.com.cn) API 接口封装
支持Token生成、刷新，以及三大财务报表的请求
"""

import time
import json
import hashlib
import logging
import requests
from typing import Dict, List, Optional, Any
from config import config

logger = logging.getLogger(__name__)


class CninfoAPI:
    """深证信数据服务平台API客户端"""

    # Token端点
    TOKEN_URL = "http://webapi.cninfo.com.cn/api-cloud-platform/oauth2/token"

    # 财务报表API端点（中文版）
    # 非金融企业资产负债表
    BALANCE_SHEET_URL = "http://webapi.cninfo.com.cn/api/stock/p_finance0001"
    # 非金融企业利润表
    INCOME_STATEMENT_URL = "http://webapi.cninfo.com.cn/api/stock/p_finance0002"
    # 非金融企业现金流量表
    CASH_FLOW_URL = "http://webapi.cninfo.com.cn/api/stock/p_finance0003"

    # 财务统计API
    DIVIDEND_URL = "http://webapi.cninfo.com.cn/api/stock/p_stock2201"

    # 备选API：东方财富数据源（通过akshare等免费方式）
    EASTMONEY_BALANCE = "https://datacenter-web.eastmoney.com/api/data/v1/get"
    EASTMONEY_INCOME = "https://datacenter-web.eastmoney.com/api/data/v1/get"
    EASTMONEY_CASHFLOW = "https://datacenter-web.eastmoney.com/api/data/v1/get"

    def __init__(self):
        self.access_key = config.access_key
        self.access_secret = config.access_secret
        self.base_url = config.base_url
        self.timeout = config.timeout
        self._token: Optional[str] = None
        self._token_expire_time: float = 0.0
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "zh-CN,zh;q=0.9",
        })

    # ==================== Token管理 ====================

    def _generate_signature(self, timestamp: str) -> str:
        """生成API签名"""
        raw = f"{self.access_key}{timestamp}{self.access_secret}"
        return hashlib.md5(raw.encode()).hexdigest().upper()

    def get_token(self) -> str:
        """
        获取访问Token（有效期为2小时，过期自动刷新）
        返回: token字符串
        """
        if self._token and time.time() < self._token_expire_time:
            return self._token

        if not config.is_credential_set():
            raise ValueError(
                "未配置深证信API凭证！\n"
                "请在 config.py 中设置 ACCESS_KEY 和 ACCESS_SECRET，\n"
                "或设置环境变量 CNINFO_ACCESS_KEY 和 CNINFO_ACCESS_SECRET。\n"
                "注册地址: http://webapi.cninfo.com.cn/"
            )

        timestamp = str(int(time.time() * 1000))
        signature = self._generate_signature(timestamp)

        headers = {
            "Content-Type": "application/json",
            "X-Ca-Key": self.access_key,
            "X-Ca-Signature": signature,
            "X-Ca-Timestamp": timestamp,
        }

        try:
            resp = self.session.post(self.TOKEN_URL, headers=headers, json={}, timeout=self.timeout)
            resp.raise_for_status()
            data = resp.json()
            self._token = data.get("access_token")
            if not self._token:
                raise ValueError(f"Token获取失败: {data}")
            # Token有效期2小时，提前5分钟刷新
            expires_in = data.get("expires_in", 7200)
            self._token_expire_time = time.time() + expires_in - 300
            logger.info("深证信API Token获取成功")
            return self._token
        except requests.RequestException as e:
            raise ConnectionError(f"Token请求失败: {e}")

    def _get_auth_headers(self) -> Dict[str, str]:
        """获取带认证的请求头"""
        token = self.get_token()
        return {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }

    # ==================== API通用请求方法 ====================

    def _api_request(
        self,
        url: str,
        method: str = "POST",
        params: Optional[Dict] = None,
        data: Optional[Dict] = None,
        json_data: Optional[Dict] = None,
        retry: int = 2,
    ) -> Dict[str, Any]:
        """
        统一API请求方法，带重试和Token自动刷新
        """
        headers = self._get_auth_headers()

        for attempt in range(retry + 1):
            try:
                resp = self.session.request(
                    method=method,
                    url=url,
                    params=params,
                    data=data,
                    json=json_data,
                    headers=headers,
                    timeout=self.timeout,
                )
                if resp.status_code == 401:
                    # Token过期，刷新后重试
                    self._token = None
                    self._token_expire_time = 0
                    headers = self._get_auth_headers()
                    continue
                if resp.status_code == 429:
                    wait = (attempt + 1) * 5
                    logger.warning(f"请求频率限制，等待{wait}秒...")
                    time.sleep(wait)
                    continue
                resp.raise_for_status()
                return resp.json()
            except requests.RequestException as e:
                if attempt < retry:
                    wait = (attempt + 1) * 2
                    logger.warning(f"请求失败({e})，{wait}秒后重试...")
                    time.sleep(wait)
                else:
                    raise ConnectionError(f"API请求失败: {e}")

        return {}

    # ==================== 财务报表数据获取 ====================

    def fetch_financial_report(
        self,
        stock_code: str,
        report_type: str,
        start_date: str,
        end_date: str,
        org_id: str = "",
    ) -> List[Dict[str, Any]]:
        """
        获取公司财务报表数据

        参数:
            stock_code: 股票代码，如 "000001" 或 "600519"（非上市公司可为空字符串）
            report_type: 报表类型，"balance"/"income"/"cashflow"
            start_date: 开始日期，格式 "YYYY-MM-DD" 或 "YYYYMMDD"
            end_date: 截止日期，格式 "YYYY-MM-DD" 或 "YYYYMMDD"
            org_id: 机构ID（非上市公司使用，替代 stock_code）

        返回:
            财务报表数据列表
        """
        url_map = {
            "balance": self.BALANCE_SHEET_URL,
            "income": self.INCOME_STATEMENT_URL,
            "cashflow": self.CASH_FLOW_URL,
        }
        url = url_map.get(report_type)
        if not url:
            raise ValueError(f"不支持的报表类型: {report_type}，请使用 balance/income/cashflow")

        # 标准化日期格式
        sd = start_date.replace("-", "")
        ed = end_date.replace("-", "")

        all_records = []
        page_num = 1

        while True:
            params = {
                "scode": stock_code,
                "sdate": sd,
                "edate": ed,
                "reportdate": "4",  # 4=年报
                "@limit": str(config.max_rows_per_page),
                "@orderby": "scode:asc,f001d:asc",
            }
            # 非上市公司：使用 orgId 替代 scode
            if org_id and not stock_code:
                params["scode"] = ""
                params["orgid"] = org_id
                params["@orderby"] = "orgid:asc,f001d:asc"

            result = self._api_request(url, method="POST", params=params)

            if not result:
                break

            records = result.get("records", [])
            if not records:
                break

            all_records.extend(records)

            # 检查是否还有更多数据
            total = result.get("total", 0)
            fetched = page_num * config.max_rows_per_page
            if fetched >= total:
                break
            page_num += 1
            time.sleep(config.request_interval)

        logger.info(f"获取{stock_code or org_id} {report_type} 报表数据: {len(all_records)} 条")
        return all_records

    # ==================== 备选数据源：东方财富（免费） ====================

    def fetch_from_eastmoney(
        self,
        stock_code: str,
        report_type: str,
        start_year: int,
        end_year: int,
    ) -> List[Dict[str, Any]]:
        """
        从东方财富网获取财务数据（免费公开数据源，作为备选方案）

        参数:
            stock_code: 股票代码，如 "000001"
            report_type: 报表类型名
            start_year: 开始年份
            end_year: 截止年份
        """
        report_map = {
            "balance": "RPT_DMSK_FN_BALANCE",
            "income": "RPT_DMSK_FN_INCOME",
            "cashflow": "RPT_DMSK_FN_CASHFLOW",
        }
        report_name = report_map.get(report_type)
        if not report_name:
            raise ValueError(f"不支持的报表类型: {report_type}")

        # 确定交易所代码
        if stock_code.startswith(("6", "68")):
            market_code = "SH"
        elif stock_code.startswith(("0", "3", "2")):
            market_code = "SZ"
        else:
            market_code = "SH"

        symbol = f"{market_code}{stock_code}"

        all_records = []
        for year in range(start_year, end_year + 1):
            params = {
                "reportName": report_name,
                "columns": "ALL",
                "pageSize": 50,
                "pageNumber": 1,
                "sortColumns": "NOTICE_DATE",
                "sortTypes": "-1",
                "filter": (
                    f'(SECURITY_CODE="{stock_code}")'
                    f'(REPORT_DATE>=\'{year}-01-01\')'
                    f'(REPORT_DATE<=\'{year}-12-31\')'
                ),
            }

            try:
                resp = self.session.get(
                    self.EASTMONEY_BALANCE,
                    params=params,
                    timeout=self.timeout,
                )
                if resp.status_code == 200:
                    data = resp.json()
                    if data.get("success") and data.get("result"):
                        records = data["result"].get("data", [])
                        all_records.extend(records)
                time.sleep(0.5)
            except Exception as e:
                logger.warning(f"东方财富数据获取失败({symbol} {year}): {e}")

        logger.info(f"东方财富获取 {stock_code} {report_type}: {len(all_records)} 条")
        return all_records

    # ==================== 备选数据源：新浪财经（免费，完整科目） ====================

    @staticmethod
    def _parse_sina_finance_html(html: str) -> List[Dict[str, Any]]:
        """解析新浪财经财务报表 HTML 表格，返回记录列表

        新浪财经表格结构：
          Row 0: 空
          Row 1: 表头（报表日期 | 2026-03-31 | 2025-12-31 | ...）
          Row 2: 空
          Row 3+: 数据行（科目名称 | val1 | val2 | ...）或分组标题行

        返回格式（与 cninfo API 一致）:
          [{"报告期": "2025-12-31", "货币资金": 123, ...}, ...]

        注：新浪财经数值单位为千元，返回时乘以 1000 转换为元
        """
        import re

        # 提取表格
        table_match = re.search(
            r'<table[^>]*id="[^"]*Table0"[^>]*>(.*?)</table>',
            html, re.DOTALL | re.IGNORECASE,
        )
        if not table_match:
            return []

        table_html = table_match.group(1)
        rows = re.findall(r'<tr[^>]*>(.*?)</tr>', table_html, re.DOTALL)
        if len(rows) < 4:
            return []

        # 解析表头行，获取日期列
        header_tds = re.findall(r'<td[^>]*>(.*?)</td>', rows[1], re.DOTALL)
        dates = []
        for td in header_tds[1:]:  # 跳过第一列"报表日期"
            date_text = re.sub(r'<[^>]+>', '', td).strip()
            dates.append(date_text)

        if not dates:
            return []

        # 按日期分组收集科目数据
        # records[date_index] = {科目: 值, "报告期": date}
        records = []
        for i, date in enumerate(dates):
            records.append({"报告期": date})

        for row in rows[2:]:
            tds = re.findall(r'<td[^>]*>(.*?)</td>', row, re.DOTALL)
            if len(tds) < 2:
                continue

            # 第一列是科目名称
            subject = re.sub(r'<[^>]+>', '', tds[0]).strip()
            if not subject:
                continue

            # 跳过纯分组标题行（如"流动资产"、"非流动负债"等）：后续列全部为空
            has_data = any(
                re.sub(r'<[^>]+>', '', td).strip()
                for td in tds[1:]
            )
            if not has_data:
                continue

            # 解析各期数值
            for i, td in enumerate(tds[1:]):
                if i >= len(dates):
                    break
                val_text = re.sub(r'<[^>]+>', '', td).strip()
                if val_text and val_text != '--':
                    try:
                        # 新浪财经单位是千元，转换为元
                        val = float(val_text.replace(',', '')) * 1000
                        records[i][subject] = val
                    except (ValueError, TypeError):
                        pass

        return records

    def fetch_from_sina(
        self,
        stock_code: str,
        report_type: str,
        start_year: int,
        end_year: int,
    ) -> List[Dict[str, Any]]:
        """从新浪财经获取财务报表数据（免费，科目完整）

        参数:
            stock_code: 股票代码
            report_type: 报表类型，"balance"/"income"/"cashflow"
            start_year: 开始年份
            end_year: 截止年份

        返回:
            财务报表数据列表，格式与 cninfo API 一致
        """
        url_map = {
            "balance": "vFD_BalanceSheet",
            "income": "vFD_ProfitStatement",
            "cashflow": "vFD_CashFlow",
        }
        endpoint = url_map.get(report_type)
        if not endpoint:
            raise ValueError(f"不支持的报表类型: {report_type}")

        url = (
            f"https://money.finance.sina.com.cn/corp/go.php/"
            f"{endpoint}/stockid/{stock_code}/ctrl/part/displaytype/4.phtml"
        )

        try:
            resp = self.session.get(
                url,
                headers={
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                },
                timeout=self.timeout,
            )
            if resp.status_code != 200:
                logger.warning(f"新浪财经请求失败 {url}: HTTP {resp.status_code}")
                return []

            resp.encoding = "gbk"
            html = resp.text

            all_records = self._parse_sina_finance_html(html)

            # 按年份过滤
            filtered = []
            for rec in all_records:
                report_date = rec.get("报告期", "")
                if not report_date:
                    continue
                try:
                    year = int(report_date[:4])
                except (ValueError, TypeError):
                    continue
                if start_year <= year <= end_year:
                    filtered.append(rec)

            logger.info(f"新浪财经获取 {stock_code} {report_type}: {len(filtered)} 条")
            return filtered

        except Exception as e:
            logger.warning(f"新浪财经数据获取失败({stock_code} {report_type}): {e}")
            return []

    # ==================== 分红数据 ====================

    def fetch_dividend_data(
        self,
        stock_code: str,
        start_date: str = "2010-01-01",
        end_date: str = "2099-12-31",
    ) -> List[Dict[str, Any]]:
        """
        获取公司分红转增信息（含派息日）

        数据源: 巨潮资讯 p_stock2201 接口
        返回字段包含: cash_divi_date(派现日), ex_divi_date(除权除息日),
                     right_reg_date(股权登记日), cash_divi_rmb(派现含税),
                     advance_date(预案公布日), event_procedure(事件进程)

        参数:
            stock_code: 股票代码
            start_date: 开始日期，格式 "YYYY-MM-DD"
            end_date: 截止日期，格式 "YYYY-MM-DD"

        返回:
            分红数据列表
        """
        sd = start_date.replace("-", "") if start_date else "20100101"
        ed = end_date.replace("-", "") if end_date else "20991231"

        all_records = []
        page_num = 1

        while True:
            params = {
                "scode": stock_code,
                "sdate": sd,
                "edate": ed,
                "@limit": str(config.max_rows_per_page),
                "@orderby": "advance_date:desc",
            }

            try:
                result = self._api_request(self.DIVIDEND_URL, method="POST", params=params)
            except Exception as e:
                logger.warning(f"获取{stock_code}分红数据失败: {e}")
                break

            if not result:
                break

            records = result.get("records", [])
            if not records:
                break

            all_records.extend(records)

            total = result.get("total", 0)
            fetched = page_num * config.max_rows_per_page
            if fetched >= total:
                break
            page_num += 1
            time.sleep(config.request_interval)

        logger.info(f"获取{stock_code}分红数据: {len(all_records)} 条")
        return all_records

    # ==================== 股票搜索 ====================

    def search_stock(self, keyword: str) -> List[Dict[str, str]]:
        """
        搜索股票信息

        参数:
            keyword: 股票代码或名称关键词

        返回:
            股票信息列表，每项包含 code、name、orgId
        """
        url = "http://www.cninfo.com.cn/new/information/topSearch/detailOfQuery"
        data = {
            "keyWord": keyword,
            "maxSecNum": 10,
            "maxListNum": 5,
        }
        headers = {
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Accept": "application/json,text/plain,*/*",
            "Origin": "http://www.cninfo.com.cn",
            "Referer": "http://www.cninfo.com.cn/",
        }

        try:
            resp = self.session.post(url, data=data, headers=headers, timeout=self.timeout)
            if resp.status_code == 200:
                body = resp.json()
                # 实际响应结构: {"keyBoardList": [...], "classifiedAnnouncements": [...]}
                results = []
                items = body.get("keyBoardList") or []
                for item in items:
                    results.append({
                        "code": item.get("code", ""),
                        "name": item.get("zwjc", ""),
                        "orgId": item.get("orgId", ""),
                        "market": item.get("plate", ""),
                    })
                return results
            else:
                logger.warning(f"股票搜索HTTP错误: {resp.status_code}")
        except Exception as e:
            logger.warning(f"股票搜索失败: {e}")
        return []


# 全局API实例
_api_instance: Optional[CninfoAPI] = None


def get_api() -> CninfoAPI:
    """获取API单例"""
    global _api_instance
    if _api_instance is None:
        _api_instance = CninfoAPI()
    return _api_instance