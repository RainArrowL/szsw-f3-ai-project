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

    # 金融企业相关API（备用）
    # 按点时间查询的API（Point-in-Time）
    BALANCE_SHEET_PT_URL = "http://webapi.cninfo.com.cn/api/stock/p_finance0010"

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