#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
财务数据获取协调层
- 字段中英文映射
- 多渠道数据获取策略（优先cninfo API，回退到免费数据源）
- 数据清洗和标准化
"""

import time
import logging
from typing import Dict, List, Optional, Tuple, Any
from cninfo_api import CninfoAPI, get_api

logger = logging.getLogger(__name__)

# ==================== 三大财务报表字段中文映射 ====================

# 资产负债表（Balance Sheet）字段映射
BALANCE_SHEET_FIELDS_CN = {
    # 基本信息字段
    "F001D": "报告期",
    "F002V": "公告日期",
    "F003V": "公司代码",
    "F004V": "公司简称",
    "F005V": "报表类型",
    "F006V": "公告年份",
    # 资产类
    "F007N": "货币资金",
    "F008N": "交易性金融资产",
    "F009N": "应收票据",
    "F010N": "应收账款",
    "F011N": "预付款项",
    "F012N": "其他应收款",
    "F013N": "存货",
    "F014N": "其他流动资产",
    "F015N": "流动资产合计",
    "F016N": "可供出售金融资产",
    "F017N": "持有至到期投资",
    "F018N": "长期股权投资",
    "F019N": "投资性房地产",
    "F020N": "固定资产净额",
    "F021N": "在建工程",
    "F022N": "无形资产",
    "F023N": "商誉",
    "F024N": "长期待摊费用",
    "F025N": "递延所得税资产",
    "F026N": "其他非流动资产",
    "F027N": "非流动资产合计",
    "F028N": "资产总计",
    # 负债类
    "F029N": "短期借款",
    "F030N": "应付票据",
    "F031N": "应付账款",
    "F032N": "预收款项",
    "F033N": "应付职工薪酬",
    "F034N": "应交税费",
    "F035N": "应付利息",
    "F036N": "其他应付款",
    "F037N": "一年内到期的非流动负债",
    "F038N": "其他流动负债",
    "F039N": "流动负债合计",
    "F040N": "长期借款",
    "F041N": "应付债券",
    "F042N": "长期应付款",
    "F043N": "递延所得税负债",
    "F044N": "其他非流动负债",
    "F045N": "非流动负债合计",
    "F046N": "负债合计",
    # 所有者权益
    "F047N": "实收资本（或股本）",
    "F048N": "资本公积",
    "F049N": "盈余公积",
    "F050N": "未分配利润",
    "F051N": "归属于母公司股东权益合计",
    "F052N": "少数股东权益",
    "F053N": "股东权益合计",
    "F054N": "负债和股东权益合计",
    # 东方财富（EastMoney）实际字段
    "SECUCODE": "证券代码",
    "SECURITY_CODE": "股票代码",
    "SECURITY_NAME_ABBR": "证券简称",
    "INDUSTRY_CODE": "行业代码",
    "ORG_CODE": "机构代码",
    "INDUSTRY_NAME": "行业名称",
    "MARKET": "交易市场",
    "SECURITY_TYPE_CODE": "证券类型代码",
    "TRADE_MARKET_CODE": "交易市场代码",
    "DATE_TYPE_CODE": "日期类型代码",
    "REPORT_TYPE_CODE": "报告类型代码",
    "DATA_STATE": "数据状态",
    "REPORT_DATE": "报告日期",
    "NOTICE_DATE": "公告日期",
    "TOTAL_ASSETS": "资产总计",
    "FIXED_ASSET": "固定资产",
    "MONETARYFUNDS": "货币资金",
    "MONETARYFUNDS_RATIO": "货币资金同比",
    "ACCOUNTS_RECE": "应收账款",
    "ACCOUNTS_RECE_RATIO": "应收账款同比",
    "INVENTORY": "存货",
    "INVENTORY_RATIO": "存货同比",
    "TOTAL_LIABILITIES": "负债合计",
    "ACCOUNTS_PAYABLE": "应付账款",
    "ACCOUNTS_PAYABLE_RATIO": "应付账款同比",
    "ADVANCE_RECEIVABLES": "预收款项",
    "ADVANCE_RECEIVABLES_RATIO": "预收款项同比",
    "TOTAL_EQUITY": "股东权益合计",
    "TOTAL_EQUITY_RATIO": "股东权益合计同比",
    "TOTAL_ASSETS_RATIO": "资产总计同比",
    "TOTAL_LIAB_RATIO": "负债合计同比",
    "CURRENT_RATIO": "流动比率",
    "DEBT_ASSET_RATIO": "资产负债率",
    "CASH_DEPOSIT_PBC": "现金及存放央行款项",
    "CDP_RATIO": "现金及存放央行款项同比",
    "LOAN_ADVANCE": "发放贷款及垫款",
    "LOAN_ADVANCE_RATIO": "发放贷款及垫款同比",
    "AVAILABLE_SALE_FINASSET": "可供出售金融资产",
    "ASF_RATIO": "可供出售金融资产同比",
    "LOAN_PBC": "向央行借款",
    "LOAN_PBC_RATIO": "向央行借款同比",
    "ACCEPT_DEPOSIT": "吸收存款",
    "ACCEPT_DEPOSIT_RATIO": "吸收存款同比",
    "SELL_REPO_FINASSET": "卖出回购金融资产款",
    "SRF_RATIO": "卖出回购金融资产款同比",
    "SETTLE_EXCESS_RESERVE": "结算备付金",
    "SER_RATIO": "结算备付金同比",
    "BORROW_FUND": "拆入资金",
    "BORROW_FUND_RATIO": "拆入资金同比",
    "AGENT_TRADE_SECURITY": "代理买卖证券款",
    "ATS_RATIO": "代理买卖证券款同比",
    "PREMIUM_RECE": "应收保费",
    "PREMIUM_RECE_RATIO": "应收保费同比",
    "SHORT_LOAN": "短期借款",
    "SHORT_LOAN_RATIO": "短期借款同比",
    "ADVANCE_PREMIUM": "预收保费",
    "ADVANCE_PREMIUM_RATIO": "预收保费同比",
    "OPERATE_EXPENSE": "营业支出",
    "SALE_EXPENSE": "销售费用",
    "MANAGE_EXPENSE": "管理费用",
    "FINANCE_EXPENSE": "财务费用",
    "INVEST_INCOME": "投资收益",
    "SURRENDER_VALUE": "退保金",
    "COMPENSATE_EXPENSE": "赔付支出",
    "INTEREST_NI": "利息净收入",
    "FEE_COMMISSION_NI": "手续费及佣金净收入",
    "EARNED_PREMIUM": "已赚保费",
    "TOTAL_OPERATE_INCOME": "营业总收入",
    "TOTAL_OPERATE_COST": "营业总成本",
    "OPERATE_COST": "营业成本",
    "OPERATE_PROFIT": "营业利润",
    "TOTAL_PROFIT": "利润总额",
    "NETPROFIT": "净利润",
    "INCOME_TAX": "所得税费用",
    "PARENT_NETPROFIT": "归属于母公司股东的净利润",
    "MINORITY_INTEREST": "少数股东损益",
    "BASIC_EPS": "基本每股收益",
    "DILUTED_EPS": "稀释每股收益",
    "OPERATE_TAX_ADD": "税金及附加",
    "SUM_CI": "综合收益总额",
    "MONETARYFUNDS_ORG": "货币资金(原始)",
    "TOTAL_NONCUR_ASSETS": "非流动资产合计",
    "TOTAL_CUR_LIAB": "流动负债合计",
    "TOTAL_NONCUR_LIAB": "非流动负债合计",
    # 增长率/比率字段保留
    "MONETARYFUNDS_RATIO": "货币资金增长率(%)",
    "INVENTORY_RATIO": "存货增长率(%)",
}

# 利润表（Income Statement）字段映射
INCOME_STATEMENT_FIELDS_CN = {
    # 基本信息
    "F001D": "报告期",
    "F002V": "公告日期",
    "F003V": "公司代码",
    "F004V": "公司简称",
    "F005V": "报表类型",
    "F006V": "公告年份",
    # 收入
    "F007N": "营业总收入",
    "F008N": "营业收入",
    "F009N": "利息净收入",
    "F010N": "手续费及佣金净收入",
    # 成本
    "F011N": "营业总成本",
    "F012N": "营业成本",
    "F013N": "利息支出",
    "F014N": "税金及附加",
    "F015N": "销售费用",
    "F016N": "管理费用",
    "F017N": "研发费用",
    "F018N": "财务费用",
    # 投资收益
    "F019N": "公允价值变动收益",
    "F020N": "投资收益",
    "F021N": "汇兑收益",
    # 利润
    "F022N": "营业利润",
    "F023N": "营业外收入",
    "F024N": "营业外支出",
    "F025N": "利润总额",
    "F026N": "所得税费用",
    "F027N": "净利润",
    "F028N": "归属于母公司股东的净利润",
    "F029N": "少数股东损益",
    "F030N": "基本每股收益",
    "F031N": "稀释每股收益",
    "F032N": "综合收益总额",
    # 东方财富字段（通用元数据）
    "SECUCODE": "证券代码",
    "SECURITY_CODE": "股票代码",
    "SECURITY_NAME_ABBR": "证券简称",
    "INDUSTRY_CODE": "行业代码",
    "ORG_CODE": "机构代码",
    "INDUSTRY_NAME": "行业名称",
    "MARKET": "交易市场",
    "SECURITY_TYPE_CODE": "证券类型代码",
    "TRADE_MARKET_CODE": "交易市场代码",
    "DATE_TYPE_CODE": "日期类型代码",
    "REPORT_TYPE_CODE": "报告类型代码",
    "DATA_STATE": "数据状态",
    "REPORT_DATE": "报告日期",
    "NOTICE_DATE": "公告日期",
    # 核心字段
    "PARENT_NETPROFIT": "归属于母公司股东的净利润",
    "TOTAL_OPERATE_INCOME": "营业总收入",
    "TOTAL_OPERATE_COST": "营业总成本",
    "TOE_RATIO": "营业总成本同比",
    "OPERATE_COST": "营业成本",
    "OPERATE_EXPENSE": "营业支出",
    "OPERATE_EXPENSE_RATIO": "营业支出同比",
    "SALE_EXPENSE": "销售费用",
    "MANAGE_EXPENSE": "管理费用",
    "FINANCE_EXPENSE": "财务费用",
    "OPERATE_PROFIT": "营业利润",
    "TOTAL_PROFIT": "利润总额",
    "INCOME_TAX": "所得税费用",
    "OPERATE_INCOME": "营业收入",
    "INTEREST_NI": "利息净收入",
    "INTEREST_NI_RATIO": "利息净收入同比",
    "FEE_COMMISSION_NI": "手续费及佣金净收入",
    "FCN_RATIO": "手续费及佣金净收入同比",
    "OPERATE_TAX_ADD": "税金及附加",
    "MANAGE_EXPENSE_BANK": "业务及管理费(银行)",
    "FCN_CALCULATE": "手续费及佣金净收入(计算)",
    "INTEREST_NI_CALCULATE": "利息净收入(计算)",
    "EARNED_PREMIUM": "已赚保费",
    "EARNED_PREMIUM_RATIO": "已赚保费同比",
    "INVEST_INCOME": "投资收益",
    "SURRENDER_VALUE": "退保金",
    "COMPENSATE_EXPENSE": "赔付支出",
    "TOI_RATIO": "营业总收入同比",
    "OPERATE_PROFIT_RATIO": "营业利润同比",
    "PARENT_NETPROFIT_RATIO": "归属母公司净利润同比",
    "DEDUCT_PARENT_NETPROFIT": "扣非归属母公司净利润",
    "DPN_RATIO": "扣非归属母公司净利润同比",
}

# 现金流量表（Cash Flow Statement）字段映射
CASH_FLOW_FIELDS_CN = {
    # 基本信息
    "F001D": "报告期",
    "F002V": "公告日期",
    "F003V": "公司代码",
    "F004V": "公司简称",
    "F005V": "报表类型",
    "F006V": "公告年份",
    # 经营活动
    "F007N": "销售商品、提供劳务收到的现金",
    "F008N": "收到的税费返还",
    "F009N": "收到其他与经营活动有关的现金",
    "F010N": "经营活动现金流入小计",
    "F011N": "购买商品、接受劳务支付的现金",
    "F012N": "支付给职工以及为职工支付的现金",
    "F013N": "支付的各项税费",
    "F014N": "支付其他与经营活动有关的现金",
    "F015N": "经营活动现金流出小计",
    "F016N": "经营活动产生的现金流量净额",
    # 投资活动
    "F017N": "收回投资收到的现金",
    "F018N": "取得投资收益收到的现金",
    "F019N": "处置固定资产、无形资产和其他长期资产收回的现金净额",
    "F020N": "收到其他与投资活动有关的现金",
    "F021N": "投资活动现金流入小计",
    "F022N": "购建固定资产、无形资产和其他长期资产支付的现金",
    "F023N": "投资支付的现金",
    "F024N": "支付其他与投资活动有关的现金",
    "F025N": "投资活动现金流出小计",
    "F026N": "投资活动产生的现金流量净额",
    # 筹资活动
    "F027N": "吸收投资收到的现金",
    "F028N": "取得借款收到的现金",
    "F029N": "收到其他与筹资活动有关的现金",
    "F030N": "筹资活动现金流入小计",
    "F031N": "偿还债务支付的现金",
    "F032N": "分配股利、利润或偿付利息支付的现金",
    "F033N": "支付其他与筹资活动有关的现金",
    "F034N": "筹资活动现金流出小计",
    "F035N": "筹资活动产生的现金流量净额",
    # 汇总
    "F036N": "汇率变动对现金及现金等价物的影响",
    "F037N": "现金及现金等价物净增加额",
    "F038N": "期初现金及现金等价物余额",
    "F039N": "期末现金及现金等价物余额",
    # 东方财富字段（通用元数据）
    "SECUCODE": "证券代码",
    "SECURITY_CODE": "股票代码",
    "SECURITY_NAME_ABBR": "证券简称",
    "INDUSTRY_CODE": "行业代码",
    "ORG_CODE": "机构代码",
    "INDUSTRY_NAME": "行业名称",
    "MARKET": "交易市场",
    "SECURITY_TYPE_CODE": "证券类型代码",
    "TRADE_MARKET_CODE": "交易市场代码",
    "DATE_TYPE_CODE": "日期类型代码",
    "REPORT_TYPE_CODE": "报告类型代码",
    "DATA_STATE": "数据状态",
    "REPORT_DATE": "报告日期",
    "NOTICE_DATE": "公告日期",
    # 东方财富实际字段
    "NETCASH_OPERATE": "经营活动产生的现金流量净额",
    "NETCASH_OPERATE_RATIO": "经营活动现金流同比",
    "SALES_SERVICES": "销售商品、提供劳务收到的现金",
    "SALES_SERVICES_RATIO": "销售商品提供劳务收到的现金同比",
    "PAY_STAFF_CASH": "支付给职工以及为职工支付的现金",
    "PSC_RATIO": "支付给职工的现金同比",
    "NETCASH_INVEST": "投资活动产生的现金流量净额",
    "NETCASH_INVEST_RATIO": "投资活动现金流同比",
    "RECEIVE_INVEST_INCOME": "取得投资收益收到的现金",
    "RII_RATIO": "取得投资收益收到的现金同比",
    "CONSTRUCT_LONG_ASSET": "购建固定资产无形资产和其他长期资产支付的现金",
    "CLA_RATIO": "购建长期资产支付的现金同比",
    "NETCASH_FINANCE": "筹资活动产生的现金流量净额",
    "NETCASH_FINANCE_RATIO": "筹资活动现金流同比",
    "CCE_ADD": "现金及现金等价物净增加额",
    "CCE_ADD_RATIO": "现金及现金等价物净增加额同比",
    "CUSTOMER_DEPOSIT_ADD": "客户存款净增加额",
    "CDA_RATIO": "客户存款净增加额同比",
    "DEPOSIT_IOFI_OTHER": "存放同业和拆出资金净增加额",
    "DIO_RATIO": "存放同业和拆出资金同比",
    "LOAN_ADVANCE_ADD": "发放贷款及垫款净增加额",
    "LAA_RATIO": "发放贷款及垫款净增加额同比",
    "RECEIVE_INTEREST_COMMISSION": "收取利息和手续费净增加额",
    "RIC_RATIO": "收取利息和手续费同比",
    "INVEST_PAY_CASH": "投资支付的现金",
    "IPC_RATIO": "投资支付的现金同比",
    "BEGIN_CCE": "期初现金及现金等价物余额",
    "BEGIN_CCE_RATIO": "期初现金同比",
    "END_CCE": "期末现金及现金等价物余额",
    "END_CCE_RATIO": "期末现金同比",
    "RECEIVE_ORIGIC_PREMIUM": "收到原保险合同保费",
    "ROP_RATIO": "收到原保险合同保费同比",
    "PAY_ORIGIC_COMPENSATE": "支付原保险合同赔付款项",
    "POC_RATIO": "支付原保险合同赔付款项同比",
}

# 报表类型字段映射汇总
FIELD_MAPS = {
    "balance": BALANCE_SHEET_FIELDS_CN,
    "income": INCOME_STATEMENT_FIELDS_CN,
    "cashflow": CASH_FLOW_FIELDS_CN,
}

# 报表中文名称
REPORT_NAMES_CN = {
    "balance": "资产负债表",
    "income": "利润表",
    "cashflow": "现金流量表",
}


# ==================== 数据获取协调器 ====================

class FinancialDataFetcher:
    """财务数据获取协调器"""

    def __init__(self, api: Optional[CninfoAPI] = None):
        self.api = api or get_api()
        self._use_cninfo = True

    def set_use_cninfo(self, use: bool):
        """设置是否尝试使用cninfo API"""
        self._use_cninfo = use

    def fetch_company_data(
        self,
        stock_code: str,
        start_year: int,
        end_year: int,
    ) -> Dict[str, Dict[str, list]]:
        """
        获取单个公司指定年份范围的三大财务报表数据

        参数:
            stock_code: 股票代码
            start_year: 开始年份
            end_year: 截止年份

        返回:
            {
                "资产负债表": [{"报告期": "...", "货币资金": ..., ...}, ...],
                "利润表": [{"报告期": "...", "营业收入": ..., ...}, ...],
                "现金流量表": [{"报告期": "...", "经营活动...": ..., ...}, ...],
            }
        """
        start_date = f"{start_year}-01-01"
        end_date = f"{end_year}-12-31"

        result = {}

        for report_type in ["balance", "income", "cashflow"]:
            cn_name = REPORT_NAMES_CN[report_type]
            logger.info(f"正在获取 {stock_code} 的{cn_name}数据...")

            try:
                if self._use_cninfo:
                    try:
                        records = self.api.fetch_financial_report(
                            stock_code, report_type, start_date, end_date
                        )
                        translated = self._translate_fields(
                            records, report_type, is_eastmoney=False
                        )
                        result[cn_name] = translated
                        time.sleep(0.5)
                        continue
                    except Exception as e:
                        logger.warning(
                            f"cninfo API获取{cn_name}失败({e})，切换到免费数据源..."
                        )

                # 回退到东方财富免费数据源
                records = self.api.fetch_from_eastmoney(
                    stock_code, report_type, start_year, end_year
                )
                translated = self._translate_fields(
                    records, report_type, is_eastmoney=True
                )
                result[cn_name] = translated

            except Exception as e:
                logger.error(f"获取{stock_code} {cn_name}失败: {e}")
                result[cn_name] = []

            time.sleep(0.5)

        return result

    def fetch_multiple_companies(
        self,
        stock_codes: List[Tuple[str, str]],
        start_year: int,
        end_year: int,
        progress_callback=None,
    ) -> Dict[str, Dict]:
        """
        批量获取多个公司的财务数据

        参数:
            stock_codes: [(股票代码, 公司名称), ...]
            start_year: 开始年份
            end_year: 截止年份
            progress_callback: 进度回调函数

        返回:
            {公司名称: {"资产负债表": [...], "利润表": [...], "现金流量表": [...]}}
        """
        all_data = {}
        total = len(stock_codes)

        for idx, (code, name) in enumerate(stock_codes):
            display_name = f"{name}({code})"
            logger.info(f"正在处理 {idx + 1}/{total}: {display_name}")

            if progress_callback:
                progress_callback(idx + 1, total, display_name)

            data = self.fetch_company_data(code, start_year, end_year)
            all_data[display_name] = data

        return all_data

    def _translate_fields(
        self,
        records: list,
        report_type: str,
        is_eastmoney: bool = False,
    ) -> list:
        """
        将字段名翻译为中文

        参数:
            records: 原始数据记录
            report_type: 报表类型
            is_eastmoney: 是否东方财富数据

        返回:
            中文字段名后的数据
        """
        field_map = FIELD_MAPS.get(report_type, {})
        if not field_map:
            return records

        translated = []
        for record in records:
            new_record = {}
            for key, value in record.items():
                cn_key = field_map.get(key, key)
                # 处理数值类型转换
                if isinstance(value, str) and value.replace(".", "").replace("-", "").isdigit():
                    try:
                        value = float(value)
                    except (ValueError, TypeError):
                        pass
                new_record[cn_key] = value
            translated.append(new_record)

        return translated

    def get_stock_info(self, keyword: str) -> List[Dict[str, str]]:
        """搜索股票信息"""
        return self.api.search_stock(keyword)


# 全局数据获取器实例
_fetcher: Optional[FinancialDataFetcher] = None


def get_fetcher() -> FinancialDataFetcher:
    """获取数据获取器单例"""
    global _fetcher
    if _fetcher is None:
        _fetcher = FinancialDataFetcher()
    return _fetcher


# ==================== 企业名称解析（独立函数） ====================

def resolve_companies(companies: list) -> list:
    """
    解析企业名单，返回 (股票代码, 公司名称) 列表

    自动识别：
    - 纯数字 -> 股票代码
    - 非纯数字 -> 公司名称（搜索获取代码）
    """
    fetcher = FinancialDataFetcher()
    resolved = []

    for item in companies:
        # 如果是纯数字/字母，视为股票代码
        if item.replace(".", "").replace("-", "").replace(" ", "").isalnum() and any(c.isdigit() for c in item):
            # 去掉后缀如 .SH .SZ
            code = item.split(".")[0].split("-")[0].strip()
            # 尝试搜索确认存在
            try:
                results = fetcher.get_stock_info(code)
                if results:
                    resolved.append((results[0]["code"], results[0]["name"]))
                else:
                    resolved.append((code, code))
            except Exception:
                resolved.append((code, code))
        else:
            # 视为公司名称
            try:
                results = fetcher.get_stock_info(item)
                if results:
                    best = results[0]
                    resolved.append((best["code"], best["name"]))
            except Exception:
                pass  # 跳过无法识别的

    return resolved