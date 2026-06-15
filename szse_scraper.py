#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
深交所日度概况爬取模块
从 东方财富网 push2his API 获取深证综指(399106)日K线数据
提取深市的成交量和成交金额
"""

import time
import logging
from datetime import date, timedelta, datetime
from typing import List, Tuple, Optional, Callable
from collections import OrderedDict

import requests
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

logger = logging.getLogger(__name__)

# 东方财富 push2his API 端点
PUSH_API_URL = "https://push2his.eastmoney.com/api/qt/stock/kline/get"

# 请求头
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Referer": "https://quote.eastmoney.com/",
    "Accept": "application/json, text/plain, */*",
}


def fetch_year_data(year: int, progress_callback: Optional[Callable] = None) -> List[Tuple[str, float, float]]:
    """
    从东方财富 push2his API 获取深证综指(399106)日K线数据
    深证综指涵盖深市全部股票，其成交量和成交额近似深市合计
    
    fields2解析:
      f51: 日期, f52: 开盘, f53: 收盘, f54: 最高, f55: 最低
      f56: 成交量(手), f57: 成交额(元)
    """
    results = []

    params = {
        "secid": "0.399106",          # 深证综指
        "fields1": "f1,f2,f3,f4,f5,f6",
        "fields2": "f51,f52,f53,f54,f55,f56,f57",
        "klt": "101",                 # 日线
        "fqt": "0",                   # 不复权
        "beg": f"{year}0101",
        "end": f"{year}1231",
        "lmt": "300",
    }

    try:
        if progress_callback:
            progress_callback(0, 1, "正在连接东方财富行情接口...")

        resp = requests.get(PUSH_API_URL, params=params, headers=HEADERS, timeout=30)

        if resp.status_code != 200:
            logger.error(f"push2his API 返回状态码: {resp.status_code}")
            return results

        body = resp.json()
        if not body or "data" not in body or not body["data"]:
            logger.error("push2his API 返回数据为空")
            return results

        klines = body["data"].get("klines", [])
        if not klines:
            logger.warning("push2his API 返回空K线列表")
            return results

        total = len(klines)
        logger.info(f"获取到 {total} 条日K线数据")

        for idx, line in enumerate(klines):
            parts = line.split(",")
            if len(parts) < 7:
                continue

            date_str = parts[0]
            vol_shou = float(parts[5]) if parts[5] and parts[5] != "-" else 0
            amt_yuan = float(parts[6]) if parts[6] and parts[6] != "-" else 0

            if amt_yuan <= 0:
                continue

            # 单位转换
            # 成交量: 手 → 亿股  (1手=100股, 1亿=1e8)
            vol_yi = round(vol_shou * 100 / 1e8, 2)
            # 成交额: 元 → 亿元
            amt_yi = round(amt_yuan / 1e8, 2)

            results.append((date_str, vol_yi, amt_yi))

            if progress_callback:
                progress_callback(idx + 1, total, f"已获取 {date_str}")

        logger.info(f"处理完成，共 {len(results)} 个有效交易日")

    except requests.RequestException as e:
        logger.error(f"请求 push2his API 失败: {e}")
    except (ValueError, KeyError, IndexError) as e:
        logger.error(f"解析 push2his API 数据失败: {e}")
    except Exception as e:
        logger.error(f"未知错误: {e}")

    return results


def write_szse_excel(year: int, data: List[Tuple[str, float, float]], output_dir: str = "output") -> str:
    """
    将深交所日度概况数据写入Excel文件

    输出格式：
    - 行：日期（按周分组，每周后插入小计行）
    - 列：成交量（亿股）、成交金额（亿元）、日均成交金额、成交额×0.05%
    - 日均成交金额仅在每周小计行显示，计算公式：当周成交金额小计 ÷ 当周实际交易日天数
    """
    import os
    from datetime import date as date_type

    os.makedirs(output_dir, exist_ok=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{year}年深市日度概况"

    # ---- 样式定义 ----
    header_font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1A5276", end_color="1A5276", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    cell_alignment = Alignment(horizontal="center", vertical="center")
    date_alignment = Alignment(horizontal="center", vertical="center")

    subtotal_font = Font(name="微软雅黑", size=10, bold=True)
    subtotal_fill = PatternFill(start_color="EBF5FB", end_color="EBF5FB", fill_type="solid")
    subtotal_alignment = Alignment(horizontal="center", vertical="center")

    summary_font = Font(name="微软雅黑", size=10, bold=True)
    summary_fill = PatternFill(start_color="D5F5E3", end_color="D5F5E3", fill_type="solid")

    thin_border = Border(
        left=Side(style="thin", color="BDC3C7"),
        right=Side(style="thin", color="BDC3C7"),
        top=Side(style="thin", color="BDC3C7"),
        bottom=Side(style="thin", color="BDC3C7"),
    )

    # ---- 标题行 ----
    ws.merge_cells("A1:E1")
    title_cell = ws["A1"]
    title_cell.value = f"{year}年 深圳证券交易所 日度概况（深市合计）"
    title_cell.font = Font(name="微软雅黑", size=14, bold=True, color="1A5276")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    # ---- 表头 ----
    headers = ["日期", "成交量（亿股）", "成交金额（亿元）", "日均成交金额", "成交额×0.05%"]
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    ws.row_dimensions[2].height = 28

    # ---- 按周分组 ----
    weeks = OrderedDict()
    for d_str, vol, amt in data:
        dt = date_type.fromisoformat(d_str)
        iso_year, iso_week, _ = dt.isocalendar()
        week_key = f"{iso_year}-W{iso_week:02d}"
        if week_key not in weeks:
            weeks[week_key] = []
        weeks[week_key].append((d_str, vol, amt))

    # ---- 写入数据 ----
    current_row = 3
    total_vol = 0.0
    total_amt = 0.0
    total_days = 0

    for week_key, week_data in weeks.items():
        week_vol = 0.0
        week_amt = 0.0
        week_days = len(week_data)

        for d_str, vol, amt in week_data:
            fee_005 = round(amt * 0.0005, 2)

            ws.cell(row=current_row, column=1, value=d_str).alignment = date_alignment
            ws.cell(row=current_row, column=1).border = thin_border

            ws.cell(row=current_row, column=2, value=vol).alignment = cell_alignment
            ws.cell(row=current_row, column=2).number_format = "#,##0.00"
            ws.cell(row=current_row, column=2).border = thin_border

            ws.cell(row=current_row, column=3, value=amt).alignment = cell_alignment
            ws.cell(row=current_row, column=3).number_format = "#,##0.00"
            ws.cell(row=current_row, column=3).border = thin_border

            # 日均成交金额：每日行留空
            ws.cell(row=current_row, column=4).alignment = cell_alignment
            ws.cell(row=current_row, column=4).border = thin_border

            ws.cell(row=current_row, column=5, value=fee_005).alignment = cell_alignment
            ws.cell(row=current_row, column=5).number_format = "#,##0.00"
            ws.cell(row=current_row, column=5).border = thin_border

            ws.row_dimensions[current_row].height = 20

            week_vol += vol
            week_amt += amt
            current_row += 1

        # ---- 本周小计行 ----
        avg_daily_amt = round(week_amt / week_days, 2) if week_days > 0 else 0
        fee_005_sub = round(week_amt * 0.0005, 2)

        for col in range(1, 6):
            ws.cell(row=current_row, column=col).fill = subtotal_fill
            ws.cell(row=current_row, column=col).font = subtotal_font
            ws.cell(row=current_row, column=col).alignment = subtotal_alignment
            ws.cell(row=current_row, column=col).border = thin_border

        ws.cell(row=current_row, column=1, value=f"{week_key} 小计（{week_days}个交易日）")

        ws.cell(row=current_row, column=2, value=round(week_vol, 2)).number_format = "#,##0.00"
        ws.cell(row=current_row, column=3, value=round(week_amt, 2)).number_format = "#,##0.00"
        ws.cell(row=current_row, column=4, value=avg_daily_amt).number_format = "#,##0.00"
        ws.cell(row=current_row, column=5, value=fee_005_sub).number_format = "#,##0.00"

        ws.row_dimensions[current_row].height = 22

        total_vol += week_vol
        total_amt += week_amt
        total_days += week_days
        current_row += 1

    # ---- 年度汇总行 ----
    year_avg_amt = round(total_amt / total_days, 2) if total_days > 0 else 0
    year_fee_005 = round(total_amt * 0.0005, 2)

    for col in range(1, 6):
        ws.cell(row=current_row, column=col).fill = summary_fill
        ws.cell(row=current_row, column=col).font = summary_font
        ws.cell(row=current_row, column=col).alignment = subtotal_alignment
        ws.cell(row=current_row, column=col).border = thin_border

    ws.cell(row=current_row, column=1, value=f"{year}年 合计（{total_days}个交易日）")
    ws.cell(row=current_row, column=2, value=round(total_vol, 2)).number_format = "#,##0.00"
    ws.cell(row=current_row, column=3, value=round(total_amt, 2)).number_format = "#,##0.00"
    ws.cell(row=current_row, column=4, value=year_avg_amt).number_format = "#,##0.00"
    ws.cell(row=current_row, column=5, value=year_fee_005).number_format = "#,##0.00"
    ws.row_dimensions[current_row].height = 24

    # ---- 列宽 ----
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 16

    # ---- 冻结窗格 ----
    ws.freeze_panes = "A3"

    # 保存
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"深交所日度概况_{year}年_{timestamp}.xlsx"
    filepath = os.path.join(output_dir, filename)
    wb.save(filepath)
    logger.info(f"Excel已保存: {filepath} (共{len(data)}个交易日, {len(weeks)}周)")

    return filepath


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

    def progress_cb(cur, total, msg):
        print(f"\r[{cur}/{total}] {msg}", end="")

    data = fetch_year_data(2025, progress_callback=progress_cb)
    print(f"\n共获取 {len(data)} 个交易日数据")
    if data:
        path = write_szse_excel(2025, data)
        print(f"已输出: {path}")