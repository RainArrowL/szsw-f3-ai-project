#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
金融机构处罚信息爬取模块
从三大金融监管机构爬取行政处罚信息：
  - NFRA（金融监管总局）
  - PBC（人民银行）
  - CSRC（证监会）

功能：
  - fetch_all_penalty(max_per_source=5)  爬取三来源处罚数据
  - write_penalty_excel(all_data, output_dir)  写入Excel
"""

import re
import os
import logging
import requests
from typing import Dict, List, Optional
from urllib.parse import urljoin

logger = logging.getLogger(__name__)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9",
}

# ── NFRA 处罚信息 ───────────────────────────────────
NFRA_BASE_URL = "https://www.nfra.gov.cn"
NFRA_PENALTY_ITEM_ID = "915"  # 行政处罚栏目ID
NFRA_PENALTY_PARENT_ID = "913"

# ── PBC 处罚信息 ───────────────────────────────────
PBC_BASE_URL = "https://www.pbc.gov.cn"
# PBC 行政处罚公示入口页面（可能因网站改版而变化，代码会做多重尝试）
PBC_PENALTY_URLS = [
    "https://www.pbc.gov.cn/zhengwugongkai/4081330/4081340/4081350/index.html",
    "https://www.pbc.gov.cn/zhengwugongkai/4081330/4081340/4081390/index.html",
    "https://www.pbc.gov.cn/",
]

# ── CSRC 处罚信息 ───────────────────────────────────
CSRC_BASE_URL = "https://www.csrc.gov.cn"
CSRC_PENALTY_URL = "https://www.csrc.gov.cn/csrc/c100028/common_list.shtml"


# ==================== 通用请求工具 ====================

def _fetch_html(url: str, timeout: int = 30) -> Optional[str]:
    """获取HTML页面内容"""
    try:
        resp = requests.get(url, headers=HEADERS, timeout=timeout)
        resp.encoding = "utf-8"
        if resp.status_code == 200:
            return resp.text
        logger.warning(f"请求失败 {url}: HTTP {resp.status_code}")
    except requests.RequestException as e:
        logger.warning(f"请求异常 {url}: {e}")
    return None


def _fetch_json(url: str, timeout: int = 30) -> Optional[dict]:
    """获取JSON API响应"""
    try:
        json_headers = {**HEADERS, "Accept": "application/json, text/javascript, */*; q=0.01"}
        resp = requests.get(url, headers=json_headers, timeout=timeout)
        resp.encoding = "utf-8"
        if resp.status_code == 200:
            return resp.json()
        logger.warning(f"JSON API 请求失败 {url}: HTTP {resp.status_code}")
    except requests.RequestException as e:
        logger.warning(f"JSON API 请求异常 {url}: {e}")
    except ValueError as e:
        logger.warning(f"JSON 解析失败 {url}: {e}")
    return None


# ==================== 内容提取工具 ====================

def _strip_html_tags(text: str) -> str:
    """去除HTML标签，保留纯文本（兼容Word生成的HTML）"""
    if not text:
        return ""
    # 移除 Word 特有的 XML 声明和条件注释
    text = re.sub(r"<!--\[if[^>]*>.*?<!\[endif\]-->", "", text, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r"<!--.*?-->", "", text, flags=re.DOTALL)
    text = re.sub(r"<\?xml[^>]*\?>", "", text, flags=re.IGNORECASE)
    # 移除 script 和 style
    text = re.sub(r"<script[^>]*>.*?</script>", "", text, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r"<style[^>]*>.*?</style>", "", text, flags=re.DOTALL | re.IGNORECASE)
    # 移除 head 段（含 meta、title 等）
    text = re.sub(r"<head[^>]*>.*?</head>", "", text, flags=re.DOTALL | re.IGNORECASE)
    # 块级标签换行
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</?(?:p|div|tr|h\d|li|table)[^>]*>", "\n", text, flags=re.IGNORECASE)
    # 移除所有其余标签
    text = re.sub(r"<[^>]+>", "", text)
    # HTML 实体
    text = re.sub(r"&nbsp;", " ", text)
    text = re.sub(r"&lt;", "<", text)
    text = re.sub(r"&gt;", ">", text)
    text = re.sub(r"&amp;", "&", text)
    text = re.sub(r"&quot;", '"', text)
    # 移除多余空行
    text = re.sub(r"\n{3,}", "\n\n", text)
    # 移除行首尾空白
    text = re.sub(r"^[ \t]+", "", text, flags=re.MULTILINE)
    return text.strip()


def _extract_title_from_html(html: str) -> str:
    """从HTML中提取页面标题"""
    m = re.search(r"<title>(.*?)</title>", html, re.IGNORECASE | re.DOTALL)
    if m:
        return _strip_html_tags(m.group(1))
    return ""


def _extract_text_from_html(html: str) -> str:
    """从HTML中提取正文内容"""
    # 尝试提取主要内容区域
    for pattern in [
        r'<div[^>]*class="[^"]*content[^"]*"[^>]*>(.*?)</div>',
        r'<div[^>]*class="[^"]*article[^"]*"[^>]*>(.*?)</div>',
        r'<div[^>]*id="content"[^>]*>(.*?)</div>',
        r'<div[^>]*class="[^"]*main[^"]*"[^>]*>(.*?)</div>',
        r'<article[^>]*>(.*?)</article>',
    ]:
        m = re.search(pattern, html, re.IGNORECASE | re.DOTALL)
        if m:
            return _strip_html_tags(m.group(1))
    return _strip_html_tags(html)


# ==================== NFRA 处罚爬取 ====================

def _fetch_nfra_penalty_list(max_count: int) -> List[Dict[str, str]]:
    """
    从 NFRA 官网爬取行政处罚列表

    使用 NFRA 的 JSON API 获取文档列表，再逐个获取详情页内容
    """
    results: List[Dict[str, str]] = []
    page_size = min(max_count, 20)
    page_index = 1

    try:
        list_url = (
            f"{NFRA_BASE_URL}/cbircweb/DocInfo/SelectDocByItemIdAndChild"
            f"?itemId={NFRA_PENALTY_ITEM_ID}&pageSize={page_size}&pageIndex={page_index}"
        )
        data = _fetch_json(list_url)
        if not data or "data" not in data:
            logger.warning("NFRA 处罚列表 API 返回为空")
            return results

        rows = data["data"].get("rows", [])
        total = data["data"].get("total", 0)
        logger.info(f"NFRA 处罚: 共 {total} 条记录，当前获取 {len(rows)} 条")

        # 只取前 max_count 条
        rows = rows[:max_count]

        for row in rows:
            doc_id = str(row.get("docId", ""))
            title = row.get("docTitle", "").strip()
            publish_date = row.get("publishDate", "") or row.get("createDate", "")

            # 获取详情页内容
            content = ""
            if doc_id:
                detail_url = f"{NFRA_BASE_URL}/cbircweb/DocInfo/SelectByDocId?docId={doc_id}"
                detail = _fetch_json(detail_url)
                if detail and "data" in detail:
                    doc_data = detail["data"]
                    # 内容来源: docClob（Word文档转HTML）、docSummary（摘要）
                    content = doc_data.get("docClob", "") or doc_data.get("docSummary", "")
                    if content:
                        content = _strip_html_tags(content)
                    # 格式化日期
                    if not publish_date:
                        publish_date = doc_data.get("publishDate", "") or doc_data.get("docDate", "")

            # 构建详情页URL（供用户浏览器访问）
            page_url = f"{NFRA_BASE_URL}/cn/view/pages/ItemDetail.html?docId={doc_id}"

            results.append({
                "title": title,
                "date": str(publish_date)[:10] if publish_date else "",
                "url": page_url,
                "content": content[:5000] if content else "",
            })

        logger.info(f"NFRA 处罚: 成功获取 {len(results)} 条")

    except Exception as e:
        logger.error(f"NFRA 处罚爬取失败: {e}")

    return results


# ==================== PBC 处罚爬取 ====================

def _fetch_pbc_penalty_list(max_count: int) -> List[Dict[str, str]]:
    """
    从 PBC（人民银行）官网爬取行政处罚列表

    尝试多个已知的入口URL，解析页面中的链接获取处罚决定
    """
    results: List[Dict[str, str]] = []
    html = None
    base_url = PBC_BASE_URL

    # 尝试多个入口URL
    for url in PBC_PENALTY_URLS:
        html = _fetch_html(url)
        if html:
            base_url = url
            logger.info(f"PBC 处罚: 成功访问入口页面 {url}")
            break
    else:
        logger.warning("PBC 处罚页面请求失败（所有入口URL均不可用）")
        return results

    # 解析列表页面中的链接
    matches = []
    # 匹配多种可能的链接模式
    link_patterns = [
        # 处罚相关路径
        re.compile(
            r'<a[^>]*href="([^"]*(?:4081340|4081350|4081390|chufa|penalty)[^"]*\.html)"[^>]*>(.*?)</a>',
            re.IGNORECASE | re.DOTALL,
        ),
        # 政务公开下的链接
        re.compile(
            r'<a[^>]*href="([^"]*zhengwugongkai[^"]*\.html)"[^>]*>(.*?)</a>',
            re.IGNORECASE | re.DOTALL,
        ),
        # 相对路径
        re.compile(
            r'<a[^>]*href="(\.\./[^"]*\.html)"[^>]*>(.*?)</a>',
            re.IGNORECASE | re.DOTALL,
        ),
        # 通用链接
        re.compile(
            r'<a[^>]*href="([^"]*\.html)"[^>]*>(.*?)</a>',
            re.IGNORECASE | re.DOTALL,
        ),
    ]

    for pattern in link_patterns:
        found = pattern.findall(html)
        if found:
            # 过滤：优先保留包含"处罚"关键字的链接
            penalty_matches = [
                (href, title) for href, title in found
                if any(kw in (href + title) for kw in ["处罚", "行政", "罚", "chufa", "penalty"])
            ]
            if penalty_matches:
                matches = penalty_matches
            else:
                matches = found
            break

    logger.info(f"PBC 处罚: 在列表页找到 {len(matches)} 个链接")

    count = 0
    seen_urls = set()

    for href, title_text in matches:
        if count >= max_count:
            break

        full_url = href if href.startswith("http") else urljoin(PBC_BASE_URL, href)
        if full_url in seen_urls:
            continue
        seen_urls.add(full_url)

        title = _strip_html_tags(title_text)
        if not title:
            title = "行政处罚决定"

        # 尝试获取详情页内容
        content = ""
        date_str = ""
        detail_html = _fetch_html(full_url)
        if detail_html:
            content = _extract_text_from_html(detail_html)
            # 尝试从内容中提取日期
            date_match = re.search(
                r"(\d{4}[-年]\d{1,2}[-月]\d{1,2})",
                content[:500] if content else "",
            )
            if date_match:
                date_str = date_match.group(1).replace("年", "-").replace("月", "-").replace("日", "")

        results.append({
            "title": title,
            "date": date_str,
            "url": full_url,
            "content": content[:5000] if content else "",
        })
        count += 1

    logger.info(f"PBC 处罚: 成功获取 {len(results)} 条")

    return results


# ==================== CSRC 处罚爬取 ====================

def _fetch_csrc_penalty_list(max_count: int) -> List[Dict[str, str]]:
    """
    从 CSRC（证监会）官网爬取行政处罚列表

    CSRC 的行政处罚决定在：
    https://www.csrc.gov.cn/csrc/c100028/common_list.shtml
    """
    results: List[Dict[str, str]] = []

    try:
        html = _fetch_html(CSRC_PENALTY_URL)
        if not html:
            logger.warning("CSRC 处罚页面请求失败")
            return results

        # CSRC 的列表结构通常包含 <li> 或 <a> 指向处罚决定
        # 匹配链接模式
        patterns = [
            # 匹配绝对路径或相对路径的行政处罚链接
            re.compile(
                r'<a[^>]*href="([^"]*c100028[^"]*\.s?html)"[^>]*title="([^"]*)"[^>]*>',
                re.IGNORECASE,
            ),
            re.compile(
                r'<a[^>]*href="([^"]*c100028[^"]*\.s?html)"[^>]*>(.*?)</a>',
                re.IGNORECASE | re.DOTALL,
            ),
            re.compile(
                r'<a[^>]*href="(\.\./[^"]*\.s?html)"[^>]*>(.*?)</a>',
                re.IGNORECASE | re.DOTALL,
            ),
        ]

        matches = []
        for pattern in patterns:
            found = pattern.findall(html)
            if found:
                matches = found
                break

        # 如果上面的模式都没匹配到，尝试匹配所有看起来像处罚链接的URL
        if not matches:
            link_pattern = re.compile(
                r'<a[^>]*href="([^"]*(?:xingzhengchufa|penalty|punish)[^"]*\.s?html)"[^>]*>(.*?)</a>',
                re.IGNORECASE | re.DOTALL,
            )
            matches = link_pattern.findall(html)

        if not matches:
            # 最后的兜底：匹配所有可能的链接
            link_pattern = re.compile(
                r'<a[^>]*href="([^"]*\.s?html)"[^>]*>(.*?)</a>',
                re.IGNORECASE | re.DOTALL,
            )
            all_matches = link_pattern.findall(html)
            # 过滤出看起来像处罚相关的
            matches = [
                (href, title) for href, title in all_matches
                if any(kw in (href + title).lower() for kw in ["处罚", "决定", "chufa", "penalty"])
            ]

        logger.info(f"CSRC 处罚: 在列表页找到 {len(matches)} 个链接")

        count = 0
        seen_urls = set()

        for href, title_text in matches:
            if count >= max_count:
                break

            full_url = href if href.startswith("http") else urljoin(CSRC_BASE_URL, href)
            if full_url in seen_urls:
                continue
            seen_urls.add(full_url)

            title = _strip_html_tags(title_text)
            if not title:
                title = "行政处罚决定书"

            # 获取详情页
            content = ""
            date_str = ""
            detail_html = _fetch_html(full_url)
            if detail_html:
                content = _extract_text_from_html(detail_html)
                # 尝试从内容中提取日期
                date_match = re.search(
                    r"(\d{4}[-年]\d{1,2}[-月]\d{1,2})",
                    content[:500] if content else "",
                )
                if date_match:
                    date_str = date_match.group(1).replace("年", "-").replace("月", "-").replace("日", "")

            results.append({
                "title": title,
                "date": date_str,
                "url": full_url,
                "content": content[:5000] if content else "",
            })
            count += 1

        logger.info(f"CSRC 处罚: 成功获取 {len(results)} 条")

    except Exception as e:
        logger.error(f"CSRC 处罚爬取失败: {e}")

    return results


# ==================== 公开接口 ====================

def fetch_all_penalty(max_per_source: int = 5) -> Dict[str, List[Dict[str, str]]]:
    """
    从三大金融监管机构爬取行政处罚信息

    参数:
        max_per_source: 每个来源最多爬取条数，默认 5

    返回:
        {
            "nfra": [{"title": "...", "date": "...", "url": "...", "content": "..."}, ...],
            "pbc": [...],
            "csrc": [...]
        }

    说明:
        - 如果某个来源爬取失败，对应列表为空，不影响其他来源
        - 每条记录包含 title（标题）、date（日期）、url（链接）、content（正文摘要）
    """
    all_data: Dict[str, List[Dict[str, str]]] = {}

    # NFRA
    logger.info("=" * 50)
    logger.info("开始爬取 NFRA（金融监管总局）处罚信息...")
    try:
        all_data["nfra"] = _fetch_nfra_penalty_list(max_per_source)
    except Exception as e:
        logger.error(f"NFRA 爬取异常: {e}")
        all_data["nfra"] = []

    # PBC
    logger.info("=" * 50)
    logger.info("开始爬取 PBC（人民银行）处罚信息...")
    try:
        all_data["pbc"] = _fetch_pbc_penalty_list(max_per_source)
    except Exception as e:
        logger.error(f"PBC 爬取异常: {e}")
        all_data["pbc"] = []

    # CSRC
    logger.info("=" * 50)
    logger.info("开始爬取 CSRC（证监会）处罚信息...")
    try:
        all_data["csrc"] = _fetch_csrc_penalty_list(max_per_source)
    except Exception as e:
        logger.error(f"CSRC 爬取异常: {e}")
        all_data["csrc"] = []

    total = sum(len(v) for v in all_data.values())
    logger.info(f"全部爬取完成，共获取 {total} 条处罚信息")
    return all_data


def _sanitize_excel_text(text: str) -> str:
    """清理文本中的非法字符，确保可以写入Excel单元格"""
    if not text:
        return ""
    # 移除 openpyxl 不接受的非法控制字符（保留换行符、制表符）
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
    # 移除多余空白
    text = re.sub(r"\n{4,}", "\n\n\n", text)
    return text.strip()


def write_penalty_excel(
    all_data: Dict[str, List[Dict[str, str]]],
    output_dir: str = "output",
) -> str:
    """
    将处罚信息写入 Excel 文件

    参数:
        all_data: fetch_all_penalty 返回的数据
        output_dir: 输出目录，默认 "output"

    返回:
        生成的文件路径

    说明:
        - 每个来源（nfra、pbc、csrc）占一个 sheet
        - sheet 名称分别为：金融监管总局、人民银行、证监会
        - 列：序号、标题、日期、链接、正文摘要
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    os.makedirs(output_dir, exist_ok=True)

    filepath = os.path.join(output_dir, "金融机构处罚信息.xlsx")

    wb = Workbook()
    wb.remove(wb.active)

    # 样式定义
    header_font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="C0392B", end_color="C0392B", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    data_font = Font(name="微软雅黑", size=10)
    data_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    center_align = Alignment(horizontal="center", vertical="center")

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    title_font = Font(name="微软雅黑", size=14, bold=True, color="C0392B")

    # sheet 映射
    sheet_map = {
        "nfra": "金融监管总局",
        "pbc": "人民银行",
        "csrc": "证监会",
    }

    columns = ["序号", "标题", "日期", "链接", "正文摘要"]

    for key, sheet_name in sheet_map.items():
        records = all_data.get(key, [])
        ws = wb.create_sheet(title=sheet_name)

        # 标题行
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(columns))
        title_cell = ws.cell(row=1, column=1, value=f"{sheet_name} - 行政处罚信息")
        title_cell.font = title_font
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 36

        # 表头
        for col_idx, header in enumerate(columns, 1):
            cell = ws.cell(row=2, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border
        ws.row_dimensions[2].height = 28

        # 数据行
        for row_idx, record in enumerate(records, 3):
            ws.cell(row=row_idx, column=1, value=row_idx - 2).alignment = center_align
            ws.cell(row=row_idx, column=1).font = data_font
            ws.cell(row=row_idx, column=1).border = thin_border

            ws.cell(row=row_idx, column=2, value=_sanitize_excel_text(record.get("title", ""))).alignment = data_align
            ws.cell(row=row_idx, column=2).font = data_font
            ws.cell(row=row_idx, column=2).border = thin_border

            ws.cell(row=row_idx, column=3, value=_sanitize_excel_text(record.get("date", ""))).alignment = center_align
            ws.cell(row=row_idx, column=3).font = data_font
            ws.cell(row=row_idx, column=3).border = thin_border

            ws.cell(row=row_idx, column=4, value=_sanitize_excel_text(record.get("url", ""))).alignment = data_align
            ws.cell(row=row_idx, column=4).font = Font(name="微软雅黑", size=10, color="2980B9", underline="single")
            ws.cell(row=row_idx, column=4).border = thin_border

            ws.cell(row=row_idx, column=5, value=_sanitize_excel_text(record.get("content", ""))).alignment = data_align
            ws.cell(row=row_idx, column=5).font = data_font
            ws.cell(row=row_idx, column=5).border = thin_border

            ws.row_dimensions[row_idx].height = 60

        # 列宽
        col_widths = [6, 45, 14, 55, 80]
        for i, w in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

        # 冻结窗格
        ws.freeze_panes = "A3"

        logger.info(f"Sheet '{sheet_name}' 已写入 {len(records)} 条记录")

    wb.save(filepath)
    logger.info(f"处罚信息Excel已保存: {filepath}")
    return filepath


# ==================== 自测入口 ====================

if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )

    print("正在爬取金融机构处罚信息...")
    data = fetch_all_penalty(max_per_source=3)

    for source, records in data.items():
        print(f"\n{'=' * 60}")
        print(f"来源: {source.upper()} 共 {len(records)} 条")
        print(f"{'=' * 60}")
        for i, rec in enumerate(records, 1):
            print(f"  [{i}] {rec['title']}")
            print(f"      日期: {rec['date']}")
            print(f"      链接: {rec['url']}")
            print(f"      内容摘要: {rec['content'][:100]}...")
            print()

    if any(len(v) > 0 for v in data.values()):
        path = write_penalty_excel(data)
        print(f"\nExcel 已保存: {path}")
    else:
        print("\n未获取到任何处罚信息")