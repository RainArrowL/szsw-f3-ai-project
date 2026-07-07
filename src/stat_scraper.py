"""
统计信息爬虫模块
爬取4个网址的金融统计数据Excel并合并为一个文件

1. NFRA深圳局统计信息 - 深圳市银行业保险业统计数据
2. NFRA全国统计信息 - 银行业金融机构总资产、原保险保费收入等
3. CSRC证券市场月报 - 最新月报
4. CSRC深圳局市场数据 - 市场数据
"""
import re
import logging
import requests
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional, Tuple
from urllib.parse import urljoin, urlsplit

logger = logging.getLogger(__name__)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9",
}

# ── NFRA 统计信息 ───────────────────────────────────
NFRA_BASE_URL = "https://www.nfra.gov.cn"
NFRA_SZ_STATS_ITEM_ID = "1050"
NFRA_NATIONAL_STATS_ITEM_ID = "954"

# ── CSRC 统计信息 ───────────────────────────────────
CSRC_SECURITIES_MONTHLY_URL = "http://www.csrc.gov.cn/csrc/c100120/c7636492/content.shtml"
CSRC_FUTURES_MONTHLY_URL = "http://www.csrc.gov.cn/csrc/c100120/c7636494/content.shtml"
CSRC_SHENZHEN_URL = "http://www.csrc.gov.cn/shenzhen/index.shtml"


def _fetch_html(url: str, timeout: int = 30) -> Optional[str]:
    try:
        resp = requests.get(url, headers=HEADERS, timeout=timeout)
        resp.encoding = "utf-8"
        if resp.status_code == 200:
            return resp.text
        logger.warning(f"请求失败 {url}: HTTP {resp.status_code}")
    except Exception as e:
        logger.warning(f"请求异常 {url}: {e}")
    return None


def _fetch_json(url: str, timeout: int = 30) -> Optional[dict]:
    try:
        resp = requests.get(url, headers=HEADERS, timeout=timeout)
        resp.encoding = "utf-8"
        if resp.status_code == 200:
            return resp.json()
        logger.warning(f"JSON API 请求失败 {url}: HTTP {resp.status_code}")
    except Exception as e:
        logger.warning(f"JSON API 请求异常 {url}: {e}")
    return None


def _download_file(url: str, timeout: int = 120) -> Optional[bytes]:
    try:
        resp = requests.get(url, headers=HEADERS, timeout=timeout)
        if resp.status_code == 200:
            return resp.content
        logger.warning(f"下载失败 {url}: HTTP {resp.status_code}")
    except Exception as e:
        logger.warning(f"下载异常 {url}: {e}")
    return None


def _find_xlsx_urls(html: str, base_url: str) -> List[str]:
    """从HTML页面中查找所有附件xlsx/xls链接"""
    urls = []
    for pattern in [r'href="([^"]+\.xlsx)"', r'href="([^"]+\.xls)"']:
        for match in re.finditer(pattern, html, re.IGNORECASE):
            urls.append(urljoin(base_url, match.group(1)))
    return urls


def _find_pdf_urls(html: str, base_url: str) -> List[str]:
    """从HTML页面中查找所有PDF链接"""
    urls = []
    pattern = re.compile(r'href="([^"]+\.pdf)"', re.IGNORECASE)
    for match in pattern.finditer(html):
        urls.append(urljoin(base_url, match.group(1)))
    return urls


def _get_doc_detail_url(doc_id: str) -> str:
    """获取文档详情页URL"""
    return f"{NFRA_BASE_URL}/cbircweb/DocInfo/SelectByDocId?docId={doc_id}"


def _get_doc_list_url(item_id: str, page_index: int = 1) -> str:
    """获取文档列表API URL"""
    return (
        f"{NFRA_BASE_URL}/cbircweb/DocInfo/SelectDocByItemIdAndChild"
        f"?itemId={item_id}&pageSize=50&pageIndex={page_index}"
    )


def _extract_attachment_urls(doc_detail: dict) -> List[Tuple[str, str]]:
    """从文档详情中提取附件URL列表"""
    results = []
    doc_data = doc_detail.get("data", {})
    attachments = doc_data.get("attachmentInfoVOList", [])
    for att in attachments:
        url = att.get("urlOtherName", "")
        name = att.get("attachmentName", "")
        if url:
            results.append((name, urljoin(NFRA_BASE_URL, url)))
    return results


def _download_nfra_stats(item_id: str, item_name: str) -> List[Tuple[str, bytes]]:
    """下载NFRA指定栏目下的统计数据Excel文件"""
    logger.info(f"正在获取 {item_name}...")
    results = []

    list_url = _get_doc_list_url(item_id)
    data = _fetch_json(list_url)
    if not data or "data" not in data:
        logger.warning(f"无法获取 {item_name} 文档列表")
        return results

    rows = data["data"].get("rows", [])
    logger.info(f"{item_name} 共 {len(rows)} 条文档")

    for row in rows:
        title = row.get("docTitle", "")
        doc_id = str(row.get("docId", ""))
        if not doc_id:
            continue

        detail_url = _get_doc_detail_url(doc_id)
        detail = _fetch_json(detail_url)
        if not detail:
            continue

        attachments = _extract_attachment_urls(detail)
        for name, url in attachments:
            ext = Path(url.split("?")[0]).suffix.lower()
            if ext in (".xlsx", ".xls"):
                logger.info(f"  下载: {title} - {name}")
                file_bytes = _download_file(url, timeout=120)
                if file_bytes:
                    results.append((title + "_" + name, file_bytes))

    return results


def _download_csrc_page_stats(page_url: str, page_name: str) -> List[Tuple[str, bytes]]:
    """下载CSRC页面中的统计数据Excel文件"""
    logger.info(f"正在获取 {page_name}: {page_url}")
    results = []

    html = _fetch_html(page_url)
    if not html:
        logger.warning(f"无法访问 {page_name}")
        return results

    xlsx_urls = _find_xlsx_urls(html, page_url)
    logger.info(f"{page_name} 找到 {len(xlsx_urls)} 个Excel文件")

    for url in xlsx_urls:
        filename = Path(urlsplit(url).path).name
        logger.info(f"  下载: {filename}")
        file_bytes = _download_file(url, timeout=120)
        if file_bytes:
            results.append((page_name + "_" + filename, file_bytes))

    return results


def download_all_stats(output_dir: str = "output") -> str:
    """下载所有4个网址的统计数据，合并为一个Excel文件

    Returns:
        合并后的xlsx文件路径，失败返回空字符串
    """
    from io import BytesIO
    from openpyxl import Workbook
    from openpyxl.utils import get_column_letter
    from copy import copy

    Path(output_dir).mkdir(parents=True, exist_ok=True)

    all_files = []

    # 1. NFRA深圳局统计信息
    logger.info("=" * 50)
    logger.info("1/4: NFRA深圳局统计信息")
    sz_files = _download_nfra_stats(NFRA_SZ_STATS_ITEM_ID, "深圳局统计信息")
    all_files.extend(sz_files)

    # 2. NFRA全国统计信息
    logger.info("=" * 50)
    logger.info("2/4: NFRA全国统计信息")
    national_files = _download_nfra_stats(NFRA_NATIONAL_STATS_ITEM_ID, "全国统计信息")
    all_files.extend(national_files)

    # 3. CSRC证券市场月报
    logger.info("=" * 50)
    logger.info("3/4: CSRC证券市场月报")
    sec_files = _download_csrc_page_stats(CSRC_SECURITIES_MONTHLY_URL, "证券市场月报")
    all_files.extend(sec_files)

    # 4. CSRC期货市场月报
    logger.info("=" * 50)
    logger.info("4/5: CSRC期货市场月报")
    fut_files = _download_csrc_page_stats(CSRC_FUTURES_MONTHLY_URL, "期货市场月报")
    all_files.extend(fut_files)

    # 5. CSRC深圳局市场数据
    logger.info("=" * 50)
    logger.info("5/5: CSRC深圳局市场数据")
    sz_csrc_files = _download_csrc_page_stats(CSRC_SHENZHEN_URL, "深圳局市场数据")
    all_files.extend(sz_csrc_files)

    if not all_files:
        logger.warning("没有成功下载任何统计数据文件")
        return ""

    # 合并到一个Excel
    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    for name, data in all_files:
        try:
            ext = name.split(".")[-1].lower() if "." in name else ".xlsx"

            if ext == ".xlsx":
                from openpyxl import load_workbook

                wb_src = load_workbook(BytesIO(data), read_only=True, data_only=True)
                ws_src = wb_src.active

                sheet_name = name.replace(".xlsx", "")[:31]
                ws_out = wb_out.create_sheet(title=sheet_name)

                for merged_range in ws_src.merged_cells.ranges:
                    ws_out.merge_cells(str(merged_range))

                for row_idx in range(1, ws_src.max_row + 1):
                    if ws_src.row_dimensions[row_idx].height:
                        ws_out.row_dimensions[row_idx].height = ws_src.row_dimensions[row_idx].height
                for col_idx in range(1, ws_src.max_column + 1):
                    col_letter = get_column_letter(col_idx)
                    if ws_src.column_dimensions[col_letter].width:
                        ws_out.column_dimensions[col_letter].width = ws_src.column_dimensions[col_letter].width

                for row in ws_src.iter_rows():
                    for cell in row:
                        new_cell = ws_out.cell(row=cell.row, column=cell.column, value=cell.value)
                        if cell.has_style:
                            new_cell.border = copy(cell.border)
                            new_cell.alignment = copy(cell.alignment)
                            new_cell.number_format = cell.number_format

                wb_src.close()
                logger.info(f"已合并: {sheet_name} ({ws_src.max_row}行)")

            elif ext == ".xls":
                import xlrd

                wb_src = xlrd.open_workbook(file_contents=data)
                ws_src = wb_src.sheet_by_index(0)

                sheet_name = name.replace(".xls", "")[:31]
                ws_out = wb_out.create_sheet(title=sheet_name)

                for r in range(ws_src.nrows):
                    for c in range(ws_src.ncols):
                        cell_value = ws_src.cell_value(r, c)
                        ws_out.cell(row=r + 1, column=c + 1, value=cell_value if cell_value != "" else None)

                logger.info(f"已合并: {sheet_name} ({ws_src.nrows}行)")

        except Exception as e:
            logger.warning(f"合并文件失败 {name}: {e}")
            continue

    if not wb_out.sheetnames:
        logger.warning("合并后没有任何Sheet")
        return ""

    timestamp = datetime.now().strftime("%Y%m%d")
    filename = f"金融统计数据_{timestamp}.xlsx"
    filepath = str(Path(output_dir) / filename)
    wb_out.save(filepath)
    logger.info(f"合并统计数据已保存: {filepath}")

    return filepath
