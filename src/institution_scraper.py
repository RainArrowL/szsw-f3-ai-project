"""
金融机构法人名录爬取模块
获取国家金融监督管理总局的银行保险法人名单、证监会的证券基金公司名单
"""
import re
import logging
import requests
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional, Tuple
from urllib.parse import urljoin

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

# ── NFRA 银行保险法人名单 ───────────────────────────────────
# NFRA 网站是 AngularJS SPA，requests 无法直接获取页面内容。
# 直接使用 CDN 上的 PDF 文件（截至2025年6月末）。
# 注意：这些 URL 可能随文件更新而变化，需要定期维护。
NFRA_BANK_PDF_URL = (
    "https://www.nfra.gov.cn/chinese/docfile/2025/"
    "86c58b1ad810422c8fa6c6d0107f1626.pdf"
)
NFRA_INSURANCE_PDF_URL = (
    "https://www.nfra.gov.cn/chinese/docfile/2025/"
    "2a78efc6d162484f8dfb8d0388b00320.pdf"
)

# ── CSRC 证券基金期货公司名录 ───────────────────────────────────
# 上海辖区证券公司名录（上海局汇总全国证券公司）
CSRC_SECURITIES_URL = (
    "http://www.csrc.gov.cn/shanghai/c103854/c7637721/content.shtml"
)
# 上海辖区基金管理公司名录
CSRC_FUND_URL = (
    "http://www.csrc.gov.cn/shanghai/c103856/c7639412/content.shtml"
)
# 证监会公开的期货公司名录（最新）
CSRC_FUTURES_URL = (
    "http://www.csrc.gov.cn/csrc/c101920/c1039268/content.shtml"
)


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


def _download_file(url: str, timeout: int = 60) -> Optional[bytes]:
    try:
        resp = requests.get(url, headers=HEADERS, timeout=timeout)
        if resp.status_code == 200:
            return resp.content
        logger.warning(f"下载失败 {url}: HTTP {resp.status_code}")
    except Exception as e:
        logger.warning(f"下载异常 {url}: {e}")
    return None


def _parse_html_table(html: str) -> List[List[str]]:
    """从HTML解析表格，返回二维数组"""
    rows = []
    tr_pattern = re.compile(r"<tr[^>]*>(.*?)</tr>", re.DOTALL)
    td_pattern = re.compile(r"<t[dh][^>]*>(.*?)</t[dh]>", re.DOTALL)
    tag_pattern = re.compile(r"<[^>]+>")
    ws_pattern = re.compile(r"\s+")

    trs = tr_pattern.findall(html)
    for tr in trs:
        cells = td_pattern.findall(tr)
        if not cells:
            continue
        row = [ws_pattern.sub(" ", tag_pattern.sub("", c)).strip() for c in cells]
        if row and any(c for c in row):
            rows.append(row)
    return rows


def _find_xlsx_url(html: str, base_url: str) -> Optional[str]:
    """从HTML页面中查找附件xlsx/xls链接"""
    for pattern in [r'href="([^"]+\.xlsx)"', r'href="([^"]+\.xls)"']:
        match = re.search(pattern, html, re.IGNORECASE)
        if match:
            return urljoin(base_url, match.group(1))
    return None


def _find_pdf_url(html: str, base_url: str) -> Optional[str]:
    """从HTML页面中查找附件PDF链接"""
    pattern = re.compile(r'href="([^"]+\.pdf)"', re.IGNORECASE)
    for match in pattern.finditer(html):
        return urljoin(base_url, match.group(1))
    return None


def _parse_pdf_table(pdf_bytes: bytes) -> List[List[str]]:
    """从PDF二进制数据中提取表格"""
    from io import BytesIO

    # 尝试 pdfplumber
    try:
        import pdfplumber

        rows = []
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        if row and any(c for c in row if c):
                            rows.append([str(c).strip() if c else "" for c in row])
        return rows
    except ImportError:
        logger.warning("pdfplumber未安装，尝试PyPDF2")

    # 降级：PyPDF2
    try:
        from PyPDF2 import PdfReader

        reader = PdfReader(BytesIO(pdf_bytes))
        text = ""
        for page in reader.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"

        lines = text.strip().split("\n")
        rows = []
        for line in lines:
            parts = re.split(r"\s{2,}", line.strip())
            if parts and any(p for p in parts):
                rows.append(parts)
        return rows
    except ImportError:
        logger.warning("PyPDF2也未安装，无法解析PDF")
        return []


def _parse_xlsx(data: bytes) -> List[List[str]]:
    """解析Excel二进制数据（支持xlsx和xls格式）"""
    from io import BytesIO

    # 尝试 openpyxl (xlsx)
    try:
        from openpyxl import load_workbook
        wb = load_workbook(BytesIO(data), read_only=True, data_only=True)
        ws = wb.active
        rows = []
        for row in ws.iter_rows(values_only=True):
            rows.append([str(c) if c is not None else "" for c in row])
        wb.close()
        return rows
    except Exception:
        pass

    # 尝试 xlrd (xls)
    try:
        import xlrd
        wb = xlrd.open_workbook(file_contents=data)
        ws = wb.sheet_by_index(0)
        rows = []
        for r in range(ws.nrows):
            rows.append([str(ws.cell_value(r, c)) if ws.cell_value(r, c) != "" else ""
                         for c in range(ws.ncols)])
        return rows
    except ImportError:
        logger.warning("xlrd未安装，无法解析.xls文件")
    except Exception as e:
        logger.warning(f"解析Excel失败: {e}")

    return []


def _is_header_row(row: List[str], header_keywords: List[str] = None,
                   name_col_index: int = 1) -> bool:
    """判断是否为表头行或无效行（需要跳过）

    name_col_index: 名称列在行中的索引（默认第1列即索引1）
    """
    if header_keywords is None:
        header_keywords = ["序号", "中文全称", "机构名称", "公司名称", "单位名称",
                           "名称", "序号", "英文全称"]
    if not row:
        return True
    name_col = row[name_col_index] if len(row) > name_col_index else row[0] if row else ""
    # 空名称
    if not name_col or not name_col.strip():
        return True
    # 表头关键词
    for kw in header_keywords:
        if kw in name_col:
            return True
    # 注释行（如"本月无变化"、"注："等）
    if re.match(r'^(本月|注[：:]|说明|备注|截止)', name_col):
        return True
    # 非中文名称（纯英文/数字/特殊字符）
    if not re.search(r'[\u4e00-\u9fff]', name_col):
        return True
    return False


def _chunk_tables(tables: List[List[str]]) -> List[List[List[str]]]:
    """将表格拆分为多个逻辑块"""
    if not tables:
        return []
    chunks = []
    current = []
    header_keywords = ["序号", "中文全称", "机构名称", "公司名称", "名称"]
    for row in tables:
        row_str = "".join(row)
        if any(kw in row_str for kw in header_keywords) and current:
            if len(current) > 1:
                chunks.append(current)
            current = [row]
        else:
            current.append(row)
    if current:
        chunks.append(current)
    return chunks


def _download_nfra_pdf(url: str) -> Optional[bytes]:
    """下载 NFRA PDF"""
    logger.info(f"NFRA PDF: {url}")
    return _download_file(url, timeout=120)


def fetch_bank_insurance_list() -> Dict[str, List[Dict]]:
    """获取银行保险法人名单（通过 NFRA CDN PDF 直接下载）"""
    result = {"bank": [], "insurance": []}

    # ── 银行名单 ──
    logger.info("正在获取银行业金融机构法人名单...")
    pdf_bytes = _download_nfra_pdf(NFRA_BANK_PDF_URL)
    if pdf_bytes:
        rows = _parse_pdf_table(pdf_bytes)
        for row in rows:
            if _is_header_row(row):
                continue
            if len(row) >= 2:
                result["bank"].append({
                    "name": row[1] if len(row) > 1 else row[0] if row else "",
                    "code": row[3] if len(row) > 3 else "",
                    "type": row[4] if len(row) > 4 else "",
                })
    logger.info(f"银行法人名单: {len(result['bank'])} 家")

    # ── 保险名单 ──
    logger.info("正在获取保险机构法人名单...")
    pdf_bytes = _download_nfra_pdf(NFRA_INSURANCE_PDF_URL)
    if pdf_bytes:
        rows = _parse_pdf_table(pdf_bytes)
        for row in rows:
            if _is_header_row(row):
                continue
            if len(row) >= 2:
                result["insurance"].append({
                    "name": row[1] if len(row) > 1 else row[0] if row else "",
                    "code": row[3] if len(row) > 3 else "",
                    "type": row[4] if len(row) > 4 else "",
                })
    logger.info(f"保险法人名单: {len(result['insurance'])} 家")
    return result


def fetch_securities_fund_list() -> Dict[str, List[Dict]]:
    """获取证券基金期货公司名单"""
    result = {"securities": [], "funds": [], "futures": []}

    # 获取证券公司名录
    logger.info("正在获取证券公司名录...")
    html = _fetch_html(CSRC_SECURITIES_URL)
    if html:
        xlsx_url = _find_xlsx_url(html, CSRC_SECURITIES_URL)
        if xlsx_url:
            data = _download_file(xlsx_url)
            if data:
                rows = _parse_xlsx(data)
                for row in rows:
                    if _is_header_row(row, ["单位名称", "公司名称", "序号"]):
                        continue
                    if len(row) >= 2:
                        result["securities"].append({
                            "name": row[1] if len(row) > 1 else row[0] if row else "",
                            "addr": row[2] if len(row) > 2 else "",
                        })
        if not result["securities"]:
            tables = _parse_html_table(html)
            for chunk in _chunk_tables(tables):
                for row in chunk[1:]:
                    if len(row) >= 2 and not _is_header_row(row):
                        result["securities"].append({
                            "name": row[1] if len(row) > 1 else row[0] if row else "",
                            "addr": row[2] if len(row) > 2 else "",
                        })
    logger.info(f"证券公司名录: {len(result['securities'])} 家")

    # 获取基金公司名录
    logger.info("正在获取基金管理公司名录...")
    html = _fetch_html(CSRC_FUND_URL)
    if html:
        xlsx_url = _find_xlsx_url(html, CSRC_FUND_URL)
        if xlsx_url:
            data = _download_file(xlsx_url)
            if data:
                rows = _parse_xlsx(data)
                for row in rows:
                    if _is_header_row(row, ["公司名称", "单位名称", "序号"]):
                        continue
                    if len(row) >= 2:
                        result["funds"].append({
                            "name": row[1] if len(row) > 1 else row[0] if row else "",
                            "addr": row[2] if len(row) > 2 else "",
                        })
        if not result["funds"]:
            tables = _parse_html_table(html)
            for chunk in _chunk_tables(tables):
                for row in chunk[1:]:
                    if len(row) >= 2 and not _is_header_row(row):
                        result["funds"].append({
                            "name": row[1] if len(row) > 1 else row[0] if row else "",
                            "addr": row[2] if len(row) > 2 else "",
                        })
    logger.info(f"基金公司名录: {len(result['funds'])} 家")

    # 获取期货公司名录
    logger.info("正在获取期货公司名录...")
    html = _fetch_html(CSRC_FUTURES_URL)
    if html:
        xlsx_url = _find_xlsx_url(html, CSRC_FUTURES_URL)
        if xlsx_url:
            data = _download_file(xlsx_url)
            if data:
                rows = _parse_xlsx(data)
                for row in rows:
                    if _is_header_row(row, ["期货公司名称", "序号", "辖区"], name_col_index=2):
                        continue
                    if len(row) >= 2:
                        # 期货公司表格：序号 | 辖区 | 期货公司名称
                        name = row[2] if len(row) > 2 else row[1] if len(row) > 1 else ""
                        result["futures"].append({
                            "name": name,
                            "region": row[1] if len(row) > 1 else "",
                        })
        if not result["futures"]:
            tables = _parse_html_table(html)
            for row in tables[1:]:
                if not _is_header_row(row, ["期货公司名称", "序号", "辖区"], name_col_index=2):
                    if len(row) >= 2:
                        name = row[2] if len(row) > 2 else row[1] if len(row) > 1 else ""
                        region = row[1] if len(row) > 1 else ""
                        result["futures"].append({"name": name, "region": region})
    logger.info(f"期货公司名录: {len(result['futures'])} 家")
    return result


def fetch_all_institution_lists() -> Dict[str, List[Dict]]:
    """获取所有法人名单"""
    bank_insurance = fetch_bank_insurance_list()
    sec_fund = fetch_securities_fund_list()

    return {
        "bank": bank_insurance["bank"],
        "insurance": bank_insurance["insurance"],
        "securities": sec_fund["securities"],
        "funds": sec_fund["funds"],
        "futures": sec_fund["futures"],
    }


def write_institution_excel(
    data: Dict[str, List[Dict]],
    output_dir: str = "output",
    prefix: str = "金融机构法人名录",
) -> str:
    """将法人名单写入Excel文件"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    Path(output_dir).mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    hf = Font(name="微软雅黑", bold=True, size=11, color="FFFFFF")
    hfill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    cf = Font(name="微软雅黑", size=10)
    align = Alignment(wrap_text=True, vertical="top")
    border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    def _write_sheet(ws, title, headers, rows_data):
        ws.title = title
        for col, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=col, value=h)
            c.font = hf
            c.fill = hfill
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border
        for i, row in enumerate(rows_data, 1):
            for col, val in enumerate(row, 1):
                c = ws.cell(row=i + 1, column=col, value=val)
                c.font = cf
                c.alignment = align
                c.border = border

    # 根据提供的数据类型写入不同Sheet
    sheet_specs = [
        ("bank", "银行法人名单", ["序号", "机构名称", "机构编码", "机构类型"]),
        ("insurance", "保险法人名单", ["序号", "机构名称", "机构编码", "机构类型"]),
        ("securities", "证券公司名单", ["序号", "公司名称", "地址"]),
        ("funds", "基金公司名单", ["序号", "公司名称", "地址"]),
        ("futures", "期货公司名单", ["序号", "辖区", "期货公司名称"]),
    ]

    first = True
    for key, title, headers in sheet_specs:
        items = data.get(key, [])
        if not items:
            continue
        if first:
            ws = wb.active
            first = False
        else:
            ws = wb.create_sheet()

        rows_data = []
        for i, item in enumerate(items, 1):
            if key == "securities":
                rows_data.append([i, item.get("name", ""), item.get("addr", "")])
            elif key == "funds":
                rows_data.append([i, item.get("name", ""), item.get("addr", "")])
            elif key == "futures":
                rows_data.append([i, item.get("region", ""), item.get("name", "")])
            else:
                rows_data.append([i, item.get("name", ""), item.get("code", ""), item.get("type", "")])

        _write_sheet(ws, title, headers, rows_data)

        # 列宽
        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 40
        if key in ("bank", "insurance"):
            ws.column_dimensions["C"].width = 20
            ws.column_dimensions["D"].width = 18
        elif key in ("securities", "funds"):
            ws.column_dimensions["C"].width = 50
        elif key == "futures":
            ws.column_dimensions["C"].width = 40

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{prefix}_{timestamp}.xlsx"
    filepath = str(Path(output_dir) / filename)
    wb.save(filepath)
    logger.info(f"法人名录Excel已保存: {filepath}")
    return filepath