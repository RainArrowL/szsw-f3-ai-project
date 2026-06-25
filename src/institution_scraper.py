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
# 银行业金融机构法人名单
NFRA_BANK_LIST_URL = (
    "https://www.nfra.gov.cn/cn/view/pages/governmentDetail.html"
    "?docId=1228300&generaltype=1&itemId=863"
)
# 保险机构法人名单
NFRA_INSURANCE_LIST_URL = (
    "https://www.nfra.gov.cn/cn/view/pages/governmentDetail.html"
    "?docId=1228301&generaltype=1&itemId=863"
)

# ── CSRC 证券基金公司名单 ───────────────────────────────────
# 上海辖区证券公司名录（全国性汇总可从上海局获取）
CSRC_SECURITIES_URL = (
    "http://www.csrc.gov.cn/shanghai/c103854/c7637721/content.shtml"
)
# 上海辖区基金管理公司名录
CSRC_FUND_URL = (
    "http://www.csrc.gov.cn/shanghai/c103856/c7639412/content.shtml"
)


def _fetch_html(url: str, timeout: int = 30) -> Optional[str]:
    """获取网页HTML"""
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
    """下载文件"""
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


def _find_xlsx_url(html: str) -> Optional[str]:
    """从HTML页面中查找.xlsx附件链接"""
    match = re.search(r'href="([^"]+\.xlsx)"', html, re.IGNORECASE)
    if match:
        url = match.group(1)
        if not url.startswith("http"):
            url = "http://www.csrc.gov.cn" + url
        return url
    return None


def _find_pdf_url(html: str) -> Optional[str]:
    """从HTML页面中查找.pdf附件链接"""
    pattern = re.compile(r'href="([^"]+\.pdf)"', re.IGNORECASE)
    for match in pattern.finditer(html):
        url = match.group(1)
        if not url.startswith("http"):
            url = "https://www.nfra.gov.cn" + url
        return url
    return None


def _parse_pdf_table(pdf_bytes: bytes) -> List[List[str]]:
    """从PDF二进制数据中提取表格（简易正则方式）"""
    try:
        # 尝试用 pdfplumber
        import pdfplumber
        from io import BytesIO

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
        logger.warning("pdfplumber未安装，尝试纯文本提取")
        # 降级：纯文本提取
        try:
            from io import StringIO
            from PyPDF2 import PdfReader

            reader = PdfReader(BytesIO(pdf_bytes))
            text = ""
            for page in reader.pages:
                text += page.extract_text() or ""

            lines = text.strip().split("\n")
            rows = []
            for line in lines:
                # 按多个空格分割
                parts = re.split(r"\s{2,}", line.strip())
                if parts and any(p for p in parts):
                    rows.append(parts)
            return rows
        except ImportError:
            logger.warning("PyPDF2也未安装，无法解析PDF")
            return []


def fetch_bank_insurance_list() -> Dict[str, List[Dict]]:
    """获取银行保险法人名单"""
    result = {"bank": [], "insurance": []}

    # 获取银行名单
    logger.info("正在获取银行业金融机构法人名单...")
    html = _fetch_html(NFRA_BANK_LIST_URL)
    if html:
        tables = _parse_html_table(html)
        if tables:
            # 找到表头确定列
            for table_chunk in _chunk_tables(tables):
                for row in table_chunk:
                    if len(row) >= 3:
                        result["bank"].append({
                            "name": row[1] if len(row) > 1 else "",
                            "code": row[2] if len(row) > 2 else "",
                            "type": row[3] if len(row) > 3 else "",
                        })
        if not result["bank"]:
            # 尝试下载PDF
            pdf_url = _find_pdf_url(html)
            if pdf_url:
                pdf_bytes = _download_file(pdf_url)
                if pdf_bytes:
                    rows = _parse_pdf_table(pdf_bytes)
                    for row in rows[1:]:  # 跳过表头
                        if len(row) >= 3:
                            result["bank"].append({
                                "name": row[1] if len(row) > 1 else row[0] if len(row) > 0 else "",
                                "code": row[2] if len(row) > 2 else "",
                                "type": row[3] if len(row) > 3 else "",
                            })

    logger.info(f"银行法人名单: {len(result['bank'])} 家")

    # 获取保险名单
    logger.info("正在获取保险机构法人名单...")
    html = _fetch_html(NFRA_INSURANCE_LIST_URL)
    if html:
        tables = _parse_html_table(html)
        if tables:
            for table_chunk in _chunk_tables(tables):
                for row in table_chunk:
                    if len(row) >= 2:
                        result["insurance"].append({
                            "name": row[1] if len(row) > 1 else row[0] if len(row) > 0 else "",
                            "code": row[2] if len(row) > 2 else "",
                            "type": row[3] if len(row) > 3 else "",
                        })
        if not result["insurance"]:
            pdf_url = _find_pdf_url(html)
            if pdf_url:
                pdf_bytes = _download_file(pdf_url)
                if pdf_bytes:
                    rows = _parse_pdf_table(pdf_bytes)
                    for row in rows[1:]:
                        if len(row) >= 2:
                            result["insurance"].append({
                                "name": row[1] if len(row) > 1 else row[0] if len(row) > 0 else "",
                                "code": row[2] if len(row) > 2 else "",
                                "type": row[3] if len(row) > 3 else "",
                            })

    logger.info(f"保险法人名单: {len(result['insurance'])} 家")
    return result


def _chunk_tables(tables: List[List[str]]) -> List[List[List[str]]]:
    """将表格拆分为多个逻辑块（跳过空行后的新表头）"""
    if not tables:
        return []
    chunks = []
    current = []
    header_keywords = ["序号", "中文全称", "机构名称", "公司名称"]
    for row in tables:
        if any(kw in "".join(row) for kw in header_keywords) and current:
            if current:
                chunks.append(current)
            current = [row]
        else:
            current.append(row)
    if current:
        chunks.append(current)
    return chunks


def fetch_securities_fund_list() -> Dict[str, List[Dict]]:
    """获取证券基金公司名单"""
    result = {"securities": [], "funds": []}

    # 获取证券公司名录
    logger.info("正在获取证券公司名录...")
    html = _fetch_html(CSRC_SECURITIES_URL)
    if html:
        xlsx_url = _find_xlsx_url(html)
        if xlsx_url:
            data = _download_file(xlsx_url)
            if data:
                rows = _parse_xlsx(data)
                for row in rows[1:]:
                    if len(row) >= 2:
                        result["securities"].append({
                            "name": row[1] if len(row) > 1 else row[0] if len(row) > 0 else "",
                            "addr": row[2] if len(row) > 2 else "",
                        })
        if not result["securities"]:
            tables = _parse_html_table(html)
            for row in tables[1:]:
                if len(row) >= 2:
                    result["securities"].append({
                        "name": row[1] if len(row) > 1 else row[0] if len(row) > 0 else "",
                        "addr": row[2] if len(row) > 2 else "",
                    })

    logger.info(f"证券公司名录: {len(result['securities'])} 家")

    # 获取基金公司名录
    logger.info("正在获取基金管理公司名录...")
    html = _fetch_html(CSRC_FUND_URL)
    if html:
        xlsx_url = _find_xlsx_url(html)
        if xlsx_url:
            data = _download_file(xlsx_url)
            if data:
                rows = _parse_xlsx(data)
                for row in rows[1:]:
                    if len(row) >= 2:
                        result["funds"].append({
                            "name": row[1] if len(row) > 1 else row[0] if len(row) > 0 else "",
                            "addr": row[2] if len(row) > 2 else "",
                        })
        if not result["funds"]:
            tables = _parse_html_table(html)
            for row in tables[1:]:
                if len(row) >= 2:
                    result["funds"].append({
                        "name": row[1] if len(row) > 1 else row[0] if len(row) > 0 else "",
                        "addr": row[2] if len(row) > 2 else "",
                    })

    logger.info(f"基金公司名录: {len(result['funds'])} 家")
    return result


def _parse_xlsx(data: bytes) -> List[List[str]]:
    """解析Excel二进制数据"""
    try:
        from io import BytesIO
        from openpyxl import load_workbook

        wb = load_workbook(BytesIO(data), read_only=True, data_only=True)
        ws = wb.active
        rows = []
        for row in ws.iter_rows(values_only=True):
            rows.append([str(c) if c is not None else "" for c in row])
        wb.close()
        return rows
    except Exception as e:
        logger.warning(f"解析Excel失败: {e}")
        return []


def fetch_all_institution_lists() -> Dict[str, List[Dict]]:
    """获取所有法人名单"""
    bank_insurance = fetch_bank_insurance_list()
    sec_fund = fetch_securities_fund_list()

    return {
        "bank": bank_insurance["bank"],
        "insurance": bank_insurance["insurance"],
        "securities": sec_fund["securities"],
        "funds": sec_fund["funds"],
    }


def write_institution_excel(
    data: Dict[str, List[Dict]],
    output_dir: str = "output",
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

    # Sheet 1: 银行法人
    ws1 = wb.active
    bank_rows = [
        [i, d.get("name", ""), d.get("code", ""), d.get("type", "")]
        for i, d in enumerate(data.get("bank", []), 1)
    ]
    _write_sheet(ws1, "银行法人名单", ["序号", "机构名称", "机构编码", "机构类型"], bank_rows)
    ws1.column_dimensions["A"].width = 6
    ws1.column_dimensions["B"].width = 40
    ws1.column_dimensions["C"].width = 20
    ws1.column_dimensions["D"].width = 18

    # Sheet 2: 保险法人
    ws2 = wb.create_sheet()
    ins_rows = [
        [i, d.get("name", ""), d.get("code", ""), d.get("type", "")]
        for i, d in enumerate(data.get("insurance", []), 1)
    ]
    _write_sheet(ws2, "保险法人名单", ["序号", "机构名称", "机构编码", "机构类型"], ins_rows)
    ws2.column_dimensions["A"].width = 6
    ws2.column_dimensions["B"].width = 40
    ws2.column_dimensions["C"].width = 20
    ws2.column_dimensions["D"].width = 18

    # Sheet 3: 证券公司
    ws3 = wb.create_sheet()
    sec_rows = [
        [i, d.get("name", ""), d.get("addr", "")]
        for i, d in enumerate(data.get("securities", []), 1)
    ]
    _write_sheet(ws3, "证券公司名单", ["序号", "公司名称", "地址"], sec_rows)
    ws3.column_dimensions["A"].width = 6
    ws3.column_dimensions["B"].width = 40
    ws3.column_dimensions["C"].width = 50

    # Sheet 4: 基金公司
    ws4 = wb.create_sheet()
    fund_rows = [
        [i, d.get("name", ""), d.get("addr", "")]
        for i, d in enumerate(data.get("funds", []), 1)
    ]
    _write_sheet(ws4, "基金公司名单", ["序号", "公司名称", "地址"], fund_rows)
    ws4.column_dimensions["A"].width = 6
    ws4.column_dimensions["B"].width = 40
    ws4.column_dimensions["C"].width = 50

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"金融机构法人名录_{timestamp}.xlsx"
    filepath = str(Path(output_dir) / filename)
    wb.save(filepath)
    logger.info(f"法人名录Excel已保存: {filepath}")
    return filepath