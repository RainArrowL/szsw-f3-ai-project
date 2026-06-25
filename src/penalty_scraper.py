"""
处罚信息爬取模块
从国家金融监督管理总局(NFRA)和证监会(CSRC)官网获取行政处罚信息
"""
import re
import logging
import requests
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

# ── NFRA (国家金融监督管理总局) ─────────────────────────────────
NFRA_PENALTY_LIST_URL = (
    "https://www.nfra.gov.cn/cn/view/pages/ItemList.html"
    "?itemPId=923&itemId=4113&itemUrl=ItemListRightList.html"
    "&itemName=总局机关&itemsubPId=931"
)
NFRA_DETAIL_BASE = "https://www.nfra.gov.cn/cn/view/pages/ItemDetail.html"

# ── CSRC (证监会) ─────────────────────────────────────────────
CSRC_PENALTY_LIST_URL = "http://www.csrc.gov.cn/csrc/c101928/"
CSRC_BASE = "http://www.csrc.gov.cn"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9",
}


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


def _extract_nfra_penalty_items(html: str) -> List[Dict[str, str]]:
    """从NFRA行政处罚列表页面提取处罚条目"""
    records = []
    # 匹配列表项链接: <a href="...ItemDetail.html?docId=XXXXXX&itemId=4113">标题</a> 日期
    pattern = re.compile(
        r'<a[^>]*href="([^"]*ItemDetail\.html\?docId=(\d+)[^"]*)"[^>]*>'
        r'([^<]+)</a>'
        r'[^<]*'
        r'(\d{4}-\d{2}-\d{2})?',
        re.DOTALL,
    )
    matches = pattern.findall(html)
    seen = set()
    for full_url, doc_id, title, date_str in matches:
        if doc_id in seen:
            continue
        seen.add(doc_id)
        if not full_url.startswith("http"):
            full_url = "https://www.nfra.gov.cn" + full_url
        records.append({
            "source": "国家金融监督管理总局",
            "doc_id": doc_id,
            "title": title.strip(),
            "date": date_str.strip() if date_str else "",
            "url": full_url,
        })
    return records


def _parse_nfra_detail_table(html: str) -> List[Dict[str, str]]:
    """解析NFRA行政处罚详情页的表格"""
    records = []
    # 提取表格行
    tr_pattern = re.compile(r"<tr[^>]*>(.*?)</tr>", re.DOTALL)
    td_pattern = re.compile(r"<td[^>]*>(.*?)</td>", re.DOTALL)
    tag_pattern = re.compile(r"<[^>]+>")

    # 找到表格区域
    table_match = re.search(r"<table[^>]*>(.*?)</table>", html, re.DOTALL)
    if not table_match:
        return records

    rows = tr_pattern.findall(table_match.group(1))
    for row in rows:
        cells = td_pattern.findall(row)
        if len(cells) < 5:
            continue
        content = [tag_pattern.sub("", c).strip() for c in cells]
        # 跳过表头
        if content[0] in ("序号", ""):
            continue
        if not content[0].isdigit():
            continue
        records.append({
            "source": "国家金融监督管理总局",
            "seq": content[0],
            "entity": content[1] if len(content) > 1 else "",
            "violation": content[2] if len(content) > 2 else "",
            "penalty": content[3] if len(content) > 3 else "",
            "authority": content[4] if len(content) > 4 else "",
        })
    return records


def fetch_nfra_penalty_list() -> List[Dict[str, str]]:
    """获取NFRA总局机关行政处罚列表"""
    logger.info("正在获取国家金融监督管理总局处罚信息...")
    html = _fetch_html(NFRA_PENALTY_LIST_URL)
    if not html:
        logger.warning("NFRA处罚列表页面获取失败")
        return []

    items = _extract_nfra_penalty_items(html)
    logger.info(f"NFRA处罚列表获取到 {len(items)} 条")

    # 只获取第一页的详情（列表页本身可能包含表格，也可能需要点进详情页）
    # 尝试从列表页直接提取表格内容
    # 如果列表页没有表格，则逐个获取详情页
    records = _parse_nfra_detail_table(html)
    if records:
        logger.info(f"从NFRA列表页直接解析到 {len(records)} 条处罚记录")
        return records

    # 逐个获取详情页
    all_records = []
    for item in items[:20]:  # 限制页数
        detail_html = _fetch_html(item["url"])
        if detail_html:
            detail_records = _parse_nfra_detail_table(detail_html)
            all_records.extend(detail_records)
            logger.info(f"NFRA详情页 {item['doc_id']}: {len(detail_records)} 条")

    return all_records


def _extract_csrc_penalty_links(html: str) -> List[Dict[str, str]]:
    """从CSRC行政处罚列表页提取链接"""
    records = []
    # 匹配行政处罚决定书链接
    pattern = re.compile(
        r'<a[^>]*href="(/csrc/c\d+/(c\d+/)?content\.shtml)"[^>]*>'
        r'([^<]*行政处罚决定书[^<]*)</a>',
        re.DOTALL,
    )
    matches = pattern.findall(html)
    seen = set()
    for url, _, title in matches:
        full_url = CSRC_BASE + url
        if full_url in seen:
            continue
        seen.add(full_url)
        records.append({
            "source": "证监会",
            "title": title.strip(),
            "url": full_url,
        })
    return records


def _parse_csrc_penalty_detail(html: str, url: str) -> Optional[Dict[str, str]]:
    """解析CSRC行政处罚决定书详情页"""
    # 提取发布日期
    date_match = re.search(r"发文日期[：:]\s*(\d{4})年(\d{2})月(\d{2})日", html)
    if not date_match:
        date_match = re.search(r"(\d{4})年(\d{2})月(\d{2})日", html)
    date_str = ""
    if date_match:
        date_str = f"{date_match.group(1)}-{date_match.group(2)}-{date_match.group(3)}"

    # 提取文号
    wenhao = ""
    wh_match = re.search(r"〔(\d{4})〕(\d+)号", html)
    if wh_match:
        wenhao = f"〔{wh_match.group(1)}〕{wh_match.group(2)}号"

    # 提取当事人
    entity = ""
    entity_match = re.search(r"当事人[：:]\s*(.+?)(?:[，,]|依据)", html, re.DOTALL)
    if entity_match:
        entity = entity_match.group(1).strip()
        # 限制长度
        if len(entity) > 200:
            entity = entity[:200] + "..."

    # 提取标题
    title_match = re.search(r"<title>([^<]+)</title>", html)
    title = title_match.group(1).strip() if title_match else "行政处罚决定书"

    # 提取处罚内容摘要 (从页面文本中提取关键段落)
    text = re.sub(r"<[^>]+>", "", html)
    text = re.sub(r"\s+", " ", text)

    # 提取处罚金额
    penalty_amount = ""
    amount_patterns = [
        r"罚款\s*(\d[\d,.]*\s*万?[元亿]?)",
        r"罚没款\s*(\d[\d,.]*\s*万?[元亿]?)",
        r"没收违法所得\s*(\d[\d,.]*\s*万?[元亿]?)",
    ]
    for pat in amount_patterns:
        am = re.search(pat, text)
        if am:
            penalty_amount = am.group(1).strip()
            break

    return {
        "source": "证监会",
        "date": date_str,
        "wenhao": wenhao,
        "entity": entity,
        "title": title,
        "penalty_amount": penalty_amount,
        "url": url,
    }


def fetch_csrc_penalty_list() -> List[Dict[str, str]]:
    """获取证监会行政处罚决定列表"""
    logger.info("正在获取证监会处罚信息...")
    html = _fetch_html(CSRC_PENALTY_LIST_URL)
    if not html:
        # 尝试备用URL
        html = _fetch_html("http://www.csrc.gov.cn/csrc/c101928/index.shtml")
    if not html:
        logger.warning("CSRC处罚列表页面获取失败")
        return []

    links = _extract_csrc_penalty_links(html)
    logger.info(f"CSRC处罚列表获取到 {len(links)} 条链接")

    records = []
    for link in links[:15]:  # 限制数量
        detail_html = _fetch_html(link["url"])
        if detail_html:
            record = _parse_csrc_penalty_detail(detail_html, link["url"])
            if record:
                records.append(record)
                logger.info(f"CSRC处罚: {record.get('wenhao', '')} {record.get('entity', '')[:50]}")

    return records


def fetch_all_penalties() -> Tuple[List[Dict], List[Dict]]:
    """获取所有处罚信息"""
    nfra_records = fetch_nfra_penalty_list()
    csrc_records = fetch_csrc_penalty_list()
    return nfra_records, csrc_records


def write_penalty_excel(
    nfra_records: List[Dict],
    csrc_records: List[Dict],
    output_dir: str = "output",
) -> str:
    """将处罚信息写入Excel文件"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    Path(output_dir).mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    header_font = Font(name="微软雅黑", bold=True, size=11)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font_white = Font(name="微软雅黑", bold=True, size=11, color="FFFFFF")
    cell_font = Font(name="微软雅黑", size=10)
    wrap_align = Alignment(wrap_text=True, vertical="top")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    # ── Sheet 1: NFRA处罚 ──
    ws1 = wb.active
    ws1.title = "金监总局处罚"
    headers1 = ["序号", "当事人名称", "主要违法违规行为", "行政处罚内容", "作出决定机关"]
    for col, h in enumerate(headers1, 1):
        cell = ws1.cell(row=1, column=col, value=h)
        cell.font = header_font_white
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    for i, rec in enumerate(nfra_records, 1):
        row_data = [
            i, rec.get("entity", ""), rec.get("violation", ""),
            rec.get("penalty", ""), rec.get("authority", ""),
        ]
        for col, val in enumerate(row_data, 1):
            cell = ws1.cell(row=i + 1, column=col, value=val)
            cell.font = cell_font
            cell.alignment = wrap_align
            cell.border = thin_border

    ws1.column_dimensions["A"].width = 6
    ws1.column_dimensions["B"].width = 30
    ws1.column_dimensions["C"].width = 45
    ws1.column_dimensions["D"].width = 60
    ws1.column_dimensions["E"].width = 15

    # ── Sheet 2: CSRC处罚 ──
    ws2 = wb.create_sheet("证监会处罚")
    headers2 = ["序号", "文号", "当事人", "处罚日期", "处罚金额", "标题", "链接"]
    for col, h in enumerate(headers2, 1):
        cell = ws2.cell(row=1, column=col, value=h)
        cell.font = header_font_white
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    for i, rec in enumerate(csrc_records, 1):
        row_data = [
            i, rec.get("wenhao", ""), rec.get("entity", ""),
            rec.get("date", ""), rec.get("penalty_amount", ""),
            rec.get("title", ""), rec.get("url", ""),
        ]
        for col, val in enumerate(row_data, 1):
            cell = ws2.cell(row=i + 1, column=col, value=val)
            cell.font = cell_font
            cell.alignment = wrap_align
            cell.border = thin_border

    ws2.column_dimensions["A"].width = 6
    ws2.column_dimensions["B"].width = 18
    ws2.column_dimensions["C"].width = 35
    ws2.column_dimensions["D"].width = 12
    ws2.column_dimensions["E"].width = 18
    ws2.column_dimensions["F"].width = 30
    ws2.column_dimensions["G"].width = 50

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"金融机构处罚信息_{timestamp}.xlsx"
    filepath = str(Path(output_dir) / filename)
    wb.save(filepath)
    logger.info(f"处罚信息Excel已保存: {filepath}")
    return filepath