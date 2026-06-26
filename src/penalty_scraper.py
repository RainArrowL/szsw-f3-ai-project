"""
处罚信息爬取模块
从证监会(CSRC)官网获取行政处罚信息（CSRC为服务器端渲染，可直接爬取）
从国家金融监督管理总局(NFRA)官网获取行政处罚信息（NFRA为AngularJS动态渲染，
requests无法直接获取，尝试读取预缓存数据或扫描详情页）
"""
import re
import json
import logging
import requests
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

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
    try:
        resp = requests.get(url, headers=HEADERS, timeout=timeout)
        # CSRC 使用 UTF-8
        if resp.encoding and resp.encoding.lower() != "utf-8":
            resp.encoding = "utf-8"
        if resp.status_code == 200:
            return resp.text
        logger.warning(f"请求失败 {url}: HTTP {resp.status_code}")
    except Exception as e:
        logger.warning(f"请求异常 {url}: {e}")
    return None


# ══════════════════════════════════════════════════════════════════
#  CSRC 爬取
# ══════════════════════════════════════════════════════════════════

def _extract_csrc_penalty_links(html: str) -> List[Dict[str, str]]:
    """从CSRC行政处罚列表页提取链接"""
    # CSRC列表页中的处罚决定书链接格式: /csrc/c101928/cXXXXXXX/content.shtml
    pattern = re.compile(
        r'href="(/csrc/c101928/c\d+/content\.shtml)"[^>]*>'
        r'([^<]*)</a>',
        re.DOTALL,
    )
    matches = pattern.findall(html)
    seen = set()
    records = []
    for url, title in matches:
        if url in seen:
            continue
        seen.add(url)
        title = title.strip()
        if not title:
            title = "行政处罚决定书"
        records.append({
            "source": "证监会",
            "title": title,
            "url": CSRC_BASE + url,
        })
    return records


def _parse_csrc_penalty_detail(html: str, url: str) -> Optional[Dict[str, str]]:
    """解析CSRC行政处罚决定书详情页"""
    # 提取标题
    title = "行政处罚决定书"
    title_match = re.search(r"<title>([^<]+)</title>", html)
    if title_match:
        t = title_match.group(1).strip()
        if t:
            title = t

    # 提取发文日期
    date_str = ""
    date_match = re.search(r"发文日期[：:]\s*(\d{4})年(\d{2})月(\d{2})日", html)
    if date_match:
        date_str = f"{date_match.group(1)}-{date_match.group(2)}-{date_match.group(3)}"
    else:
        date_match = re.search(r"(\d{4})年(\d{2})月(\d{2})日", html)
        if date_match:
            date_str = f"{date_match.group(1)}-{date_match.group(2)}-{date_match.group(3)}"

    # 提取文号
    wenhao = ""
    wh_match = re.search(r"〔(\d{4})〕(\d+)号", html)
    if wh_match:
        wenhao = f"〔{wh_match.group(1)}〕{wh_match.group(2)}号"

    # 提取当事人（取第一个"当事人：xxx"）
    entity = ""
    # 多种当事人格式
    for pat in [
        r"当事人[：:]\s*([^，,。\n]+?)(?:[，,。\n]|依据|涉嫌)",
        r"当事人[：:]\s*(.+?)(?:，|。|\n|依据)",
    ]:
        entity_match = re.search(pat, html, re.DOTALL)
        if entity_match:
            entity = entity_match.group(1).strip()
            entity = re.sub(r"<[^>]+>", "", entity)
            entity = re.sub(r"\s+", " ", entity)
            if len(entity) > 200:
                entity = entity[:200] + "..."
            break

    # 提取处罚金额
    penalty_amount = ""
    text = re.sub(r"<[^>]+>", " ", html)
    text = re.sub(r"\s+", " ", text)
    for pat in [
        r"罚款\s*([\d,.]+\s*万?[元亿]?)",
        r"罚没款\s*([\d,.]+\s*万?[元亿]?)",
        r"没收违法所得\s*([\d,.]+\s*万?[元亿]?)",
        r"处以\s*(\d[\d,.]*\s*万?[元亿]?)\s*(?:罚款|的罚款)",
    ]:
        am = re.search(pat, text)
        if am:
            penalty_amount = am.group(1).strip()
            break

    # 提取处罚摘要
    summary = ""
    # 提取"我会认为"或"我会决定"附近的段落
    for key in ["我会认为", "我会决定", "处罚决定", "处罚如下"]:
        idx = text.find(key)
        if idx > 0:
            summary = text[idx:idx + 300].strip()
            break

    return {
        "source": "证监会",
        "date": date_str,
        "wenhao": wenhao,
        "entity": entity,
        "title": title,
        "penalty_amount": penalty_amount,
        "summary": summary,
        "url": url,
    }


def fetch_csrc_penalty_list() -> List[Dict[str, str]]:
    """获取证监会行政处罚决定列表"""
    logger.info("正在获取证监会处罚信息...")
    html = _fetch_html(CSRC_PENALTY_LIST_URL)
    if not html:
        html = _fetch_html("http://www.csrc.gov.cn/csrc/c101928/index.shtml")
    if not html:
        logger.warning("CSRC处罚列表页面获取失败")
        return []

    links = _extract_csrc_penalty_links(html)
    logger.info(f"CSRC处罚列表获取到 {len(links)} 条链接")

    records = []
    for link in links[:20]:
        detail_html = _fetch_html(link["url"])
        if detail_html:
            record = _parse_csrc_penalty_detail(detail_html, link["url"])
            if record:
                records.append(record)
                logger.info(f"CSRC处罚: {record.get('wenhao', '')} {record.get('entity', '')[:50]}")

    return records


# ══════════════════════════════════════════════════════════════════
#  NFRA 爬取（尝试多种方式）
# ══════════════════════════════════════════════════════════════════

# 预缓存的NFRA处罚数据文件
_NFRA_CACHE_FILE = Path(__file__).resolve().parent / "nfra_penalty_cache.json"


def _try_load_nfra_cache() -> Optional[List[Dict[str, str]]]:
    """尝试加载预缓存的NFRA处罚数据"""
    if _NFRA_CACHE_FILE.exists():
        try:
            with open(_NFRA_CACHE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            logger.info(f"从缓存加载NFRA处罚数据: {len(data)} 条")
            return data
        except Exception as e:
            logger.warning(f"NFRA缓存加载失败: {e}")
    return None


def _try_scan_nfra_detail_pages() -> List[Dict[str, str]]:
    """尝试扫描NFRA详情页获取处罚数据（限已知docId范围）"""
    records = []
    # 已知的NFRA总局行政处罚详情页docId范围（从WebFetch获取的列表）
    # 这些docId来自NFRA总局机关行政处罚列表页
    known_doc_ids = [
        # 2025-2026年批次
        "1246789",  # 洛阳监管分局
        "1257546",  # 陕西
    ]
    # 扩展扫描范围
    base = 1246780
    for i in range(50):
        doc_id = str(base + i)
        if doc_id in known_doc_ids:
            continue
        known_doc_ids.append(doc_id)

    for doc_id in known_doc_ids[:30]:
        url = f"https://www.nfra.gov.cn/cn/view/pages/ItemDetail.html?docId={doc_id}&itemId=4113"
        html = _fetch_html(url)
        if not html:
            continue
        # 检查是否是AngularJS模板（无数据）
        if "ng-app" in html.lower() and "行政处罚" not in html:
            continue
        # 尝试解析表格
        recs = _parse_nfra_detail_table(html)
        if recs:
            records.extend(recs)
            logger.info(f"NFRA详情页 {doc_id}: {len(recs)} 条")

    return records


def _parse_nfra_detail_table(html: str) -> List[Dict[str, str]]:
    """解析NFRA行政处罚详情页的表格"""
    records = []
    tr_pattern = re.compile(r"<tr[^>]*>(.*?)</tr>", re.DOTALL)
    td_pattern = re.compile(r"<t[dh][^>]*>(.*?)</t[dh]>", re.DOTALL)
    tag_pattern = re.compile(r"<[^>]+>")

    table_match = re.search(r"<table[^>]*>(.*?)</table>", html, re.DOTALL)
    if not table_match:
        return records

    rows = tr_pattern.findall(table_match.group(1))
    for row in rows:
        cells = td_pattern.findall(row)
        if len(cells) < 5:
            continue
        content = [tag_pattern.sub("", c).strip() for c in cells]
        content = [re.sub(r"\s+", " ", c) for c in content]
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
    """获取NFRA处罚信息（尝试缓存、扫描详情页）"""
    logger.info("正在获取国家金融监督管理总局处罚信息...")

    # 1. 尝试加载缓存
    cached = _try_load_nfra_cache()
    if cached:
        return cached

    # 2. 尝试扫描详情页
    records = _try_scan_nfra_detail_pages()
    if records:
        logger.info(f"NFRA详情页扫描到 {len(records)} 条")
        return records

    logger.warning("NFRA处罚数据获取失败（网站为AngularJS动态渲染，CDN不可用）")
    return []


def fetch_all_penalties() -> Tuple[List[Dict], List[Dict]]:
    """获取所有处罚信息"""
    nfra_records = fetch_nfra_penalty_list()
    csrc_records = fetch_csrc_penalty_list()
    return nfra_records, csrc_records


# ══════════════════════════════════════════════════════════════════
#  Excel 输出
# ══════════════════════════════════════════════════════════════════

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
    hf = Font(name="微软雅黑", bold=True, size=11, color="FFFFFF")
    hfill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    cf = Font(name="微软雅黑", size=10)
    wrap_align = Alignment(wrap_text=True, vertical="top")
    border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    def _write_header(ws, headers):
        for col, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=col, value=h)
            c.font = hf
            c.fill = hfill
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border

    def _write_rows(ws, start_row, rows_data):
        for i, row in enumerate(rows_data):
            for col, val in enumerate(row, 1):
                c = ws.cell(row=start_row + i, column=col, value=val)
                c.font = cf
                c.alignment = wrap_align
                c.border = border

    # Sheet 1: NFRA
    ws1 = wb.active
    ws1.title = "金监总局处罚"
    _write_header(ws1, ["序号", "当事人名称", "主要违法违规行为", "行政处罚内容", "作出决定机关"])
    if nfra_records:
        rows = [[i, r.get("entity", ""), r.get("violation", ""),
                 r.get("penalty", ""), r.get("authority", "")]
                for i, r in enumerate(nfra_records, 1)]
        _write_rows(ws1, 2, rows)
    else:
        ws1.cell(row=2, column=1, value="（NFRA网站为AngularJS动态渲染，需浏览器JavaScript执行，暂无法直接爬取）").font = cf
    ws1.column_dimensions["A"].width = 6
    ws1.column_dimensions["B"].width = 30
    ws1.column_dimensions["C"].width = 45
    ws1.column_dimensions["D"].width = 60
    ws1.column_dimensions["E"].width = 15

    # Sheet 2: CSRC
    ws2 = wb.create_sheet("证监会处罚")
    _write_header(ws2, ["序号", "文号", "当事人", "处罚日期", "涉嫌处罚金额", "标题", "链接"])
    if csrc_records:
        rows = [[i, r.get("wenhao", ""), r.get("entity", ""),
                 r.get("date", ""), r.get("penalty_amount", ""),
                 r.get("title", ""), r.get("url", "")]
                for i, r in enumerate(csrc_records, 1)]
        _write_rows(ws2, 2, rows)
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