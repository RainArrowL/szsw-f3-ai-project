#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
金融机构名录爬取模块

数据来源：
  - 中国证券投资基金业协会 (AMAC): 公募基金管理人名录
  - 国家金融监督管理总局 (NFRA): 银行/保险机构法人名单 (PDF)
  - 中国证监会 (CSRC): 证券公司/基金公司/期货公司名录 (Excel)

功能：
  1. fetch_all_institution_lists()      — 获取全部机构名录
  2. write_institution_excel()          — 将机构数据写入 Excel
  3. download_nfra_pdfs()               — 从 NFRA 下载银行/保险机构 PDF
  4. download_csrc_to_combined_xlsx()   — 从 CSRC 下载并合并为 xlsx
"""

import os
import re
import logging
from typing import Dict, List, Optional, Tuple
from datetime import datetime
from urllib.parse import urljoin, urlparse, parse_qs

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

logger = logging.getLogger(__name__)

# ==================== 常量 ====================

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9",
}

# AMAC 公募基金管理人 API
AMAC_API_URL = "https://www.amac.org.cn/portal/front/mutualFund/findMutualFundHousePage"

# AMAC 字段映射
AMAC_FIELD_CN = {
    "lineId": "序号",
    "houseName": "公司名称",
    "registerAddr": "注册地",
    "officeAddr": "辖区",
    "website": "官方网址",
    "phone": "客服电话",
}

# NFRA 银行法人机构列表 PDF 搜索关键词
NFRA_BANK_SEARCH_URL = "https://www.nfra.gov.cn/cn/view/pages/governmentDetail.html"
# NFRA 保险法人机构列表 PDF 搜索关键词
NFRA_INSURANCE_SEARCH_URL = "https://www.nfra.gov.cn/cn/view/pages/governmentDetail.html"

# NFRA 已知的银行/保险机构名单 PDF 页面 (通过搜索接口获取)
NFRA_SEARCH_API = "https://www.nfra.gov.cn/cn/search/Search.json"
NFRA_DOC_LIST_API = "https://www.nfra.gov.cn/cbircweb/DocInfo/SelectDocByItemIdAndChild"
NFRA_DOC_DETAIL_API = "https://www.nfra.gov.cn/cbircweb/DocInfo/SelectByDocId"
NFRA_BANK_ITEM_ID = "863"  # 银行业金融机构法人名单栏目ID
NFRA_BANK_KEYWORD = "银行业金融机构法人名单"
NFRA_INSURANCE_KEYWORD = "保险机构法人名单"

# CSRC 机构名录 Excel 下载地址
CSRC_SECURITIES_LIST_URL = "http://www.csrc.gov.cn/csrc/c100028/common_list.shtml"
CSRC_FUND_LIST_URL = "http://www.csrc.gov.cn/csrc/c100029/common_list.shtml"
CSRC_FUTURES_LIST_URL = "http://www.csrc.gov.cn/csrc/c100030/common_list.shtml"

# CSRC Excel 文件下载页面 API
CSRC_LIST_API = "http://www.csrc.gov.cn/csrc/{path}/list.shtml"

# ==================== 内部辅助函数 ====================

def _safe_request(url: str, params: dict = None, stream: bool = False,
                  timeout: int = 30, **kwargs) -> Optional[requests.Response]:
    """发送 HTTP 请求，统一错误处理"""
    try:
        resp = requests.get(url, params=params, headers=HEADERS,
                            timeout=timeout, stream=stream, **kwargs)
        return resp
    except requests.RequestException as e:
        logger.error(f"请求失败 [{url}]: {e}")
        return None


def _fetch_amac() -> Optional[List[Dict[str, str]]]:
    """从 AMAC 获取公募基金管理人名录"""
    try:
        resp = _safe_request(AMAC_API_URL, params={"pageNo": 1, "pageSize": 500})
        if resp is None:
            return None

        resp.raise_for_status()
        body = resp.json()

        if body.get("code") != 200:
            logger.error(f"AMAC API 返回错误: {body}")
            return None

        data = body.get("data", {})
        if data.get("errcode") != 0:
            logger.error(f"AMAC API 业务错误: {data.get('msg')}")
            return None

        inner = data.get("data", {})
        data_list = inner.get("dataList", [])
        total = inner.get("total", 0)

        logger.info(f"AMAC 获取到 {len(data_list)} 条公募基金管理人记录 (共 {total} 条)")

        result = []
        for item in data_list:
            row = {}
            for en_key, cn_key in AMAC_FIELD_CN.items():
                val = item.get(en_key, "")
                if isinstance(val, str):
                    val = val.replace("\n", " / ").replace("\r", "")
                row[cn_key] = val
            result.append(row)

        return result

    except requests.RequestException as e:
        logger.error(f"请求 AMAC API 失败: {e}")
        return None
    except Exception as e:
        logger.error(f"解析 AMAC 数据失败: {e}")
        return None


def _fetch_nfra_search(keyword: str) -> Optional[str]:
    """
    通过 NFRA 文档列表 API 查找指定关键词的最新公告详情页 URL

    返回：
        第一个匹配的公告详情页 URL，失败返回 None
    """
    try:
        params = {
            "itemId": NFRA_BANK_ITEM_ID,
            "pageSize": 20,
            "pageIndex": 1,
        }
        json_headers = {
            **HEADERS,
            "Accept": "application/json, text/javascript, */*; q=0.01",
        }
        resp = _safe_request(NFRA_DOC_LIST_API, params=params)
        if resp is None:
            return None

        body = resp.json()
        rows = body.get("data", {}).get("rows", [])
        if not rows:
            logger.warning(f"NFRA 文档列表无结果: {keyword}")
            return None

        # 按关键词匹配，取第一个
        for row in rows:
            title = row.get("docTitle", "")
            if keyword in title:
                doc_id = row.get("docId", "")
                if doc_id:
                    detail_url = f"{NFRA_BANK_SEARCH_URL}?docId={doc_id}&itemId={NFRA_BANK_ITEM_ID}"
                    logger.info(f"NFRA 找到 [{keyword}]: {detail_url}")
                    return detail_url

        logger.warning(f"NFRA 文档列表中未找到匹配 [{keyword}] 的文档")
        return None

    except Exception as e:
        logger.error(f"NFRA 搜索失败 [{keyword}]: {e}")
        return None


def _find_pdf_links_on_page(page_url: str) -> List[str]:
    """
    从 NFRA 公告详情页中提取 PDF 附件链接
    优先使用 API 获取附件，失败时回退到 HTML 解析

    返回：
        PDF URL 列表
    """
    try:
        # 从 URL 中提取 docId
        parsed = urlparse(page_url)
        params = parse_qs(parsed.query)
        doc_id = params.get("docId", [None])[0]

        if doc_id:
            # 优先使用文档详情 API 获取附件信息
            json_headers = {
                **HEADERS,
                "Accept": "application/json, text/javascript, */*; q=0.01",
            }
            detail_url = f"{NFRA_DOC_DETAIL_API}?docId={doc_id}"
            resp = _safe_request(detail_url, headers=json_headers)
            if resp is not None:
                try:
                    body = resp.json()
                    detail_data = body.get("data", {})
                    attachments = detail_data.get("attachmentInfoVOList", [])
                    pdf_urls = []
                    for att in attachments:
                        name = att.get("attachmentName", "")
                        url = att.get("urlOtherName", "")
                        if url and (name.lower().endswith(".pdf") or url.lower().endswith(".pdf")):
                            full_url = urljoin("https://www.nfra.gov.cn", url)
                            if full_url not in pdf_urls:
                                pdf_urls.append(full_url)
                    if pdf_urls:
                        logger.info(f"通过 API 找到 {len(pdf_urls)} 个 PDF 链接")
                        return pdf_urls
                except Exception:
                    logger.warning("API 获取附件失败，回退到 HTML 解析")

        # 回退到 HTML 解析
        resp = _safe_request(page_url)
        if resp is None:
            return []

        html = resp.text
        pdf_urls = []
        patterns = [
            r'href=["\']([^"\']*\.pdf[^"\']*)["\']',
            r'src=["\']([^"\']*\.pdf[^"\']*)["\']',
            r'["\'](/[^"\']*\.pdf[^"\']*)["\']',
        ]

        for pattern in patterns:
            matches = re.findall(pattern, html, re.IGNORECASE)
            for match in matches:
                if match.startswith("//"):
                    match = "https:" + match
                elif match.startswith("/"):
                    match = "https://www.nfra.gov.cn" + match
                elif not match.startswith("http"):
                    continue
                if match not in pdf_urls:
                    pdf_urls.append(match)

        logger.info(f"通过 HTML 找到 {len(pdf_urls)} 个 PDF 链接")
        return pdf_urls

    except Exception as e:
        logger.error(f"解析页面 PDF 链接失败: {e}")
        return []


def _download_file(url: str, output_dir: str, filename: str = None) -> Optional[str]:
    """
    下载文件到本地

    参数：
        url: 下载地址
        output_dir: 输出目录
        filename: 保存文件名（不含路径），为 None 时从 URL 提取

    返回：
        本地文件路径，失败返回 None
    """
    os.makedirs(output_dir, exist_ok=True)

    if filename is None:
        # 从 URL 提取文件名
        filename = url.split("/")[-1].split("?")[0]
        if not filename or "." not in filename:
            filename = f"download_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"

    filepath = os.path.join(output_dir, filename)

    try:
        resp = _safe_request(url, stream=True, timeout=60)
        if resp is None:
            return None

        resp.raise_for_status()

        total_size = int(resp.headers.get("content-length", 0))
        downloaded = 0

        with open(filepath, "wb") as f:
            for chunk in resp.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
                    downloaded += len(chunk)

        if total_size > 0 and downloaded < total_size * 0.9:
            logger.warning(f"文件下载不完整: {downloaded}/{total_size} bytes")

        logger.info(f"文件已下载: {filepath} ({downloaded} bytes)")
        return filepath

    except Exception as e:
        logger.error(f"下载文件失败 [{url}]: {e}")
        # 删除不完整的文件
        if os.path.exists(filepath):
            try:
                os.remove(filepath)
            except OSError:
                pass
        return None


def _fetch_csrc_list_page(csrc_path: str, page: int = 1) -> Optional[dict]:
    """
    获取 CSRC 机构名录页面数据

    参数：
        csrc_path: CSRC 路径，如 'c100028'
        page: 页码

    返回：
        JSON 响应体
    """
    url = CSRC_LIST_API.format(path=csrc_path)
    try:
        params = {
            "pageNo": page,
            "pageSize": 50,
        }
        resp = _safe_request(url, params=params)
        if resp is None:
            return None

        resp.raise_for_status()
        return resp.json()

    except Exception as e:
        logger.error(f"获取 CSRC 列表页失败 [{csrc_path}]: {e}")
        return None


def _find_excel_links_from_csrc(csrc_path: str) -> List[Tuple[str, str]]:
    """
    从 CSRC 机构名录页面查找 Excel 文件下载链接

    参数：
        csrc_path: CSRC 路径，如 'c100028'

    返回：
        [(下载URL, 文件名), ...] 列表
    """
    page_url = CSRC_LIST_API.format(path=csrc_path)
    try:
        resp = _safe_request(page_url)
        if resp is None:
            return []

        html = resp.text

        excel_links = []
        # 匹配 xls/xlsx 链接
        patterns = [
            r'href=["\']([^"\']*\.xlsx?[^"\']*)["\']',
            r'href=["\']([^"\']*\.xlsx?[^"\']*)["\']',
        ]

        seen = set()
        for pattern in patterns:
            matches = re.findall(pattern, html, re.IGNORECASE)
            for match in matches:
                if match in seen:
                    continue
                seen.add(match)

                # 处理相对路径
                if match.startswith("//"):
                    full_url = "https:" + match
                elif match.startswith("/"):
                    full_url = "http://www.csrc.gov.cn" + match
                elif match.startswith("http"):
                    full_url = match
                else:
                    full_url = f"http://www.csrc.gov.cn/csrc/{csrc_path}/{match}"

                filename = match.split("/")[-1].split("?")[0]
                excel_links.append((full_url, filename))

        logger.info(f"CSRC [{csrc_path}] 找到 {len(excel_links)} 个 Excel 链接")
        return excel_links

    except Exception as e:
        logger.error(f"解析 CSRC 页面 Excel 链接失败 [{csrc_path}]: {e}")
        return []


def _download_csrc_excel(url: str, output_dir: str, filename: str) -> Optional[str]:
    """下载 CSRC 的 Excel 文件"""
    os.makedirs(output_dir, exist_ok=True)
    filepath = os.path.join(output_dir, filename)

    try:
        resp = _safe_request(url, timeout=60)
        if resp is None:
            return None

        resp.raise_for_status()

        # 检查是否为 Excel 内容
        content_type = resp.headers.get("content-type", "").lower()
        if "excel" in content_type or "spreadsheet" in content_type or \
           filename.endswith((".xls", ".xlsx")):
            with open(filepath, "wb") as f:
                f.write(resp.content)
            logger.info(f"CSRC Excel 已下载: {filepath}")
            return filepath
        else:
            # 可能是 HTML 页面，尝试保存内容
            logger.warning(f"响应可能不是 Excel 文件: content-type={content_type}")
            with open(filepath, "wb") as f:
                f.write(resp.content)
            return filepath

    except Exception as e:
        logger.error(f"下载 CSRC Excel 失败 [{url}]: {e}")
        return None


# ==================== 公开接口 ====================

def fetch_all_institution_lists() -> Dict[str, List[Dict[str, str]]]:
    """
    获取全部金融机构名录

    返回：
        {
            'amac':          公募基金管理人列表,
            'bank_insurance': 银行保险机构列表,
            'securities':     证券公司列表,
            'fund':           基金公司列表,
            'futures':        期货公司列表,
        }
        每个列表为 [{"列名": 值, ...}, ...]
        获取失败的来源对应空列表
    """
    result = {
        "amac": [],
        "bank_insurance": [],
        "securities": [],
        "fund": [],
        "futures": [],
    }

    # 1. AMAC 公募基金管理人
    logger.info("正在获取 AMAC 公募基金管理人名录...")
    try:
        amac_data = _fetch_amac()
        if amac_data:
            result["amac"] = amac_data
        else:
            logger.warning("AMAC 数据获取失败，将跳过")
    except Exception as e:
        logger.error(f"AMAC 获取异常: {e}")

    # 2. NFRA 银行保险机构 (通过搜索获取 PDF 链接，解析为基本信息)
    logger.info("正在获取 NFRA 银行保险机构信息...")
    try:
        bank_url = _fetch_nfra_search(NFRA_BANK_KEYWORD)
        insurance_url = _fetch_nfra_search(NFRA_INSURANCE_KEYWORD)
        bank_insurance = []
        if bank_url:
            bank_insurance.append({
                "机构类型": "银行",
                "来源": "NFRA",
                "公告链接": bank_url,
                "获取时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            })
        if insurance_url:
            bank_insurance.append({
                "机构类型": "保险",
                "来源": "NFRA",
                "公告链接": insurance_url,
                "获取时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            })
        if bank_insurance:
            result["bank_insurance"] = bank_insurance
        else:
            logger.warning("NFRA 银行保险数据获取失败，将跳过")
    except Exception as e:
        logger.error(f"NFRA 获取异常: {e}")

    # 3. CSRC 证券/基金/期货公司
    logger.info("正在获取 CSRC 机构名录...")
    csrc_categories = {
        "securities": ("c100028", "证券公司"),
        "fund": ("c100029", "基金公司"),
        "futures": ("c100030", "期货公司"),
    }

    for key, (csrc_path, label) in csrc_categories.items():
        try:
            links = _find_excel_links_from_csrc(csrc_path)
            if links:
                for url, filename in links:
                    result[key].append({
                        "机构类型": label,
                        "来源": "CSRC",
                        "文件名": filename,
                        "下载链接": url,
                        "获取时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    })
            else:
                logger.warning(f"CSRC {label} 数据获取失败，将跳过")
        except Exception as e:
            logger.error(f"CSRC {label} 获取异常: {e}")

    total = sum(len(v) for v in result.values())
    logger.info(f"机构名录获取完成，共 {total} 条记录")
    return result


def write_institution_excel(data: Dict[str, List[Dict[str, str]]],
                            output_dir: str = "output") -> str:
    """
    将机构名录数据写入 Excel 文件

    每个数据来源一个 sheet：AMAC、银行保险、证券公司、基金公司、期货公司

    参数：
        data: fetch_all_institution_lists() 的返回值
        output_dir: 输出目录

    返回：
        生成的 Excel 文件路径
    """
    os.makedirs(output_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filepath = os.path.join(output_dir, f"金融机构名录_{timestamp}.xlsx")

    wb = Workbook()
    # 删除默认 sheet
    wb.remove(wb.active)

    # 样式定义
    header_font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1A5276", end_color="1A5276", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    data_font = Font(name="微软雅黑", size=10)
    data_align = Alignment(horizontal="left", vertical="center")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    # Sheet 配置
    sheets_config = [
        ("amac", "公募基金管理人 (AMAC)", "公募基金管理人"),
        ("bank_insurance", "银行保险机构 (NFRA)", "银行保险机构"),
        ("securities", "证券公司 (CSRC)", "证券公司"),
        ("fund", "基金公司 (CSRC)", "基金公司"),
        ("futures", "期货公司 (CSRC)", "期货公司"),
    ]

    for key, sheet_title, _ in sheets_config:
        records = data.get(key, [])

        ws = wb.create_sheet(title=sheet_title[:31])

        if not records:
            ws.cell(row=1, column=1, value="暂无数据").font = data_font
            ws.column_dimensions["A"].width = 20
            continue

        # 收集所有字段名
        headers = []
        seen = set()
        for record in records:
            for k in record:
                if k not in seen:
                    headers.append(k)
                    seen.add(k)

        # 写表头
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border
        ws.row_dimensions[1].height = 28

        # 写数据
        for row_idx, record in enumerate(records, 2):
            for col_idx, header in enumerate(headers, 1):
                value = record.get(header, "")
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.font = data_font
                cell.alignment = data_align
                cell.border = thin_border
            ws.row_dimensions[row_idx].height = 20

        # 列宽自适应
        for col_idx, header in enumerate(headers, 1):
            col_letter = get_column_letter(col_idx)
            width = max(len(str(header)) * 2, 12)
            ws.column_dimensions[col_letter].width = min(width, 50)

        # 冻结首行 + 自动筛选
        ws.freeze_panes = "A2"
        if len(records) > 0:
            ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(records) + 1}"

        logger.info(f"Sheet [{sheet_title}] 写入 {len(records)} 条记录")

    wb.save(filepath)
    logger.info(f"机构名录 Excel 已保存: {filepath}")
    return filepath


def download_nfra_pdfs(output_dir: str = "output") -> List[str]:
    """
    从 NFRA 网站下载银行和保险机构法人名单 PDF 文件

    流程：
      1. 通过 NFRA 搜索接口，分别搜索"银行业金融机构法人名单"和"保险机构法人名单"
      2. 获取对应公告详情页
      3. 从详情页中提取 PDF 附件链接
      4. 下载 PDF 到本地

    参数：
        output_dir: 输出目录

    返回：
        已下载的 PDF 文件路径列表
    """
    os.makedirs(output_dir, exist_ok=True)
    downloaded = []

    targets = [
        (NFRA_BANK_KEYWORD, "银行业金融机构法人名单.pdf"),
        (NFRA_INSURANCE_KEYWORD, "保险机构法人名单.pdf"),
    ]

    for keyword, default_filename in targets:
        try:
            logger.info(f"正在搜索 NFRA: {keyword}")

            # 搜索公告
            detail_url = _fetch_nfra_search(keyword)
            if not detail_url:
                logger.warning(f"NFRA 搜索无结果: {keyword}")
                continue

            # 从详情页获取 PDF 链接
            pdf_urls = _find_pdf_links_on_page(detail_url)
            if not pdf_urls:
                logger.warning(f"NFRA 详情页未找到 PDF 链接: {detail_url}")
                continue

            # 下载找到的 PDF（取第一个）
            pdf_url = pdf_urls[0]
            filename = pdf_url.split("/")[-1].split("?")[0]
            if not filename.endswith(".pdf"):
                filename = default_filename

            filepath = _download_file(pdf_url, output_dir, filename)
            if filepath:
                downloaded.append(filepath)

        except Exception as e:
            logger.error(f"NFRA PDF 下载异常 [{keyword}]: {e}")

    logger.info(f"NFRA PDF 下载完成，共 {len(downloaded)} 个文件")
    return downloaded


def download_csrc_to_combined_xlsx(output_dir: str = "output") -> str:
    """
    从 CSRC 网站下载证券公司、基金公司、期货公司的 Excel 名录，
    并合并为一个 xlsx 文件（每个机构类型一个 sheet）

    流程：
      1. 分别访问 CSRC 证券公司(c100028)、基金公司(c100029)、期货公司(c100030)页面
      2. 从页面中提取 Excel 文件下载链接
      3. 下载 Excel 文件
      4. 合并到一个 xlsx 中，每个机构类型一个 sheet

    参数：
        output_dir: 输出目录

    返回：
        合并后的 xlsx 文件路径
    """
    os.makedirs(output_dir, exist_ok=True)

    # 需要下载的 CSRC 名录
    csrc_targets = [
        ("c100028", "证券公司"),
        ("c100029", "基金公司"),
        ("c100030", "期货公司"),
    ]

    # 先下载各个 Excel 文件
    downloaded_files: Dict[str, str] = {}  # {label: filepath}

    for csrc_path, label in csrc_targets:
        try:
            logger.info(f"正在获取 CSRC {label} 名录...")

            links = _find_excel_links_from_csrc(csrc_path)
            if not links:
                logger.warning(f"CSRC {label} 页面未找到 Excel 链接")
                continue

            # 下载第一个 Excel 文件
            url, filename = links[0]
            if not filename.endswith((".xls", ".xlsx")):
                filename = f"{label}名录.xlsx"

            filepath = _download_csrc_excel(url, output_dir, filename)
            if filepath:
                downloaded_files[label] = filepath

        except Exception as e:
            logger.error(f"CSRC {label} 下载异常: {e}")

    if not downloaded_files:
        logger.warning("CSRC 未下载到任何文件，生成空合并文件")

    # 合并为一个 xlsx
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    combined_path = os.path.join(output_dir, f"CSRC机构名录_合并_{timestamp}.xlsx")

    wb = Workbook()
    wb.remove(wb.active)

    for label, src_path in downloaded_files.items():
        try:
            # 尝试读取下载的 Excel 文件
            from openpyxl import load_workbook

            if not os.path.exists(src_path):
                logger.warning(f"CSRC 文件不存在: {src_path}")
                continue

            src_wb = load_workbook(src_path, read_only=True, data_only=True)

            for src_ws in src_wb.worksheets:
                # Sheet 名称
                sheet_title = f"{label}_{src_ws.title}"[:31]
                dst_ws = wb.create_sheet(title=sheet_title)

                # 复制数据
                for row in src_ws.iter_rows(values_only=True):
                    dst_ws.append(list(row))

                logger.info(f"合并 sheet [{sheet_title}]: {dst_ws.max_row} 行")

            src_wb.close()

        except Exception as e:
            logger.error(f"合并 CSRC {label} 文件失败: {e}")
            # 创建占位 sheet
            ws = wb.create_sheet(title=f"{label}(合并失败)"[:31])
            ws.cell(row=1, column=1, value=f"合并失败: {e}")

    # 如果没有任何 sheet，创建一个占位 sheet
    if len(wb.worksheets) == 0:
        ws = wb.create_sheet(title="暂无数据")
        ws.cell(row=1, column=1, value="未获取到 CSRC 机构名录数据")

    wb.save(combined_path)
    logger.info(f"CSRC 合并 Excel 已保存: {combined_path}")
    return combined_path


# ==================== 模块自测 ====================

if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )

    print("=" * 60)
    print("  金融机构名录爬取工具 - 自测")
    print("=" * 60)

    # 测试 1: 获取所有机构名录
    print("\n[1] 获取所有机构名录...")
    data = fetch_all_institution_lists()
    for key, records in data.items():
        print(f"  - {key}: {len(records)} 条记录")

    # 测试 2: 写入 Excel
    print("\n[2] 写入机构名录 Excel...")
    excel_path = write_institution_excel(data, output_dir="output")
    print(f"  已保存: {excel_path}")

    # 测试 3: 下载 NFRA PDF
    print("\n[3] 下载 NFRA PDF...")
    pdf_paths = download_nfra_pdfs(output_dir="output")
    for p in pdf_paths:
        print(f"  已下载: {p}")

    # 测试 4: 下载并合并 CSRC Excel
    print("\n[4] 下载并合并 CSRC Excel...")
    csrc_path = download_csrc_to_combined_xlsx(output_dir="output")
    print(f"  已保存: {csrc_path}")

    print("\n" + "=" * 60)
    print("  自测完成")
    print("=" * 60)