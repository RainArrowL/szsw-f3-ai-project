#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
年报财务数据获取工具 - 交互式主入口

功能：
1. 从深证信数据服务平台 (webapi.cninfo.com.cn) / 东方财富 获取上市公司年报财务数据
2. 支持资产负债表、利润表、现金流量表
3. 输出格式：一家企业一个Excel，每年度每个报表一个sheet
4. 列名自动翻译为中文

使用方式：
    python main.py                        # 交互式模式
    python main.py --help                 # 查看帮助
"""

import sys
import logging
from datetime import datetime

from config import config
from cninfo_api import get_api
from cninfo_fin_data import FinancialDataFetcher, REPORT_NAMES_CN
from excel_writer import write_all_companies

# 日志配置
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)


# ==================== 辅助函数 ====================

def print_banner():
    """打印程序Banner"""
    print("\n" + "=" * 60)
    print("    上市公司年报财务数据获取工具")
    print("    Annual Report Financial Data Fetcher")
    print("=" * 60)
    print("    数据源: deepsearch_data_platform (webapi.cninfo.com.cn)")
    print("    备用数据源: 东方财富网 (免费公开)")
    print(f"    输出目录: {config.output_dir}/")
    print("=" * 60)


def input_year_range() -> tuple:
    """交互式输入年份范围"""
    current_year = datetime.now().year

    while True:
        try:
            print(f"\n请输入年度范围（当前年份: {current_year}）")
            start = input("  开始年份（如 2020）: ").strip()
            end = input("  截止年份（如 2024）: ").strip()

            start_year = int(start)
            end_year = int(end)

            if start_year > end_year:
                start_year, end_year = end_year, start_year
                print(f"  [提示] 已自动调换顺序: {start_year} ~ {end_year}")

            if start_year < 1990:
                print("  [警告] 开始年份较早，部分数据可能不完整")
            if end_year > current_year:
                print("  [警告] 截止年份超出当前年份，最新年份数据可能尚未发布")

            confirm = input(f"\n  确认年度范围: {start_year} ~ {end_year} 年? (y/n): ").strip().lower()
            if confirm in ("y", "yes", ""):
                return start_year, end_year

        except ValueError:
            print("  [错误] 请输入有效的年份数字")
        except KeyboardInterrupt:
            print("\n\n已取消")
            sys.exit(0)


def input_company_list() -> list:
    """交互式输入企业名单"""
    print("\n" + "-" * 40)
    print("请输入企业名单（支持以下输入方式）:")
    print("  1. 股票代码: 600519,000001,002415")
    print("  2. 公司简称: 贵州茅台,平安银行,海康威视")
    print("  3. 从文件读取: file:companies.txt")
    print("     (每行一个代码/名称，逗号或换行分隔)")
    print("-" * 40)

    while True:
        raw = input("\n请输入: ").strip()
        if not raw:
            print("  [错误] 输入不能为空")
            continue

        # 从文件读取
        if raw.startswith("file:"):
            filepath = raw[5:].strip()
            try:
                with open(filepath, "r", encoding="utf-8") as f:
                    content = f.read()
                # 支持换行或逗号分隔
                items = []
                for line in content.split("\n"):
                    line = line.strip()
                    if not line or line.startswith("#"):
                        continue
                    for item in line.split(","):
                        item = item.strip()
                        if item:
                            items.append(item)
                raw = ",".join(items)
                print(f"  从文件读取到 {len(items)} 条记录")
            except FileNotFoundError:
                print(f"  [错误] 文件不存在: {filepath}")
                continue
            except Exception as e:
                print(f"  [错误] 读取文件失败: {e}")
                continue

        # 解析企业列表
        companies = []
        for item in raw.replace("，", ",").split(","):
            item = item.strip()
            if item:
                companies.append(item)

        if not companies:
            print("  [错误] 未解析到有效企业")
            continue

        print(f"\n  已识别 {len(companies)} 家企业:")
        for i, c in enumerate(companies):
            print(f"    {i + 1}. {c}")

        confirm = input(f"\n  确认以上企业名单? (y/n): ").strip().lower()
        if confirm in ("y", "yes", ""):
            return companies

        print("  请重新输入企业名单")


def resolve_companies(companies: list) -> list:
    """
    解析企业名单，返回 (股票代码, 公司名称) 列表

    自动识别：
    - 纯数字 -> 股票代码
    - 非纯数字 -> 公司名称（需要搜索获取代码）
    """
    fetcher = FinancialDataFetcher()
    resolved = []

    for item in companies:
        # 如果是纯数字/字母，视为股票代码
        if item.replace(".", "").replace("-", "").isalnum() and any(c.isdigit() for c in item):
            # 去掉后缀如 .SH .SZ
            code = item.split(".")[0].split("-")[0].strip()
            # 尝试搜索确认存在
            try:
                results = fetcher.get_stock_info(code)
                if results:
                    resolved.append((results[0]["code"], results[0]["name"]))
                    print(f"  [{item}] -> {results[0]['code']} {results[0]['name']}")
                else:
                    # 直接使用输入的代码
                    resolved.append((code, code))
                    print(f"  [{item}] -> {code} (未搜索到名称，使用代码)")
            except Exception:
                resolved.append((code, code))
                print(f"  [{item}] -> {code}")
        else:
            # 视为公司名称，需要搜索
            try:
                print(f"  正在搜索: {item}...")
                results = fetcher.get_stock_info(item)
                if results:
                    best = results[0]
                    resolved.append((best["code"], best["name"]))
                    print(f"  [{item}] -> {best['code']} {best['name']}")
                else:
                    print(f"  [警告] 未找到 '{item}'，已跳过")
            except Exception as e:
                print(f"  [警告] 搜索 '{item}' 失败({e})，已跳过")

    return resolved


def check_credentials():
    """检查API凭证是否配置"""
    print("\n[检查] 深证信API凭证状态...")
    if config.is_credential_set():
        print("  [OK] 已配置深证信API凭证，将优先使用cninfo API获取数据")
        return True
    else:
        print("  [INFO] 未配置深证信API凭证")
        print("  [INFO] 将使用东方财富免费公开数据源")
        print("  [INFO] 如需使用cninfo官方API，请:")
        print("    1. 访问 http://webapi.cninfo.com.cn/ 注册")
        print("    2. 在 config.py 中设置 ACCESS_KEY 和 ACCESS_SECRET")
        return False


# ==================== 主流程 ====================

def run_interactive():
    """交互式运行模式"""
    print_banner()

    # 检查凭证
    use_cninfo = check_credentials()

    # 输入年份范围
    start_year, end_year = input_year_range()

    # 输入企业名单
    raw_companies = input_company_list()

    # 解析企业名单
    print("\n" + "-" * 40)
    print("正在解析企业信息...")
    companies = resolve_companies(raw_companies)

    if not companies:
        print("\n[错误] 没有有效的企业信息，程序退出")
        sys.exit(1)

    # 确认开始
    print("\n" + "-" * 40)
    print("将要获取以下数据:")
    print(f"  年度范围: {start_year} ~ {end_year} 年")
    print(f"  企业数量: {len(companies)}")
    print("  报表类型: 资产负债表、利润表、现金流量表")
    print("  输出格式: 每家企业一个Excel文件")
    print(f"  数据源: {'cninfo API' if use_cninfo else '东方财富(免费)'}")
    print("-" * 40)

    confirm = input("\n确认开始获取? (y/n): ").strip().lower()
    if confirm not in ("y", "yes", ""):
        print("已取消")
        sys.exit(0)

    # 执行数据获取
    print("\n" + "=" * 60)
    print("开始获取财务数据...")
    print("=" * 60)

    fetcher = FinancialDataFetcher()
    fetcher.set_use_cninfo(use_cninfo)

    start_time = datetime.now()

    # 进度显示
    def show_progress(current, total, status):
        pct = current / total * 100
        bar_len = 30
        filled = int(bar_len * current / total)
        bar = "█" * filled + "░" * (bar_len - filled)
        print(f"\r  进度: [{bar}] {current}/{total} ({pct:.1f}%) {status}", end="")

    all_data = fetcher.fetch_multiple_companies(
        companies, start_year, end_year, progress_callback=show_progress
    )
    print()  # 换行

    # 写入Excel
    print("\n" + "=" * 60)
    print("正在写入Excel文件...")
    print("=" * 60)

    files = write_all_companies(all_data, progress_callback=show_progress)
    print()

    # 完成
    elapsed = (datetime.now() - start_time).total_seconds()
    print("\n" + "=" * 60)
    print("  数据获取完成!")
    print(f"  耗时: {elapsed:.1f} 秒")
    print(f"  生成文件数: {len(files)}")
    print(f"  输出目录: {config.output_dir}/")
    print("=" * 60)

    print("\n生成的文件列表:")
    for f in files:
        print(f"  - {f}")


def run_quick_mode(companies_str: str, start_year: int, end_year: int):
    """
    快速模式：通过命令行参数直接运行

    使用示例:
        python main.py --companies "600519,000001" --start 2020 --end 2024
    """
    print_banner()
    use_cninfo = check_credentials()

    # 解析企业
    raw = [c.strip() for c in companies_str.replace("，", ",").split(",") if c.strip()]
    companies = resolve_companies(raw)

    if not companies:
        print("没有有效的企业信息")
        return

    print(f"\n开始获取: {len(companies)} 家企业, {start_year}-{end_year}年")

    fetcher = FinancialDataFetcher()
    fetcher.set_use_cninfo(use_cninfo)

    all_data = fetcher.fetch_multiple_companies(companies, start_year, end_year)
    files = write_all_companies(all_data)

    print(f"\n完成! 生成 {len(files)} 个文件")
    for f in files:
        print(f"  {f}")


# ==================== 入口 ====================

if __name__ == "__main__":
    if len(sys.argv) > 1:
        # 命令行参数模式
        import argparse

        parser = argparse.ArgumentParser(
            description="上市公司年报财务数据获取工具",
            formatter_class=argparse.RawDescriptionHelpFormatter,
            epilog="""
使用示例:
  python main.py                                              # 交互式模式
  python main.py --companies "600519,000001" --start 2020 --end 2024
  python main.py -c "贵州茅台,平安银行" -s 2022 -e 2023
  python main.py -c "file:companies.txt" -s 2020 -e 2024
            """
        )
        parser.add_argument(
            "-c", "--companies",
            type=str,
            help="企业名单，逗号分隔，支持股票代码/名称/file:文件路径"
        )
        parser.add_argument(
            "-s", "--start",
            type=int,
            help="开始年份"
        )
        parser.add_argument(
            "-e", "--end",
            type=int,
            help="截止年份"
        )
        parser.add_argument(
            "-o", "--output",
            type=str,
            default=config.output_dir,
            help=f"输出目录 (默认: {config.output_dir})"
        )

        args = parser.parse_args()

        if args.output:
            config.output_dir = args.output

        if args.companies and args.start and args.end:
            run_quick_mode(args.companies, args.start, args.end)
        else:
            parser.print_help()
    else:
        # 交互式模式
        try:
            run_interactive()
        except KeyboardInterrupt:
            print("\n\n已取消")
            sys.exit(0)