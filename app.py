#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Web界面 - Flask后端服务
年报财务数据获取工具 Web 版
"""

import os
import re
import uuid
import time
import logging
import threading
from datetime import datetime
from functools import wraps
from typing import Dict, List, Optional, Tuple
from werkzeug.utils import secure_filename
from flask import Flask, render_template, request, jsonify, send_from_directory, abort

from config import config
from cninfo_fin_data import FinancialDataFetcher, resolve_companies
from excel_writer import write_company_excel, write_industry_avg_excel
from industry_avg import compute_all_industry_averages, get_company_industry
from amac_scraper import fetch_fund_manager_list, write_amac_excel
from szse_scraper import fetch_year_data, write_szse_excel

# 日志配置
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

# 初始化Flask
app = Flask(__name__, template_folder="templates", static_folder="static")
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'annual-report-fetcher-secret')
app.config['MAX_CONTENT_LENGTH'] = 1 * 1024 * 1024  # 最大1MB文件

# 允许上传的文件类型
ALLOWED_EXTENSIONS = {'txt', 'csv'}

# 全局任务存储
tasks: Dict[str, Dict] = {}


# ==================== 工具函数 ====================

def allowed_file(filename):
    """检查文件扩展名是否允许"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def sanitize_filename(name):
    """文件名消毒，防止路径遍历"""
    name = re.sub(r'[^\w\s\-_.()（）]', '_', name)
    return secure_filename(name)


def parse_companies_from_text(text: str) -> List[str]:
    """从文本中解析企业列表"""
    companies = []
    for line in text.split('\n'):
        line = line.strip()
        if not line or line.startswith('#'):
            continue
        for item in line.replace('，', ',').split(','):
            item = item.strip()
            if item:
                companies.append(item)
    return companies


# ==================== 后台任务 ====================

def process_task(task_id: str, start_year: int, end_year: int, raw_companies: List[str], industry_avg_enabled: bool = False, amac_list_enabled: bool = False):
    """
    后台线程执行数据获取任务
    """
    task = tasks[task_id]
    task['status'] = 'processing'

    try:
        fetcher = FinancialDataFetcher()
        fetcher.set_use_cninfo(config.is_credential_set())

        # 解析企业信息
        task['progress']['message'] = "正在解析企业信息..."
        resolved = resolve_companies(raw_companies)

        if not resolved:
            raise ValueError("未能解析到有效企业信息，请检查输入")

        total = len(resolved)
        task['progress']['total'] = total

        # 获取各企业自身数据
        for idx, (code, name) in enumerate(resolved):
            current = idx + 1
            company_display = f"{name}({code})"
            task['progress']['current'] = current
            task['progress']['message'] = f"正在获取: {company_display}"

            logger.info(f"任务 {task_id}: [{current}/{total}] {company_display}")

            data = fetcher.fetch_company_data(code, start_year, end_year)

            # 写入Excel（企业自身数据，不含行业均值）
            filepath = write_company_excel(company_display, data, output_dir=config.output_dir)

            file_size = os.path.getsize(filepath) if os.path.exists(filepath) else 0
            task['files'].append({
                'name': os.path.basename(filepath),
                'path': filepath,
                'size': file_size,
                'display_name': f"{company_display}_年报财务数据.xlsx",
            })

            time.sleep(0.5)

        # 如果开启了行业均值，计算并输出到单独Excel
        if industry_avg_enabled:
            task['progress']['message'] = "正在计算行业平均值..."
            logger.info(f"任务 {task_id}: 开始计算行业平均值")

            # 收集各企业的行业信息，按行业去重
            industry_map: Dict[str, Tuple[str, str, List[str]]] = {}  # ind_name -> (ind_code, ind_name, [codes])

            for code, name in resolved:
                ind = get_company_industry(code)
                if ind and ind[1]:
                    ind_code, ind_name = ind
                    if ind_name not in industry_map:
                        industry_map[ind_name] = (ind_code, ind_name, [])
                    industry_map[ind_name][2].append(code)

            logger.info(f"共 {len(industry_map)} 个不同行业")

            # 为每个行业计算均值并写入一个Excel
            for ind_name, (ind_code, ind_full_name, codes) in industry_map.items():
                if not codes:
                    continue
                # 取第一个股票代码代表该行业
                rep_code = codes[0]
                task['progress']['message'] = f"计算行业均值: {ind_name}"

                industry_avg = compute_all_industry_averages(
                    stock_code=rep_code,
                    ind_name=ind_full_name,
                    ind_code=ind_code,
                    start_year=start_year,
                    end_year=end_year,
                )

                if industry_avg:
                    # 行业Excel文件名
                    safe_ind = ind_name.replace("/", "_").replace("\\", "_").replace(":", "_")
                    ia_filepath = write_industry_avg_excel(
                        ind_name, industry_avg, output_dir=config.output_dir
                    )
                    file_size = os.path.getsize(ia_filepath) if os.path.exists(ia_filepath) else 0
                    task['files'].append({
                        'name': os.path.basename(ia_filepath),
                        'path': ia_filepath,
                        'size': file_size,
                        'display_name': f"行业均值_{safe_ind}_年报财务数据.xlsx",
                    })

        # 如果开启了公募基金管理人名录
        if amac_list_enabled:
            task['progress']['message'] = "正在爬取公募基金管理人名录..."
            logger.info(f"任务 {task_id}: 开始爬取公募基金管理人名录")

            amac_records = fetch_fund_manager_list()
            if amac_records:
                amac_filepath = write_amac_excel(amac_records, output_dir=config.output_dir)
                file_size = os.path.getsize(amac_filepath) if os.path.exists(amac_filepath) else 0
                task['files'].append({
                    'name': os.path.basename(amac_filepath),
                    'path': amac_filepath,
                    'size': file_size,
                    'display_name': f"公募基金管理人名录.xlsx",
                })
            else:
                logger.warning(f"任务 {task_id}: 公募基金管理人名录爬取失败")

        task['status'] = 'done'
        task['progress']['message'] = f"完成! 共生成 {len(task['files'])} 个文件"
        logger.info(f"任务 {task_id} 完成: {len(task['files'])} 个文件")

    except Exception as e:
        task['status'] = 'error'
        task['error'] = str(e)
        logger.error(f"任务 {task_id} 失败: {e}", exc_info=True)


# ==================== 路由 ====================

@app.route('/')
def index():
    """首页"""
    current_year = datetime.now().year
    return render_template('index.html', current_year=current_year)


@app.route('/api/fetch', methods=['POST'])
def fetch_data():
    """提交数据获取任务"""

    # 获取表单数据
    form = request.form
    start_year = form.get('start_year', type=int)
    end_year = form.get('end_year', type=int)
    text_input = form.get('text_input', '').strip()

    # 校验年份
    current_year = datetime.now().year
    if not start_year or not end_year:
        return jsonify({'success': False, 'error': '请选择年份范围'}), 400
    if start_year < 1990 or end_year > current_year:
        return jsonify({'success': False, 'error': f'年份范围应在 1990-{current_year} 之间'}), 400
    if start_year > end_year:
        start_year, end_year = end_year, start_year

    # 解析企业列表
    companies = []
    if text_input:
        companies = parse_companies_from_text(text_input)
    # 如果有文件上传，覆盖文本输入
    if 'file' in request.files:
        file = request.files['file']
        if file and file.filename and allowed_file(file.filename):
            content = file.read().decode('utf-8')
            companies = parse_companies_from_text(content)

    if not companies:
        return jsonify({'success': False, 'error': '请输入至少一个企业'}), 400

    # 是否计算行业均值
    industry_avg_enabled = request.form.get('industry_avg', '0') == '1'
    # 是否爬取公募基金管理人名录
    amac_list_enabled = request.form.get('amac_list', '0') == '1'

    # 创建任务
    task_id = str(uuid.uuid4())[:8]
    tasks[task_id] = {
        'id': task_id,
        'start_year': start_year,
        'end_year': end_year,
        'companies': companies,
        'status': 'pending',
        'progress': {
            'current': 0,
            'total': len(companies),
            'message': '等待开始...',
        },
        'files': [],
        'error': None,
        'created_at': time.time(),
    }

    # 启动后台线程
    thread = threading.Thread(
        target=process_task,
        args=(task_id, start_year, end_year, companies, industry_avg_enabled, amac_list_enabled),
        daemon=True
    )
    thread.start()

    return jsonify({
        'success': True,
        'task_id': task_id,
    })


@app.route('/api/progress/<task_id>')
def get_progress(task_id):
    """查询任务进度"""
    if task_id not in tasks:
        return jsonify({'success': False, 'error': '任务不存在'}), 404

    return jsonify({
        'success': True,
        'task': tasks[task_id],
    })


@app.route('/download/<filename>')
def download_file(filename):
    """下载生成的Excel文件"""
    # 安全检查：防止路径遍历
    filename = os.path.basename(filename)
    full_path = os.path.join(config.output_dir, filename)

    if not os.path.exists(full_path):
        abort(404)

    return send_from_directory(
        config.output_dir,
        filename,
        as_attachment=True,
        download_name=filename
    )


@app.route('/api/status')
def get_status():
    """获取服务状态"""
    return jsonify({
        'success': True,
        'credential_configured': config.is_credential_set(),
    })


# ==================== AMAC 名录独立端点 ====================

def process_amac_task(task_id: str):
    """
    后台线程执行公募基金管理人名录爬取任务
    """
    task = tasks[task_id]
    task['status'] = 'processing'

    try:
        task['progress']['message'] = "正在爬取公募基金管理人名录..."
        logger.info(f"AMAC任务 {task_id}: 开始爬取")

        amac_records = fetch_fund_manager_list()
        if amac_records:
            amac_filepath = write_amac_excel(amac_records, output_dir=config.output_dir)
            file_size = os.path.getsize(amac_filepath) if os.path.exists(amac_filepath) else 0
            task['files'].append({
                'name': os.path.basename(amac_filepath),
                'path': amac_filepath,
                'size': file_size,
                'display_name': "公募基金管理人名录.xlsx",
            })
            task['progress']['total'] = 1
            task['progress']['current'] = 1
        else:
            raise ValueError("公募基金管理人名录爬取失败，未获取到数据")

        task['status'] = 'done'
        task['progress']['message'] = "完成! 名录已生成"
        logger.info(f"AMAC任务 {task_id} 完成")

    except Exception as e:
        task['status'] = 'error'
        task['error'] = str(e)
        logger.error(f"AMAC任务 {task_id} 失败: {e}", exc_info=True)


@app.route('/api/amac', methods=['POST'])
def fetch_amac():
    """提交公募基金管理人名录爬取任务"""
    # 创建任务
    task_id = str(uuid.uuid4())[:8]
    tasks[task_id] = {
        'id': task_id,
        'status': 'pending',
        'progress': {
            'current': 0,
            'total': 1,
            'message': '等待开始...',
        },
        'files': [],
        'error': None,
        'created_at': time.time(),
    }

    # 启动后台线程
    thread = threading.Thread(
        target=process_amac_task,
        args=(task_id,),
        daemon=True
    )
    thread.start()

    return jsonify({
        'success': True,
        'task_id': task_id,
    })


# ==================== 深交所日度概况端点 ====================

def process_szse_task(task_id: str, year: int):
    """
    后台线程执行深交所日度概况爬取任务
    """
    task = tasks[task_id]
    task['status'] = 'processing'

    try:
        def progress_callback(current, total, message):
            task['progress']['current'] = current
            task['progress']['total'] = total
            task['progress']['message'] = message

        task['progress']['message'] = f"正在爬取 {year} 年深交所日度概况..."
        logger.info(f"SZSE任务 {task_id}: 开始爬取 {year} 年数据")

        data = fetch_year_data(year, progress_callback=progress_callback)

        if data:
            filepath = write_szse_excel(year, data, output_dir=config.output_dir)
            file_size = os.path.getsize(filepath) if os.path.exists(filepath) else 0
            task['files'].append({
                'name': os.path.basename(filepath),
                'path': filepath,
                'size': file_size,
                'display_name': f"深交所日度概况_{year}年.xlsx",
            })
            task['progress']['message'] = f"完成! {year}年共 {len(data)} 个交易日"
        else:
            raise ValueError(f"{year}年未获取到任何深交所日度概况数据")

        task['status'] = 'done'
        logger.info(f"SZSE任务 {task_id} 完成: {len(data)} 条记录")

    except Exception as e:
        task['status'] = 'error'
        task['error'] = str(e)
        logger.error(f"SZSE任务 {task_id} 失败: {e}", exc_info=True)


@app.route('/api/szse', methods=['POST'])
def fetch_szse():
    """提交深交所日度概况爬取任务"""
    year = request.form.get('year', type=int)
    if not year:
        year = datetime.now().year

    task_id = str(uuid.uuid4())[:8]
    tasks[task_id] = {
        'id': task_id,
        'status': 'pending',
        'progress': {
            'current': 0,
            'total': 260,  # 大约260个交易日
            'message': '等待开始...',
        },
        'files': [],
        'error': None,
        'created_at': time.time(),
    }

    thread = threading.Thread(
        target=process_szse_task,
        args=(task_id, year),
        daemon=True
    )
    thread.start()

    return jsonify({
        'success': True,
        'task_id': task_id,
    })


# ==================== 启动 ====================

if __name__ == '__main__':
    # 创建输出目录
    os.makedirs(config.output_dir, exist_ok=True)

    # 检查依赖
    print("=" * 60)
    print("  上市公司年报财务数据获取工具 - Web 版")
    print("=" * 60)
    if config.is_credential_set():
        print("  [OK] 深证信API凭证已配置")
    else:
        print("  [INFO] 未配置深证信API凭证，将使用东方财富免费数据源")
    print(f"  [INFO] 输出目录: {config.output_dir}/")
    print(f"  [INFO] 服务将在: http://127.0.0.1:5000")
    print("=" * 60)
    print()

    app.run(host='0.0.0.0', port=5000, debug=False)
