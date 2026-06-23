#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
配置文件
深证信API凭证配置，输出路径配置等
"""

import os
from dataclasses import dataclass

# ==================== 深证信API凭证配置 ====================
# 请在 http://webapi.cninfo.com.cn/ 注册后，在个人中心-我的凭证中获取
ACCESS_KEY = ""       # 填写你的 Access Key
ACCESS_SECRET = ""    # 填写你的 Access Secret

# ==================== API基础URL ====================
BASE_URL = "http://webapi.cninfo.com.cn/api"

# ==================== 输出配置 ====================
# Excel输出目录
OUTPUT_DIR = "output"
# 是否创建输出目录
CREATE_OUTPUT_DIR = True

# ==================== 请求配置 ====================
# 请求超时时间（秒）
TIMEOUT = 30
# 请求间隔（秒），避免请求过于频繁被限制
REQUEST_INTERVAL = 1.0
# 每页最大返回记录数
MAX_ROWS_PER_PAGE = 20000

# ==================== 财务报表类型 ====================
REPORT_TYPES = {
    "balance": "资产负债表",
    "income": "利润表",
    "cashflow": "现金流量表",
}

# ==================== 财报期间类型 ====================
# 4 = 年报
PERIOD_TYPE_ANNUAL = "4"


@dataclass
class Config:
    """配置类"""
    access_key: str = ACCESS_KEY
    access_secret: str = ACCESS_SECRET
    base_url: str = BASE_URL
    output_dir: str = OUTPUT_DIR
    timeout: int = TIMEOUT
    request_interval: float = REQUEST_INTERVAL
    max_rows_per_page: int = MAX_ROWS_PER_PAGE

    def __init__(self):
        # 从环境变量读取（如果存在）
        self.access_key = os.environ.get("CNINFO_ACCESS_KEY", self.access_key)
        self.access_secret = os.environ.get("CNINFO_ACCESS_SECRET", self.access_secret)
        self.output_dir = os.environ.get("CNINFO_OUTPUT_DIR", self.output_dir)

    def is_credential_set(self) -> bool:
        """检查凭证是否已设置"""
        return bool(self.access_key and self.access_secret)


# 默认配置实例
config = Config()


def get_config() -> Config:
    """获取配置"""
    return config
