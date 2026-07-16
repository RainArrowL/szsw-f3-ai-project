#!/bin/bash
set -e
cd "$(dirname "$0")" || exit 1

echo "========================================"
echo "  税金智 - 安装依赖"
echo "========================================"
echo ""

# 检查 Python
if ! command -v python3 &>/dev/null; then
    echo "[错误] 未检测到 python3，请先安装 Python 3.9+"
    exit 1
fi

echo "[检测] Python: $(python3 --version)"
echo ""

# 创建虚拟环境
if [ ! -d ".venv" ]; then
    echo "[创建] 虚拟环境 .venv ..."
    python3 -m venv .venv
    echo "[完成] 虚拟环境创建成功"
else
    echo "[跳过] 虚拟环境 .venv 已存在"
fi

echo ""

# 激活虚拟环境并安装依赖
echo "[安装] Python 依赖包 ..."
source .venv/bin/activate
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple 2>/dev/null || pip install -r requirements.txt

echo ""
echo "========================================"
echo "  安装完成！"
echo "========================================"
echo ""
echo "运行方式: bash run.sh"
echo ""