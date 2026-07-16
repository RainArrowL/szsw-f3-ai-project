#!/bin/bash
set -e
cd "$(dirname "$0")" || exit 1

echo "========================================"
echo "  税金智 - 启动服务"
echo "========================================"
echo ""

# 检查虚拟环境
if [ ! -f ".venv/bin/activate" ]; then
    echo "[错误] 未找到虚拟环境，请先运行: bash install.sh"
    exit 1
fi

# 激活虚拟环境
source .venv/bin/activate

# 检查依赖
python3 -c "import flask" 2>/dev/null || {
    echo "[错误] 依赖未安装，请先运行: bash install.sh"
    exit 1
}

echo "[启动] Web 服务 http://localhost:5000"
echo "[提示] 按 Ctrl+C 停止服务"
echo ""

python3 src/app.py