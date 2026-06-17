#!/bin/bash
# =========================================
#  智览金融 财数贯通 - 统信UOS 启动脚本
#  首次运行自动安装依赖，后续直接启动
# =========================================

cd "$(dirname "$0")"

echo "========================================"
echo "      智览金融 财数贯通"
echo "========================================"
echo ""

# 检查 python3
if ! command -v python3 &> /dev/null; then
    echo "Python3 未安装，正在安装..."
    if command -v apt &> /dev/null; then
        sudo apt update && sudo apt install -y python3 python3-pip
    elif command -v dnf &> /dev/null; then
        sudo dnf install -y python3 python3-pip
    elif command -v yum &> /dev/null; then
        sudo yum install -y python3 python3-pip
    else
        echo "未检测到支持的包管理器，请手动安装 python3 和 pip3"
        echo "命令示例: sudo apt install python3 python3-pip"
        exit 1
    fi
fi

# 安装/更新依赖
echo "检查依赖..."
PIP_CMD=""
if command -v pip3 &> /dev/null; then
    PIP_CMD="pip3"
elif command -v pip &> /dev/null; then
    PIP_CMD="pip"
else
    echo "pip3/pip 未找到，请先安装: sudo apt install python3-pip"
    exit 1
fi

$PIP_CMD install -r requirements.txt
if [ $? -ne 0 ]; then
    echo "依赖安装失败，请检查网络后重试"
    exit 1
fi

echo ""
echo "启动服务..."
echo ""

python3 src/app.py &
APP_PID=$!
sleep 2

# 检查服务是否启动成功
if ! kill -0 $APP_PID 2>/dev/null; then
    echo "服务启动失败，请检查日志"
    exit 1
fi

# 打开浏览器
if command -v xdg-open &> /dev/null; then
    xdg-open http://127.0.0.1:5000 2>/dev/null &
elif command -v open &> /dev/null; then
    open http://127.0.0.1:5000 2>/dev/null &
elif command -v gnome-open &> /dev/null; then
    gnome-open http://127.0.0.1:5000 2>/dev/null &
fi

echo "================================================"
echo "  服务已启动: http://127.0.0.1:5000"
echo "  按 Ctrl+C 停止服务"
echo "================================================"

# 等待Flask进程结束
wait $APP_PID