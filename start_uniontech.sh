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
pip3 install -r requirements.txt -q 2>/dev/null
if [ $? -ne 0 ]; then
    echo "依赖安装失败，尝试使用 pip 安装..."
    pip install -r requirements.txt -q 2>/dev/null
    if [ $? -ne 0 ]; then
        echo "请手动安装依赖: pip3 install -r requirements.txt"
        exit 1
    fi
fi

echo ""
echo "启动服务..."
python3 src/app.py &
APP_PID=$!
sleep 3

# 打开浏览器
if command -v xdg-open &> /dev/null; then
    xdg-open http://127.0.0.1:5000
elif command -v open &> /dev/null; then
    open http://127.0.0.1:5000
fi

echo ""
echo "服务已启动: http://127.0.0.1:5000"
echo "按 Ctrl+C 停止服务"
echo ""

# 等待Flask进程结束
wait $APP_PID