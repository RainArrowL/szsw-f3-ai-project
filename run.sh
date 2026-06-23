#!/bin/bash

# 自动切换到脚本自身所在目录，彻底避免相对路径找不到Python文件
cd "$(dirname "$0")" || exit 1

# 启动虚拟环境
source .venv/bin/activate

# 执行Python脚本，把 your_script.py 替换成你实际的Python文件名
python3 src/app.py