@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo   税金智 - 安装依赖
echo ========================================
echo.

:: 检查 Python 是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到 Python，请先安装 Python 3.9+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [检测] Python 已安装:
python --version
echo.

:: 创建虚拟环境
if not exist ".venv" (
    echo [创建] 虚拟环境 .venv ...
    python -m venv .venv
    if %errorlevel% neq 0 (
        echo [错误] 虚拟环境创建失败
        pause
        exit /b 1
    )
    echo [完成] 虚拟环境创建成功
) else (
    echo [跳过] 虚拟环境 .venv 已存在
)

echo.

:: 激活虚拟环境并安装依赖
echo [安装] Python 依赖包 ...
call .venv\Scripts\activate.bat
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
if %errorlevel% neq 0 (
    echo [警告] 清华源安装失败，尝试默认源 ...
    pip install -r requirements.txt
)

echo.
echo ========================================
echo   安装完成！
echo ========================================
echo.
echo 运行方式: 双击 run.bat
echo.
pause