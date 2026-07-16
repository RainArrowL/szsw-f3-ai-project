@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo   税金智 - 启动服务
echo ========================================
echo.

:: 检查虚拟环境
if not exist ".venv\Scripts\activate.bat" (
    echo [错误] 未找到虚拟环境，请先运行 install.bat
    pause
    exit /b 1
)

:: 激活虚拟环境
call .venv\Scripts\activate.bat

:: 检查依赖
python -c "import flask" >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 依赖未安装，请先运行 install.bat
    pause
    exit /b 1
)

echo [启动] Web 服务 http://localhost:5000
echo [提示] 按 Ctrl+C 停止服务
echo.

python src/app.py

pause