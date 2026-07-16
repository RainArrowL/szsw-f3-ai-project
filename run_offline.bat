@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo   税金智 - 启动服务 (离线模式)
echo ========================================
echo.

:: 检查 Python 39
if not exist "python39\python.exe" (
    echo [错误] 未找到 python39\python.exe，请先运行 install_offline.bat
    pause
    exit /b 1
)

echo [启动] Web 服务 http://localhost:5000
echo [提示] 按 Ctrl+C 停止服务
echo.

python39\python.exe src\app.py

pause