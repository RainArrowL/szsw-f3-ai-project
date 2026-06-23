@echo off
chcp 65001 >nul
cd /d "%~dp0"

title 智览金融 财数贯通

echo ========================================
echo       智览金融 财数贯通
echo ========================================
echo.

:: 检查 embedded Python 环境
if exist "python\python.exe" goto :run

echo [1/4] 首次运行，正在下载 Python 3.12 embedded (~7MB)...
powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.8/python-3.12.8-embed-amd64.zip' -OutFile 'python.zip'"
if %errorlevel% neq 0 (
    echo 下载失败，请检查网络连接
    pause
    exit /b 1
)

echo [2/4] 解压到 python\ 目录...
powershell -Command "Expand-Archive -Path 'python.zip' -DestinationPath 'python' -Force"
del python.zip

echo [3/4] 配置 pip...
powershell -Command "(Get-Content 'python\python312._pth') -replace '#import site', 'import site' | Set-Content 'python\python312._pth'"
powershell -Command "Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile 'python\get-pip.py'"
.\python\python.exe python\get-pip.py --no-warn-script-location

echo [4/4] 安装依赖...
.\python\python.exe -m pip install -r requirements.txt --no-warn-script-location

echo.
echo 环境配置完成！
echo.

:run
echo 启动服务...
start http://127.0.0.1:5000
.\python\python.exe src\app.py
pause