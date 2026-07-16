@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo   税金智 - 离线安装 (Python 3.9 + 依赖)
echo ========================================
echo.

:: 步骤1: 解压 Python 3.9 embeddable 包
if not exist "python39" (
    echo [步骤1] 解压 Python 3.9 embeddable ...
    if exist "offline\python\python-3.9.13-embed-amd64.zip" (
        powershell -command "Expand-Archive -Path 'offline\python\python-3.9.13-embed-amd64.zip' -DestinationPath 'python39' -Force"
        echo [完成] Python 3.9 已解压到 python39\
    ) else (
        echo [错误] 未找到 offline\python\python-3.9.13-embed-amd64.zip
        pause
        exit /b 1
    )
) else (
    echo [跳过] python39\ 已存在
)

:: 配置 Python embeddable: 启用 pip
echo.
echo [步骤2] 配置 Python embeddable ...
if exist "python39\python39._pth" (
    :: 修改 _pth 文件，取消 import site 注释
    powershell -command "(Get-Content python39\python39._pth) -replace '#import site', 'import site' | Set-Content python39\python39._pth"
    echo [完成] 已启用 site-packages
)

:: 安装 pip
if not exist "python39\Scripts\pip.exe" (
    echo [步骤3] 安装 pip ...
    if exist "offline\wheels\pip-*.whl" (
        powershell -command "python39\python.exe offline\wheels\get-pip.py --no-index --find-links=offline\wheels 2>nul"
    )
    :: 如果 get-pip.py 不存在，用 ensurepip
    python39\python.exe -m ensurepip --upgrade 2>nul
    echo [完成] pip 已安装
) else (
    echo [跳过] pip 已存在
)

:: 步骤4: 离线安装所有依赖
echo.
echo [步骤4] 离线安装 Python 依赖包 ...
echo [注意] 部分包需要编译（numpy/pandas/cryptography等），请确保已安装 Visual C++ Build Tools
echo        如无编译环境，可手动运行: python39\python.exe -m pip install numpy pandas cryptography
echo.
python39\python.exe -m pip install --no-index --find-links=offline\wheels flask openpyxl requests pdfplumber PyPDF2 xlrd akshare pandas werkzeug 2>&1
if %errorlevel% neq 0 (
    echo [警告] 离线安装失败，尝试在线安装缺失的包...
    python39\python.exe -m pip install flask openpyxl requests pdfplumber PyPDF2 xlrd akshare pandas werkzeug
)

echo.
echo ========================================
echo   安装完成！
echo ========================================
echo.
echo Python 路径: python39\python.exe
echo 运行方式: python39\python.exe src\app.py
echo 或双击 run_offline.bat
echo.
pause