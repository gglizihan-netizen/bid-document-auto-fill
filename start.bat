@echo off
chcp 65001 >nul
echo =========================================
echo   资信标自动填充系统 - 启动脚本
echo =========================================
echo.

:: 检查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python 3.8+
    pause
    exit /b 1
)

:: 查找并结束占用 8000 端口的进程（解决旧代码运行问题）
echo [检查] 正在检查端口 8000...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8000.*LISTENING" 2^>nul') do (
    echo [清理] 发现占用端口的进程 PID: %%a，正在结束...
    taskkill /F /PID %%a >nul 2>&1
)
timeout /t 2 /nobreak >nul

:: 检查虚拟环境
if not exist venv (
    echo [1/4] 创建虚拟环境...
    python -m venv venv
)

:: 激活虚拟环境
echo [2/4] 激活虚拟环境...
call venv\Scripts\activate.bat

:: 安装依赖
echo [3/4] 安装依赖...
pip install -q -r requirements.txt
if errorlevel 1 (
    echo [错误] 依赖安装失败
    pause
    exit /b 1
)

:: 创建必要目录
if not exist uploads mkdir uploads
if not exist outputs mkdir outputs
if not exist static mkdir static

echo [4/4] 启动服务...
echo.
echo =========================================
echo 服务启动成功！
echo 请访问: http://localhost:8000
echo =========================================
echo.

:: 启动服务（使用 -B 禁用字节码缓存，确保运行最新代码）
python -B main.py

pause
