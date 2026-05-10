@echo off
chcp 65001 >/dev/null
echo =========================================
echo   停止资信标自动填充系统
echo =========================================
echo.

:: 结束占用 8000 端口的进程
echo 正在查找并停止服务...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8000.*LISTENING" 2^>nul') do (
    echo 正在结束进程 PID: %%a
    taskkill /F /PID %%a >/dev/null 2>&1
)

echo 服务已停止。
pause
