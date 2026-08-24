@echo off
chcp 65001 >nul
echo ============================================
echo  制函渲染验证服务 - 停止脚本
echo ============================================
echo.

REM ---- 查找并结束占用 5002 端口的进程 ----
echo [停止] 查找端口 5002 上的服务进程...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":5002" ^| findstr "LISTENING"') do (
    echo [停止] 结束进程 PID: %%a
    taskkill /f /pid %%a >nul 2>&1
)
echo.
echo [完成] 服务已停止（若提示找不到进程，说明服务本来就没在运行）
echo.
pause
