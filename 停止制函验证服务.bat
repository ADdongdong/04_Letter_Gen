@echo off
chcp 65001 >nul
REM ============================================================
REM  制函渲染验证服务 - 停止脚本
REM  结束 5002 (Flask) 与 8899 (静态托管) 端口进程
REM ============================================================
echo ============================================
echo   制函渲染验证服务 - 停止脚本
echo ============================================
echo.

echo [停止] 结束端口 5002 (Flask) 上的进程...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":5002" ^| findstr "LISTENING"') do (
    echo [停止] 结束 PID: %%a
    taskkill /f /pid %%a >nul 2>&1
)

echo [停止] 结束端口 8899 (静态托管) 上的进程...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8899" ^| findstr "LISTENING"') do (
    echo [停止] 结束 PID: %%a
    taskkill /f /pid %%a >nul 2>&1
)

echo.
echo [完成] 服务已停止（若提示找不到进程，说明本来就没在运行）
echo.
pause
