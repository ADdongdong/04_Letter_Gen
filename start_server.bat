@echo off
chcp 65001 >nul
REM ============================================================
REM  制函渲染验证服务 - 一键启动（最新版，含 React 前端 5173）
REM  服务清单:
REM    5173  React 前端 (上传 Excel / 选位置 / 生成函证) —— 最新版界面
REM    5002  Flask 后端 (app.py: /api/* 接口 + /demo 旧界面)
REM    8899  静态托管 (conversion_test/*.docx，OnlyOffice 加载用)
REM  前置依赖: OnlyOffice Docker 8080 (vite 代理 /onlyoffice -> 8080)
REM ============================================================
echo.
echo ============================================
echo   制函渲染验证服务 - 启动中（最新版）
echo ============================================
echo   最新版界面: http://127.0.0.1:5173/
echo ============================================
echo.

REM --- 1. 清理可能残留的端口进程 ---
"E:\08_Anaconda3\Anaconda3\envs\pytorch\python.exe" "e:\13_dingdian\03_demo\04_Letter_Gen\code\render_verify\_clean_ports.py"

REM --- 2. 启动 Flask 5002 + 静态 8899 (Python 启动器，后台运行) ---
start "" /B "E:\08_Anaconda3\Anaconda3\envs\pytorch\python.exe" "e:\13_dingdian\03_demo\04_Letter_Gen\code\render_verify\start_all.py"

REM --- 3. 启动 React 前端 5173 (vite dev) ---
echo [启动] React 前端 (vite dev) -> http://127.0.0.1:5173/
start "" /B cmd /c "C:\nvm4w\nodejs\npm.cmd run dev --prefix e:\13_dingdian\03_demo\04_Letter_Gen\react_app"

echo.
echo 全部服务已在后台启动，可直接打开 http://127.0.0.1:5173/ 测试
echo （关闭本窗口不会停止服务；停止请运行 停止制函验证服务.bat）
echo.
pause
