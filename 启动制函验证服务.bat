@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo ============================================
echo  制函验证服务 - 启动脚本
echo  OnlyOffice 演示: http://127.0.0.1:5002/demo
echo  渲染验证首页  : http://127.0.0.1:5002/
echo  静态资源服务  : http://192.168.100.92:8899/
echo ============================================
echo.

REM ---- 检查 Python 环境 ----
set PYTHON=E:\08_Anaconda3\Anaconda3\envs\pytorch\python.exe
if not exist "%PYTHON%" (
    echo [警告] 默认 Python 不存在，尝试使用 python 命令...
    set PYTHON=python
)

REM ---- 检查依赖 ----
echo [检查] 验证依赖库...
%PYTHON% -c "import flask, docx, openpyxl" 2>nul
if errorlevel 1 (
    echo [提示] 缺少依赖，尝试安装...
    %PYTHON% -m pip install flask python-docx openpyxl -i https://pypi.tuna.tsinghua.edu.cn/simple
)

REM ---- 启动服务 ----
echo.
echo [启动] 正在启动 OnlyOffice 演示 + 渲染验证服务...
echo [提示] 关闭本窗口即停止所有服务
echo.
"%PYTHON%" "%~dp0code\render_verify\start_all.py"

REM ---- 服务停止后 ----
echo.
echo [服务已停止]
pause
