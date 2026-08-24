@echo off
chcp 65001 >nul
echo 启动制函渲染验证服务...
start "制函渲染验证服务" cmd /k "E:\08_Anaconda3\Anaconda3\envs\pytorch\python.exe app.py"
timeout /t 2 >nul
echo 服务已启动，请打开浏览器访问: http://127.0.0.1:5002
echo 若提示端口占用，请先关闭旧实例。
pause
