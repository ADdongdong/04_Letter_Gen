"""临时脚本：检测各服务端口状态。用完删除。"""
import socket

PORTS = {
    5002: 'Flask 后端',
    5173: 'Vite 前端',
    8899: '静态文件服务',
    8080: 'OnlyOffice docker',
}
for port, name in PORTS.items():
    s = socket.socket()
    s.settimeout(2)
    try:
        s.connect(('127.0.0.1', port))
        print(f'{port:>5}  UP    {name}')
    except Exception:
        print(f'{port:>5}  DOWN  {name}')
    finally:
        s.close()
