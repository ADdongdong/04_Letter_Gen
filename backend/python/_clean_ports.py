# -*- coding: utf-8 -*-
"""清理占用 5002 / 8899 的进程（跨平台，使用 Python 避免 cmd 引号地狱）。"""
import subprocess, sys

def kill_port(port):
    try:
        out = subprocess.check_output(['netstat', '-ano'], stderr=subprocess.DEVNULL, text=True)
    except Exception as e:
        print(f"[清理] 无法读取端口列表: {e}")
        return
    for line in out.splitlines():
        if f':{port}' in line and 'LISTENING' in line:
            parts = line.split()
            # 格式: Proto  Local Address  Foreign Address  State  PID
            pid = parts[-1]
            try:
                subprocess.run(['taskkill', '/f', '/pid', pid],
                               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
                print(f"[清理] 已结束端口 {port} 的进程 PID={pid}")
            except Exception as e:
                print(f"[清理] 结束 PID={pid} 失败: {e}")

if __name__ == '__main__':
    for p in (5002, 8899, 5173):
        kill_port(p)
    print("[清理] 完成")
