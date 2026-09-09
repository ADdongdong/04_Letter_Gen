# -*- coding: utf-8 -*-
"""
一键停止制函服务（Java 后端版）

用法：
  python stop_server_java.py

行为：
  结束监听 5002（Java 后端 lettergen-java.jar）/ 5173（vite）/ 8899（静态托管）的进程。
  OnlyOffice Docker 容器保留运行（测试频繁无需反复启停）；
  如需一并停止，运行： docker stop onlyoffice

注意：Python 版后端（start_server.py）与本脚本操作同一组端口——
  停 Java 后改跑 Python 后端时，直接运行 start_server.py 即可（自带端口清理）。
"""
import os
import subprocess
import sys

# 控制台切 UTF-8
os.system('chcp 65001 >nul')
try:
    sys.stdout.reconfigure(encoding='utf-8')  # type: ignore[attr-defined]
except Exception:
    pass

SERVER_PORTS = (5002, 5173, 8899)


def kill_port(port):
    """结束监听指定端口的进程，返回结束的进程数"""
    out = subprocess.run(['netstat', '-ano'], capture_output=True, text=True).stdout
    pids = set()
    for line in out.splitlines():
        parts = line.split()
        if len(parts) >= 5 and parts[1].endswith(':%d' % port) and parts[3] == 'LISTENING':
            pids.add(parts[4])
    for pid in pids:
        subprocess.run(['taskkill', '/f', '/pid', pid], capture_output=True)
        print('  [停止] 端口 %d：已结束进程 PID=%s' % (port, pid))
    return len(pids)


def main():
    print('=' * 54)
    print('  制函服务一键停止（Java 后端）')
    print('=' * 54)
    total = 0
    for p in SERVER_PORTS:
        total += kill_port(p)
    if total == 0:
        print('  未发现运行中的服务进程')
    print()
    print('  服务已停止。OnlyOffice 容器保留运行（停止命令：docker stop onlyoffice）')


if __name__ == '__main__':
    main()
