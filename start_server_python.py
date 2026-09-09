# -*- coding: utf-8 -*-
"""
一键启动制函服务（替代 start_server.bat + backend/python/start_all.py）

用法：
  python start_server.py        # 前台模式：拉起全部服务后保持运行，Ctrl+C 停止全部
  python start_server.py --bg   # 后台模式：拉起全部服务后脚本立即退出（服务独立后台运行）
  建议用项目 venv 解释器运行： venv\\Scripts\\python.exe start_server.py

服务清单：
  5002  Flask 后端（/api/*）
  5173  React 前端（vite dev）
  8899  静态托管（backend/static-docs，OnlyOffice 加载 docx 用）
  8080  OnlyOffice（Docker 容器 onlyoffice，未运行则自动 docker start）

特性：
  - 启动前清理端口残留进程
  - winnat 端口保留区检测：5002/5173 被保留时自动请求管理员修复（UAC 弹窗），
    服务启动成功后再自动恢复 winnat
  - 启动后轮询验证端口与 OO healthcheck，打印状态汇总
"""
import ctypes
import os
import shutil
import socket
import subprocess
import sys
import time
import urllib.request

# 控制台切 UTF-8，避免中文乱码
os.system('chcp 65001 >nul')
try:
    sys.stdout.reconfigure(encoding='utf-8')  # pyright: ignore[reportAttributeAccessIssue]
except Exception:
    pass

BASE = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable
BACKEND = os.path.join(BASE, 'backend', 'python')
FRONTEND = os.path.join(BASE, 'frontend')
STATIC_DIR = os.path.join(BASE, 'backend', 'static-docs')
FLASK_LOG = os.path.join(BACKEND, 'flask_oo.log')
VITE_LOG = os.path.join(FRONTEND, 'vite_out.log')

CREATE_NO_WINDOW = subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0
SERVER_PORTS = (5002, 5173, 8899)


def is_admin():
    """当前进程是否具有管理员权限"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() == 1
    except Exception:
        return False


def port_in_use(port):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    r = s.connect_ex(('127.0.0.1', port))
    s.close()
    return r == 0


def kill_port(port):
    """结束监听指定端口的进程，返回结束的进程数"""
    out = subprocess.run(['netstat', '-ano'], capture_output=True, text=True).stdout
    pids = set()
    for line in out.splitlines():
        parts = line.split()
        # TCP 行：Proto Local Foreign State PID
        if len(parts) >= 5 and parts[1].endswith(':%d' % port) and parts[3] == 'LISTENING':
            pids.add(parts[4])
    for pid in pids:
        subprocess.run(['taskkill', '/f', '/pid', pid], capture_output=True)
        print('  [清理] 已结束端口 %d 的进程 PID=%s' % (port, pid))
    return len(pids)


def excluded_ranges():
    """解析 winnat TCP 端口保留区，返回 [(start, end), ...]"""
    out = subprocess.run(
        ['netsh', 'interface', 'ipv4', 'show', 'excludedportrange', 'protocol=tcp'],
        capture_output=True, text=True).stdout
    ranges = []
    for line in out.splitlines():
        parts = line.split()
        if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
            ranges.append((int(parts[0]), int(parts[1])))
    return ranges


def port_reserved(port, ranges):
    return any(a <= port <= b for a, b in ranges)


def elevate_netstop_winnat():
    """弹 UAC 以管理员权限执行 net stop winnat（用户批准返回 True）"""
    cmd = "Start-Process powershell -Verb RunAs -Wait -ArgumentList '-Command','net stop winnat'"
    r = subprocess.run(['powershell', '-NoProfile', '-Command', cmd], capture_output=True)
    return r.returncode == 0


def net_start_winnat():
    """恢复 winnat：管理员直接执行；否则再弹一次 UAC"""
    if is_admin():
        subprocess.run(['net', 'start', 'winnat'], capture_output=True)
        return True
    r = subprocess.run(['powershell', '-NoProfile', '-Command',
        "Start-Process powershell -Verb RunAs -Wait -ArgumentList '-Command','net start winnat'"],
        capture_output=True)
    return r.returncode == 0


def handle_winnat_conflicts():
    """检测 5002/5173 是否被 winnat 保留区占用；占用则自动修复。返回是否已停止 winnat（启动完成后需恢复）"""
    ranges = excluded_ranges()
    conflicts = [p for p in (5002, 5173) if port_reserved(p, ranges)]
    if not conflicts:
        return False
    print('[winnat] 端口 %s 被 Windows 保留区占用（Docker/WSL 重启后常见）' % conflicts)
    if is_admin():
        subprocess.run(['net', 'stop', 'winnat'], capture_output=True)
        print('[winnat] 已临时停止 winnat 释放端口，服务启动完成后将自动恢复')
        return True
    print('[winnat] 正在请求管理员权限释放端口（请在弹窗中点"是"）...')
    if elevate_netstop_winnat():
        print('[winnat] 端口已释放，服务启动完成后将自动恢复 winnat')
        return True
    print('[winnat] 未获批准。如服务启动失败，请以管理员手动执行：net stop winnat')
    return False


def docker_status():
    """返回 (daemon_ok, container_up)"""
    r = subprocess.run(['docker', 'ps', '-a', '--filter', 'name=onlyoffice',
                        '--format', '{{.Status}}'], capture_output=True, text=True)
    if r.returncode != 0:
        return False, False
    return True, ('Up' in r.stdout)


def start_docker():
    """确保 OnlyOffice 容器运行并就绪（healthcheck=true）。阻塞最多 120s。"""
    daemon_ok, up = docker_status()
    if not daemon_ok:
        print('  [OO] Docker 未运行——请先打开 Docker Desktop，稍后 OnlyOffice 会自动可用')
        return False
    if up:
        print('  [OO] OnlyOffice 容器已在运行')
    else:
        print('  [OO] 启动 OnlyOffice 容器...')
        subprocess.run(['docker', 'start', 'onlyoffice'], capture_output=True)
    # 轮询 healthcheck（OO 文档服务启动需 1-2 分钟）
    for i in range(24):
        try:
            resp = urllib.request.urlopen('http://127.0.0.1:8080/healthcheck', timeout=4)
            if resp.read() == b'true':
                print('  [OO] OnlyOffice 就绪（healthcheck=true，%ds）' % (i * 5 + 5))
                return True
        except Exception:
            pass
        if i * 5 + 5 < 120:
            print('  [OO] OnlyOffice 启动中... %ds' % (i * 5 + 5))
            time.sleep(5)
    print('  [OO] OnlyOffice 120s 内未就绪（不影响其他服务，稍后自动可用）')
    return False


def start_static():
    if port_in_use(8899):
        print('  [静态8899] 端口已占用，跳过')
        return
    subprocess.Popen(
        [PY, '-m', 'http.server', '8899', '--directory', STATIC_DIR],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=CREATE_NO_WINDOW)
    print('  [静态8899] 已启动')


def start_flask():
    if port_in_use(5002):
        print('  [Flask5002] 端口已占用，跳过')
        return
    logf = open(FLASK_LOG, 'a', encoding='utf-8')
    logf.write('\n==== 启动于 %s ====\n' % time.strftime('%Y-%m-%d %H:%M:%S'))
    subprocess.Popen(
        [PY, 'app.py'],
        cwd=BACKEND, stdout=logf, stderr=subprocess.STDOUT, creationflags=CREATE_NO_WINDOW)
    print('  [Flask5002] 已启动（日志: backend/python/flask_oo.log）')


def start_vite():
    if port_in_use(5173):
        print('  [vite5173] 端口已占用，跳过')
        return
    logf = open(VITE_LOG, 'a', encoding='utf-8')
    logf.write('\n==== 启动于 %s ====\n' % time.strftime('%Y-%m-%d %H:%M:%S'))
    # shutil.which 在 PATH 中定位 npm.cmd 绝对路径，避免 cmd /c 的引号嵌套解析问题
    npm = shutil.which('npm.cmd') or shutil.which('npm')
    if not npm:
        print('  [vite5173] 启动失败：PATH 中找不到 npm（请确认 node 已安装并加入 PATH）')
        return
    subprocess.Popen(
        ['cmd', '/c', npm, 'run', 'dev', '--prefix', FRONTEND],
        stdout=logf, stderr=subprocess.STDOUT, creationflags=CREATE_NO_WINDOW)
    print('  [vite5173] 已启动（日志: frontend/vite_out.log）')


def wait_port(port, timeout=30):
    """轮询等待端口 HTTP 可访问"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            urllib.request.urlopen('http://127.0.0.1:%d/' % port, timeout=2)
            return True
        except Exception:
            time.sleep(1)
    return False


def main():
    bg = '--bg' in sys.argv[1:]
    print('=' * 54)
    print('  制函服务一键启动')
    print('=' * 54)

    # 1. 清理端口残留
    print('[1/5] 清理端口残留...')
    for p in SERVER_PORTS:
        kill_port(p)

    # 2. winnat 保留区检测与修复
    print('[2/5] 检测 winnat 端口保留区...')
    winnat_stopped = handle_winnat_conflicts()

    # 3. 启动各服务
    print('[3/5] 启动 OnlyOffice 容器...')
    start_docker()
    print('[4/5] 启动 Flask / 静态托管 / vite...')
    start_static()
    start_flask()
    start_vite()

    # 5. 验证
    print('[5/5] 验证端口...')
    results = {}
    for p in SERVER_PORTS:
        results[p] = wait_port(p, timeout=30)
        print('  端口 %d: %s' % (p, 'OK' if results[p] else 'FAIL'))

    # 恢复 winnat（服务已绑定端口，重启 winnat 不影响）
    if winnat_stopped:
        print('[winnat] 恢复 winnat...')
        net_start_winnat()

    print()
    print('  访问入口: http://127.0.0.1:5173/')
    ok = all(results.values())
    print('  状态: %s' % ('全部就绪' if ok else '部分服务未就绪（见上方日志）'))
    print()

    if bg:
        print('后台模式：脚本退出，服务独立运行。停止服务请运行 stop_server.py')
        return
    print('前台模式：保持运行中，按 Ctrl+C 停止全部服务...')
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print('\n停止全部服务...')
        for p in SERVER_PORTS:
            kill_port(p)
        print('已停止')


if __name__ == '__main__':
    main()
