# -*- coding: utf-8 -*-
"""
启动全部验证服务：
1. render_verify Flask 服务（5002）——OnlyOffice 演示 + 渲染验证
2. 静态 HTTP 服务（8899）——托管 conversion_test 目录（OnlyOffice 需加载的 docx）
"""
import os, sys, subprocess, threading, time, io, urllib.request, socket

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

ROOT = r'e:\13_dingdian\03_demo\04_Letter_Gen'
RV_DIR = os.path.join(ROOT, 'code', 'render_verify')
CONV_DIR = os.path.join(ROOT, 'code', 'conversion_test')
PYTHON = sys.executable  # 当前解释器

RUNNING = []

def port_in_use(port):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    r = s.connect_ex(('127.0.0.1', port))
    s.close()
    return r == 0

def start_static(port):
    """启动 conversion_test 静态托管"""
    if port_in_use(port):
        print(f"[静态{port}] 端口已占用，跳过（可能已在运行）")
        return
    p = subprocess.Popen(
        [PYTHON, '-m', 'http.server', str(port), '--directory', CONV_DIR],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0,
    )
    RUNNING.append(p)
    print(f"[静态{port}] 已启动 (PID {p.pid}) -> http://127.0.0.1:{port}/")

def start_flask():
    """启动 render_verify Flask（日志写入 flask_oo.log，便于测试排查）"""
    if port_in_use(5002):
        print("[Flask5002] 端口已占用，跳过（可能已在运行）")
        return
    log_path = os.path.join(RV_DIR, 'flask_oo.log')
    logf = open(log_path, 'a', encoding='utf-8')
    logf.write("\n==== 启动于 %s ====\n" % time.strftime('%Y-%m-%d %H:%M:%S'))
    p = subprocess.Popen(
        [PYTHON, 'app.py'],
        cwd=RV_DIR,
        stdout=logf, stderr=subprocess.STDOUT,
        creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0,
    )
    RUNNING.append(p)
    print(f"[Flask5002] 已启动 (PID {p.pid}) -> http://127.0.0.1:5002/demo  (日志: code/render_verify/flask_oo.log)")

def verify():
    time.sleep(3)
    checks = [
        ('http://127.0.0.1:5002/', '渲染验证首页'),
        ('http://127.0.0.1:5002/demo', 'OnlyOffice 演示'),
        ('http://127.0.0.1:8899/demo_template.docx', '静态 docx'),
    ]
    for url, name in checks:
        try:
            r = urllib.request.urlopen(url, timeout=5)
            print(f"  [OK] {name}: HTTP {r.status}")
        except Exception as e:
            print(f"  [FAIL] {name}: {e}")

def main():
    print("=" * 50)
    print("  制函验证服务启动器")
    print("=" * 50)
    start_static(8899)
    start_flask()
    verify()
    print()
    print("  访问地址:")
    print("    OnlyOffice 演示: http://127.0.0.1:5002/demo")
    print("    渲染验证首页  : http://127.0.0.1:5002/")
    print()
    print("  服务已在后台运行，本脚本正常退出（不影响服务）")

if __name__ == '__main__':
    main()
