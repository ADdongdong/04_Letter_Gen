# -*- coding: utf-8 -*-
import subprocess

def run(cmd):
    r = subprocess.run(cmd, capture_output=True)
    return r.stdout, r.stderr

# 1. 检查 /etc/onlyoffice/documentserver/local.json 的十六进制前 60 字节
out, err = run(['docker','exec','onlyoffice','od','-c','/etc/onlyoffice/documentserver/local.json'])
print('=== /etc/.../local.json 前60字节 od ===')
print(out.decode('utf-8', errors='replace')[:500])

# 2. 找所有 local.json
out, err = run(['docker','exec','onlyoffice','bash','-c','find / -name local.json 2>/dev/null'])
print('=== 所有 local.json ===')
print(out.decode())

# 3. 检查是否有 BOM / 非法字节（用 Python 读字节）
out, err = run(['docker','exec','onlyoffice','cat','/etc/onlyoffice/documentserver/local.json'])
print('=== cat 字节数:', len(out), '前20字节:', out[:20])
