# -*- coding: utf-8 -*-
import subprocess

def run(cmd):
    r = subprocess.run(cmd, capture_output=True)
    return r.stdout, r.stderr

# 1. 检查 /tmp/local.json 内容（可能是坏的）
out, err = run(['docker','exec','onlyoffice','cat','/tmp/local.json'])
print('=== /tmp/local.json 前60字节 ===')
print(out[:60])
print('字节数:', len(out))

# 2. 检查 docservice 是否在跑
out, err = run(['docker','exec','onlyoffice','ps','-ef'])
ps = out.decode('utf-8', errors='replace')
print('=== docservice 进程 ===')
for line in ps.splitlines():
    if 'docservice' in line:
        print(line[:120])

# 3. supervisor 状态
out, err = run(['docker','exec','onlyoffice','supervisorctl','status'])
print('=== supervisor status ===')
print(out.decode('utf-8', errors='replace'))

# 4. 检查 example 的 local.json
out, err = run(['docker','exec','onlyoffice','cat','/etc/onlyoffice/documentserver-example/local.json'])
print('=== example local.json 前60字节 ===')
print(out[:60])
