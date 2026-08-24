# -*- coding: utf-8 -*-
"""从镜像提取原始 local.json，写回容器（保证 UTF-8 无 BOM）。"""
import subprocess, json

# 1. 从镜像提取原始 local.json 内容（通过 docker run）
r = subprocess.run(
    ['docker', 'run', '--rm', '--entrypoint', 'cat',
     'onlyoffice/documentserver:latest', '/etc/onlyoffice/documentserver/local.json'],
    capture_output=True
)
raw = r.stdout
print('镜像原始 local.json 字节数:', len(raw))

# 2. 校验是合法 JSON
try:
    cfg = json.loads(raw.decode('utf-8'))
    print('镜像 local.json JSON 有效')
except Exception as e:
    print('镜像 JSON 无效:', e)
    raise SystemExit(1)

# 3. 写回容器（docker exec -i ... cat >）
r2 = subprocess.run(
    ['docker', 'exec', '-i', 'onlyoffice', 'bash', '-c',
     'cat > /etc/onlyoffice/documentserver/local.json'],
    input=raw, capture_output=True
)
print('写回 rc:', r2.returncode, r2.stderr.decode()[:200])

# 4. 验证容器内文件
r3 = subprocess.run(['docker', 'exec', 'onlyoffice', 'cat', '/etc/onlyoffice/documentserver/local.json'], capture_output=True)
try:
    json.loads(r3.stdout.decode('utf-8'))
    print('容器内 local.json JSON 有效, 字节数:', len(r3.stdout))
except Exception as e:
    print('容器内 JSON 无效:', e)

# 5. 重启 docservice
r4 = subprocess.run(['docker', 'exec', 'onlyoffice', 'supervisorctl', 'restart', 'ds:docservice'], capture_output=True)
print('重启 docservice rc:', r4.returncode, r4.stdout.decode(), r4.stderr.decode()[:200])
