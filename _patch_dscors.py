# -*- coding: utf-8 -*-
"""给 OO 容器内 ds-docservice.conf 的 /cache/files/ location 加上 CORS 头。"""
import subprocess, re

# 1. 读取容器内当前 ds-docservice.conf
r = subprocess.run(['docker', 'exec', 'onlyoffice', 'cat', '/etc/nginx/includes/ds-docservice.conf'], capture_output=True)
conf = r.stdout.decode('utf-8')
print('原文件长度:', len(conf), '字节')

# 2. 在 /cache/files/ location 块中加入 CORS 头
# 找到 "add_header Content-Disposition" 那行（紧跟 alias 后），在它后面插入 CORS 头
cors_headers = (
    "    add_header Access-Control-Allow-Origin * always;\n"
    "    add_header Access-Control-Allow-Methods \"GET, OPTIONS\" always;\n"
    "    add_header Access-Control-Allow-Headers \"Origin, Content-Type, Accept, Range\" always;\n"
)
new_conf = re.sub(
    r"(location \~\* \^\(\\\/cache\\\/files\.\*\)\(\\\/.\*\) \{[^}]*?add_header Content-Disposition[^;]*;)",
    r"\1\n" + cors_headers.rstrip('\n'),
    conf,
    count=1,
    flags=re.DOTALL,
)
if new_conf == conf:
    print('WARNING: 没找到目标 location，未修改')
else:
    print('已修改，新文件长度:', len(new_conf))

# 3. 写回容器（用 docker exec -i 写）
p = subprocess.run(
    ['docker', 'exec', '-i', 'onlyoffice', 'bash', '-c',
     'cat > /etc/nginx/includes/ds-docservice.conf'],
    input=new_conf.encode('utf-8'), capture_output=True
)
print('写回 rc:', p.returncode, p.stderr.decode()[:200])

# 4. 测试配置 + 重载
r2 = subprocess.run(['docker', 'exec', 'onlyoffice', 'nginx', '-t'], capture_output=True)
print('nginx -t:', r2.stdout.decode(), r2.stderr.decode()[:300])
r3 = subprocess.run(['docker', 'exec', 'onlyoffice', 'nginx', '-s', 'reload'], capture_output=True)
print('reload:', r3.stdout.decode(), r3.stderr.decode()[:200])

# 5. 验证：检查 ds-docservice.conf 中 CORS 头已写入
r4 = subprocess.run(['docker', 'exec', 'onlyoffice', 'bash', '-c',
                    'sed -n "/cache\\/files/,/^  }/p" /etc/nginx/includes/ds-docservice.conf'],
                   capture_output=True)
print('=== /cache/files/ location 块 ===')
print(r4.stdout.decode('utf-8', errors='replace'))
