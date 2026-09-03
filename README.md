# 西点函证系统——制函功能优化项目

前后端分离结构：唯一前端（React + Vite），多后端语言并置（现有 Python 后端，预留 Java 后端目录）。

## 目录结构

```
04_Letter_Gen/
├── backend/
│   ├── python/                  # 🐍 Python 后端（Flask 5002：/api/* 接口 + 渲染引擎）
│   │   ├── app.py               #   Flask 主服务
│   │   ├── render_engine.py     #   渲染引擎（python-docx + openpyxl，占位段替换/表格注入 + 表格样式）
│   │   ├── data/                #   函证编号登记数据（letters_registry.json）
│   │   ├── templates_store/     #   模板配置存储（word + excel + 样式）
│   │   ├── uploads/             #   上传临时目录（生成）
│   │   └── outputs/             #   渲染输出目录（生成）
│   ├── static-docs/             # 📄 静态托管源（8899，与后端语言无关）
│   └── java/                    # ☕ Java 版后端（预留空目录，待开发）
├── frontend/                    # ⚛️ 唯一前端（React + Vite 5173，Python/Java 后端共用）
│   └── src/                     #   App / ListPage / GeneratePage / RightPanel / OnlyOfficeEditor
├── onlyoffice-plugin/           # 🔌 click2insert 插件源码（部署到 OO 容器 sdkjs-plugins/）
├── docs/                        # 📄 需求文档 + design.md（视觉规范）
├── data/                        # 📁 样例 / 产物 / 截图
├── design-plans/                # 🎨 UI 审查报告
├── start_server.py              # 🚀 一键启动（清理端口 → OO 容器 → Flask 5002 + 静态 8899 + vite 5173 → 端口验证）
├── stop_server.py               # 🛑 一键停止
└── README.md                    # 本说明
```

## 快速使用

### 启动全套服务

```bash
venv\Scripts\python.exe start_server.py        # 前台模式（Ctrl+C 停止全部）
venv\Scripts\python.exe start_server.py --bg   # 后台模式（服务独立运行）
```

脚本自动完成：清理端口残留 → winnat 保留区检测（5002/5173 被 Windows 保留时自动提权修复）→ 启动 OnlyOffice Docker 容器并等待就绪 → 启动 Flask 5002 + 静态 8899 + vite 5173 → 轮询验证端口。

- 最新版界面：**http://127.0.0.1:5173/**（模板列表 / 模板配置 / Excel 制函三页）
- Flask API：http://127.0.0.1:5002/
- 静态托管：http://127.0.0.1:8899/

主流程：上传 Word 模板与 Excel（多 Sheet）→ 在 OnlyOffice 编辑器光标处插入 Sheet 占位标注 → 保存模板配置（可配置表格样式：字号/字体/对齐/行高/列宽）→ Excel 批量制函（按单位组合逐封渲染，打包 ZIP 下载）。

### 停止服务

```bash
venv\Scripts\python.exe stop_server.py
```

（OnlyOffice 容器保留运行；如需停止：`docker stop onlyoffice`）

## 环境

- **Python**：项目 venv `E:\13_dingdian\03_demo\04_Letter_Gen\venv\Scripts\python.exe`（已装 flask/docx/openpyxl）
- **node**：`C:\nvm4w\nodejs`（node v24.18.0）
- **OnlyOffice**：本地 Docker 容器 `onlyoffice`（8.2-ready 定制镜像，8080 映射；vite 代理 /onlyoffice → 8080）
