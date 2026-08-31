# 西点函证系统——制函功能优化项目

本目录存放**制函功能优化**的资料与验证代码，按「数据 / 需求 / 代码」三层管理。

## 目录结构

```
04_制函优化/
├── data/                        # 📁 数据
│   ├── 截图/                    #   系统截图（制函管理、模板编辑器等 9 张）
│   ├── 模板/                    #   函证 Word 模板（可放往来函证模板等）
│   ├── 样例/                    #   示例 Excel / 转换样例
│   └── 产物/                    #   渲染/转换测试生成的 docx/xlsx
├── docs/                        # 📄 需求与方案文档
│   ├── 西点函证系统_制函功能现状说明.md      # 制函现状梳理（v1.1）
│   ├── 西点函证系统_制函功能优化方案_PRD草案.md # 优化 PRD（v0.3）
│   ├── M2_技术原型验证方案.md              # M2 技术验证方案
│   └── 转换测试_template_from_docx2markdown.md
├── code/                        # 💻 代码
│   ├── render_verify/           #   制函渲染验证服务（Flask Web）
│   └── conversion_test/         #   docx→md 转换测试脚本
├── 启动制函验证服务.bat          # 🚀 一键启动渲染验证服务
├── 停止制函验证服务.bat          # 🛑 一键停止渲染验证服务
└── README.md                    # 本说明
```

## 快速使用

### 启动制函渲染验证服务（可视化测试「Excel 直填 → 动态生成表格」）

双击 **`启动制函验证服务.bat`**，浏览器打开 **`http://127.0.0.1:5002`**。

界面操作：
1. 上传 Word 模板（.docx，可放 `data/模板/`）
2. 上传 Excel（.xlsx，多 Sheet，可放 `data/样例/`）
3. 添加绑定：选 Sheet + 选插入位置
4. 点「渲染」→ 下载结果 docx → 用 Word 打开查看

### 停止服务

双击 **`停止制函验证服务.bat`**（或直接关闭启动服务的黑色窗口）。

## 关键技术结论（摘要）

- **方案方向**：双模式制函——现有 `Output` 函数模式保留，新增「Excel 直填」模式（模板放占位锚点，按 Excel Sheet 动态生成表格）。
- **制函渲染**：后端已用 **Poi-tl** 动态建表（`WordManage.execTable`），验证服务用 Python（python-docx）验证了等价能力：列数=Sheet列数、表头/数据正确、空 Sheet 跳过、总宽锁定文本宽。
- **模板引擎**：统一 OnlyOffice（编辑 + 绑定 + 预览），取消 Markdown 方案（docx→md 对宽表/复杂表格失真）。
- **OnlyOffice**：远程部署（`113.201.2.61:14002/onlyoffice/`），文档隔离靠独立 document key。

## 环境

- **Python**：项目 venv `E:\13_dingdian\03_demo\04_Letter_Gen\venv\Scripts\python.exe`（Python 3.14，已装 flask/docx/openpyxl）
- **node**：`C:\nvm4w\nodejs`（node v24.18.0，已加入用户 PATH，重启终端生效）
