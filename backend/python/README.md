# 制函渲染验证服务（Excel 直填模式 Demo）

一个本地运行的 Web 验证服务，用于直观验证新方案的核心能力：
**在 Word 模板的"锚点位置"，按 Excel Sheet 动态生成完整表格**（列数=Sheet列数、表头正确、总宽锁定模板文本宽）。

## 功能

1. **上传 Word 模板**（.docx）：解析并展示模板结构（段落、已有表格）
2. **上传 Excel 数据**（.xlsx，多 Sheet）：解析每个 Sheet 的表头和数据
3. **选择锚点**：在模板中指定一个插入位置（支持：文档末尾 / 指定表格后 / 指定书签）
4. **选择 Sheet + 绑定**：把某个 Sheet 绑定到锚点位置
5. **渲染**：在模板锚点位置动态生成表格 → 下载结果 docx
6. **查看效果**：下载后可用 Word/OnlyOffice 打开查看表格是否正确生成

## 启动

```bash
# 在项目根（本目录）运行
E:\13_dingdian\03_demo\04_Letter_Gen\venv\Scripts\python.exe app.py
```

浏览器打开 `http://127.0.0.1:5002`

## 依赖

- flask
- python-docx
- openpyxl

（pytorch 环境已全部安装）

## 使用步骤

1. 打开页面，上传一个 Word 模板（如「往来函证模板（销售、采购）.docx」）
2. 上传一个 Excel（含至少一个有数据的 Sheet）
3. 在"模板锚点"区选一个插入位置
4. 在"Sheet 列表"勾选要渲染的 Sheet
5. 点击「渲染」→ 生成新 docx → 下载查看

## 设计说明

- 该服务仅用于**验证渲染能力**，不连接 OnlyOffice、不访问任何远程服务，纯本地操作，对生产系统零影响
- 核心逻辑 `render_engine.py`：封装"在 docx 锚点位置按 Sheet 动态建表"，后续可移植到 Java/poi-tl
