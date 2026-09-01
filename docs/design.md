# 制函工具设计规范（design.md）

> 版本 v1.0（2026-09-01）。所有前端界面开发必须遵循本规范；新增界面元素必须引用 `var(--token)`，禁止硬编码色值。
> 样式表：`react_app/src/styles.css`（唯一全局样式表，token 定义于 `:root`）。

## 1. 色板

### 主色（品牌蓝）
| Token | 值 | 用途 |
|-------|-----|------|
| `--primary` | `#6083F0` | 导航条、主按钮（次级）、badge、选中态、上传引导、表头强调、链接 |
| `--primary-hover` | `#4A6FE0` | 主色按钮/链接 hover |
| `--primary-active` | `#3D5ED0` | 按压态 |
| `--primary-bg` | `#EEF3FE` | 浅蓝底：匹配表格表头、hover 高亮、pos-item |
| `--primary-bg-light` | `#F5F8FE` | 更浅蓝底：info 信息条、保存区底、file-box hover |
| `--primary-border` | `#C7D6FB` | 蓝调边框：file-box 虚线、保存区边框 |

### 语义色
| Token | 值 | 用途 |
|-------|-----|------|
| `--success` / `--success-bg` | `#2e7d32` / `#e8f5e9` | 成功（.btn.primary 主操作、Sheet 已插入） |
| `--danger` / `--danger-hover` | `#d9534f` / `#c9302c` | 危险（.btn.danger 删除） |
| `--warn` | `#c62828` | 警示红（.status-alert、.warn-red、未匹配提醒）——**全站唯一警示红** |

### 中性色
| Token | 值 | 用途 |
|-------|-----|------|
| `--text` / `--text-2` / `--text-3` | `#333` / `#666` / `#888` | 正文 / 次级 / 提示 |
| `--border` | `#e6e9ef` | 卡片/控件边框 |
| `--bg` | `#f5f7fa` | 页面背景 |
| `--card-bg` | `#fff` | 卡片/面板背景 |
| `--shadow-card` | `0 1px 3px rgba(16,24,40,.06)` | 卡片投影 |

## 2. 字阶

| 层级 | 规格 |
|------|------|
| 页面标题（h2） | 15px / 600，右面板 h2 带 2px 主色下边框 |
| 卡片小标题（h3） | 13px / 600 主色 |
| 正文/控件 | 13px / 400 |
| 辅助提示（.hint） | 12px，--text-3 |
| badge/状态 | 11-12px |

字体：`"Microsoft YaHei", Arial, sans-serif`。

## 3. 形状与阴影

| Token/约定 | 值 |
|-----------|-----|
| 卡片圆角 | 10px，边框 1px `--border`，投影 `--shadow-card` |
| 控件圆角（按钮/info/file-box） | 6px |
| 列表项圆角 | 5px |
| badge 胶囊 | 10px |
| 过渡 | `transition: background 0.15s`（按钮/导航/列表项/上传框统一） |

间距档：**8 / 12 / 16px**（卡片内区块 12px、卡片间距 12px、页面容器 padding 16px）。

## 4. 组件约定

### 按钮（.btn）
- 默认（次级操作）：`--primary` 底白字，hover `--primary-hover`
- `.btn.primary`（主操作，绿）：`--success` 底，hover #256628——**每屏仅一个主操作**
- `.btn.danger`（危险，红）：`--danger` 底白字——删除类操作必须用此变体，禁止默认紫/蓝底+红字
- 统一 6px 圆角、0.15s 过渡；禁用态 opacity 0.5

### 卡片（.card）
白底、1px `--border` 边框、10px 圆角、`--shadow-card` 投影、16px 内边距。页面容器用 `.page`（max-width 760）。

### 上传框（.file-box）
2px 蓝调虚线边框、hover 时底色 `--primary-bg-light` + 边框实色化；内含隐藏 input + label。

### 状态条
- 配置编辑页：`.toolbar`（白底状态条，禁止紫底——避免与顶部导航形成双横条）
- 列表/制函页：`.status-line`（页面底部灰条）

### 导航（.nav）
纯色 `--primary`（**不使用渐变**）；`.nav-link` 白字，active 态 20% 白底高亮。

### 警示（.warn-red / .status-alert）
一律 `--warn`(#c62828)；匹配表格未匹配行底色 `#fdecea` + 行内文字 `.warn-red`。

## 5. 使用规则

1. 新增界面元素**必须**引用 `var(--token)`，禁止硬编码色值；
2. 换肤/调色只修改 styles.css `:root` 变量；
3. 每屏主操作（绿色 .btn.primary）唯一；删除必须 `.btn.danger`；
4. 新页面容器用 `.page`，内容块用 `.card`，状态反馈用 `.status-line`，警示用 `.warn-red`；
5. OnlyOffice iframe 内部样式独立，不受本规范影响，也不得向其注入样式。
