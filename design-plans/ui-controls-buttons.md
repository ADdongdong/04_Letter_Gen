# UI 审查报告：输入控件统一 + 主按钮蓝色化

> improve-ui 审查（2026-09-02）。Audited surface: `frontend/`（模板列表 / 模板配置 / Excel 制函三页）。
> Design sources: `docs/design.md` v1.0 → 本报告实施后升 v1.1。

## Design language
- Audited surface: frontend/ 三页（React + Vite，唯一全局样式表 styles.css）
- Design sources: docs/design.md（v1.0，token 定义于 styles.css :root）
- Documented decisions: 控件圆角 6px、0.15s 过渡、.btn.primary 绿色主操作、token 强制引用、禁止硬编码色值
- Governing owners and consumers: styles.css（唯一 owner）→ 三页 JSX
- Explicit exceptions: None documented

## Findings
| # | Problem | Evidence | Proposed change | Scope | Confidence |
| --- | --- | --- | --- | --- | --- |
| 1 | 输入控件无统一系统样式：text/number input 与 select 靠内联样式各自为政，原生外观直出（无统一圆角/边框/focus 态） | styles.css 仅 `.bind-item select`（4px 圆角）一处；GeneratePage.jsx:122 模板选择 select、RightPanel.jsx 模板名称 input 与样式区 selStyle 均内联且无边框圆角 | styles.css 新增元素选择器统一样式：白底 + 1px --border + 8px 圆角 + hover 加深 + focus 蓝边与 `0 0 0 3px --primary-bg` 光晕 + disabled 灰底 | 三页全部输入框/下拉框 | 高（规范缺失与运行时渲染均已证实） |
| 2 | 主操作按钮为绿色（.btn.primary = --success）且带 emoji 前缀，与全站主题蓝冲突 | styles.css:82-83（--success 底）；RightPanel.jsx:321（💾/⏳）、ListPage.jsx:45（＋）、GeneratePage.jsx:190（🚀/⏳） | .btn.primary 改主题蓝（--primary/hover/active 梯度）；按钮与保存卡片标题去 emoji（⏳/💾/✏️/＋/🚀）；空状态文案同步 | 三个主操作按钮 + 保存卡片标题 | 高（用户明确决策 + 实现证据一致） |

## Improve first
Finding 1（统一输入控件样式）——影响面覆盖三页所有输入/下拉控件，且元素选择器方案可覆盖未来新增控件；Finding 2 为用户明确决策，随同实施。

## 实施状态（2026-09-02 已执行）
- styles.css：新增「输入控件统一」段；.btn.primary 绿改主题蓝（hover/active 梯度）
- RightPanel/ListPage/GeneratePage：按钮与标题 emoji 清除
- docs/design.md 升 v1.1：§1 success 用途收窄、§3 输入控件 8px 圆角、§4 按钮蓝色规则 + 输入控件规范、§5 规则 3
