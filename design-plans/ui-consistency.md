# UI 一致性审查与统一计划（improve-ui）

- 审查表面：函证制函工具前端（react_app/src，三页路由：#/templates 模板列表 / #/config 配置编辑 / #/generate 制函）
- 设计来源：react_app/src/styles.css（组件类体系）+ 各页面内联样式（实际生效）
- 文档化决策：无 DESIGN.md；本文档首次记录
- 治理归属：styles.css 为唯一样式表；页面内联样式为漂移源
- 显式例外：无

## Design language
- 主题紫 #534AB7（nav/toolbar/badge/file-box label/sheet-item.active）
- 主操作绿 #2e7d32（.btn.primary）/ hover #256628
- 危险红 #d9534f（.btn.danger）/ hover #c9302c
- 警示色：统一为 #c62828（本计划前为双色并存：#c62828 内联 + #ff5722 .status-alert）
- 中性色：body #f5f7fa / 面板 #fff / 次级背景 #f7f7fb、#f0f0f5 / 边框 #e0e0e0、#ccc / 文字 #333、#888
- 字体 Microsoft YaHei；圆角体系 4-8px（badge 10px）
- 组件类：.file-box、.info、.hint、.actions、.btn(.primary/.danger)、.badge、.sheet-item、.status(-alert)

## Findings

| # | 问题 | 证据 | 修改 | 范围 | 置信 |
|---|------|------|------|------|------|
| 1 | 列表页【删除】按钮紫底红字，对比度差；已有 .btn.danger（红底白字）变体未使用 | styles.css L57-58 定义 .btn.danger；ListPage.jsx 删除按钮 `className="btn"` + inline `color:#c62828` | 删除按钮改 `className="btn danger"`，移除内联 color | ListPage | 高 |
| 2 | 配置编辑页同屏双紫条：顶部 Router 导航（#534AB7）与页内 .toolbar（#534AB7，仅承载状态栏）紧邻堆叠，视觉重复 | main.jsx nav 内联紫底；styles.css L11 `.toolbar{background:#534AB7}`；App.jsx 渲染顺序 nav → 顶部返回条 → .toolbar | .toolbar 改浅色状态条（白底 #333 字 + 下边框），与"返回+标题"白条形成层级，消除双紫条 | App.jsx/styles.css | 中 |
| 3 | 警示红双色并存：GeneratePage 匹配表格/提示内联 #c62828 vs styles.css .status-alert #ff5722，同产品两种警示红 | GeneratePage.jsx 匹配表格未匹配行/提示；styles.css L36 | 统一为 #c62828：.status-alert 改色，页面内联红色收敛为 .warn-red 公共类 | GeneratePage/styles.css | 中 |

## 实施说明（执行者须知）

1. styles.css 扩展公共类：`.nav`/`.nav-brand`/`.nav-link(.active)`（替代 main.jsx 导航内联）、`.page`（max-width 760 容器，替代 ListPage/GeneratePage 重复内联）、`.card`（白卡，替代重复 card 内联对象）、`.status-line`（页面状态条）、`.warn-red`（统一警示红）；`.status-alert` #ff5722→#c62828；`.toolbar` 紫底→白底 #333 字+下边框。
2. main.jsx / ListPage.jsx / GeneratePage.jsx 内联样式替换为公共类；组件结构、文案、功能零改动。
3. 配置编辑页（App.jsx）顶部"返回+标题"白条保留内联（与浅色 toolbar 形成层级），`.right` 面板类体系不动。
4. 验证：read_lints 全绿；浏览器硬刷新三页目测（无双紫条/删除按钮红底白字/警示红一致/功能不变）。
