import React, { useState, useEffect } from 'react';
import ReactDOM from 'react-dom';
import App from './App.jsx';
import GeneratePage from './GeneratePage.jsx';
import HelpPage from './HelpPage.jsx';
import ListPage from './ListPage.jsx';
import './styles.css';

// 极简 hash 路由（零依赖，页面常驻挂载仅 display 切换）：
//   #/templates（默认）→ 模板列表页（新增/修改/删除配置的入口）
//   #/config[?id=tpl_x] → 配置编辑页（OO 展示 + 插占位段 + 保存模板；带 id 为修改模式）
//   #/generate → 制函页（下拉选模板 + 上传批量 Excel + 制函）
//   #/help → 使用说明（应对场景 + 从配置到制函全流程）
//
// 重要：OO 所在配置页**常驻挂载**（仅 display 切换）——OO 编辑器实例永不因路由切换被卸载。
// 若用条件渲染卸载配置页，React 移除容器 div 会与 OO SDK 异步 destroyEditor 产生竞态，
// SDK 内部 getElementById 返回 null → "Cannot read properties of null (reading 'getBoundingClientRect')"。
function Router() {
  const [route, setRoute] = useState(window.location.hash || '#/templates');
  useEffect(() => {
    const onHash = () => setRoute(window.location.hash || '#/templates');
    window.addEventListener('hashchange', onHash);
    return () => window.removeEventListener('hashchange', onHash);
  }, []);
  const isGenerate = route.startsWith('#/generate');
  const isConfig = route.startsWith('#/config');
  const isHelp = route.startsWith('#/help');
  return (
    <>
      <div className="nav">
        <span className="nav-brand">函证制函工具</span>
        <a className={'nav-link' + (!isGenerate && !isConfig && !isHelp ? ' active' : '')} href="#/templates">模板列表</a>
        <a className={'nav-link' + (isGenerate ? ' active' : '')} href="#/generate">Excel 制函</a>
        <a className={'nav-link' + (isHelp ? ' active' : '')} href="#/help">使用说明</a>
        {isConfig && <span className="nav-brand">正在编辑模板配置…</span>}
      </div>
      {/* 配置页常驻挂载：切路由只切可见性，OO 实例与各页状态均保留 */}
      <div style={{ display: !isGenerate && !isConfig && !isHelp ? 'block' : 'none', height: '100%', overflowY: 'auto' }}>
        <ListPage />
      </div>
      <div style={{ display: isConfig ? 'block' : 'none', height: '100%' }}>
        <App />
      </div>
      <div style={{ display: isGenerate ? 'block' : 'none', height: '100%', overflowY: 'auto' }}>
        <GeneratePage />
      </div>
      <div style={{ display: isHelp ? 'block' : 'none', height: '100%', overflowY: 'auto' }}>
        <HelpPage />
      </div>
    </>
  );
}

ReactDOM.render(
  <React.StrictMode>
    <Router />
  </React.StrictMode>,
  document.getElementById('root')
);
