import React, { useState, useEffect } from 'react';
import ReactDOM from 'react-dom';
import App from './App.jsx';
import GeneratePage from './GeneratePage.jsx';
import ListPage from './ListPage.jsx';
import './styles.css';

// 极简 hash 路由（零依赖，三页常驻挂载仅 display 切换）：
//   #/templates（默认）→ 模板列表页（新增/修改/删除配置的入口）
//   #/config[?id=tpl_x] → 配置编辑页（OO 展示 + 插占位段 + 保存模板；带 id 为修改模式）
//   #/generate → 制函页（下拉选模板 + 上传批量 Excel + 制函）
//
// 重要：三页**常驻挂载**（仅 display 切换）——OO 编辑器实例永不因路由切换被卸载。
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
  const navLink = (href, label, active) => (
    <a
      href={href}
      style={{
        color: '#fff',
        textDecoration: 'none',
        fontWeight: active ? 700 : 400,
        padding: '4px 10px',
        borderRadius: '4px',
        background: active ? 'rgba(255,255,255,0.2)' : 'transparent',
      }}
    >
      {label}
    </a>
  );
  return (
    <>
      <div style={{ padding: '8px 16px', background: '#534AB7', display: 'flex', gap: '8px', alignItems: 'center' }}>
        <span style={{ color: 'rgba(255,255,255,0.85)', fontSize: '13px', marginRight: '8px' }}>函证制函工具</span>
        {navLink('#/templates', '模板列表', !isGenerate && !isConfig)}
        {navLink('#/generate', 'Excel 制函', isGenerate)}
        {isConfig && <span style={{ color: 'rgba(255,255,255,0.85)', fontSize: '13px', padding: '4px 10px' }}>正在编辑模板配置…</span>}
      </div>
      {/* 三页常驻挂载：切路由只切可见性，OO 实例与各页状态均保留 */}
      <div style={{ display: !isGenerate && !isConfig ? 'block' : 'none', height: '100%', overflowY: 'auto' }}>
        <ListPage />
      </div>
      <div style={{ display: isConfig ? 'block' : 'none', height: '100%' }}>
        <App />
      </div>
      <div style={{ display: isGenerate ? 'block' : 'none', height: '100%', overflowY: 'auto' }}>
        <GeneratePage />
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
