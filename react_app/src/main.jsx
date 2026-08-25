import React from 'react';
import ReactDOM from 'react-dom';
import App from './App.jsx';
import './styles.css';

// 注意：不要用 React.StrictMode —— OnlyOffice React SDK 的卸载清理逻辑
// 与 StrictMode 的双调用（mount→unmount→remount）不兼容，
// 会触发 `Failed to execute 'removeChild' on 'Node'` 导致切换模板时白屏。
ReactDOM.render(
  <App />,
  document.getElementById('root')
);
