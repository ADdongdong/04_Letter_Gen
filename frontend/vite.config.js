import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// 前端独立运行在 5173，/api 与 OnlyOffice 静态资源经代理转发到 Flask(5002)
export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
    host: '0.0.0.0',
    proxy: {
      // Flask 后端接口 + 静态 docx 托管（/static-docx）
      '/api': {
        target: 'http://127.0.0.1:5002',
        changeOrigin: true,
      },
      // OnlyOffice 模板/渲染结果 docx 静态托管（由 Flask 在 5002 提供，本地同源）
      '/static-docx': {
        target: 'http://127.0.0.1:5002',
        changeOrigin: true,
        rewrite: (p) => p.replace(/^\/static-docx/, '/static-docx'),
      },
      // OnlyOffice Document Server（本地 Docker，端口 8080；关键！让 iframe 同源，解决跨域访问 contentWindow 的问题）
      '/onlyoffice': {
        target: 'http://127.0.0.1:8080',
        changeOrigin: true,
        ws: true,  // OnlyOffice 用 WebSocket 做协作
        rewrite: (p) => p.replace(/^\/onlyoffice/, ''),
      },
      // OO 缓存文件（Editor.bin 等），编辑器加载后从此获取文档二进制数据
      '/cache': {
        target: 'http://127.0.0.1:8080',
        changeOrigin: true,
        ws: true,
      },
    },
  },
});
