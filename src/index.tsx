import React from 'react';
import { createRoot } from 'react-dom/client';
import App from './App';

// 清除旧版（整表导出/打印）遗留的 localStorage 预设，避免混淆
try {
  localStorage.removeItem('export_print_prefs_v1');
  localStorage.removeItem('export_print_templates_v1');
} catch (e) {
  // 忽略隐私模式等无法访问 localStorage 的情况
}

const container = document.getElementById('root');
if (container) {
  createRoot(container).render(<App />);
}
