import React, { useEffect, useRef, useState, useCallback } from 'react';
import { Space, Button, Tooltip, Slider } from 'antd';
import {
  ZoomInOutlined, ZoomOutOutlined, ColumnWidthOutlined, RedoOutlined,
} from '@ant-design/icons';
import { renderXlsxToHtml } from '../services/xlsxRender';

interface Props {
  blob: Blob | null;
  onError?: (msg: string) => void;
}

const MIN_SCALE = 0.5;
const MAX_SCALE = 2.5;
const STEP = 0.1;

function clamp(v: number): number {
  return Math.min(MAX_SCALE, Math.max(MIN_SCALE, Math.round(v * 100) / 100));
}

// xlsx 保真预览：用 exceljs 读取单元格值+样式，渲染成 HTML 表格。
// 与 DocxPreview 一致的缩放工具条（放大/缩小/适应宽度/重置/Ctrl滚轮）。
export default function XlsxPreview({ blob, onError }: Props) {
  const scrollRef = useRef<HTMLDivElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const [scale, setScale] = useState(1);
  const [hasContent, setHasContent] = useState(false);

  const fitWidth = useCallback(() => {
    const scroll = scrollRef.current;
    const el = containerRef.current;
    if (!scroll || !el) return;
    const table = el.querySelector('table') as HTMLElement | null;
    if (!table) return;
    const pageW = table.offsetWidth;
    if (!pageW) return;
    const avail = scroll.clientWidth - 24;
    setScale(clamp(avail / pageW));
  }, []);

  useEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    if (!blob) { el.innerHTML = ''; setHasContent(false); return; }
    let cancelled = false;
    (async () => {
      try {
        const html = await renderXlsxToHtml(blob);
        if (cancelled) return;
        el.innerHTML = html;
        setHasContent(true);
        requestAnimationFrame(() => fitWidth());
      } catch (e: any) {
        if (!cancelled && onError) onError(e?.message || 'Excel 预览渲染失败');
      }
    })();
    return () => { cancelled = true; };
  }, [blob, onError, fitWidth]);

  const zoomIn = () => setScale((s) => clamp(s + STEP));
  const zoomOut = () => setScale((s) => clamp(s - STEP));
  const reset = () => setScale(1);

  const onWheel = useCallback((e: React.WheelEvent) => {
    if (!(e.ctrlKey || e.metaKey)) return;
    e.preventDefault();
    setScale((s) => clamp(s + (e.deltaY < 0 ? STEP : -STEP)));
  }, []);

  return (
    <div style={{ border: '1px solid #e5e7eb', borderRadius: 8, overflow: 'hidden' }}>
      <div style={{
        display: 'flex', alignItems: 'center', gap: 8,
        padding: '6px 10px', background: '#fafafa', borderBottom: '1px solid #eef0f2',
      }}>
        <Space size={4}>
          <Tooltip title="缩小 (Ctrl+滚轮)"><Button size="small" icon={<ZoomOutOutlined />} onClick={zoomOut} disabled={!hasContent || scale <= MIN_SCALE} /></Tooltip>
          <Tooltip title="放大 (Ctrl+滚轮)"><Button size="small" icon={<ZoomInOutlined />} onClick={zoomIn} disabled={!hasContent || scale >= MAX_SCALE} /></Tooltip>
          <Tooltip title="适应宽度"><Button size="small" icon={<ColumnWidthOutlined />} onClick={fitWidth} disabled={!hasContent} /></Tooltip>
          <Tooltip title="重置 100%"><Button size="small" icon={<RedoOutlined />} onClick={reset} disabled={!hasContent} /></Tooltip>
        </Space>
        <div style={{ width: 120, marginLeft: 4 }}>
          <Slider min={MIN_SCALE} max={MAX_SCALE} step={STEP} value={scale} onChange={(v) => setScale(clamp(v as number))} disabled={!hasContent} tooltip={{ open: false }} />
        </div>
        <span style={{ fontSize: 12, color: '#6b7280', minWidth: 42, textAlign: 'right' }}>{Math.round(scale * 100)}%</span>
      </div>

      <div ref={scrollRef} onWheel={onWheel} style={{ background: '#f0f2f5', padding: 12, maxHeight: '55vh', overflow: 'auto' }}>
        {!blob && (
          <div style={{ textAlign: 'center', color: '#9ca3af', padding: '48px 0', fontSize: 13 }}>
            选择 Excel 模板后在此预览打印效果
          </div>
        )}
        <div ref={containerRef} style={{
          transform: `scale(${scale})`,
          transformOrigin: 'top center',
          transition: 'transform 0.12s ease-out',
          display: 'inline-block',
        }} />
      </div>
    </div>
  );
}
