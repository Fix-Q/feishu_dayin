import React, { useEffect, useRef, useState, useCallback } from 'react';
import { renderAsync } from 'docx-preview';
import { Space, Button, Tooltip, Slider } from 'antd';
import {
  ZoomInOutlined, ZoomOutOutlined, ColumnWidthOutlined, RedoOutlined,
} from '@ant-design/icons';

type Orientation = 'auto' | 'portrait' | 'landscape';

interface Props {
  blob: Blob | null;
  orientation?: Orientation;
  onError?: (msg: string) => void;
}

const MIN_SCALE = 0.5;
const MAX_SCALE = 2.5;
const STEP = 0.1;

function clamp(v: number): number {
  return Math.min(MAX_SCALE, Math.max(MIN_SCALE, Math.round(v * 100) / 100));
}

// 用 docx-preview 把填充后的 docx 渲染为纸张视图，供用户校对数据。
// 缩放通过 CSS transform scale 实现，只影响预览显示，不影响下载/打印。
export default function DocxPreview({ blob, orientation = 'auto', onError }: Props) {
  const scrollRef = useRef<HTMLDivElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const [scale, setScale] = useState(1);
  const [hasContent, setHasContent] = useState(false);
  // 用 ref 保存最新方向，供 fitWidth 读取（避免 fitWidth 依赖 orientation 导致渲染 effect 频繁重跑）
  const orientationRef = useRef<Orientation>(orientation);
  useEffect(() => { orientationRef.current = orientation; }, [orientation]);

  // 适应宽度：按容器可用宽度 / 页面实际宽度 计算缩放比
  const fitWidth = useCallback(() => {
    const scroll = scrollRef.current;
    const el = containerRef.current;
    if (!scroll || !el) return;
    const section = el.querySelector('section.docx') as HTMLElement | null;
    if (!section) return;
    // 横向时可见宽度是旋转后的宽度（即原始高度）
    const pageW = orientationRef.current === 'landscape' ? section.offsetHeight : section.offsetWidth;
    if (!pageW) return;
    const avail = scroll.clientWidth - 24; // 减去内边距
    setScale(clamp(avail / pageW));
  }, []);

  useEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    if (!blob) {
      el.innerHTML = '';
      setHasContent(false);
      return;
    }
    let cancelled = false;
    (async () => {
      try {
        el.innerHTML = '';
        await renderAsync(blob, el, undefined, {
          className: 'docx',
          inWrapper: true,
          breakPages: true,
          ignoreWidth: false,
          ignoreHeight: false,
          ignoreFonts: false,
          useBase64URL: true, // 与打印一致：图片内联，签名等图片可靠显示
          experimental: true,
        });
        if (cancelled) return;
        setHasContent(true);
        // 渲染完成后自动适应宽度一次
        requestAnimationFrame(() => fitWidth());
      } catch (e: any) {
        if (!cancelled && onError) onError(e?.message || '预览渲染失败');
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [blob, onError, fitWidth]);

  // 横向预览：把每个渲染出的 section 旋转 90°，视觉上与横版打印一致。
  // 用行内样式直接改 section，避免与外层 scale transform 冲突。
  useEffect(() => {
    const el = containerRef.current;
    if (!el || !hasContent) return;
    const sections = Array.from(el.querySelectorAll('section.docx')) as HTMLElement[];
    sections.forEach((sec) => {
      if (orientation === 'landscape') {
        const w = sec.offsetWidth;
        const h = sec.offsetHeight;
        sec.style.transform = 'rotate(90deg)';
        sec.style.transformOrigin = 'top left';
        // 旋转后向右平移一个「原高度」使其落回可见区，并把外框尺寸对调
        sec.style.left = `${h}px`;
        sec.style.position = 'relative';
        sec.style.marginBottom = `${w - h}px`; // 旋转后实际占位高度变为原宽度
      } else {
        sec.style.transform = '';
        sec.style.transformOrigin = '';
        sec.style.left = '';
        sec.style.position = '';
        sec.style.marginBottom = '';
      }
    });
    requestAnimationFrame(() => fitWidth());
  }, [orientation, hasContent, blob, fitWidth]);

  const zoomIn = () => setScale((s) => clamp(s + STEP));
  const zoomOut = () => setScale((s) => clamp(s - STEP));
  const reset = () => setScale(1);

  // Ctrl/Cmd + 滚轮缩放
  const onWheel = useCallback((e: React.WheelEvent) => {
    if (!(e.ctrlKey || e.metaKey)) return;
    e.preventDefault();
    setScale((s) => clamp(s + (e.deltaY < 0 ? STEP : -STEP)));
  }, []);

  return (
    <div style={{ border: '1px solid #e5e7eb', borderRadius: 8, overflow: 'hidden' }}>
      {/* 缩放工具条 */}
      <div
        style={{
          display: 'flex', alignItems: 'center', gap: 8,
          padding: '6px 10px', background: '#fafafa',
          borderBottom: '1px solid #eef0f2',
        }}
      >
        <Space size={4}>
          <Tooltip title="缩小 (Ctrl+滚轮)">
            <Button size="small" icon={<ZoomOutOutlined />} onClick={zoomOut} disabled={!hasContent || scale <= MIN_SCALE} />
          </Tooltip>
          <Tooltip title="放大 (Ctrl+滚轮)">
            <Button size="small" icon={<ZoomInOutlined />} onClick={zoomIn} disabled={!hasContent || scale >= MAX_SCALE} />
          </Tooltip>
          <Tooltip title="适应宽度">
            <Button size="small" icon={<ColumnWidthOutlined />} onClick={fitWidth} disabled={!hasContent} />
          </Tooltip>
          <Tooltip title="重置 100%">
            <Button size="small" icon={<RedoOutlined />} onClick={reset} disabled={!hasContent} />
          </Tooltip>
        </Space>
        <div style={{ width: 120, marginLeft: 4 }}>
          <Slider
            min={MIN_SCALE} max={MAX_SCALE} step={STEP}
            value={scale} onChange={(v) => setScale(clamp(v as number))}
            disabled={!hasContent} tooltip={{ open: false }}
          />
        </div>
        <span style={{ fontSize: 12, color: '#6b7280', minWidth: 42, textAlign: 'right' }}>
          {Math.round(scale * 100)}%
        </span>
      </div>

      {/* 预览滚动区 */}
      <div
        ref={scrollRef}
        onWheel={onWheel}
        style={{ background: '#f0f2f5', padding: 12, maxHeight: '55vh', overflow: 'auto' }}
      >
        {!blob && (
          <div style={{ textAlign: 'center', color: '#9ca3af', padding: '48px 0', fontSize: 13 }}>
            选择模板后在此预览打印效果
          </div>
        )}
        <div
          ref={containerRef}
          style={{
            transform: `scale(${scale})`,
            transformOrigin: 'top center',
            transition: 'transform 0.12s ease-out',
          }}
        />
      </div>
    </div>
  );
}
