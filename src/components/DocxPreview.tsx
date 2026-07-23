import React, { useEffect, useRef } from 'react';
import { renderAsync } from 'docx-preview';

interface Props {
  blob: Blob | null;
  onError?: (msg: string) => void;
}

// 用 docx-preview 把填充后的 docx 渲染为纸张视图，供用户校对数据。
export default function DocxPreview({ blob, onError }: Props) {
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    if (!blob) {
      el.innerHTML = '';
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
      } catch (e: any) {
        if (!cancelled && onError) onError(e?.message || '预览渲染失败');
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [blob, onError]);

  return (
    <div
      style={{
        background: '#f5f5f5',
        padding: 8,
        borderRadius: 4,
        maxHeight: '55vh',
        overflow: 'auto',
      }}
    >
      <div ref={containerRef} />
    </div>
  );
}
