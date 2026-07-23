import { renderAsync } from 'docx-preview';

// 用隐藏的同源 iframe 渲染填充后的 docx 并调用打印。
// 目标：尽量贴近 Word 打开的效果——保留签名等图片、页面尺寸跟随文档、颜色不被浏览器淡化。
//
// 关键点：
// 1. useBase64URL:true —— 图片内联为 base64，避免 blob: URL 在打印上下文里失效导致签名图丢失。
// 2. hideWrapperOnPrint:true —— 打印时去掉 docx-preview 的灰底外框。
// 3. 注入 @page + print-color-adjust —— 页面尺寸跟随文档、强制按原色打印。
// 4. 等待所有图片真正 load 完再 print()，否则可能打印出空白图位。
export async function printDocxBlob(blob: Blob): Promise<void> {
  const iframe = document.createElement('iframe');
  iframe.style.position = 'fixed';
  iframe.style.right = '0';
  iframe.style.bottom = '0';
  iframe.style.width = '0';
  iframe.style.height = '0';
  iframe.style.border = '0';
  document.body.appendChild(iframe);

  const cleanup = () => {
    setTimeout(() => {
      if (iframe.parentNode) iframe.parentNode.removeChild(iframe);
    }, 1500);
  };

  try {
    const doc = iframe.contentDocument;
    const win = iframe.contentWindow;
    if (!doc || !win) throw new Error('无法创建打印容器');

    doc.open();
    doc.write('<!doctype html><html><head><meta charset="utf-8"><title>打印</title></head><body></body></html>');
    doc.close();

    const mount = doc.body;
    await renderAsync(blob, mount, undefined, {
      className: 'docx',
      inWrapper: true,
      hideWrapperOnPrint: true, // 打印时隐藏灰色外框
      breakPages: true,
      ignoreWidth: false,
      ignoreHeight: false,
      ignoreFonts: false, // 保留文档字体声明
      useBase64URL: true, // 图片内联，避免打印时签名图丢失
      experimental: true,
    });

    injectPrintStyles(doc);
    await waitForImages(doc, 4000);
    // 再给排版/字体一点稳定时间
    await new Promise((r) => setTimeout(r, 200));

    win.focus();
    win.print();
    cleanup();
  } catch (e: any) {
    cleanup();
    throw new Error('打印被环境拦截或渲染失败，请改用「下载 Word」后在本地打印。（' + (e?.message || e) + '）');
  }
}

// 注入打印专用样式：页面尺寸跟随文档、强制原色、去掉页边空白。
function injectPrintStyles(doc: Document) {
  // 读取首个渲染出的页面 section，取其真实宽高作为 @page size
  const section = doc.querySelector('section.docx') as HTMLElement | null;
  let pageRule = '';
  if (section) {
    const w = section.style.width; // docx-preview 写入的是带单位的值，如 "595.3pt"
    const h = section.style.minHeight || section.style.height;
    if (w && h) {
      pageRule = `@page { size: ${w} ${h}; margin: 0; }`;
    } else {
      pageRule = `@page { margin: 0; }`;
    }
  } else {
    pageRule = `@page { margin: 0; }`;
  }

  const style = doc.createElement('style');
  style.textContent = `
    ${pageRule}
    html, body { margin: 0 !important; padding: 0 !important; background: #fff !important; }
    /* 强制按文档原色打印，防止背景色/浅色被浏览器淡化或丢弃 */
    * {
      -webkit-print-color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    /* 页面 section 打印时去掉阴影与外边距，避免额外空白挤动版面 */
    @media print {
      section.docx { box-shadow: none !important; margin: 0 !important; }
      .docx-wrapper { background: #fff !important; padding: 0 !important; }
      /* 图片按其在文档中的尺寸打印，不被页面宽度压缩 */
      section.docx img { max-width: none; }
    }
  `;
  doc.head.appendChild(style);
}

// 等待文档内所有 <img> 加载完成（或超时），避免打印出未加载的空图。
function waitForImages(doc: Document, timeoutMs: number): Promise<void> {
  const imgs = Array.from(doc.images || []);
  if (imgs.length === 0) return Promise.resolve();
  const pending = imgs.filter((img) => !img.complete || img.naturalWidth === 0);
  if (pending.length === 0) return Promise.resolve();

  return new Promise((resolve) => {
    let done = 0;
    let settled = false;
    const finish = () => {
      if (settled) return;
      settled = true;
      resolve();
    };
    const tick = () => {
      done += 1;
      if (done >= pending.length) finish();
    };
    pending.forEach((img) => {
      img.addEventListener('load', tick, { once: true });
      img.addEventListener('error', tick, { once: true });
    });
    setTimeout(finish, timeoutMs); // 兜底超时，避免个别图卡住打印
  });
}
