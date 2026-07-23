import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import type { PrintDataValue } from '../types';

// 用数据填充 .docx 模板，返回填充后的 Blob（预览/打印/下载三用）。
export function fillTemplate(
  templateBuffer: ArrayBuffer,
  data: Record<string, PrintDataValue>
): Blob {
  const zip = new PizZip(templateBuffer);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    // 未提供的变量渲染为空串，而非报错或残留标签
    nullGetter: () => '',
  });

  doc.render(data);

  const out = doc.getZip().generate({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE',
  });
  return out as Blob;
}

// 把 docxtemplater 的报错翻译为中文可读信息
export function explainDocxError(err: any): string[] {
  const msgs: string[] = [];
  if (err && err.properties && Array.isArray(err.properties.errors)) {
    for (const e of err.properties.errors) {
      const ctx = e?.properties?.context || e?.properties?.xtag || '';
      const id = e?.properties?.id || '';
      if (id === 'unopened_tag' || id === 'unclosed_tag') {
        msgs.push(`标签未正确闭合：${ctx}（请检查 {#字段}…{/字段} 是否成对）`);
      } else if (id === 'duplicate_open_tag' || id === 'duplicate_close_tag') {
        msgs.push(`标签重复：${ctx}`);
      } else if (id === 'unbalanced_loop_tags') {
        msgs.push(`循环标签不匹配：${ctx}`);
      } else {
        msgs.push(e?.message || `模板标签错误：${ctx}`);
      }
    }
  }
  if (msgs.length === 0) {
    msgs.push(err?.message || '模板填充失败，请检查模板标签是否正确');
  }
  return msgs;
}
