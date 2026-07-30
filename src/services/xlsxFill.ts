import PizZip from 'pizzip';
import type { PrintDataValue, LinkedRow } from '../types';
import { amountToChinese } from './money';

const MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

// ============================================================================
// 轻量 xlsx 模板填充引擎
// ----------------------------------------------------------------------------
// docxtemplater 的 xlsx 模块是付费的，这里针对「货单」这类结构自研一个：
//   - 普通占位符  {字段名}         → 直接替换单元格文本
//   - 明细循环    {#字段}…{/字段}  → 该字段值为数组，把「含循环标记的整行」按数据条数复制
//   - 自动求和    {合计数量}/{合计金额} 等 → 见 SUM_MAP 约定或显式传入
//
// 实现策略：把所有共享字符串(sharedStrings)内联进 worksheet，再对 sheet XML 做
// 纯文本占位符替换 + 循环行克隆 + 行号/合并单元格重排。这样避免共享字符串索引维护。
// ============================================================================

const XLML_ESC: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' };
function escapeXml(s: string): string {
  return String(s).replace(/[&<>"']/g, (c) => XLML_ESC[c]);
}

// 从 "12"、"1,234.5"、"￥100"、"12个" 等文本提取数字
export function parseNumber(v: PrintDataValue): number {
  if (typeof v === 'number') return v;
  if (v == null) return 0;
  const m = String(v).replace(/[^0-9.\-]/g, '').match(/-?\d+(\.\d+)?/);
  return m ? parseFloat(m[0]) : 0;
}

function formatNumber(n: number): string {
  if (!isFinite(n)) return '0';
  return Number.isInteger(n) ? String(n) : Number(n.toFixed(2)).toString();
}

// 解析共享字符串表 → 字符串数组（<si> 可能含多个 <t> run，拼接）
export function parseSharedStrings(xml: string): string[] {
  const out: string[] = [];
  const siRe = /<si>([\s\S]*?)<\/si>/g;
  let m: RegExpExecArray | null;
  while ((m = siRe.exec(xml))) {
    const runs = m[1].match(/<t[^>]*>([\s\S]*?)<\/t>/g) || [];
    const text = runs.map((r) => r.replace(/<t[^>]*>([\s\S]*?)<\/t>/, '$1')).join('');
    out.push(unescapeXml(text));
  }
  return out;
}

function unescapeXml(s: string): string {
  return s
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

// 把 worksheet 里所有 t="s"（共享字符串引用）单元格改为 t="inlineStr" 内联文本，
// 这样占位符 {字段} 就直接出现在 sheet XML 中，便于后续替换与循环克隆。
export function inlineSharedStrings(sheetXml: string, strings: string[]): string {
  return sheetXml.replace(
    /<c ([^>]*?)t="s"([^>]*?)>\s*<v>(\d+)<\/v>\s*<\/c>/g,
    (_all, pre, post, idx) => {
      const text = strings[parseInt(idx, 10)] ?? '';
      const attrs = (pre + post).replace(/\s+/g, ' ').trim();
      return `<c ${attrs} t="inlineStr"><is><t xml:space="preserve">${escapeXml(text)}</t></is></c>`;
    }
  );
}

// 提取模板中的全部占位符名（{字段}、{#字段}、{/字段}）
export function extractPlaceholders(strings: string[]): string[] {
  const set = new Set<string>();
  for (const s of strings) {
    const ms = s.match(/\{[#/]?[^{}]+\}/g) || [];
    for (const raw of ms) {
      const name = raw.replace(/[{}#/]/g, '').trim();
      if (name) set.add(name);
    }
  }
  return Array.from(set);
}

// 找到循环字段名：形如 {#字段} … {/字段}。返回第一个匹配的字段名。
export function findLoopField(sheetXml: string): string | null {
  const m = sheetXml.match(/\{#([^{}]+)\}/);
  return m ? m[1].trim() : null;
}

interface RowInfo { r: number; xml: string; start: number; end: number; }

// 按 <row r="N"> 拆出所有行
function splitRows(sheetData: string): RowInfo[] {
  const rows: RowInfo[] = [];
  const re = /<row r="(\d+)"[\s\S]*?<\/row>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(sheetData))) {
    rows.push({ r: parseInt(m[1], 10), xml: m[0], start: m.index, end: m.index + m[0].length });
  }
  return rows;
}

// 把一行 XML 的行号从 old 改为 nn（包括 <row r> 和内部 <c r="A5"> 引用）
function renumberRow(rowXml: string, oldR: number, newR: number): string {
  return rowXml
    .replace(new RegExp(`<row r="${oldR}"`), `<row r="${newR}"`)
    .replace(new RegExp(`(<c r="[A-Z]+)${oldR}"`, 'g'), `$1${newR}"`);
}

// 用一条数据行的值替换循环行 XML 里的占位符（去掉 {#..}{/..} 标记）
function fillLoopRow(rowXml: string, row: LinkedRow): string {
  let xml = rowXml.replace(/\{#[^{}]+\}/g, '').replace(/\{\/[^{}]+\}/g, '');
  xml = xml.replace(/\{([^{}]+)\}/g, (_all, name) => {
    const v = row[name.trim()];
    return escapeXml(v == null ? '' : String(v));
  });
  return xml;
}

// 替换普通（非循环）占位符
function fillScalars(xml: string, data: Record<string, PrintDataValue>): string {
  return xml.replace(/\{([^#/{}][^{}]*)\}/g, (all, name) => {
    const key = name.trim();
    if (Array.isArray(data[key])) return all; // 循环字段不在这里处理
    const v = data[key];
    return v == null || Array.isArray(v) ? '' : escapeXml(String(v));
  });
}

// 求和：默认对明细里的「数量」求和写入 {合计数量}，「金额」/「产品总价」写入 {合计金额}。
// 也可由调用方在 data 里预置这些字段来覆盖。
function computeAutoSums(rows: LinkedRow[]): Record<string, string> {
  const pick = (cands: string[]): number => {
    for (const c of cands) {
      if (rows.some((r) => r[c] != null && r[c] !== '')) {
        return rows.reduce((s, r) => s + parseNumber(r[c]), 0);
      }
    }
    return 0;
  };
  return {
    合计数量: formatNumber(pick(['数量'])),
    合计金额: formatNumber(pick(['金额', '产品总价', '总价'])),
  };
}

// 平移合并单元格：循环行克隆后，原本在 loopRow 之下的所有行都下移 delta 行。
// 合并区间形如 "B11:J11"、"A10:E10"；对行号 >= loopRow+1 的端点整体 +delta。
function shiftMergeCells(sheetXml: string, loopRow: number, delta: number): string {
  if (delta === 0) return sheetXml;
  return sheetXml.replace(/ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/g, (all, c1, r1, c2, r2) => {
    const nr1 = parseInt(r1, 10);
    const nr2 = parseInt(r2, 10);
    if (nr1 > loopRow) {
      return `ref="${c1}${nr1 + delta}:${c2}${nr2 + delta}"`;
    }
    return all;
  });
}

// 用数据填充 xlsx 模板，返回填充后的 Blob（下载/预览三用）。
export function fillXlsx(
  templateBuffer: ArrayBuffer,
  data: Record<string, PrintDataValue>
): Blob {
  const zip = new PizZip(templateBuffer);
  const ssFile = zip.file('xl/sharedStrings.xml');
  const strings = ssFile ? parseSharedStrings(ssFile.asText()) : [];

  // 找第一个 worksheet
  const sheetName = Object.keys(zip.files).find((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n));
  if (!sheetName) throw new Error('xlsx 模板缺少工作表');
  let sheetXml = zip.file(sheetName)!.asText();

  // 1) 内联共享字符串，占位符进入 sheet
  sheetXml = inlineSharedStrings(sheetXml, strings);

  // 2) 处理明细循环
  const loopField = findLoopField(sheetXml);
  let delta = 0;
  let loopRowNum = 0;
  if (loopField) {
    const rows: LinkedRow[] = Array.isArray(data[loopField]) ? (data[loopField] as LinkedRow[]) : [];
    const sdMatch = sheetXml.match(/<sheetData>([\s\S]*?)<\/sheetData>/);
    if (sdMatch) {
      const sheetData = sdMatch[1];
      const allRows = splitRows(sheetData);
      const loopRow = allRows.find((r) => /\{#[^{}]+\}/.test(r.xml));
      if (loopRow) {
        loopRowNum = loopRow.r;
        const dataRows = rows.length > 0 ? rows : [];
        const n = dataRows.length;
        delta = Math.max(0, n - 1); // 原本占 1 行，n 行则新增 n-1 行
        // 生成填充后的循环行组
        const cloned = (n === 0 ? [{} as LinkedRow] : dataRows).map((row, i) => {
          const nr = loopRow.r + i;
          const filled = fillLoopRow(loopRow.xml, row);
          return renumberRow(filled, loopRow.r, nr);
        }).join('');
        // 循环行之后的行整体下移 delta，并重新编号
        const after = allRows.filter((r) => r.r > loopRow.r)
          .map((r) => renumberRow(r.xml, r.r, r.r + delta)).join('');
        const before = allRows.filter((r) => r.r < loopRow.r).map((r) => r.xml).join('');
        const newSheetData = before + cloned + after;
        sheetXml = sheetXml.replace(/<sheetData>[\s\S]*?<\/sheetData>/, `<sheetData>${newSheetData}</sheetData>`);
        // 合并单元格随之下移
        sheetXml = shiftMergeCells(sheetXml, loopRowNum, delta);
        // 注入自动求和（若调用方未提供）
        const sums = computeAutoSums(dataRows);
        for (const [k, v] of Object.entries(sums)) {
          if (data[k] == null) data[k] = v;
        }
        // 合计金额大写（若模板用到 {合计金额大写}/{金额大写}）
        if (data['合计金额大写'] == null) data['合计金额大写'] = amountToChinese(data['合计金额']);
        if (data['金额大写'] == null) data['金额大写'] = amountToChinese(data['合计金额']);
      }
    }
  }

  // 3) 替换普通占位符
  sheetXml = fillScalars(sheetXml, data);

  zip.file(sheetName, sheetXml);
  return zip.generate({ type: 'blob', mimeType: MIME, compression: 'DEFLATE' }) as Blob;
}

export function isXlsxName(name: string): boolean {
  return /\.xlsx$/i.test(name);
}
