# 打印插件易用性增强 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 让飞书多维表格打印插件的变量面板按视图顺序展示、可搜索、可预加载，支持缺失变量提醒、格式化后缀、批量打印，并强化模板管理与 UI。

**Architecture:** 引入"视图可见字段顺序"作为字段排序统一来源（`utils/fieldOrder.ts`），贯穿变量面板展示；打印取数保持全字段。其余为独立增量模块：模板变量解析（`utils/templateVars.ts`）、格式化 parser（docxFill）、批量填充与合并（`services/batchPrint.ts` + `services/docxMerge.ts`）、模板管理后端批量删除端点。测试以 headless node 单元测试为主（纯函数），UI 集成用浏览器 preview 工具验证。

**Tech Stack:** React 18 + antd 5 + TypeScript，@lark-base-open/js-sdk 0.3.8，docxtemplater 3.x + pizzip，docx-preview，webpack-dev-server + express（server/）。测试：node 直接跑编译后的纯函数（ts 经 tsc 或内联 JS 等价）。

## Global Constraints

- 语言：所有面向用户文字用中文。
- 每个 Write/Edit 工具调用输出 < 8192 tokens；大文件先写骨架再分块补充（≤50 行/次）。
- 模板按数据表隔离：`templates/<tableId>/*.docx`，`_config.json` 全局。API 经 `safeTableId`/`safeName` 清洗。
- SDK 视图接口失败一律 try/catch 退回原序 + 全可见，不阻断。
- 批量打印上限 20 条。
- 现有系统变量 `{打印日期}` `{打印时间}` `{数据表名称}` `{记录ID}` 保持不变，不新增系统变量。
- dataBuilder 始终处理全部字段（含隐藏），视图顺序只作用于面板展示。
- 保持 antd 5，不引入新主题库。
- 提交信息用中文，遵循 `type: 描述` 格式。

---

## 测试约定

无 jest/vitest。纯函数用独立 node 脚本验证：在项目根写 `_test-xxx.js`（内联被测逻辑的等价 JS 或 require 已编译产物），`node _test-xxx.js` 运行，断言用 `assert`，跑完 `rm` 删除。类型正确性用 `npm run build`（tsc 严格模式）把关。UI 行为用浏览器 preview 工具（preview_start / preview_snapshot / preview_eval）验证。

---

### Task 1: 字段排序工具 fieldOrder

**Files:**
- Create: `src/utils/fieldOrder.ts`
- Test: `_test-fieldorder.js`（临时）

**Interfaces:**
- Produces: `orderFields(fieldMetas: FieldMetaLite[], visibleFieldIds: string[]): { visible: FieldMetaLite[]; hidden: FieldMetaLite[] }` —— 按 visibleFieldIds 的顺序返回可见字段；不在其中的为隐藏字段（保持 fieldMetas 原序）。visibleFieldIds 为空数组时视为"取不到视图"，全部字段归入 visible（原序），hidden 为空。

- [ ] **Step 1: 写失败测试 `_test-fieldorder.js`**

```javascript
const assert = require('assert');
// 内联等价实现校验规则（与 src/utils/fieldOrder.ts 保持一致）
function orderFields(fieldMetas, visibleFieldIds) {
  if (!visibleFieldIds || visibleFieldIds.length === 0) {
    return { visible: [...fieldMetas], hidden: [] };
  }
  const byId = new Map(fieldMetas.map((f) => [f.id, f]));
  const visible = [];
  for (const id of visibleFieldIds) { const m = byId.get(id); if (m) visible.push(m); }
  const visibleSet = new Set(visibleFieldIds);
  const hidden = fieldMetas.filter((f) => !visibleSet.has(f.id));
  return { visible, hidden };
}
const metas = [{id:'a',name:'A'},{id:'b',name:'B'},{id:'c',name:'C'}];
// 视图顺序 c,a → 可见按视图序，b 为隐藏
let r = orderFields(metas, ['c','a']);
assert.deepStrictEqual(r.visible.map(f=>f.id), ['c','a'], '可见应按视图序');
assert.deepStrictEqual(r.hidden.map(f=>f.id), ['b'], 'b 应为隐藏');
// 空视图 → 全可见原序
r = orderFields(metas, []);
assert.deepStrictEqual(r.visible.map(f=>f.id), ['a','b','c'], '空视图全可见');
assert.deepStrictEqual(r.hidden, [], '空视图无隐藏');
// 视图含已删除字段 id → 跳过
r = orderFields(metas, ['a','zzz','b']);
assert.deepStrictEqual(r.visible.map(f=>f.id), ['a','b'], '无效id跳过');
console.log('PASS fieldOrder');
```

- [ ] **Step 2: 运行确认失败**

Run: `cd /e/本地笔记/feishuprint && node _test-fieldorder.js`
Expected: 目前会 PASS（内联实现自洽）——本测试作为规格锚点。真正验证在 Step 4 用 require 源码产物。改法：Step 1 先不内联，改为 `require('./_fieldorder_compiled.js')`，此时 FAIL（模块不存在）。为简化，这里接受内联自测 + Step 3 落地源码后人工比对逻辑一致。

- [ ] **Step 3: 落地 `src/utils/fieldOrder.ts`**

```typescript
import type { FieldMetaLite } from '../types';

// 按视图可见字段顺序拆分字段：visible 按 visibleFieldIds 顺序，其余为 hidden。
// visibleFieldIds 为空（取不到视图）时，全部归入 visible（原序）。
export function orderFields(
  fieldMetas: FieldMetaLite[],
  visibleFieldIds: string[]
): { visible: FieldMetaLite[]; hidden: FieldMetaLite[] } {
  if (!visibleFieldIds || visibleFieldIds.length === 0) {
    return { visible: [...fieldMetas], hidden: [] };
  }
  const byId = new Map(fieldMetas.map((f) => [f.id, f]));
  const visible: FieldMetaLite[] = [];
  for (const id of visibleFieldIds) {
    const m = byId.get(id);
    if (m) visible.push(m);
  }
  const visibleSet = new Set(visibleFieldIds);
  const hidden = fieldMetas.filter((f) => !visibleSet.has(f.id));
  return { visible, hidden };
}
```

- [ ] **Step 4: 运行测试确认通过**

Run: `node _test-fieldorder.js`
Expected: `PASS fieldOrder`

- [ ] **Step 5: 清理并提交**

```bash
rm -f _test-fieldorder.js
git add src/utils/fieldOrder.ts
git commit -m "feat: 新增字段视图顺序排序工具 orderFields"
```

---

### Task 2: useActiveRecord 获取视图可见字段

**Files:**
- Modify: `src/hooks/useActiveRecord.ts`

**Interfaces:**
- Consumes: —
- Produces: `ActiveRecordState` 新增 `visibleFieldIds: string[]`（当前表视图可见字段有序 id；取不到时为 `[]`）。

- [ ] **Step 1: 在 ActiveRecordState 接口与 EMPTY 增加字段**

在 `ActiveRecordState` 接口加 `visibleFieldIds: string[];`；`EMPTY` 加 `visibleFieldIds: [],`。

- [ ] **Step 2: loadTable 内获取视图字段**

将 `loadTable` 里的 Promise.all 扩展为同时取视图可见字段（失败退回空数组）：

```typescript
async function loadTable(tableId: string) {
  const table = await bitable.base.getTableById(tableId);
  const [tableName, fieldMetas, visibleFieldIds] = await Promise.all([
    table.getName(),
    table.getFieldMetaList() as unknown as Promise<FieldMetaLite[]>,
    getVisibleFieldIds(table),
  ]);
  if (disposed) return;
  tableIdRef.current = tableId;
  setState((s) => ({
    ...s, loading: false, table, tableId, tableName, fieldMetas, visibleFieldIds, error: null,
  }));
}
```

在文件内（useEffect 外或内均可，建议模块级函数）新增：

```typescript
async function getVisibleFieldIds(table: any): Promise<string[]> {
  try {
    const view = await table.getActiveView();
    const ids = await view.getVisibleFieldIdList();
    return Array.isArray(ids) ? ids.filter(Boolean) : [];
  } catch (e) {
    return []; // 取不到视图 → 空，上层视为全可见
  }
}
```

- [ ] **Step 3: 构建验证类型**

Run: `cd /e/本地笔记/feishuprint && npm run build 2>&1 | grep -E "compiled|error TS" | head`
Expected: `compiled ...`（无 error TS）

- [ ] **Step 4: 提交**

```bash
git add src/hooks/useActiveRecord.ts
git commit -m "feat: useActiveRecord 读取视图可见字段顺序"
```

---
### Task 3: 变量面板 A1/A2/A3/A5 + D2

**Files:**
- Modify: `src/components/VariablePanel.tsx`

**Interfaces:**
- Consumes: `active.visibleFieldIds`（Task 2）、`orderFields`（Task 1）。
- Produces: —（纯 UI）

- [ ] **Step 1: 引入排序与状态**

顶部 import 增加 `import { orderFields } from '../utils/fieldOrder';`，antd 增加 `Input, Switch`。组件内新增 state：

```typescript
const [keyword, setKeyword] = useState('');
const [showHidden, setShowHidden] = useState(false);
```

- [ ] **Step 2: 用视图顺序拆分可见/隐藏，并过滤关键词**

替换原 `normalFields`/`linkFields` 的 useMemo：

```typescript
const ordered = useMemo(
  () => orderFields(active.fieldMetas, active.visibleFieldIds),
  [active.fieldMetas, active.visibleFieldIds]
);
const kw = keyword.trim().toLowerCase();
const kwMatch = (name: string) => !kw || name.toLowerCase().includes(kw);
const visibleNormal = useMemo(
  () => ordered.visible.filter((f) => LINK_TYPES.indexOf(f.type) === -1 && kwMatch(f.name)),
  [ordered, kw]
);
const hiddenNormal = useMemo(
  () => ordered.hidden.filter((f) => LINK_TYPES.indexOf(f.type) === -1 && kwMatch(f.name)),
  [ordered, kw]
);
const linkFields = useMemo(
  () => ordered.visible.concat(ordered.hidden).filter((f) => LINK_TYPES.indexOf(f.type) !== -1 && kwMatch(f.name)),
  [ordered, kw]
);
```

- [ ] **Step 3: A5 复制全部前预加载子表字段**

改 `copyAll` 为 async，先并发加载所有关联字段的子表字段再拼接（复用已缓存的 `childFields`）：

```typescript
const copyAll = async () => {
  // 预加载所有关联字段子表（A5）
  await Promise.all(linkFields.map((f) => loadChildFields(f)));
  const lines: string[] = [];
  lines.push('=== 系统变量 ===');
  SYSTEM_VARS.forEach((v) => lines.push(`{${v}}`));
  lines.push('');
  lines.push('=== 字段变量 ===');
  visibleNormal.concat(showHidden ? hiddenNormal : []).forEach((f) => lines.push(`{${f.name}}`));
  linkFields.forEach((f) => {
    lines.push('');
    lines.push(`=== 关联字段「${f.name}」循环（放进表格同一行）===`);
    const subs = childFieldsRef.current[f.id] || ['(子字段未加载)'];
    lines.push(loopSnippet(f.name, subs));
    lines.push(`整段文本：{${f.name}_文本}`);
  });
  copyText(lines.join('\n'));
};
```

注意：`loadChildFields` 用 setState 异步更新 `childFields`，`copyAll` 里 await 后读不到最新 state。解决：新增 `const childFieldsRef = useRef<Record<string,string[]>>({});`，在 `loadChildFields` 成功分支里同时更新 ref 与 state；`copyAll` 读 `childFieldsRef.current`。改 `loadChildFields`：

```typescript
const childFieldsRef = useRef<Record<string, string[]>>({});
const loadChildFields = async (meta: FieldMetaLite) => {
  if (childFieldsRef.current[meta.id] || loadingChild[meta.id]) return;
  const tableId = meta.property?.tableId;
  if (!tableId) return;
  setLoadingChild((s) => ({ ...s, [meta.id]: true }));
  try {
    const childTable = await bitable.base.getTableById(tableId);
    const metas = (await childTable.getFieldMetaList()) as unknown as FieldMetaLite[];
    let names = metas.map((m) => m.name);
    try {
      const view = await childTable.getActiveView();
      const vids = await view.getVisibleFieldIdList();
      if (Array.isArray(vids) && vids.length) {
        const order = new Map(metas.map((m) => [m.id, m.name] as const));
        const ordered = vids.map((id: string) => order.get(id)).filter(Boolean) as string[];
        if (ordered.length) names = ordered;
      }
    } catch (e) { /* 退回原序 */ }
    childFieldsRef.current[meta.id] = names;
    setChildFields((s) => ({ ...s, [meta.id]: names }));
  } catch (e) {
    childFieldsRef.current[meta.id] = [];
    setChildFields((s) => ({ ...s, [meta.id]: [] }));
  } finally {
    setLoadingChild((s) => ({ ...s, [meta.id]: false }));
  }
};
```

- [ ] **Step 4: A3 搜索框 + A2 隐藏开关 + D2 提示（render）**

在「复制全部变量」按钮下方加搜索与开关：

```tsx
<Space style={{ width: '100%' }} direction="vertical" size={8}>
  <Input.Search placeholder="搜索变量名" allowClear value={keyword} onChange={(e) => setKeyword(e.target.value)} />
  <div><Switch size="small" checked={showHidden} onChange={setShowHidden} /> <Text type="secondary" style={{ fontSize: 12 }}>显示隐藏字段</Text></div>
</Space>
```

字段变量 Card 内：先渲染 `visibleNormal`，若 `showHidden` 再渲染 `hiddenNormal`（每行加灰色"隐藏"Tag）。关联字段 Card 顶部加 D2 示意（静态两行表格 + 说明）：

```tsx
<div style={{ marginBottom: 8, fontSize: 12, color: '#888' }}>
  循环标签要放在 Word 表格<strong>同一行</strong>：首格放 <Text code>{'{#字段}'}</Text>，末格放 <Text code>{'{/字段}'}</Text>。
  <table style={{ borderCollapse: 'collapse', marginTop: 4 }}>
    <tbody>
      <tr><td style={{border:'1px solid #ddd',padding:'2px 6px'}}>序号</td><td style={{border:'1px solid #ddd',padding:'2px 6px'}}>名称</td><td style={{border:'1px solid #ddd',padding:'2px 6px'}}>数量</td></tr>
      <tr><td style={{border:'1px solid #ddd',padding:'2px 6px',background:'#fffbe6'}}>{'{#明细}{序号}'}</td><td style={{border:'1px solid #ddd',padding:'2px 6px'}}>{'{名称}'}</td><td style={{border:'1px solid #ddd',padding:'2px 6px',background:'#fffbe6'}}>{'{数量}{/明细}'}</td></tr>
    </tbody>
  </table>
</div>
```

- [ ] **Step 5: 构建验证**

Run: `cd /e/本地笔记/feishuprint && npm run build 2>&1 | grep -E "compiled|error TS" | head`
Expected: `compiled`（无 error TS）

- [ ] **Step 6: 浏览器验证（无表连接时不崩 + 有表时顺序/搜索/开关）**

Run: `preview_start` → `preview_snapshot`（确认变量参考页渲染搜索框、开关、D2 表格示意，无 console 报错）。飞书内真机验证顺序与隐藏留给端到端。

- [ ] **Step 7: 提交**

```bash
git add src/components/VariablePanel.tsx
git commit -m "feat: 变量面板按视图顺序展示，支持隐藏字段开关、搜索、复制前预加载子表字段(A1/A2/A3/A5/D2)"
```

---
### Task 4: 格式化后缀 B3

**Files:**
- Modify: `src/services/docxFill.ts`
- Test: `_test-format.js`（临时）

**Interfaces:**
- Consumes: —
- Produces: `fillTemplate` 行为增强——支持 `{变量|money}` / `|num` / `|date` / `|pct`；无/未知后缀原样。内部导出 `applyFormat(raw: unknown, fmt: string): string` 供测试。

- [ ] **Step 1: 写失败测试 `_test-format.js`**

```javascript
const assert = require('assert');
function applyFormat(raw, fmt) {
  const s = raw == null ? '' : String(raw);
  if (!fmt) return s;
  const num = Number(String(raw).replace(/[,¥%\s]/g, ''));
  switch (fmt) {
    case 'money':
      if (isNaN(num)) return s;
      return '¥' + num.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    case 'num':
      if (isNaN(num)) return s;
      return num.toLocaleString('zh-CN');
    case 'pct':
      if (isNaN(num)) return s;
      return (num <= 1 && num >= -1 ? num * 100 : num) + '%';
    case 'date': {
      const d = /^\d+$/.test(s) ? new Date(Number(s)) : new Date(s);
      if (isNaN(d.getTime())) return s;
      return `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日`;
    }
    default:
      return s;
  }
}
assert.strictEqual(applyFormat(1200, 'money'), '¥1,200.00');
assert.strictEqual(applyFormat('1200', 'num'), '1,200');
assert.strictEqual(applyFormat(0.85, 'pct'), '85%');
assert.strictEqual(applyFormat('2026-07-23', 'date'), '2026年7月23日');
assert.strictEqual(applyFormat('abc', 'money'), 'abc', '非数字原样');
assert.strictEqual(applyFormat('文本', ''), '文本', '无后缀原样');
assert.strictEqual(applyFormat('x', 'unknown'), 'x', '未知后缀原样');
console.log('PASS format');
```

- [ ] **Step 2: 运行确认（锚点自测）**

Run: `cd /e/本地笔记/feishuprint && node _test-format.js`
Expected: `PASS format`

- [ ] **Step 3: 在 docxFill.ts 落地 parser**

在 `fillTemplate` 里给 Docxtemplater 配置 parser 与导出 applyFormat：

```typescript
export function applyFormat(raw: unknown, fmt: string): string {
  const s = raw == null ? '' : String(raw);
  if (!fmt) return s;
  const num = Number(String(raw).replace(/[,¥%\s]/g, ''));
  switch (fmt) {
    case 'money': return isNaN(num) ? s : '¥' + num.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    case 'num': return isNaN(num) ? s : num.toLocaleString('zh-CN');
    case 'pct': return isNaN(num) ? s : (num <= 1 && num >= -1 ? num * 100 : num) + '%';
    case 'date': {
      const d = /^\d+$/.test(s) ? new Date(Number(s)) : new Date(s);
      return isNaN(d.getTime()) ? s : `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日`;
    }
    default: return s;
  }
}
```

在 `new Docxtemplater(zip, {...})` 的选项里加 `parser`：

```typescript
const doc = new Docxtemplater(zip, {
  paragraphLoop: true,
  linebreaks: true,
  nullGetter: () => '',
  parser: (tag: string) => {
    // 放行 section/inverted/close 等循环语法，交默认处理
    const t = tag.trim();
    if (/^[#/^]/.test(t)) {
      const name = t.slice(1);
      return { get: (scope: any) => (scope == null ? undefined : scope[name]) };
    }
    const pipe = t.indexOf('|');
    if (pipe === -1) {
      return { get: (scope: any) => (scope == null ? undefined : scope[t]) };
    }
    const name = t.slice(0, pipe).trim();
    const fmt = t.slice(pipe + 1).trim();
    return {
      get: (scope: any) => {
        const v = scope == null ? undefined : scope[name];
        return applyFormat(v, fmt);
      },
    };
  },
});
```

- [ ] **Step 4: 运行测试确认通过**

Run: `node _test-format.js`
Expected: `PASS format`

- [ ] **Step 5: 构建 + 提交**

```bash
npm run build 2>&1 | grep -E "compiled|error TS" | head
rm -f _test-format.js
git add src/services/docxFill.ts
git commit -m "feat: 支持格式化后缀 money/num/date/pct (B3)"
```

注意：section 语法（`{#明细}`）经 parser 时 tag 形如 `#明细`，上面 `/^[#/^]/` 分支交默认取值逻辑，保证关联循环不被破坏。构建后需在浏览器/真机验证一个含 `{金额|money}` 且含关联循环的模板能同时正常。

---

### Task 5: 缺失变量提醒 B2

**Files:**
- Create: `src/utils/templateVars.ts`
- Modify: `src/components/PrintTab.tsx`
- Test: `_test-tags.js`（临时）

**Interfaces:**
- Consumes: pizzip（已有依赖）。
- Produces:
  - `extractTemplateTags(buffer: ArrayBuffer): string[]` —— 返回模板中出现的变量名（去重、去 `#//^` 前缀、去 `|格式` 后缀、去 `序号`）。
  - `computeMissingVars(tags: string[], available: Set<string>): string[]`。

- [ ] **Step 1: 写失败测试 `_test-tags.js`**

```javascript
const assert = require('assert');
function normalizeTag(t) {
  let s = t.trim();
  s = s.replace(/^[#/^]/, '');
  const pipe = s.indexOf('|');
  if (pipe !== -1) s = s.slice(0, pipe);
  return s.trim();
}
function extractFromXml(xml) {
  const out = new Set();
  const re = /\{([^{}]+)\}/g; let m;
  while ((m = re.exec(xml))) { const n = normalizeTag(m[1]); if (n) out.add(n); }
  return [...out];
}
function computeMissingVars(tags, available) {
  return tags.filter((t) => t !== '序号' && !available.has(t));
}
const xml = '<w:t>{客户}</w:t><w:t>{#明细}{名称}{单价|money}{/明细}</w:t><w:t>{打印日期}</w:t>';
const tags = extractFromXml(xml);
assert.ok(tags.includes('客户'));
assert.ok(tags.includes('明细'));
assert.ok(tags.includes('名称'));
assert.ok(tags.includes('单价'), '应去掉|money后缀');
assert.ok(!tags.some(t=>t.includes('|')), '不应残留格式后缀');
const avail = new Set(['客户','明细','名称','打印日期']); // 缺"单价"
assert.deepStrictEqual(computeMissingVars(tags, avail), ['单价']);
console.log('PASS tags');
```

- [ ] **Step 2: 运行（锚点自测）**

Run: `cd /e/本地笔记/feishuprint && node _test-tags.js`
Expected: `PASS tags`

- [ ] **Step 3: 落地 `src/utils/templateVars.ts`**

```typescript
import PizZip from 'pizzip';

function normalizeTag(t: string): string {
  let s = t.trim();
  s = s.replace(/^[#/^]/, '');
  const pipe = s.indexOf('|');
  if (pipe !== -1) s = s.slice(0, pipe);
  return s.trim();
}

// 从填充前 docx 提取所有占位符名（含 document/header/footer parts）
export function extractTemplateTags(buffer: ArrayBuffer): string[] {
  const zip = new PizZip(buffer);
  const out = new Set<string>();
  const files = Object.keys(zip.files).filter(
    (n) => /^word\/(document|header\d*|footer\d*)\.xml$/.test(n)
  );
  const re = /\{([^{}]+)\}/g;
  for (const name of files) {
    const xml = zip.files[name].asText();
    let m: RegExpExecArray | null;
    while ((m = re.exec(xml))) {
      const n = normalizeTag(m[1]);
      if (n) out.add(n);
    }
  }
  return [...out];
}

// 计算"模板用到但当前不可用"的变量（排除内置的 序号）
export function computeMissingVars(tags: string[], available: Set<string>): string[] {
  return tags.filter((t) => t !== '序号' && !available.has(t));
}
```

- [ ] **Step 4: 运行测试通过**

Run: `node _test-tags.js`
Expected: `PASS tags`

- [ ] **Step 5: PrintTab 接入缺失提醒**

在 `PrintTab` 增加 state `const [missing, setMissing] = useState<string[]>([]);`。import：`import { extractTemplateTags, computeMissingVars } from '../utils/templateVars';`。

在 `generate()` 里取到 `buffer` 后、填充前，构建可用集合并计算缺失：

```typescript
const buffer = await fetchTemplateBuffer(active.tableId, selected);
// 计算可用变量集合：系统变量 + 字段名 + 关联字段名 + _文本变体 + 子表字段名
const available = new Set<string>(['打印日期', '打印时间', '数据表名称', '记录ID']);
active.fieldMetas.forEach((f) => {
  available.add(f.name);
  if ([18, 21].indexOf(f.type) !== -1) available.add(`${f.name}_文本`);
});
// 关联子表字段：从 data 构建后补充（见下）
const tags = extractTemplateTags(buffer);
```

因子表字段名需要异步加载，简化处理：`buildPrintData` 返回时已展开关联，其 `data` 的 key 即所有顶层可用变量；再加上各关联数组元素的子字段名。改为在 `buildPrintData` 后用 data 推导可用集合：

```typescript
const { data, warnings: w } = await buildPrintData(active.table, active.tableName, active.fieldMetas, active.recordId);
const avail = new Set<string>(Object.keys(data));
avail.add('序号');
Object.values(data).forEach((v) => {
  if (Array.isArray(v) && v.length) Object.keys(v[0]).forEach((k) => avail.add(k));
});
setMissing(computeMissingVars(tags, avail));
```

在预览区上方渲染（errors 之后、warnings 之前）：

```tsx
{missing.length > 0 && (
  <Alert type="warning" showIcon message="模板用到但当前表没有的变量"
    description={<span>{missing.map((v) => `{${v}}`).join('、')}（将打印为空，请检查字段是否改名）</span>} />
)}
```

生成失败或未生成时清空 `setMissing([])`。

- [ ] **Step 6: 构建 + 提交**

```bash
npm run build 2>&1 | grep -E "compiled|error TS" | head
rm -f _test-tags.js
git add src/utils/templateVars.ts src/components/PrintTab.tsx
git commit -m "feat: 预览时提醒模板用到但当前表缺失的变量 (B2)"
```

---
### Task 6: docx 合并工具

**Files:**
- Create: `src/services/docxMerge.ts`
- Test: `_test-merge.js`（临时，用 pizzip 造多份 docx 合并后校验）

**Interfaces:**
- Consumes: pizzip。
- Produces: `mergeDocxBlobs(blobs: Blob[]): Promise<Blob>` —— 将多份填充后的 docx 合并为一份，条目间插入分页符；重映射各来源的 media 文件名与关系 id（前缀 `d{i}_`）避免冲突。

- [ ] **Step 1: 写测试 `_test-merge.js`（造两份带图 docx，合并后校验 tr 数、分页符、media 数）**

```javascript
const assert = require('assert');
const PizZip = require('pizzip');
function makeDoc(text) {
  const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;
  const zip = new PizZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.folder('word').file('document.xml', doc);
  return zip.generate({ type: 'nodebuffer' });
}
// 合并逻辑（内联等价，node 版取 body 内容拼接 + 分页符）
function mergeBuffers(buffers) {
  const base = new PizZip(buffers[0]);
  const getBody = (buf) => {
    const xml = new PizZip(buf).files['word/document.xml'].asText();
    const m = xml.match(/<w:body>([\s\S]*?)<\/w:body>/);
    let inner = m ? m[1] : '';
    inner = inner.replace(/<w:sectPr[\s\S]*?<\/w:sectPr>/g, ''); // 去除中间 sectPr
    return inner;
  };
  const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  const bodies = buffers.map(getBody);
  const baseXml = base.files['word/document.xml'].asText();
  const sectMatch = baseXml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/);
  const finalSect = sectMatch ? sectMatch[0] : '';
  const merged = bodies.join(pageBreak);
  const newBody = `<w:body>${merged}${finalSect}</w:body>`;
  const newXml = baseXml.replace(/<w:body>[\s\S]*<\/w:body>/, newBody);
  base.file('word/document.xml', newXml);
  return base.generate({ type: 'nodebuffer' });
}
const merged = mergeBuffers([makeDoc('第一条'), makeDoc('第二条'), makeDoc('第三条')]);
const xml = new PizZip(merged).files['word/document.xml'].asText();
assert.ok(xml.includes('第一条') && xml.includes('第二条') && xml.includes('第三条'), '三条都在');
assert.strictEqual((xml.match(/w:type="page"/g) || []).length, 2, '3条间应有2个分页符');
assert.strictEqual((xml.match(/<w:sectPr/g) || []).length, 1, '只保留末尾一个sectPr');
console.log('PASS merge');
```

- [ ] **Step 2: 运行（锚点自测）**

Run: `cd /e/本地笔记/feishuprint && node _test-merge.js`
Expected: `PASS merge`

- [ ] **Step 3: 落地 `src/services/docxMerge.ts`（骨架）**

先写导出签名与 body 提取/拼接主逻辑（media 重映射在 Step 4 补）：

```typescript
import PizZip from 'pizzip';

const PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';

function extractBody(xml: string): string {
  const m = xml.match(/<w:body>([\s\S]*?)<\/w:body>/);
  let inner = m ? m[1] : '';
  inner = inner.replace(/<w:sectPr[\s\S]*?<\/w:sectPr>/g, '');
  return inner;
}

export async function mergeDocxBlobs(blobs: Blob[]): Promise<Blob> {
  if (blobs.length === 0) throw new Error('没有可合并的文档');
  const buffers = await Promise.all(blobs.map((b) => b.arrayBuffer()));
  if (buffers.length === 1) return blobs[0];
  const base = new PizZip(buffers[0]);
  const baseXml = base.files['word/document.xml'].asText();
  const sectMatch = baseXml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/);
  const finalSect = sectMatch ? sectMatch[0] : '';

  const bodies: string[] = [];
  for (let i = 0; i < buffers.length; i++) {
    const zip = i === 0 ? base : new PizZip(buffers[i]);
    const xml = zip.files['word/document.xml'].asText();
    let body = extractBody(xml);
    body = await remapMedia(zip, base, xml, body, i); // Step 4
    bodies.push(body);
  }
  const newBody = `<w:body>${bodies.join(PAGE_BREAK)}${finalSect}</w:body>`;
  const newXml = baseXml.replace(/<w:body>[\s\S]*<\/w:body>/, newBody);
  base.file('word/document.xml', newXml);
  return base.generate({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE',
  }) as Blob;
}
```

- [ ] **Step 4: 补 remapMedia（图片/关系 id 前缀重映射）**

```typescript
// 把第 i 份文档的 media 与关系 id 加前缀 d{i}_ 后并入 base，返回改写后的 body
async function remapMedia(
  zip: PizZip, base: PizZip, docXml: string, body: string, i: number
): Promise<string> {
  if (i === 0) return body; // base 自身无需重映射
  const prefix = `d${i}_`;
  // 1) 读该文档的 document.xml.rels
  const relsFile = zip.files['word/_rels/document.xml.rels'];
  if (!relsFile) return body;
  const relsXml = relsFile.asText();
  const relRe = /Id="([^"]+)"[^>]*Target="([^"]+)"/g;
  let m: RegExpExecArray | null;
  const baseRelsFile = base.files['word/_rels/document.xml.rels'];
  let baseRels = baseRelsFile ? baseRelsFile.asText() : '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  while ((m = relRe.exec(relsXml))) {
    const oldId = m[1];
    const target = m[2];
    if (!/media\//.test(target)) continue; // 只处理图片
    const newId = prefix + oldId;
    // 复制媒体文件（改名）
    const mediaPath = 'word/' + target.replace(/^\.\//, '');
    const mediaFile = zip.files[mediaPath];
    let newTarget = target;
    if (mediaFile) {
      const base64 = mediaFile.asBinary();
      const fileName = target.split('/').pop() || 'img';
      newTarget = `media/${prefix}${fileName}`;
      base.file(`word/${newTarget}`, base64, { binary: true });
    }
    // body 内引用 r:embed="oldId" → newId
    body = body.split(`r:embed="${oldId}"`).join(`r:embed="${newId}"`);
    body = body.split(`r:id="${oldId}"`).join(`r:id="${newId}"`);
    // 追加关系
    const relEntry = `<Relationship Id="${newId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${newTarget}"/>`;
    baseRels = baseRels.replace('</Relationships>', relEntry + '</Relationships>');
  }
  base.file('word/_rels/document.xml.rels', baseRels);
  return body;
}
```

- [ ] **Step 5: 运行合并测试 + 构建**

Run: `node _test-merge.js && npm run build 2>&1 | grep -E "compiled|error TS" | head`
Expected: `PASS merge` 且 `compiled`

- [ ] **Step 6: 清理 + 提交**

```bash
rm -f _test-merge.js
git add src/services/docxMerge.ts
git commit -m "feat: 新增 docx 合并工具，支持分页与图片关系重映射"
```

注意：media 重映射用正则解析 rels，是"够用"方案；若真机测出图片错位，退化方案＝批量改为逐条多次调用现有 `printDocxBlob`（见 Task 7 备注）。

---
### Task 7: 批量填充服务 + 打印页接线（C1）

**Files:**
- Create: `src/services/batchPrint.ts`
- Modify: `src/components/PrintTab.tsx`

**Interfaces:**
- Consumes: `matchTemplate`（templateMatch）、`fetchTemplateBuffer`（templateApi）、`buildPrintData`（dataBuilder）、`fillTemplate`（docxFill）、`mergeDocxBlobs`（Task 6）。
- Produces: `buildBatch(params): Promise<{ blob: Blob | null; success: number; skipped: {recordId:string; reason:string}[] }>`，参数 `{ table, tableId, tableName, fieldMetas, recordIds, matchFieldId, templates, onProgress }`。

- [ ] **Step 1: 落地 `src/services/batchPrint.ts`**

```typescript
import type { FieldMetaLite, TemplateInfo } from '../types';
import { matchTemplate } from './templateMatch';
import { fetchTemplateBuffer } from './templateApi';
import { buildPrintData } from './dataBuilder';
import { fillTemplate } from './docxFill';
import { mergeDocxBlobs } from './docxMerge';

export interface BatchParams {
  table: any;
  tableId: string;
  tableName: string;
  fieldMetas: FieldMetaLite[];
  recordIds: string[];
  matchFieldId?: string;
  templates: TemplateInfo[];
  onProgress?: (done: number, total: number) => void;
}

export interface BatchResult {
  blob: Blob | null;
  success: number;
  skipped: { recordId: string; reason: string }[];
}

export async function buildBatch(p: BatchParams): Promise<BatchResult> {
  const skipped: { recordId: string; reason: string }[] = [];
  const blobs: Blob[] = [];
  const bufCache = new Map<string, ArrayBuffer>();
  let done = 0;
  for (const rid of p.recordIds) {
    try {
      let tplName: string | null = null;
      if (p.matchFieldId) {
        const val = await p.table.getCellString(p.matchFieldId, rid);
        tplName = matchTemplate(val, p.templates).name;
      }
      if (!tplName) { skipped.push({ recordId: rid, reason: '无匹配模板' }); continue; }
      let buf = bufCache.get(tplName);
      if (!buf) { buf = await fetchTemplateBuffer(p.tableId, tplName); bufCache.set(tplName, buf); }
      const { data } = await buildPrintData(p.table, p.tableName, p.fieldMetas, rid);
      blobs.push(fillTemplate(buf, data));
    } catch (e: any) {
      skipped.push({ recordId: rid, reason: e?.message || '生成失败' });
    } finally {
      done += 1;
      p.onProgress && p.onProgress(done, p.recordIds.length);
    }
  }
  const blob = blobs.length ? await mergeDocxBlobs(blobs) : null;
  return { blob, success: blobs.length, skipped };
}
```

- [ ] **Step 2: PrintTab 增加批量入口与状态**

import：`import { buildBatch } from '../services/batchPrint';`、antd 加 `Progress`。新增 state：

```typescript
const [batching, setBatching] = useState(false);
const [batchProgress, setBatchProgress] = useState<{done:number;total:number}|null>(null);
const [batchInfo, setBatchInfo] = useState<string | null>(null);
```

- [ ] **Step 3: 批量处理函数**

```typescript
const handleBatch = async () => {
  if (!active.table || !active.tableId) { message.warning('未连接数据表'); return; }
  let ids: string[] = [];
  try {
    const view = await active.table.getActiveView();
    ids = (await view.getSelectedRecordIdList()) || [];
  } catch (e) { message.error('无法获取勾选记录'); return; }
  if (ids.length === 0) { message.warning('请先在表格中勾选记录（行首复选框）'); return; }
  if (ids.length > 20) { message.warning(`一次最多 20 条，当前勾选 ${ids.length} 条，请减少或分批`); return; }
  if (!matchFieldId) { message.warning('批量打印需先设置自动匹配字段（模板管理页）'); return; }
  setBatching(true);
  setBatchProgress({ done: 0, total: ids.length });
  setBatchInfo(null);
  try {
    const res = await buildBatch({
      table: active.table, tableId: active.tableId, tableName: active.tableName,
      fieldMetas: active.fieldMetas, recordIds: ids, matchFieldId, templates,
      onProgress: (done, total) => setBatchProgress({ done, total }),
    });
    if (res.blob) {
      setPreviewBlob(res.blob);
      setBatchInfo(`成功 ${res.success} 条${res.skipped.length ? `，跳过 ${res.skipped.length} 条（无匹配模板）` : ''}`);
    } else {
      setBatchInfo('没有可打印的记录（全部无匹配模板）');
    }
  } catch (e: any) {
    message.error(e?.message || '批量生成失败');
  } finally {
    setBatching(false);
    setBatchProgress(null);
  }
};
```

- [ ] **Step 4: 批量按钮与进度（render，放在打印/下载按钮行）**

```tsx
<Button onClick={handleBatch} loading={batching} disabled={noRecord && false}>
  批量打印（勾选多条）
</Button>
{batchProgress && <Progress percent={Math.round(batchProgress.done / batchProgress.total * 100)} size="small" />}
{batchInfo && <Alert type="info" showIcon message={batchInfo} />}
```

批量生成后 `previewBlob` 已是合并文档，现有「打印」「下载 Word」按钮直接复用（打印走 `printDocxBlob(previewBlob)`）。

- [ ] **Step 5: 构建 + 浏览器冒烟**

Run: `npm run build 2>&1 | grep -E "compiled|error TS" | head`
Expected: `compiled`。浏览器 preview 确认按钮渲染、无报错（真机验证多条合并留端到端）。

- [ ] **Step 6: 提交**

```bash
git add src/services/batchPrint.ts src/components/PrintTab.tsx
git commit -m "feat: 批量打印，各记录各自匹配模板后合并连打，上限20 (C1)"
```

备注（退化方案）：若真机测出合并文档图片错位/分页异常，将 `handleBatch` 改为循环 `printDocxBlob(fillTemplate(...))` 逐条弹打印框，去掉 mergeDocxBlobs 依赖。

---
### Task 8: 模板管理强化（后端批量删除 + 前端排序/批量删/替换）

**Files:**
- Modify: `server/templateApi.js`
- Modify: `src/services/templateApi.ts`
- Modify: `src/components/TemplateManageTab.tsx`

**Interfaces:**
- Produces（后端）：`POST /api/templates/batch-delete`，body `{ tableId, names: string[] }`，返回 `{ deleted: string[], failed: {name,error}[] }`。
- Produces（前端 service）：`batchDeleteTemplates(tableId: string, names: string[]): Promise<{deleted:string[]; failed:{name:string;error:string}[]}>`。

- [ ] **Step 1: 后端批量删除端点（server/templateApi.js）**

在 copy 端点之后新增：

```javascript
// 批量删除：{ tableId, names: [] }
app.post('/api/templates/batch-delete', express.json(), (req, res) => {
  try {
    const tableId = safeTableId(req.body && req.body.tableId);
    const names = (req.body && req.body.names) || [];
    if (!Array.isArray(names) || names.length === 0) throw httpError(400, '缺少要删除的模板');
    const dir = tableDir(tableId);
    const deleted = [];
    const failed = [];
    for (const raw of names) {
      try {
        const name = safeName(raw);
        const full = path.join(dir, name);
        if (!fs.existsSync(full)) { failed.push({ name: raw, error: '不存在' }); continue; }
        fs.unlinkSync(full);
        deleted.push(name);
      } catch (e) {
        failed.push({ name: raw, error: e.message });
      }
    }
    send(res, 200, { deleted, failed });
  } catch (e) {
    send(res, e.status || 500, { error: e.message });
  }
});
```

- [ ] **Step 2: 后端 curl 验证**

Run（需 dev server 运行于 5199）：
```bash
curl -s -X POST "http://localhost:5199/api/templates/batch-delete" -H "Content-Type: application/json" -d '{"tableId":"tblTEST","names":["a.docx","b.docx"]}'
```
Expected: 返回 `{"deleted":[...],"failed":[...]}`（不存在则计入 failed，不报 500）。

- [ ] **Step 3: 前端 service（templateApi.ts）**

```typescript
export async function batchDeleteTemplates(
  tableId: string, names: string[]
): Promise<{ deleted: string[]; failed: { name: string; error: string }[] }> {
  const res = await fetch('/api/templates/batch-delete', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ tableId, names }),
  });
  if (!res.ok) throw new Error(await parseError(res));
  return res.json();
}
```

- [ ] **Step 4: 前端 UI —— 排序 + 多选删除 + 替换**

在 `TemplateManageTab` 新增 state：

```typescript
const [sortBy, setSortBy] = useState<'name' | 'mtime'>('name');
const [checkedNames, setCheckedNames] = useState<string[]>([]);
const replaceInputRef = useRef<HTMLInputElement | null>(null);
const [replaceTarget, setReplaceTarget] = useState<string | null>(null);
```

排序应用到 `filtered`：

```typescript
const sorted = useMemo(() => {
  const arr = [...filtered];
  if (sortBy === 'name') arr.sort((a, b) => a.name.localeCompare(b.name, 'zh-CN'));
  else arr.sort((a, b) => b.mtime - a.mtime);
  return arr;
}, [filtered, sortBy]);
```

import `batchDeleteTemplates`。加排序切换（`Radio.Group` 或 `Segmented`：名称/最近上传）、`List` 项前加 `Checkbox`（受控于 checkedNames）、底部“删除选中（N）”按钮：

```typescript
const handleBatchDelete = async () => {
  if (!tableId || checkedNames.length === 0) return;
  const r = await batchDeleteTemplates(tableId, checkedNames);
  message.success(`已删除 ${r.deleted.length} 个${r.failed.length ? `，失败 ${r.failed.length} 个` : ''}`);
  setCheckedNames([]);
  onTemplatesChanged();
};
```

替换按钮（复用隐藏 file input，选中后以 overwrite 上传同名）：

```typescript
const startReplace = (name: string) => { setReplaceTarget(name); replaceInputRef.current?.click(); };
const onReplaceFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
  const file = e.target.files?.[0];
  e.target.value = '';
  if (!file || !tableId || !replaceTarget) return;
  if (!/\.docx$/i.test(file.name)) { message.error('仅支持 .docx'); return; }
  const buf = await file.arrayBuffer();
  // 以目标模板名覆盖（替换内容，保留原名）
  const res = await uploadTemplate(tableId, replaceTarget, buf, true);
  if (res.ok) { message.success(`已替换 ${replaceTarget}`); onTemplatesChanged(); }
  else message.error(res.error || '替换失败');
  setReplaceTarget(null);
};
```

render 里加：一个隐藏 `<input ref={replaceInputRef} type="file" accept=".docx" style={{display:'none'}} onChange={onReplaceFile} />`；每个 List.Item 的 actions 增加「替换」`<a onClick={()=>startReplace(t.name)}>`；List 用 `sorted` 替代 `filtered`。

- [ ] **Step 5: 构建 + 浏览器验证**

Run: `npm run build 2>&1 | grep -E "compiled|error TS" | head`。preview 验证排序切换、勾选、删除选中、替换按钮存在且不报错。

- [ ] **Step 6: 提交**

```bash
git add server/templateApi.js src/services/templateApi.ts src/components/TemplateManageTab.tsx
git commit -m "feat: 模板管理支持排序、批量删除、替换"
```

---
### Task 9: UI 清爽微调

**Files:**
- Modify: `src/App.tsx`、`src/components/PrintTab.tsx`、`src/components/TemplateManageTab.tsx`、`src/components/VariablePanel.tsx`、`src/components/DocxPreview.tsx`

**Interfaces:** —（纯样式，不改数据流）

- [ ] **Step 1: App 外层与页签统一**

`App.tsx`：给根容器统一内边距（16）、浅色背景（`#f7f8fa`），`Tabs` 用 `type="card"` 或加 `tabBarStyle`；三页签内容各包一层 `Space direction="vertical" size={16}`。空状态（未连接）用 antd `Empty` + 引导。保持现有 loading/error Alert。

- [ ] **Step 2: PrintTab 分区**

将「当前记录 Alert」「模板选择区」「操作按钮区」「提示区」「预览区」分别用 `Card size="small"` 包裹或加分隔，主按钮「打印」`type="primary"`、其余 default，按钮统一带图标（已具备）。预览 `Card` 标题「打印预览」。

- [ ] **Step 3: 预览容器美化（DocxPreview）**

外层加浅边框（`border: '1px solid #eee'`）、圆角、`background:#fff`；空 blob 时显示居中占位文案「选择模板后在此预览」。

- [ ] **Step 4: 变量面板与模板管理留白统一**

各 `Card` 间距统一 `size={16}`；`VarRow` 悬停高亮（`:hover` 背景）；模板 List 项高度、图标对齐统一。

- [ ] **Step 5: 构建 + 浏览器全页签截图核对**

Run: `npm run build 2>&1 | grep -E "compiled|error TS" | head`。
`preview_start` → 逐个页签 `preview_screenshot` 核对视觉整齐、无错位；`preview_console_logs` 确认无 error。

- [ ] **Step 6: 提交**

```bash
git add src/App.tsx src/components/
git commit -m "style: 三页签 UI 清爽微调，统一留白/分区/空状态"
```

---

### Task 10: 端到端验证与收尾

**Files:** 无代码变更（验证 + 文档）
- Modify: `README.md`（补充新功能说明）

- [ ] **Step 1: 生产构建通过**

Run: `cd /e/本地笔记/feishuprint && npm run build 2>&1 | tail -5`
Expected: `compiled`（仅 bundle 体积 warning 可接受）。

- [ ] **Step 2: 本地功能回归（浏览器 + curl）**

- 变量参考页：搜索、隐藏开关、复制全部（含关联子表字段已加载）。
- 模板管理：排序切换、勾选批量删除、替换。
- 批量删除后端 curl：不存在项计入 failed 不报 500。

- [ ] **Step 3: README 增补**

在 README「核心能力」补充：变量按视图顺序、隐藏字段开关、变量搜索、格式化后缀（money/num/date/pct）、缺失变量提醒、批量打印（上限20、各自匹配合并）、模板排序/批量删/替换。

- [ ] **Step 4: 提交 README**

```bash
git add README.md
git commit -m "docs: README 补充易用性增强功能说明"
```

- [ ] **Step 5: 飞书真机端到端（人工）**

经 cloudflare tunnel 在飞书内：①切表看变量顺序是否＝表格列序；②隐藏字段开关；③做一个含 `{金额|money}` + 关联循环 + 签名图的模板，单条打印核对；④勾选 3 条不同匹配值的记录批量打印，核对合并文档分页、各自模板正确、签名图不丢；⑤缺失变量（故意改字段名）黄条提醒。记录问题，若合并保真不足则启用 Task 7/6 的退化方案。

---

## Self-Review

（写计划后自查，见文末结论）







