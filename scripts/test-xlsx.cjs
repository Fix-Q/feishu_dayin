/* xlsx 填充引擎自测：对真实货单模板验证占位符、循环、求和、结构完整性。
 * 运行：node scripts/test-xlsx.cjs
 * 注意：本文件用 CommonJS，直接内联复制 xlsxFill 的核心逻辑做黑盒验证不现实，
 * 因此改为通过 ts 编译产物思路——这里用 require 加载 pizzip 直接跑引擎的等价实现。
 * 为避免 TS/ESM 差异，本测试通过 esbuild 动态转译 xlsxFill.ts 后执行。
 */
const path = require('path');
const fs = require('fs');

const ROOT = path.resolve(__dirname, '..');
const PizZip = require(path.resolve(ROOT, 'node_modules', 'pizzip'));

let pass = 0, fail = 0;
function assert(name, cond) {
  if (cond) { pass++; console.log('  \u2713', name); }
  else { fail++; console.log('  \u2717', name); }
}

// 用 TypeScript 编译器把 xlsxFill.ts 转成 CJS 再 require（不依赖 esbuild）
function loadEngine() {
  const ts = require(path.resolve(ROOT, 'node_modules', 'typescript'));
  const compile = (rel) => {
    let source = fs.readFileSync(path.resolve(ROOT, 'src', rel), 'utf8');
    source = source.replace(/import\s+type\s+[\s\S]*?;/g, '');
    return ts.transpileModule(source, {
      compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2019 },
    }).outputText;
  };
  // 先编译依赖 money，再编译 xlsxFill；用一个 require 解析 './money' 到已编译产物
  const moneyMod = { exports: {} };
  new Function('module', 'exports', 'require', compile('services/money.ts'))(moneyMod, moneyMod.exports, require);

  const m = { exports: {} };
  const req = (id) => {
    if (id === 'pizzip') return PizZip;
    if (id === './money') return moneyMod.exports;
    return require(id);
  };
  new Function('module', 'exports', 'require', compile('services/xlsxFill.ts'))(m, m.exports, req);
  return m.exports;
}

const TEMPLATE = path.resolve(ROOT, 'templates', '迪新货单.xlsx');

function main() {
  const eng = loadEngine();
  const buf = fs.readFileSync(TEMPLATE);
  const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);

  // 读原始共享字符串
  const zip0 = new PizZip(ab);
  const ss = eng.parseSharedStrings(zip0.file('xl/sharedStrings.xml').asText());

  console.log('测试1：占位符提取');
  const ph = eng.extractPlaceholders(ss);
  assert('包含 客户名称', ph.includes('客户名称'));
  assert('包含 品名(循环)', ph.includes('品名'));
  assert('包含 合计金额', ph.includes('合计金额'));

  console.log('测试2：普通占位符填充');
  const data = {
    发货时间: '2026-07-30',
    客户名称: '测试客户<A&B>',
    合同单号: 'HT-001',
    '内部货单编号（自动赋码）': 'DX20260730',
    品名: [
      { 出货名: '苯氧乙醇', 数量: '10', 产品单价: '25', 产品总价: '250', 单位: 'kg' },
      { 出货名: '对羟基苯乙酮', 数量: '5', 产品单价: '40', 产品总价: '200', 单位: 'kg' },
      { 出货名: '乙基己基甘油', 数量: '3', 产品单价: '100', 产品总价: '300', 单位: 'kg' },
    ],
  };
  const blob = eng.fillXlsx(ab, { ...data });
  assert('返回 Blob 且有内容', blob && blob.size > 0);

  // 把 blob 转回 buffer 重新解析
  return blob.arrayBuffer().then((filledAb) => {
    const z = new PizZip(filledAb);
    const sheetName = Object.keys(z.files).find((n) => /worksheets\/sheet\d+\.xml$/.test(n));
    const sheet = z.file(sheetName).asText();

    assert('客户名称已填且 XSS 转义', sheet.includes('测试客户&lt;A&amp;B&gt;'));
    assert('单号已填', sheet.includes('DX20260730'));
    assert('循环标记已清除', !sheet.includes('{#品名}') && !sheet.includes('{/品名}'));

    console.log('测试3：明细循环 3 行');
    assert('出货名1 苯氧乙醇', sheet.includes('苯氧乙醇'));
    assert('出货名2 对羟基苯乙酮', sheet.includes('对羟基苯乙酮'));
    assert('出货名3 乙基己基甘油', sheet.includes('乙基己基甘油'));

    console.log('测试4：自动求和');
    assert('合计数量=18', sheet.includes('>18<') || sheet.includes('18</t>'));
    assert('合计金额=750', sheet.includes('>750<') || sheet.includes('750</t>'));

    console.log('测试5：结构完整（可被 PizZip 重新打开且含关键部件）');
    assert('含 workbook.xml', !!z.file('xl/workbook.xml'));
    assert('含 sheetData', sheet.includes('<sheetData>') && sheet.includes('</sheetData>'));
    assert('无残留占位符', !/\{[^{}]+\}/.test(sheet));

    console.log('测试6：0 条明细不报错');
    const blob0 = eng.fillXlsx(ab, { 客户名称: '空单', 品名: [] });
    assert('空明细返回 Blob', blob0 && blob0.size > 0);

    console.log(`\n结果：${pass} 通过，${fail} 失败`);
    process.exit(fail > 0 ? 1 : 0);
  });
}

main();
