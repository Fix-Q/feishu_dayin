# 飞书多维表格打印插件 · 易用性增强设计

日期：2026-07-23
状态：已确认，待实现

## 背景

插件已实现按数据表隔离的 Word 模板打印（选中记录 → 填充模板 → 预览/打印/下载）。健哥在实际使用中提出一批易用性痛点，核心是"变量面板不好用、找数据源费劲"，并希望增加更强的模板管理、优化 UI。本设计覆盖确认要做的改进项，逐条落地。

## 范围（已与用户逐项确认）

纳入：A1 字段顺序、A2 隐藏字段开关、A3 变量搜索、A5 复制前预加载子表字段、B2 缺失变量提醒、B3 格式化后缀、C1 批量打印（上限 20、各自匹配合并连打）、D2 循环片段防错提示、模板管理强化（排序/批量删除/替换）、UI 清爽微调。

**明确排除**：
- A4 变量行显示当前记录实际值 —— 会导致每次切记录对全字段拉值，字段多时卡顿。
- B1 扩充系统变量 & 打印人 —— 用户数据表本身有时间字段，不需要插件重复造；打印人不显示、不录入。
- 模板分组/标签 —— 按表隔离后单表模板量可控，暂不需要。

**保持不变**：现有系统变量 `{打印日期}` `{打印时间}` `{数据表名称}` `{记录ID}` 原样保留（模板不引用则不出现，零成本）。

## 架构改动概览

引入"视图可见字段顺序"作为字段排序的统一来源，贯穿变量面板与打印取数；其余为独立增量。

### 1. 视图字段顺序（A1/A2）—— 基础改动

`hooks/useActiveRecord.ts`：在加载表元信息时，额外调用 `table.getActiveView()` → `view.getVisibleFieldIdList()`，得到可见字段的有序 id 列表。

- 新增 state 字段：`visibleFieldIds: string[]`（视图可见字段，有序）。
- `fieldMetas` 仍全量返回（隐藏字段打印时仍可用），但派生一个有序视图。
- 提供一个工具：`orderFields(fieldMetas, visibleFieldIds)` → 返回 `{ visible: FieldMetaLite[](按视图序), hidden: FieldMetaLite[] }`。放在 `utils/fieldOrder.ts`。
- 失败兜底：`getActiveView`/`getVisibleFieldIdList` 异常时，退回 `fieldMetas` 原序，全部视为可见（不阻断）。

影响点：
- `VariablePanel`：普通字段、关联字段按 visible 顺序展示；隐藏字段按开关追加显示（A2）。顺序仅影响面板展示。
- `dataBuilder.buildPrintData`：**始终处理全部字段（可见+隐藏）**，数据对象按字段名存储，遍历顺序对填充结果无功能影响。隐藏字段仍写入数据，保证模板引用隐藏字段能正常取值。视图顺序不改变此处行为，只用于面板。
- 关联子表字段顺序：`expandLinkField` 仍读取子表全部字段（保证循环数据完整）；仅 `VariablePanel.loadChildFields` 展示时按子表 `getActiveView().getVisibleFieldIdList()` 排序；失败退回原序。

### 2. 变量面板改造（A1/A2/A3/A5）

`components/VariablePanel.tsx`：
- **A1**：字段按 `visible` 顺序渲染。
- **A2**：顶部加 `Switch`「显示隐藏字段」，默认关。开启后在可见字段下方追加隐藏字段区（灰色标注"隐藏"）。
- **A3**：顶部加搜索框，实时过滤变量名（系统变量、普通字段、关联字段名都参与过滤）。
- **A5**：`copyAll()` 执行前，`await` 预加载所有关联字段的子表字段（并发用 mapLimit），保证复制出的循环片段完整；加载期间按钮显示 loading。

### 3. 缺失变量提醒（B2）

新增 `utils/templateVars.ts`：
- `extractTemplateTags(buffer): string[]` —— 从填充前的 docx 解析出用到的占位符名（含循环标签 `#`/`/`）。用 pizzip 读 `word/document.xml`（含 header/footer parts），正则提取 `{[#/]?([^{}|]+)(\|[^{}]+)?}`，去重、去格式后缀。
- 在 `PrintTab.generate()` 里：填充成功后，用可用变量集合（系统变量 + 当前表字段名 + 关联字段名 + 各关联子表字段名 + `_文本` 变体）比对模板标签，算出"模板用到但不存在"的集合，存入 state，在预览区上方以 `Alert warning` 列出。
- 不阻断打印（缺失变量已被 `nullGetter` 渲染为空）。

### 4. 格式化后缀（B3）

`services/docxFill.ts`：给 Docxtemplater 配置自定义 parser，支持 `{变量|格式}`：
- `money` → `¥1,200.00`（两位小数、千分位、¥ 前缀）
- `num` → `1,200`（千分位）
- `date` → `2026年7月23日`（解析常见日期串/时间戳）
- `pct` → `85%`（0.85→85%，已是百分数字符串则原样）
- 无后缀或未知后缀 → 原样返回（当作普通文本，不报错）。
- parser 取值仍走 `scope.get(name)`，格式化只作用于最终字符串；解析失败静默退回原值。
- 变量面板补充一小段"可用格式"说明。

### 5. 批量打印（C1）

`components/PrintTab.tsx` + 新增 `services/batchPrint.ts`：
- 打印区加「批量打印」按钮（次级）。点击后读取 `view.getSelectedRecordIdList()`。
- 上限 **20 条**：超出提示"一次最多 20 条，请减少勾选或分批"。
- 对每条记录：用其匹配字段值 `matchTemplate` 各自选模板（复用现有逻辑）；未命中的记录跳过并计入"跳过列表"。
- 命中的记录：`fetchTemplateBuffer`（按 tableId+模板名，带缓存避免同模板重复下载）→ `buildPrintData` → `fillTemplate` 得到各自 docx。
- 合并：将各 docx 的 `word/document.xml` body 内容顺序拼接，条目间插入分页符（`<w:p><w:r><w:br w:type="page"/></w:r></w:p>`），媒体/关系需重映射（见风险）。合并实现放 `services/docxMerge.ts`。
- 进度：filling 期间显示「正在生成 x/N」。
- 结果：合并后的 Blob 复用现有预览/打印/下载三通道；顶部提示"成功 M 条，跳过 K 条（无匹配模板）"。

### 6. 循环片段防错提示（D2）

`components/VariablePanel.tsx`：关联字段区顶部加一个静态示意（用简单表格 + 高亮单元格）说明 `{#字段}` 与 `{/字段}` 要放在 Word 表格同一行的首格与末格。纯展示。

### 7. 模板管理强化

`components/TemplateManageTab.tsx` + `services/templateApi.ts` + `server/templateApi.js`：
- **排序**：列表加「名称 / 上传时间」切换（前端排序，`mtime` 已有）。
- **批量删除**：列表项加多选，底部「删除选中」；逐个调用现有 DELETE（或后端加一个批量删除端点，见下）。
- **替换**：每个模板项加「替换」按钮，触发选文件 → 以 `overwrite=1` 上传同名，免去先删再传。
- 后端：新增 `POST /api/templates/batch-delete`（body `{ tableId, names: [] }`），逐个安全删除，返回成功/失败清单。沿用 `safeTableId`/`safeName`。

### 8. UI 清爽微调

保持 antd 5，不引入主题库：
- 统一卡片间距（12）、圆角、留白；三页签结构对齐。
- 打印页分区：当前记录（Alert）/ 模板选择区 / 操作按钮 / 提示 / 预览，视觉分组清晰；预览区加浅边框 + "预览"标题。
- 空状态（无模板、未选记录、未连接）统一图标 + 引导文案（含跳转链接）。
- 操作按钮统一加图标，主次分明（打印=primary，其余 default）。
- 匹配状态 Tag 保留强化（精确/包含/手动/未匹配 配色）。

## 涉及文件

新增：`utils/fieldOrder.ts`、`utils/templateVars.ts`、`services/batchPrint.ts`、`services/docxMerge.ts`
修改：`hooks/useActiveRecord.ts`、`components/VariablePanel.tsx`、`components/PrintTab.tsx`、`components/TemplateManageTab.tsx`、`services/dataBuilder.ts`、`services/docxFill.ts`、`services/templateApi.ts`、`server/templateApi.js`、`types.ts`

## 风险与对策

1. **视图接口失败**：`getActiveView`/`getVisibleFieldIdList` 在某些视图类型（如非表格视图）可能异常 → 一律 try/catch 退回原序 + 全可见，不阻断。
2. **docx 合并复杂度（C1 最大风险）**：多个 docx 合并需处理各自的 `word/media` 图片与 `document.xml.rels` 关系 id 冲突。对策：合并时对每个来源的关系 id、媒体文件名加前缀重命名（如 `d0_rId1`、`media/d0_img1.png`），再拼接 body。若合并保真度不达标，退化方案＝逐条独立打印（循环调用现有单条打印，多次弹打印框），在设计评审后按实测决定。
3. **B3 parser 与循环并存**：格式化 parser 必须正确放行 `#`/`/`/`^` 等 section 语法，只对普通变量套格式，避免破坏关联循环。实现时用 docxtemplater 官方 parser 扩展写法，section 交回默认处理。
4. **批量性能**：20 条 × 关联展开，最坏几十次 bridge 调用。用 mapLimit 限流 + 模板 buffer 缓存 + 进度提示。
5. **缺失变量误报**：header/footer/多 section 里的标签需一并解析，否则漏解析导致误报。extractTemplateTags 覆盖 document/header/footer parts。

## 验证

- 单元级（headless node）：fieldOrder 排序、extractTemplateTags 提取、B3 各 formatter、docxMerge 合并后 `<w:tr>`/分页符/图片数正确。
- 集成（浏览器 + 本地 server）：变量面板顺序与隐藏开关、搜索、复制全部；预览缺失变量黄条；批量选 3 条不同模板合并连打；模板管理排序/批量删/替换。
- 端到端：飞书内经 tunnel 用真实带签名模板复测（尤其 C1 合并后签名图与分页）。
