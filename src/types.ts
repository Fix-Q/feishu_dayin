// 共享类型定义

// 模板文件信息（服务端 /api/templates 返回）
export interface TemplateInfo {
  name: string; // 含 .docx 后缀
  size: number; // 字节
  mtime: number; // 毫秒时间戳
}

// 单个数据表的匹配配置
export interface TableMatchConfig {
  matchFieldId: string;
  matchFieldName: string;
}

// 全局匹配配置（/api/config）
export interface MatchConfig {
  tables: Record<string, TableMatchConfig>;
}

// 字段元信息（对 SDK IFieldMeta 的最小约束，含 link 字段的 property.tableId）
export interface FieldMetaLite {
  id: string;
  type: number;
  name: string;
  isPrimary?: boolean;
  property?: {
    tableId?: string; // SingleLink/DuplexLink 的关联子表 ID
    multiple?: boolean;
    [k: string]: unknown;
  } | null;
}

// 关联字段展开后的一条子记录：{ 序号: 1, 子字段名: '值', ... }
export type LinkedRow = Record<string, string | number>;

// docxtemplater 填充数据：普通字段为字符串，关联字段为子记录数组
export type PrintDataValue = string | number | LinkedRow[];

// buildPrintData 的产物
export interface PrintDataResult {
  data: Record<string, PrintDataValue>;
  warnings: string[];
}

// 模板匹配结果
export type MatchKind = 'exact' | 'contains' | 'reverse' | 'none';
export interface MatchResult {
  name: string | null; // 匹配到的模板文件名，无则 null
  kind: MatchKind;
}
