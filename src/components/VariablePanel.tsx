import React, { useMemo, useState } from 'react';
import { Alert, Button, Card, Collapse, Space, Tag, Typography, message } from 'antd';
import { CopyOutlined } from '@ant-design/icons';
import { bitable, FieldType } from '@lark-base-open/js-sdk';
import type { FieldMetaLite } from '../types';
import type { ActiveRecordState } from '../hooks/useActiveRecord';

const { Text, Paragraph } = Typography;

const LINK_TYPES = [FieldType.SingleLink, FieldType.DuplexLink];

const SYSTEM_VARS = ['打印日期', '打印时间', '数据表名称', '记录ID'];

interface Props {
  active: ActiveRecordState;
}

async function copyText(text: string) {
  try {
    await navigator.clipboard.writeText(text);
    message.success('已复制');
  } catch (e) {
    // 退化方案
    const ta = document.createElement('textarea');
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand('copy'); message.success('已复制'); }
    catch (e2) { message.error('复制失败，请手动选择'); }
    document.body.removeChild(ta);
  }
}

// 一行可复制的变量
function VarRow({ tag }: { tag: string }) {
  return (
    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '2px 0' }}>
      <Text code style={{ fontSize: 13 }}>{tag}</Text>
      <Button size="small" type="text" icon={<CopyOutlined />} onClick={() => copyText(tag)} />
    </div>
  );
}

export default function VariablePanel({ active }: Props) {
  // 关联字段的子表字段缓存：{ [fieldId]: 子字段名[] }
  const [childFields, setChildFields] = useState<Record<string, string[]>>({});
  const [loadingChild, setLoadingChild] = useState<Record<string, boolean>>({});

  const normalFields = useMemo(
    () => active.fieldMetas.filter((f) => LINK_TYPES.indexOf(f.type) === -1),
    [active.fieldMetas]
  );
  const linkFields = useMemo(
    () => active.fieldMetas.filter((f) => LINK_TYPES.indexOf(f.type) !== -1),
    [active.fieldMetas]
  );

  const loadChildFields = async (meta: FieldMetaLite) => {
    if (childFields[meta.id] || loadingChild[meta.id]) return;
    const tableId = meta.property?.tableId;
    if (!tableId) return;
    setLoadingChild((s) => ({ ...s, [meta.id]: true }));
    try {
      const childTable = await bitable.base.getTableById(tableId);
      const metas = (await childTable.getFieldMetaList()) as unknown as FieldMetaLite[];
      setChildFields((s) => ({ ...s, [meta.id]: metas.map((m) => m.name) }));
    } catch (e) {
      setChildFields((s) => ({ ...s, [meta.id]: [] }));
    } finally {
      setLoadingChild((s) => ({ ...s, [meta.id]: false }));
    }
  };

  // 生成关联字段的循环片段
  const loopSnippet = (fieldName: string, subFields: string[]): string => {
    const cols = ['序号', ...subFields];
    const cells = cols.map((c) => `{${c}}`).join('\t');
    return `{#${fieldName}}${cells}{/${fieldName}}`;
  };

  // 复制全部变量（多行文本，粘进 Word 即为起始模板）
  const copyAll = () => {
    const lines: string[] = [];
    lines.push('=== 系统变量 ===');
    SYSTEM_VARS.forEach((v) => lines.push(`{${v}}`));
    lines.push('');
    lines.push('=== 字段变量 ===');
    normalFields.forEach((f) => lines.push(`{${f.name}}`));
    linkFields.forEach((f) => {
      lines.push('');
      lines.push(`=== 关联字段「${f.name}」循环（放进表格同一行）===`);
      const subs = childFields[f.id] || ['(展开该字段以加载子字段)'];
      lines.push(loopSnippet(f.name, subs));
      lines.push(`整段文本：{${f.name}_文本}`);
    });
    copyText(lines.join('\n'));
  };

  const collapseItems = linkFields.map((f) => ({
    key: f.id,
    label: <span>{f.name} <Tag color="purple">关联</Tag></span>,
    children: (
      <div>
        <Text type="secondary" style={{ fontSize: 12 }}>
          循环片段（把整段放进 Word 表格的一行里，每条关联记录会复制成一行）：
        </Text>
        <Paragraph
          code
          copyable={{ text: loopSnippet(f.name, childFields[f.id] || []) }}
          style={{ fontSize: 12, marginTop: 6, whiteSpace: 'pre-wrap' }}
        >
          {loopSnippet(f.name, childFields[f.id] || [])}
        </Paragraph>
        <Text type="secondary" style={{ fontSize: 12 }}>子表字段：</Text>
        {(childFields[f.id] || []).length === 0 ? (
          <div><Text type="secondary" style={{ fontSize: 12 }}>{loadingChild[f.id] ? '加载中…' : '无'}</Text></div>
        ) : (
          (childFields[f.id] || []).map((name) => <VarRow key={name} tag={`{${name}}`} />)
        )}
        <VarRow tag={`{${f.name}_文本}`} />
      </div>
    ),
  }));

  return (
    <Space direction="vertical" size={12} style={{ width: '100%' }}>
      <Alert
        type="info"
        showIcon
        message="如何制作模板"
        description={
          <div style={{ fontSize: 12 }}>
            <p style={{ margin: '2px 0' }}>1. 在 Word 里排好版式，把下面的变量粘贴到需要填数据的位置。</p>
            <p style={{ margin: '2px 0' }}>2. 关联字段用循环片段 <Text code>{'{#字段}…{/字段}'}</Text>，开始和结束标签要放在 Word 表格的同一行，每条关联记录会自动复制成一行。</p>
            <p style={{ margin: '2px 0' }}>3. 保存为 .docx，到「模板管理」上传。文件名（去掉 .docx）等于匹配字段的值时会自动匹配。</p>
            <p style={{ margin: '2px 0' }}>4. 附件字段只输出文件名，不插入图片。字段与系统变量重名时以真实字段为准。</p>
          </div>
        }
      />

      <Button icon={<CopyOutlined />} onClick={copyAll} block>复制全部变量</Button>

      <Card size="small" title="系统变量">
        {SYSTEM_VARS.map((v) => <VarRow key={v} tag={`{${v}}`} />)}
      </Card>

      <Card size="small" title={`字段变量（${normalFields.length}）`}>
        {normalFields.length === 0 ? (
          <Text type="secondary">当前表无普通字段</Text>
        ) : (
          normalFields.map((f) => <VarRow key={f.id} tag={`{${f.name}}`} />)
        )}
      </Card>

      {linkFields.length > 0 && (
        <Card size="small" title={`关联字段（${linkFields.length}）`}>
          <Collapse
            items={collapseItems}
            onChange={(keys) => {
              const arr = Array.isArray(keys) ? keys : [keys];
              arr.forEach((k) => {
                const meta = linkFields.find((f) => f.id === k);
                if (meta) loadChildFields(meta);
              });
            }}
          />
        </Card>
      )}
    </Space>
  );
}
