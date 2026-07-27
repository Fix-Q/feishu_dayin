import React, { useEffect, useRef, useState } from 'react';
import { Alert, Button, Card, Empty, Input, Select, Space, Spin, Typography, message } from 'antd';
import { FileTextOutlined, SaveOutlined, ReloadOutlined } from '@ant-design/icons';
import type { TemplateInfo } from '../types';
import type { ActiveRecordState } from '../hooks/useActiveRecord';
import { fetchTemplateBuffer, uploadTemplate } from '../services/templateApi';
import { listTemplateParagraphs, applyParagraphEdits } from '../services/docxText';
import type { ParagraphInfo } from '../services/docxText';

const { Text } = Typography;
const { TextArea } = Input;

interface Props {
  active: ActiveRecordState;
  templates: TemplateInfo[];
  onTemplatesChanged: () => void;
}

export default function TemplateEditorTab({ active, templates, onTemplatesChanged }: Props) {
  const [selected, setSelected] = useState<string | null>(null);
  const [paras, setParas] = useState<ParagraphInfo[]>([]);
  const [texts, setTexts] = useState<string[]>([]);
  const [original, setOriginal] = useState<string[]>([]);
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const bufferRef = useRef<ArrayBuffer | null>(null);

  const tableId = active.tableId;

  // 选中模板后自动加载段落
  useEffect(() => {
    if (!selected || !tableId) {
      setParas([]);
      setTexts([]);
      setOriginal([]);
      bufferRef.current = null;
      return;
    }
    let cancelled = false;
    (async () => {
      setLoading(true);
      try {
        const buf = await fetchTemplateBuffer(tableId, selected);
        if (cancelled) return;
        const list = listTemplateParagraphs(buf);
        bufferRef.current = buf;
        setParas(list);
        setTexts(list.map((p) => p.text));
        setOriginal(list.map((p) => p.text));
      } catch (e: any) {
        if (!cancelled) message.error(e?.message || '加载模板失败');
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selected, tableId]);

  const changedCount = texts.filter((t, i) => t !== original[i]).length;

  const handleSave = async () => {
    if (!selected || !tableId || !bufferRef.current) return;
    const edits = paras
      .map((p, i) => ({ index: p.index, text: texts[i] }))
      .filter((_, i) => texts[i] !== original[i]);
    if (edits.length === 0) {
      message.info('没有改动');
      return;
    }
    setSaving(true);
    try {
      const blob = applyParagraphEdits(bufferRef.current, edits);
      const res = await uploadTemplate(tableId, selected, blob, true);
      if (res.ok) {
        message.success(`已保存 ${selected}（${edits.length} 处改动）`);
        onTemplatesChanged();
      } else {
        message.error(res.error || '保存失败');
      }
    } catch (e: any) {
      message.error(e?.message || '保存失败');
    } finally {
      setSaving(false);
    }
  };

  const handleReset = () => {
    setTexts([...original]);
  };

  const templateOptions = templates.map((t) => ({ label: t.name, value: t.name }));

  return (
    <Space direction="vertical" size={12} style={{ width: '100%' }}>
      <Alert
        type="info"
        showIcon
        message="模板正文编辑"
        description={
          <div style={{ fontSize: 12 }}>
            <p style={{ margin: '2px 0' }}>在此直接修改模板的<b>文字与占位符</b>（如改标签、改 {'{字段}'} 名称、调整分页指令行），保存即写回 .docx，无需重新上传 Word。</p>
            <p style={{ margin: '2px 0' }}>每个文本框对应模板里的一个段落（含表格单元格）。仅“有改动的段落”会被写回，未改的保持不变。注意：被改写的段落内部多段格式会被合并为单一格式。</p>
          </div>
        }
      />

      <Card size="small" title="选择要编辑的模板">
        <Space wrap>
          <Select
            style={{ width: 320 }}
            placeholder={templates.length ? '选择模板' : '暂无模板'}
            showSearch
            allowClear
            value={selected}
            onChange={setSelected}
            options={templateOptions}
            optionFilterProp="label"
          />
          {selected && (
            <>
              <Button icon={<SaveOutlined />} type="primary" loading={saving} onClick={handleSave} disabled={changedCount === 0}>
                保存{changedCount > 0 ? `（${changedCount} 处）` : ''}
              </Button>
              <Button icon={<ReloadOutlined />} onClick={handleReset} disabled={changedCount === 0}>
                还原
              </Button>
            </>
          )}
        </Space>
      </Card>

      {!selected && (
        <Empty description="请选择上方模板开始编辑" />
      )}

      {selected && loading && (
        <div style={{ textAlign: 'center', padding: 32 }}><Spin /></div>
      )}

      {selected && !loading && (
        <Card size="small" title={<span><FileTextOutlined /> 段落（共 {paras.length} 个，可编辑）</span>}>
          <div style={{ maxHeight: '56vh', overflow: 'auto', paddingRight: 4 }}>
            {paras.map((p, i) => {
              const dirty = texts[i] !== original[i];
              return (
                <div key={p.index} style={{ marginBottom: 10 }}>
                  <Text type="secondary" style={{ fontSize: 11 }}>段落 #{p.index}{dirty ? '（已改）' : ''}</Text>
                  <TextArea
                    value={texts[i]}
                    onChange={(e) => {
                      const next = [...texts];
                      next[i] = e.target.value;
                      setTexts(next);
                    }}
                    autoSize={{ minRows: 1, maxRows: 6 }}
                    style={dirty ? { borderColor: '#fa8c16' } : undefined}
                  />
                </div>
              );
            })}
          </div>
        </Card>
      )}
    </Space>
  );
}
