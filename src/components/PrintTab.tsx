import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Alert, Button, Select, Space, Spin, Tag, Typography, message } from 'antd';
import { PrinterOutlined, DownloadOutlined, ReloadOutlined } from '@ant-design/icons';
import { saveAs } from 'file-saver';

import type { TemplateInfo, MatchConfig, MatchKind } from '../types';
import type { ActiveRecordState } from '../hooks/useActiveRecord';
import { fetchTemplateBuffer } from '../services/templateApi';
import { buildPrintData } from '../services/dataBuilder';
import { fillTemplate, explainDocxError } from '../services/docxFill';
import { matchTemplate } from '../services/templateMatch';
import { printDocxBlob } from '../utils/print';
import DocxPreview from './DocxPreview';

const { Text } = Typography;

interface Props {
  active: ActiveRecordState;
  templates: TemplateInfo[];
  matchConfig: MatchConfig;
  onNeedTemplates: () => void;
  goManage: () => void;
}

const KIND_LABEL: Record<MatchKind, string> = {
  exact: '精确匹配',
  contains: '包含匹配',
  reverse: '包含匹配',
  none: '',
};

export default function PrintTab({ active, templates, matchConfig, onNeedTemplates, goManage }: Props) {
  const [selected, setSelected] = useState<string | null>(null);
  const [manual, setManual] = useState(false);
  const [matchKind, setMatchKind] = useState<MatchKind>('none');
  const [previewBlob, setPreviewBlob] = useState<Blob | null>(null);
  const [rendering, setRendering] = useState(false);
  const [errors, setErrors] = useState<string[]>([]);
  const [warnings, setWarnings] = useState<string[]>([]);
  const debounceRef = useRef<any>(null);

  const matchFieldId = active.tableId ? matchConfig.tables[active.tableId]?.matchFieldId : undefined;

  // 自动匹配：读取匹配字段值 → 选模板。切记录时重置手动标记。
  const runAutoMatch = useCallback(async () => {
    if (!active.table || !active.recordId || !matchFieldId) {
      setMatchKind('none');
      return;
    }
    try {
      const value = await active.table.getCellString(matchFieldId, active.recordId);
      const res = matchTemplate(value, templates);
      setMatchKind(res.kind);
      if (res.name) setSelected(res.name);
    } catch (e) {
      setMatchKind('none');
    }
  }, [active.table, active.recordId, matchFieldId, templates]);

  // 切表后模板列表已换：若当前选中的模板不在新表模板里，清空选择，避免取到不存在的文件
  useEffect(() => {
    if (selected && !templates.some((t) => t.name === selected)) {
      setSelected(null);
      setPreviewBlob(null);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [templates]);

  // 记录变化 → 重置手动标记并触发自动匹配（debounce 300ms）
  useEffect(() => {
    setManual(false);
    if (debounceRef.current) clearTimeout(debounceRef.current);
    debounceRef.current = setTimeout(() => {
      runAutoMatch();
    }, 300);
    return () => {
      if (debounceRef.current) clearTimeout(debounceRef.current);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [active.recordId, matchFieldId, templates]);

  // 生成预览：取模板 → 构建数据 → 填充 → 出 Blob
  const generate = useCallback(async (): Promise<Blob | null> => {
    if (!selected) { message.warning('请先选择模板'); return null; }
    if (!active.table || !active.recordId) { message.warning('请先在表格中选中一条记录'); return null; }
    if (!active.tableId) { message.warning('未获取到当前数据表'); return null; }
    setRendering(true);
    setErrors([]);
    setWarnings([]);
    try {
      const buffer = await fetchTemplateBuffer(active.tableId, selected);
      const { data, warnings: w } = await buildPrintData(
        active.table, active.tableName, active.fieldMetas, active.recordId
      );
      setWarnings(w);
      const blob = fillTemplate(buffer, data);
      setPreviewBlob(blob);
      return blob;
    } catch (e: any) {
      setPreviewBlob(null);
      setErrors(explainDocxError(e));
      return null;
    } finally {
      setRendering(false);
    }
  }, [selected, active.tableId, active.table, active.recordId, active.tableName, active.fieldMetas]);

  // 选中模板或记录变化后自动刷新预览
  useEffect(() => {
    if (selected && active.recordId) {
      generate();
    } else {
      setPreviewBlob(null);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selected, active.recordId]);

  const handlePrint = async () => {
    const blob = previewBlob || (await generate());
    if (!blob) return;
    try {
      await printDocxBlob(blob);
    } catch (e: any) {
      message.error(e?.message || '打印失败');
    }
  };

  const handleDownload = async () => {
    const blob = previewBlob || (await generate());
    if (!blob) return;
    const base = selected ? selected.replace(/\.docx$/i, '') : '打印';
    const suffix = active.primaryText ? `-${active.primaryText}` : '';
    saveAs(blob, `${base}${suffix}.docx`);
  };

  const templateOptions = useMemo(
    () => templates.map((t) => ({ label: t.name, value: t.name })),
    [templates]
  );

  const noRecord = !active.recordId;

  return (
    <Space direction="vertical" size={12} style={{ width: '100%' }}>
      {active.error && <Alert type="error" showIcon message={`连接多维表格出错：${active.error}`} />}

      <Alert
        type={noRecord ? 'warning' : 'info'}
        showIcon
        message={
          noRecord
            ? '请在左侧表格中选中一条记录'
            : <span>当前记录：<Text strong>{active.primaryText || '(空)'}</Text>　表：{active.tableName}</span>
        }
      />

      <div>
        <Space wrap>
          <Select
            style={{ width: 260 }}
            placeholder={templates.length ? '选择模板' : '暂无模板，请先在“模板管理”上传'}
            showSearch
            allowClear
            value={selected}
            onChange={(v) => { setSelected(v ?? null); setManual(true); }}
            options={templateOptions}
            optionFilterProp="label"
            notFoundContent={
              <a onClick={goManage}>去模板管理上传</a>
            }
          />
          {selected && !manual && matchKind !== 'none' && (
            <Tag color="green">{KIND_LABEL[matchKind]}</Tag>
          )}
          {selected && manual && <Tag color="blue">手动选择</Tag>}
          {selected && !manual && matchKind === 'none' && matchFieldId && (
            <Tag>未自动匹配</Tag>
          )}
        </Space>
        {!matchFieldId && (
          <div style={{ marginTop: 6 }}>
            <Text type="secondary" style={{ fontSize: 12 }}>
              未设置自动匹配字段，<a onClick={goManage}>去设置</a>后可按记录自动选模板。
            </Text>
          </div>
        )}
      </div>

      <Space>
        <Button type="primary" icon={<PrinterOutlined />} onClick={handlePrint} disabled={!selected || noRecord}>
          打印
        </Button>
        <Button icon={<DownloadOutlined />} onClick={handleDownload} disabled={!selected || noRecord}>
          下载 Word
        </Button>
        <Button icon={<ReloadOutlined />} onClick={generate} disabled={!selected || noRecord}>
          刷新预览
        </Button>
      </Space>

      {errors.length > 0 && (
        <Alert
          type="error"
          showIcon
          message="模板填充失败"
          description={<ul style={{ margin: 0, paddingLeft: 18 }}>{errors.map((e, i) => <li key={i}>{e}</li>)}</ul>}
        />
      )}
      {warnings.length > 0 && (
        <Alert
          type="warning"
          showIcon
          message="数据提示"
          description={<ul style={{ margin: 0, paddingLeft: 18 }}>{warnings.map((w, i) => <li key={i}>{w}</li>)}</ul>}
        />
      )}

      <Spin spinning={rendering} tip="正在生成预览…">
        <DocxPreview blob={previewBlob} onError={(m) => setErrors([m])} />
      </Spin>
    </Space>
  );
}
