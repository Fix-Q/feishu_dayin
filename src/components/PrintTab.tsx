import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Alert, Button, Card, Segmented, Select, Space, Spin, Switch, Tag, Tooltip, Typography, message } from 'antd';
import { PrinterOutlined, DownloadOutlined, ReloadOutlined, FileWordOutlined, RotateRightOutlined } from '@ant-design/icons';
import { saveAs } from 'file-saver';

import type { TemplateInfo, MatchConfig, MatchKind } from '../types';
import type { ActiveRecordState } from '../hooks/useActiveRecord';
import { fetchTemplateBuffer } from '../services/templateApi';
import { buildPrintData } from '../services/dataBuilder';
import { fillTemplate, explainDocxError } from '../services/docxFill';
import { fillXlsx, isXlsxName } from '../services/xlsxFill';
import { matchTemplate } from '../services/templateMatch';
import { printDocxBlob, printHtmlTable, printCopies, type PrintOrientation } from '../utils/print';
import { renderXlsxToHtml } from '../services/xlsxRender';
import DocxPreview from './DocxPreview';
import XlsxPreview from './XlsxPreview';

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
  const [orientation, setOrientation] = useState<PrintOrientation>('auto');
  // 多联打印：默认关闭（单份）；开启后按联名列表连打
  const [multiCopy, setMultiCopy] = useState(false);
  const DEFAULT_COPIES = ['生产部', '销售部', '客户', '财务部', '开票'];
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
      const blob = isXlsxName(selected) ? fillXlsx(buffer, data) : fillTemplate(buffer, data);
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

  const isXlsx = !!selected && isXlsxName(selected);

  const handlePrint = async () => {
    const blob = previewBlob || (await generate());
    if (!blob) return;
    try {
      if (multiCopy) {
        const html = isXlsx ? await renderXlsxToHtml(blob) : undefined;
        await printCopies({
          copies: DEFAULT_COPIES,
          orientation,
          docxBlob: isXlsx ? undefined : blob,
          htmlTable: html,
        });
      } else if (isXlsx) {
        const html = await renderXlsxToHtml(blob);
        await printHtmlTable(html, orientation);
      } else {
        await printDocxBlob(blob, orientation);
      }
    } catch (e: any) {
      message.error(e?.message || '打印失败');
    }
  };

  const handleDownload = async () => {
    const blob = previewBlob || (await generate());
    if (!blob) return;
    const ext = isXlsx ? '.xlsx' : '.docx';
    const base = selected ? selected.replace(/\.(docx|xlsx)$/i, '') : '打印';
    const suffix = active.primaryText ? `-${active.primaryText}` : '';
    saveAs(blob, `${base}${suffix}${ext}`);
  };

  const templateOptions = useMemo(
    () => templates.map((t) => ({ label: t.name, value: t.name })),
    [templates]
  );

  const noRecord = !active.recordId;

  return (
    <Space direction="vertical" size={16} style={{ width: '100%' }}>
      {active.error && <Alert type="error" showIcon message={`连接多维表格出错：${active.error}`} />}

      <Alert
        type={noRecord ? 'warning' : 'success'}
        showIcon
        message={
          noRecord
            ? '请在左侧表格中选中一条记录'
            : <span>当前记录：<Text strong>{active.primaryText || '(空)'}</Text>　·　表：{active.tableName}</span>
        }
      />

      <Card size="small" title="选择模板" styles={{ body: { paddingBottom: 12 } }}>
        <Space wrap>
          <Select
            style={{ width: 280 }}
            placeholder={templates.length ? '选择模板' : '暂无模板，请先在“模板管理”上传'}
            showSearch
            allowClear
            value={selected}
            onChange={(v) => { setSelected(v ?? null); setManual(true); }}
            options={templateOptions}
            optionFilterProp="label"
            notFoundContent={<a onClick={goManage}>去模板管理上传</a>}
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
          <div style={{ marginTop: 8 }}>
            <Text type="secondary" style={{ fontSize: 12 }}>
              未设置自动匹配字段，<a onClick={goManage}>去设置</a>后可按记录自动选模板。
            </Text>
          </div>
        )}
        <div style={{ marginTop: 14 }}>
          <div style={{ marginBottom: 10 }}>
            <Space size={8} align="center">
              <Tooltip title="五联货单等横向内容，若 Word 是竖版排版导致打印被裁切，选「横向」会自动旋转 90° 并切换横版纸张">
                <Text type="secondary" style={{ fontSize: 12 }}>
                  <RotateRightOutlined /> 打印方向
                </Text>
              </Tooltip>
              <Segmented
                size="small"
                value={orientation}
                onChange={(v) => setOrientation(v as PrintOrientation)}
                options={[
                  { label: '跟随文档', value: 'auto' },
                  { label: '竖向', value: 'portrait' },
                  { label: '横向', value: 'landscape' },
                ]}
              />
              <Tooltip title="开启后一次连打 5 份（生产部/销售部/客户/财务部/开票），每份右上角标注联名">
                <Space size={4} align="center">
                  <Switch size="small" checked={multiCopy} onChange={setMultiCopy} />
                  <Text type="secondary" style={{ fontSize: 12 }}>多联(5联)</Text>
                </Space>
              </Tooltip>
            </Space>
          </div>
        </div>
      </Card>

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

      <Card
        size="small"
        title={<span><FileWordOutlined /> 打印预览</span>}
        styles={{ body: { padding: 12 } }}
      >
        <Spin spinning={rendering} tip="正在生成预览…">
          {isXlsx
            ? <XlsxPreview blob={previewBlob} onError={(m) => setErrors([m])} />
            : <DocxPreview blob={previewBlob} orientation={orientation} onError={(m) => setErrors([m])} />}
        </Spin>
      </Card>

      {/* 底部固定操作栏：打印/下载常驻，不必滚到底 */}
      <div
        style={{
          position: 'sticky', bottom: 0, margin: '4px -16px -16px',
          padding: '10px 16px', background: '#fff',
          borderTop: '1px solid #eef0f3',
          display: 'flex', gap: 8, alignItems: 'center',
          boxShadow: '0 -2px 8px rgba(0,0,0,0.03)',
        }}
      >
        <Button type="primary" icon={<PrinterOutlined />} onClick={handlePrint} disabled={!selected || noRecord}>
          {multiCopy ? '打印(5联)' : '打印'}
        </Button>
        <Button icon={<DownloadOutlined />} onClick={handleDownload} disabled={!selected || noRecord}>
          {isXlsx ? '下载 Excel' : '下载 Word'}
        </Button>
        <Button type="text" icon={<ReloadOutlined />} onClick={generate} disabled={!selected || noRecord} style={{ marginLeft: 'auto' }}>
          刷新预览
        </Button>
      </div>
    </Space>
  );
}
