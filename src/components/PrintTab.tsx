import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Alert, Button, Select, Slider, Space, Spin, Tag, Tooltip, Typography, message } from 'antd';
import {
  PrinterOutlined, DownloadOutlined, ReloadOutlined, RotateRightOutlined,
  ZoomInOutlined, ZoomOutOutlined, ColumnWidthOutlined,
} from '@ant-design/icons';
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
import DocxPreview, {
  MIN_SCALE, MAX_SCALE, SCALE_STEP, clampScale, type PreviewHandle,
} from './DocxPreview';
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

const ORIENTATION_NEXT: Record<PrintOrientation, PrintOrientation> = {
  auto: 'portrait',
  portrait: 'landscape',
  landscape: 'auto',
};

const ORIENTATION_LABEL: Record<PrintOrientation, string> = {
  auto: '跟随',
  portrait: '竖向',
  landscape: '横向',
};

export default function PrintTab({ active, templates, matchConfig, onNeedTemplates, goManage }: Props) {
  const [selected, setSelected] = useState<string | null>(null);
  const [manual, setManual] = useState(false);
  const [matchKind, setMatchKind] = useState<MatchKind>('none');
  const [previewBlob, setPreviewBlob] = useState<Blob | null>(null);
  const [orientation, setOrientation] = useState<PrintOrientation>('auto');
  const [multiCopy, setMultiCopy] = useState(false);
  const DEFAULT_COPIES = ['生产部', '销售部', '客户', '财务部', '开票'];
  const [rendering, setRendering] = useState(false);
  const [errors, setErrors] = useState<string[]>([]);
  const [warnings, setWarnings] = useState<string[]>([]);
  const [scale, setScale] = useState(1);
  const previewRef = useRef<PreviewHandle>(null);
  const debounceRef = useRef<any>(null);

  const handleScaleChange = useCallback((s: number) => setScale(clampScale(s)), []);
  const zoomIn = () => setScale((s) => clampScale(s + SCALE_STEP));
  const zoomOut = () => setScale((s) => clampScale(s - SCALE_STEP));
  const fitWidth = () => previewRef.current?.fitWidth();

  const matchFieldId = active.tableId ? matchConfig.tables[active.tableId]?.matchFieldId : undefined;

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

  useEffect(() => {
    if (selected && !templates.some((t) => t.name === selected)) {
      setSelected(null);
      setPreviewBlob(null);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [templates]);

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
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
      {/* 可滚动内容区 */}
      <div style={{ flex: 1, overflow: 'auto', padding: 12, display: 'flex', flexDirection: 'column', gap: 12 }}>
        {active.error && (
          <Alert type="error" showIcon message={`连接多维表格出错：${active.error}`} style={{ flexShrink: 0 }} />
        )}
        {noRecord && (
          <Alert type="warning" showIcon message="请在左侧表格中选中一条记录" style={{ flexShrink: 0 }} />
        )}

        {/* 模板选择行 */}
        <div style={{ background: '#fff', borderRadius: 10, boxShadow: '0 1px 4px rgba(31,35,41,.06)', padding: 12, display: 'flex', alignItems: 'center', gap: 8, flexShrink: 0 }}>
          <Select
            style={{ flex: 1 }}
            placeholder={templates.length ? '选择模板' : '暂无模板，请先在"模板管理"上传'}
            showSearch
            allowClear
            value={selected}
            onChange={(v) => { setSelected(v ?? null); setManual(true); }}
            options={templateOptions}
            optionFilterProp="label"
            notFoundContent={<a onClick={goManage}>去模板管理上传</a>}
          />
          {selected && !manual && matchKind !== 'none' && (
            <Tag color="green" style={{ margin: 0 }}>{KIND_LABEL[matchKind]}</Tag>
          )}
          {selected && manual && <Tag color="blue" style={{ margin: 0 }}>手动选择</Tag>}
          {selected && !manual && matchKind === 'none' && matchFieldId && (
            <Tag style={{ margin: 0 }}>未自动匹配</Tag>
          )}
        </div>
        {!matchFieldId && active.recordId && (
          <Text type="secondary" style={{ fontSize: 12, lineHeight: 1.5 }}>
            未设置自动匹配字段，<a onClick={goManage}>去设置</a>后可按记录自动选模板。
          </Text>
        )}

        {/* 错误 / 警告 */}
        {errors.length > 0 && (
          <Alert
            type="error"
            showIcon
            message="模板填充失败"
            style={{ flexShrink: 0 }}
            description={
              <ul style={{ margin: 0, paddingLeft: 18 }}>
                {errors.map((e, i) => <li key={i}>{e}</li>)}
              </ul>
            }
          />
        )}
        {warnings.length > 0 && (
          <Alert
            type="warning"
            showIcon
            message="数据提示"
            style={{ flexShrink: 0 }}
            description={
              <ul style={{ margin: 0, paddingLeft: 18 }}>
                {warnings.map((w, i) => <li key={i}>{w}</li>)}
              </ul>
            }
          />
        )}

        {/* 预览卡 */}
        <div style={{ background: '#fff', borderRadius: 10, boxShadow: '0 1px 4px rgba(31,35,41,.06)', flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden', minHeight: 200 }}>
          {/* Toolbar */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '8px 12px', borderBottom: '1px solid #e5e6eb', background: '#fafbfc', flexShrink: 0 }}>
            <span style={{ fontSize: 13, fontWeight: 600, color: '#1f2329', marginRight: 'auto' }}>打印预览</span>

            {/* 缩放控件：放在预览面板外，不会被 zoom 影响 */}
            <Space size={2}>
              <Tooltip title="缩小 (Ctrl+滚轮)">
                <Button size="small" icon={<ZoomOutOutlined />} onClick={zoomOut} disabled={!previewBlob || scale <= MIN_SCALE} />
              </Tooltip>
              <Tooltip title="放大 (Ctrl+滚轮)">
                <Button size="small" icon={<ZoomInOutlined />} onClick={zoomIn} disabled={!previewBlob || scale >= MAX_SCALE} />
              </Tooltip>
              <Tooltip title="适应宽度">
                <Button size="small" icon={<ColumnWidthOutlined />} onClick={fitWidth} disabled={!previewBlob} />
              </Tooltip>
            </Space>
            <div style={{ width: 90 }}>
              <Slider
                min={MIN_SCALE} max={MAX_SCALE} step={SCALE_STEP}
                value={scale} onChange={(v) => setScale(clampScale(v as number))}
                disabled={!previewBlob} tooltip={{ open: false }}
              />
            </div>
            <span style={{ fontSize: 12, color: '#6b7280', minWidth: 38, textAlign: 'right' }}>{Math.round(scale * 100)}%</span>

            <div style={{ width: 1, height: 16, background: '#e5e6eb', margin: '0 4px' }} />

            <Tooltip title="刷新预览">
              <Button
                size="small"
                icon={<ReloadOutlined />}
                onClick={generate}
                disabled={!selected || noRecord}
              />
            </Tooltip>

            <div style={{ width: 1, height: 16, background: '#e5e6eb', margin: '0 4px' }} />

            <Tooltip title="打印方向：五联货单等横向内容，若 Word 是竖版排版导致打印被裁切，选「横向」会自动旋转 90°">
              <Button
                size="small"
                icon={<RotateRightOutlined />}
                onClick={() => setOrientation((o) => ORIENTATION_NEXT[o])}
              >
                {ORIENTATION_LABEL[orientation]}
              </Button>
            </Tooltip>

            <Tooltip title={multiCopy ? '一次连打 5 份（生产部/销售部/客户/财务部/开票）' : '单份打印'}>
              <Button
                size="small"
                type={multiCopy ? 'primary' : 'default'}
                onClick={() => setMultiCopy((v) => !v)}
              >
                {multiCopy ? '5联' : '单联'}
              </Button>
            </Tooltip>
          </div>

          {/* Preview body */}
          <div style={{ flex: 1, background: '#eceef1', overflow: 'hidden', padding: 16, display: 'flex' }}>
            <Spin spinning={rendering} tip="正在生成预览…" wrapperClassName="preview-spin" style={{ width: '100%' }}>
              {isXlsx
                ? <XlsxPreview ref={previewRef} blob={previewBlob} scale={scale} onScaleChange={handleScaleChange} onError={(m) => setErrors([m])} />
                : <DocxPreview ref={previewRef} blob={previewBlob} orientation={orientation} scale={scale} onScaleChange={handleScaleChange} onError={(m) => setErrors([m])} />}
            </Spin>
          </div>
        </div>
      </div>

      {/* 底部固定操作栏 */}
      <div style={{ background: '#fff', borderTop: '1px solid #e5e6eb', padding: '10px 16px', display: 'flex', gap: 10, flexShrink: 0 }}>
        <Button
          type="primary"
          style={{ flex: 1, height: 38, borderRadius: 8, fontSize: 14, fontWeight: 500 }}
          onClick={handlePrint}
          disabled={!selected || noRecord}
        >
          <PrinterOutlined /> {multiCopy ? '打印(5联)' : '打印'}
        </Button>
        <Button
          style={{ flex: 1, height: 38, borderRadius: 8, border: '1px solid #e5e6eb', fontSize: 14, fontWeight: 500 }}
          onClick={handleDownload}
          disabled={!selected || noRecord}
        >
          <DownloadOutlined /> 下载
        </Button>
        <Button
          style={{ width: 38, height: 38, borderRadius: 8, border: '1px solid #e5e6eb', padding: 0 }}
        >
          ⋯
        </Button>
      </div>
    </div>
  );
}
