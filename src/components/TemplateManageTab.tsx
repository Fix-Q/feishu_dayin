import React, { useMemo, useState } from 'react';
import {
  Card, Input, List, Button, Upload, Space, Tag, Popconfirm, Modal, message, Select, Typography, Empty,
} from 'antd';
import { UploadOutlined, DeleteOutlined, CopyOutlined, DownloadOutlined } from '@ant-design/icons';
import type { UploadProps } from 'antd';
import type { TemplateInfo, MatchConfig } from '../types';
import type { ActiveRecordState } from '../hooks/useActiveRecord';
import { uploadTemplate, deleteTemplate, copyTemplate, putConfig, templateDownloadUrl } from '../services/templateApi';

const { Text } = Typography;

interface Props {
  active: ActiveRecordState;
  templates: TemplateInfo[];
  matchConfig: MatchConfig;
  onTemplatesChanged: () => void;
  onConfigChanged: (cfg: MatchConfig) => void;
}

function formatSize(n: number): string {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / 1024 / 1024).toFixed(1)} MB`;
}

export default function TemplateManageTab({
  active, templates, matchConfig, onTemplatesChanged, onConfigChanged,
}: Props) {
  const [keyword, setKeyword] = useState('');
  const [copySource, setCopySource] = useState<string | null>(null);
  const [copyTarget, setCopyTarget] = useState('');
  const [savingConfig, setSavingConfig] = useState(false);

  const filtered = useMemo(() => {
    const k = keyword.trim().toLowerCase();
    if (!k) return templates;
    return templates.filter((t) => t.name.toLowerCase().includes(k));
  }, [templates, keyword]);

  const currentMatchFieldId = active.tableId
    ? matchConfig.tables[active.tableId]?.matchFieldId
    : undefined;

  const tableId = active.tableId;

  // 上传：customRequest 自行调用 API，处理 409 覆盖确认
  const doUpload = async (file: File, overwrite: boolean): Promise<boolean> => {
    if (!tableId) { message.error('未连接到数据表，无法上传'); return false; }
    const buf = await file.arrayBuffer();
    const res = await uploadTemplate(tableId, file.name, buf, overwrite);
    if (res.ok) {
      message.success(`已上传 ${file.name}`);
      onTemplatesChanged();
      return true;
    }
    if (res.conflict) {
      Modal.confirm({
        title: '同名模板已存在',
        content: `模板「${file.name}」已存在，是否覆盖？`,
        okText: '覆盖',
        cancelText: '取消',
        onOk: async () => {
          const r2 = await uploadTemplate(tableId, file.name, buf, true);
          if (r2.ok) {
            message.success(`已覆盖 ${file.name}`);
            onTemplatesChanged();
          } else {
            message.error(r2.error || '覆盖失败');
          }
        },
      });
      return false;
    }
    message.error(res.error || '上传失败');
    return false;
  };

  const uploadProps: UploadProps = {
    accept: '.docx',
    showUploadList: false,
    multiple: true,
    beforeUpload: (file) => {
      if (!tableId) {
        message.error('未连接到数据表，无法上传');
        return Upload.LIST_IGNORE;
      }
      if (!/\.docx$/i.test(file.name)) {
        message.error('仅支持 .docx 文件');
        return Upload.LIST_IGNORE;
      }
      if (file.size > 15 * 1024 * 1024) {
        message.warning('文件较大（>15MB），上传可能较慢');
      }
      doUpload(file, false);
      return Upload.LIST_IGNORE; // 阻止 antd 默认上传
    },
  };

  const handleDelete = async (name: string) => {
    if (!tableId) return;
    try {
      await deleteTemplate(tableId, name);
      message.success(`已删除 ${name}`);
      onTemplatesChanged();
    } catch (e: any) {
      message.error(e?.message || '删除失败');
    }
  };

  const openCopy = (name: string) => {
    setCopySource(name);
    // 默认新名：原名 + " 副本"
    const base = name.replace(/\.docx$/i, '');
    setCopyTarget(`${base} 副本.docx`);
  };

  const handleCopy = async () => {
    if (!copySource) return;
    let target = copyTarget.trim();
    if (!target) { message.error('请输入新模板名称'); return; }
    if (!/\.docx$/i.test(target)) target += '.docx';
    if (!tableId) { message.error('未连接到数据表'); return; }
    const res = await copyTemplate(tableId, copySource, target);
    if (res.ok) {
      message.success(`已复制为 ${target}`);
      setCopySource(null);
      onTemplatesChanged();
    } else if (res.conflict) {
      message.error('目标名称已存在，请换一个');
    } else {
      message.error(res.error || '复制失败');
    }
  };

  const handleSaveMatchField = async (fieldId: string | undefined) => {
    if (!active.tableId) return;
    setSavingConfig(true);
    try {
      const next: MatchConfig = { tables: { ...matchConfig.tables } };
      if (!fieldId) {
        delete next.tables[active.tableId];
      } else {
        const meta = active.fieldMetas.find((f) => f.id === fieldId);
        next.tables[active.tableId] = {
          matchFieldId: fieldId,
          matchFieldName: meta?.name || '',
        };
      }
      const saved = await putConfig(next);
      onConfigChanged(saved);
      message.success('已保存自动匹配字段');
    } catch (e: any) {
      message.error(e?.message || '保存失败');
    } finally {
      setSavingConfig(false);
    }
  };

  return (
    <Space direction="vertical" size={12} style={{ width: '100%' }}>
      <Card size="small" title="自动匹配设置">
        <Text type="secondary" style={{ fontSize: 12 }}>
          选择当前表「{active.tableName || '未连接'}」用于匹配模板的字段。打开插件时，会用该字段的值去匹配同名模板（文件名去掉 .docx）。
        </Text>
        <div style={{ marginTop: 8 }}>
          <Select
            style={{ width: '100%' }}
            placeholder="选择匹配字段（可清空）"
            allowClear
            loading={savingConfig}
            value={currentMatchFieldId}
            onChange={(v) => handleSaveMatchField(v)}
            options={active.fieldMetas.map((f) => ({ label: f.name, value: f.id }))}
            showSearch
            optionFilterProp="label"
          />
        </div>
      </Card>

      <Card
        size="small"
        title={`模板库（${templates.length}）`}
        extra={
          <Upload {...uploadProps}>
            <Button type="primary" icon={<UploadOutlined />} size="small">上传模板</Button>
          </Upload>
        }
      >
        <Input.Search
          placeholder="搜索模板名称"
          allowClear
          value={keyword}
          onChange={(e) => setKeyword(e.target.value)}
          style={{ marginBottom: 12 }}
        />
        {filtered.length === 0 ? (
          <Empty description={keyword ? '无匹配模板' : '还没有模板，点右上角上传'} />
        ) : (
          <List
            size="small"
            dataSource={filtered}
            renderItem={(t) => (
              <List.Item
                actions={[
                  <a key="dl" href={tableId ? templateDownloadUrl(tableId, t.name) : undefined} download title="下载">
                    <DownloadOutlined />
                  </a>,
                  <a key="cp" onClick={() => openCopy(t.name)} title="复制"><CopyOutlined /></a>,
                  <Popconfirm
                    key="del"
                    title={`删除「${t.name}」？`}
                    okText="删除"
                    cancelText="取消"
                    onConfirm={() => handleDelete(t.name)}
                  >
                    <a style={{ color: '#cf1322' }} title="删除"><DeleteOutlined /></a>
                  </Popconfirm>,
                ]}
              >
                <List.Item.Meta
                  title={<Text ellipsis style={{ maxWidth: 220 }}>{t.name}</Text>}
                  description={<Text type="secondary" style={{ fontSize: 12 }}>{formatSize(t.size)}</Text>}
                />
              </List.Item>
            )}
          />
        )}
      </Card>

      <Modal
        open={!!copySource}
        title={`复制模板：${copySource || ''}`}
        okText="复制"
        cancelText="取消"
        onOk={handleCopy}
        onCancel={() => setCopySource(null)}
      >
        <Input
          value={copyTarget}
          onChange={(e) => setCopyTarget(e.target.value)}
          placeholder="新模板名称（.docx）"
          onPressEnter={handleCopy}
        />
      </Modal>
    </Space>
  );
}
