import React, { useCallback, useEffect, useState } from 'react';
import { ConfigProvider, Tabs, Spin, Alert, theme } from 'antd';
import { PrinterOutlined, FolderOpenOutlined, TagsOutlined, FileTextOutlined } from '@ant-design/icons';
import zhCN from 'antd/locale/zh_CN';
import 'antd/dist/reset.css';

import type { TemplateInfo, MatchConfig } from './types';
import { listTemplates, getConfig } from './services/templateApi';
import { useActiveRecord } from './hooks/useActiveRecord';
import PrintTab from './components/PrintTab';
import TemplateManageTab from './components/TemplateManageTab';
import TemplateEditorTab from './components/TemplateEditorTab';
import VariablePanel from './components/VariablePanel';

// 主题：飞书风格。蓝色主色、克制圆角、紧凑控件、浅色卡片，贴近多维表格原生观感。
const THEME = {
  token: {
    colorPrimary: '#3370FF',      // 飞书蓝
    colorInfo: '#3370FF',
    borderRadius: 6,
    colorBgLayout: '#f7f8fa',
    colorBorderSecondary: '#eef0f3',
    fontSize: 13,
  },
  components: {
    Card: { headerFontSize: 14, paddingLG: 12 },
    Tabs: { horizontalItemPadding: '8px 4px', titleFontSize: 14 },
    Segmented: { itemSelectedBg: '#3370FF', itemSelectedColor: '#fff' },
  },
};

export default function App() {
  const active = useActiveRecord();
  const [templates, setTemplates] = useState<TemplateInfo[]>([]);
  const [matchConfig, setMatchConfig] = useState<MatchConfig>({ tables: {} });
  const [loadError, setLoadError] = useState<string | null>(null);
  const [activeKey, setActiveKey] = useState('print');

  // 刷新当前表的模板列表（上传/删除/复制后、切表时调用）
  const refreshTemplates = useCallback(async () => {
    if (!active.tableId) {
      setTemplates([]);
      return;
    }
    try {
      const list = await listTemplates(active.tableId);
      setTemplates(list);
      setLoadError(null);
    } catch (e: any) {
      setLoadError('无法连接本地模板服务：' + (e?.message || e));
    }
  }, [active.tableId]);

  const refreshConfig = useCallback(async () => {
    try {
      const cfg = await getConfig();
      setMatchConfig(cfg);
    } catch (e) {
      // 配置读取失败不阻塞界面，保留空配置
    }
  }, []);

  // 切表时（tableId 变化）自动重新拉取该表的模板，实现按表隔离
  useEffect(() => {
    refreshTemplates();
  }, [refreshTemplates]);

  useEffect(() => {
    refreshConfig();
  }, [refreshConfig]);

  const items = [
    {
      key: 'print',
      label: (<span><PrinterOutlined /> 打印</span>),
      children: (
        <PrintTab
          active={active}
          templates={templates}
          matchConfig={matchConfig}
          onNeedTemplates={refreshTemplates}
          goManage={() => setActiveKey('manage')}
        />
      ),
    },
    {
      key: 'manage',
      label: (<span><FolderOpenOutlined /> 模板管理</span>),
      children: (
        <TemplateManageTab
          active={active}
          templates={templates}
          matchConfig={matchConfig}
          onTemplatesChanged={refreshTemplates}
          onConfigChanged={(cfg) => setMatchConfig(cfg)}
        />
      ),
    },
    {
      key: 'edit',
      label: (<span><FileTextOutlined /> 模板编辑</span>),
      children: (
        <TemplateEditorTab
          active={active}
          templates={templates}
          onTemplatesChanged={refreshTemplates}
        />
      ),
    },
    {
      key: 'vars',
      label: (<span><TagsOutlined /> 变量参考</span>),
      children: <VariablePanel active={active} />,
    },
  ];

  return (
    <ConfigProvider locale={zhCN} theme={THEME}>
      <div
        style={{
          height: '100vh',
          boxSizing: 'border-box',
          display: 'flex',
          flexDirection: 'column',
          background: THEME.token.colorBgLayout,
        }}
      >
        {/* 顶部品牌条 */}
        <div
          style={{
            display: 'flex', alignItems: 'center', gap: 8,
            padding: '10px 16px', background: '#fff',
            borderBottom: '1px solid #eef0f3', flexShrink: 0,
          }}
        >
          <span
            style={{
              display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
              width: 24, height: 24, borderRadius: 6, background: '#3370FF', color: '#fff',
            }}
          >
            <PrinterOutlined style={{ fontSize: 14 }} />
          </span>
          <span style={{ fontWeight: 600, fontSize: 14, color: '#1f2329' }}>单据打印</span>
          <span style={{ marginLeft: 'auto', fontSize: 12, color: '#8f959e' }}>
            {active.tableName ? `表：${active.tableName}` : '未连接'}
          </span>
        </div>

        <div style={{ flex: 1, overflow: 'auto', padding: 16 }}>
          {loadError && (
            <Alert
              type="warning"
              showIcon
              style={{ marginBottom: 16, borderRadius: 8 }}
              message={loadError}
              description="模板功能依赖本地 dev server 的 /api 接口，请确认服务已启动。"
            />
          )}
          {active.loading ? (
            <div style={{ textAlign: 'center', paddingTop: 96 }}>
              <Spin size="large" />
              <div style={{ marginTop: 16, color: '#8c8c8c', fontSize: 13 }}>正在连接多维表格…</div>
            </div>
          ) : (
            <Tabs
              activeKey={activeKey}
              onChange={setActiveKey}
              items={items}
              tabBarStyle={{ marginBottom: 16, fontWeight: 500 }}
            />
          )}
        </div>
      </div>
    </ConfigProvider>
  );
}
