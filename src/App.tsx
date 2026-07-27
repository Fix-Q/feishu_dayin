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

// 主题色：贴合品牌的克制蓝绿，统一圆角与控件密度
const THEME = {
  token: {
    colorPrimary: '#2b7a78',
    borderRadius: 8,
    colorBgLayout: '#f5f6f8',
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
          overflow: 'auto',
          background: THEME.token.colorBgLayout,
          padding: 16,
        }}
      >
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
            size="large"
            tabBarStyle={{ marginBottom: 16, fontWeight: 500 }}
          />
        )}
      </div>
    </ConfigProvider>
  );
}
