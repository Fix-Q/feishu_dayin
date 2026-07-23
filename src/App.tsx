import React, { useCallback, useEffect, useState } from 'react';
import { ConfigProvider, Tabs, Spin, Alert, message } from 'antd';
import zhCN from 'antd/locale/zh_CN';
import 'antd/dist/reset.css';

import type { TemplateInfo, MatchConfig } from './types';
import { listTemplates, getConfig } from './services/templateApi';
import { useActiveRecord } from './hooks/useActiveRecord';
import PrintTab from './components/PrintTab';
import TemplateManageTab from './components/TemplateManageTab';
import VariablePanel from './components/VariablePanel';

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
      setLoadError('无法加载模板列表：' + (e?.message || e));
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
      label: '打印',
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
      label: '模板管理',
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
      key: 'vars',
      label: '变量参考',
      children: <VariablePanel active={active} />,
    },
  ];

  return (
    <ConfigProvider locale={zhCN}>
      <div style={{ padding: 12, height: '100vh', boxSizing: 'border-box', overflow: 'auto' }}>
        {loadError && (
          <Alert
            type="warning"
            showIcon
            style={{ marginBottom: 12 }}
            message={loadError}
            description="模板功能依赖本地 dev server 的 /api 接口，请确认服务已启动。"
          />
        )}
        {active.loading ? (
          <div style={{ textAlign: 'center', paddingTop: 80 }}>
            <Spin />
            <div style={{ marginTop: 12, color: '#888' }}>正在连接多维表格…</div>
          </div>
        ) : (
          <Tabs activeKey={activeKey} onChange={setActiveKey} items={items} />
        )}
      </div>
    </ConfigProvider>
  );
}
