// 模板文件与匹配配置的服务端 API。
// 由 webpack.config.js 的 devServer.setupMiddlewares 挂载到 express app。
// 模板按数据表隔离：templates/<tableId>/*.docx，每张表只看到自己目录下的模板。
// 匹配配置 templates/_config.json 全局，按 tableId 记录。全部纳入 git，本地部署直接调用。
const fs = require('fs');
const path = require('path');
const express = require('express');

const TEMPLATES_DIR = path.resolve(__dirname, '..', 'templates');
const CONFIG_PATH = path.join(TEMPLATES_DIR, '_config.json');

function ensureStore() {
  fs.mkdirSync(TEMPLATES_DIR, { recursive: true });
  if (!fs.existsSync(CONFIG_PATH)) {
    fs.writeFileSync(CONFIG_PATH, JSON.stringify({ tables: {} }, null, 2), 'utf8');
  }
}

// 校验 tableId：飞书表 ID 形如 tbl8SmrWcQadcF9k，仅字母数字与短横线，
// 禁止路径成分与下划线开头（下划线前缀留给 _config.json / _legacy 等保留项）。
function safeTableId(raw) {
  if (typeof raw !== 'string' || !raw) throw httpError(400, '缺少数据表标识');
  const id = raw; // Express 已解码 query 参数，不可再次 decodeURIComponent。
  if (!/^[A-Za-z0-9-]{1,64}$/.test(id)) throw httpError(400, '数据表标识非法');
  return id;
}

// 返回并按需创建某张表的模板目录
function tableDir(tableId) {
  const dir = path.join(TEMPLATES_DIR, tableId);
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

// 文件名清洗：所有涉及文件名的路由必经此关。
// 返回合法 basename，或抛出带 status 的错误。
function safeName(raw) {
  if (typeof raw !== 'string' || !raw) throw httpError(400, '缺少文件名');
  let name = raw; // Express 已解码 query/route 参数，二次解码会破坏含 % 的合法文件名。
  // 显式拒绝含路径成分的输入（不静默改名，避免 a/b.docx 悄悄变 b.docx 或路径穿越尝试）
  if (name.includes('/') || name.includes('\\') || name.includes('\0')) {
    throw httpError(400, '文件名不能含路径分隔符');
  }
  name = path.basename(name); // 双保险：去掉任何残留路径成分
  if (!/\.(docx|xlsx)$/i.test(name)) throw httpError(400, '仅支持 .docx / .xlsx 文件');
  if (name.startsWith('_')) throw httpError(400, '文件名不能以下划线开头');
  if (name.startsWith('~$')) throw httpError(400, '非法的临时文件名');
  // 拒绝控制字符与 Windows 保留字符（basename 后不应再有 / \，双保险）
  if (/[<>:"/\|?*]/.test(name)) throw httpError(400, '文件名含非法字符');
  if (/[\x00-\x1f]/.test(name)) throw httpError(400, '文件名含控制字符');
  if (name === '.docx' || name === '.xlsx') throw httpError(400, '文件名不能为空');
  const full = path.resolve(TEMPLATES_DIR, name);
  if (path.dirname(full) !== TEMPLATES_DIR) throw httpError(400, '路径越界');
  return name;
}

function httpError(status, message) {
  const err = new Error(message);
  err.status = status;
  return err;
}

function send(res, status, body) {
  res.status(status).json(body);
}

module.exports = function attach(app) {
  ensureStore();

  // 列表：某张表目录下的 .docx，中文排序。?tableId=xxx
  app.get('/api/templates', (req, res) => {
    try {
      const tableId = safeTableId(req.query.tableId);
      const dir = tableDir(tableId);
      const files = fs.readdirSync(dir)
        .filter((f) => /\.(docx|xlsx)$/i.test(f) && !f.startsWith('~$'))
        .map((f) => {
          const st = fs.statSync(path.join(dir, f));
          return { name: f, size: st.size, mtime: st.mtimeMs };
        })
        .sort((a, b) => a.name.localeCompare(b.name, 'zh-CN'));
      send(res, 200, { templates: files });
    } catch (e) {
      send(res, e.status || 500, { error: e.status ? e.message : '读取模板列表失败：' + e.message });
    }
  });

  // 上传：raw body，校验 PK 魔数，重名需 overwrite=1。?tableId=xxx&name=yyy
  app.post('/api/templates', express.raw({ type: () => true, limit: '20mb' }), (req, res) => {
    try {
      const tableId = safeTableId(req.query.tableId);
      const name = safeName(req.query.name);
      const buf = req.body;
      if (!Buffer.isBuffer(buf) || buf.length === 0) throw httpError(400, '上传内容为空');
      if (buf[0] !== 0x50 || buf[1] !== 0x4b) throw httpError(400, '不是有效的 Office 文件（docx/xlsx）');
      const full = path.join(tableDir(tableId), name);
      const overwrite = req.query.overwrite === '1';
      if (fs.existsSync(full) && !overwrite) {
        return send(res, 409, { error: '同名模板已存在', name });
      }
      fs.writeFileSync(full, buf);
      const st = fs.statSync(full);
      send(res, 200, { name, size: st.size, mtime: st.mtimeMs });
    } catch (e) {
      send(res, e.status || 500, { error: e.message });
    }
  });

  // 删除：?tableId=xxx
  app.delete('/api/templates/:name', (req, res) => {
    try {
      const tableId = safeTableId(req.query.tableId);
      const name = safeName(req.params.name);
      const full = path.join(tableDir(tableId), name);
      if (!fs.existsSync(full)) return send(res, 404, { error: '模板不存在' });
      fs.unlinkSync(full);
      send(res, 200, { name });
    } catch (e) {
      send(res, e.status || 500, { error: e.message });
    }
  });

  // 复制（同表内）：{ tableId, source, target }
  app.post('/api/templates/copy', express.json(), (req, res) => {
    try {
      const tableId = safeTableId(req.body && req.body.tableId);
      const source = safeName(req.body && req.body.source);
      const target = safeName(req.body && req.body.target);
      const dir = tableDir(tableId);
      const src = path.join(dir, source);
      const dst = path.join(dir, target);
      if (!fs.existsSync(src)) return send(res, 404, { error: '源模板不存在' });
      if (fs.existsSync(dst)) return send(res, 409, { error: '目标模板已存在' });
      fs.copyFileSync(src, dst);
      const st = fs.statSync(dst);
      send(res, 200, { name: target, size: st.size, mtime: st.mtimeMs });
    } catch (e) {
      send(res, e.status || 500, { error: e.message });
    }
  });

  // 匹配配置读写
  app.get('/api/config', (req, res) => {
    try {
      ensureStore();
      const raw = fs.readFileSync(CONFIG_PATH, 'utf8');
      send(res, 200, JSON.parse(raw));
    } catch (e) {
      send(res, 500, { error: '读取配置失败：' + e.message });
    }
  });

  app.put('/api/config', express.json(), (req, res) => {
    try {
      const body = req.body;
      if (!body || typeof body !== 'object' || typeof body.tables !== 'object' || body.tables === null) {
        throw httpError(400, '配置格式错误，需为 { tables: {...} }');
      }
      fs.writeFileSync(CONFIG_PATH, JSON.stringify(body, null, 2), 'utf8');
      send(res, 200, body);
    } catch (e) {
      send(res, e.status || 500, { error: e.message });
    }
  });
};
