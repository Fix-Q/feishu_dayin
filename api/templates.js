const fs = require('fs');
const path = require('path');

const TEMPLATES_DIR = path.join(process.cwd(), 'public', 'templates');

function safeTableId(raw) {
  if (typeof raw !== 'string' || !raw) {
    const err = new Error('缺少数据表标识');
    err.status = 400;
    throw err;
  }
  let id;
  try {
    id = decodeURIComponent(raw);
  } catch (e) {
    const err = new Error('数据表标识编码错误');
    err.status = 400;
    throw err;
  }
  if (!/^[A-Za-z0-9-]{1,64}$/.test(id)) {
    const err = new Error('数据表标识非法');
    err.status = 400;
    throw err;
  }
  return id;
}

function send(res, status, body) {
  res.status(status).json(body);
}

module.exports = (req, res) => {
  if (req.method === 'GET') {
    try {
      const tableId = safeTableId(req.query.tableId);
      const dir = path.join(TEMPLATES_DIR, tableId);

      if (!fs.existsSync(dir)) {
        return send(res, 200, { templates: [] });
      }

      const files = fs.readdirSync(dir)
        .filter((f) => /\.docx$/i.test(f) && !f.startsWith('~$'))
        .map((f) => {
          const st = fs.statSync(path.join(dir, f));
          return { name: f, size: st.size, mtime: st.mtimeMs };
        })
        .sort((a, b) => a.name.localeCompare(b.name, 'zh-CN'));

      send(res, 200, { templates: files });
    } catch (e) {
      send(res, e.status || 500, { error: e.status ? e.message : '读取模板列表失败：' + e.message });
    }
  } else if (req.method === 'POST') {
    res.status(405).json({
      error: '静态部署不支持在线上传模板。请将 .docx 文件放入 public/templates/<tableId>/ 目录后重新部署。',
    });
  } else {
    res.status(405).json({ error: 'Method Not Allowed' });
  }
};
