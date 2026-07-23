const fs = require('fs');
const path = require('path');

const CONFIG_PATH = path.join(process.cwd(), 'public', 'config.json');

module.exports = (req, res) => {
  if (req.method === 'GET') {
    try {
      if (!fs.existsSync(CONFIG_PATH)) {
        return res.status(200).json({ tables: {} });
      }
      const raw = fs.readFileSync(CONFIG_PATH, 'utf8');
      res.status(200).json(JSON.parse(raw));
    } catch (e) {
      res.status(500).json({ error: '读取配置失败：' + e.message });
    }
  } else if (req.method === 'PUT') {
    res.status(405).json({
      error: '静态部署不支持在线保存配置。请修改 public/config.json 后重新部署。',
    });
  } else {
    res.status(405).json({ error: 'Method Not Allowed' });
  }
};
