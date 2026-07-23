module.exports = (req, res) => {
  res.status(405).json({
    error: '静态部署不支持在线复制模板。请手动复制文件到 public/templates/<tableId>/ 目录后重新部署。',
  });
};
