module.exports = (req, res) => {
  res.status(405).json({
    error: '静态部署不支持在线删除模板。请从仓库中移除文件后重新部署。',
  });
};
