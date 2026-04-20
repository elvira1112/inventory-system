// 认证中间件
function isAuthenticated(req, res, next) {
  if (req.session && req.session.user) {
    return next();
  }
  res.redirect('/login');
}

// 一级管理员权限中间件
function isAdmin(req, res, next) {
  if (req.session && req.session.user && req.session.user.role === 1) {
    return next();
  }
  res.status(403).render('error', {
    message: '权限不足，仅一级管理员可访问',
    user: req.session.user
  });
}

// 二级人员权限中间件
function isStaff(req, res, next) {
  if (req.session && req.session.user && req.session.user.role === 2) {
    return next();
  }
  res.status(403).render('error', {
    message: '权限不足',
    user: req.session.user
  });
}

module.exports = {
  isAuthenticated,
  isAdmin,
  isStaff
};
