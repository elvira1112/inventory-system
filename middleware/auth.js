function isAuthenticated(req, res, next) {
  if (req.session && req.session.user) {
    return next();
  }
  res.redirect('/login');
}

function isSuperAdmin(req, res, next) {
  if (req.session && req.session.user && req.session.user.role === 0) {
    return next();
  }
  res.status(403).render('error', {
    message: '权限不足，仅超级管理员可访问',
    user: req.session.user
  });
}

function isAdmin(req, res, next) {
  if (req.session && req.session.user && (req.session.user.role === 0 || req.session.user.role === 1)) {
    return next();
  }
  res.status(403).render('error', {
    message: '权限不足，仅管理员可访问',
    user: req.session.user
  });
}

function isPrimary(req, res, next) {
  if (req.session && req.session.user && req.session.user.role === 1) {
    return next();
  }
  res.status(403).render('error', {
    message: '权限不足，仅一级人员可访问',
    user: req.session.user
  });
}

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
  isSuperAdmin,
  isAdmin,
  isPrimary,
  isStaff
};
