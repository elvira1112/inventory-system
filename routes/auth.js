const express = require('express');
const router = express.Router();
const bcrypt = require('bcryptjs');
const db = require('../database/init');

// 登录页面
router.get('/login', (req, res) => {
  if (req.session && req.session.user) {
    // 已登录，根据角色跳转
    if (req.session.user.role === 1) {
      return res.redirect('/admin');
    } else {
      return res.redirect('/staff');
    }
  }
  res.render('login', { error: null });
});

// 登录处理
router.post('/login', (req, res) => {
  const { username, password } = req.body;

  const user = db.get('SELECT * FROM users WHERE username = ?', [username]);

  if (!user) {
    return res.render('login', { error: '用户名或密码错误' });
  }

  const isValid = bcrypt.compareSync(password, user.password);
  if (!isValid) {
    return res.render('login', { error: '用户名或密码错误' });
  }

  // 保存用户信息到session
  req.session.user = {
    id: user.id,
    username: user.username,
    name: user.name,
    role: user.role,
    department_id: user.department_id
  };

  // 根据角色跳转
  if (user.role === 1) {
    res.redirect('/admin');
  } else {
    res.redirect('/staff');
  }
});

// 登出
router.get('/logout', (req, res) => {
  req.session.destroy((err) => {
    res.redirect('/login');
  });
});

module.exports = router;
