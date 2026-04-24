const express = require('express');
const bcrypt = require('bcryptjs');
const db = require('../database/init');
const { passwordRuleError } = require('../utils/helpers');

const router = express.Router();

function homeForRole(role) {
  return role === 2 ? '/staff' : '/admin';
}

router.get('/login', (req, res) => {
  if (req.session && req.session.user) {
    return res.redirect(homeForRole(req.session.user.role));
  }
  res.render('login', { error: null });
});

router.post('/login', (req, res) => {
  const { username, password } = req.body;
  const user = db.get('SELECT * FROM users WHERE username = ?', [username]);

  if (!user || !bcrypt.compareSync(password || '', user.password)) {
    return res.render('login', { error: '用户名或密码错误' });
  }

  req.session.user = {
    id: user.id,
    username: user.username,
    name: user.name,
    role: user.role,
    department_id: user.department_id,
    must_change_password: user.must_change_password
  };

  if (user.must_change_password) {
    if (user.role === 2) return res.redirect('/staff/password?first=1');
    if (user.role === 1) return res.redirect('/admin/password?first=1');
  }

  res.redirect(homeForRole(user.role));
});

router.get('/logout', (req, res) => {
  req.session.destroy(() => {
    res.redirect('/login');
  });
});

module.exports = router;
module.exports.passwordRuleError = passwordRuleError;
