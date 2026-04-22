const express = require('express');
const router = express.Router();
const bcrypt = require('bcryptjs');
const db = require('../database/init');

function homeForRole(role) {
  return role === 2 ? '/staff' : '/admin';
}

function passwordRuleError(password, user) {
  if (!password || password.length < 6) {
    return '密码长度至少6位';
  }
  if (/^(\d)\1{3,}$/.test(password)) {
    return '密码不能是4位以上重复数字';
  }
  if (/\d{3,}/.test(password)) {
    for (let i = 0; i <= password.length - 3; i++) {
      const part = password.slice(i, i + 3);
      if (/^\d{3}$/.test(part)) {
        const nums = part.split('').map(Number);
        if (nums[1] === nums[0] + 1 && nums[2] === nums[1] + 1) {
          return '密码不能包含3位以上连续数字';
        }
      }
    }
  }
  if (user && bcrypt.compareSync(password, user.password)) {
    return '新密码不能和上一次密码相同';
  }
  return null;
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

  if (!user || !bcrypt.compareSync(password, user.password)) {
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

  if (user.role === 2 && user.must_change_password) {
    return res.redirect('/staff/password?first=1');
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
