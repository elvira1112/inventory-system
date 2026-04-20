const express = require('express');
const session = require('express-session');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// 初始化数据库
const db = require('./database/init');

// 中间件
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Session 配置
app.use(session({
  secret: 'inventory-system-secret-key-2024',
  resave: false,
  saveUninitialized: false,
  cookie: {
    maxAge: 24 * 60 * 60 * 1000 // 24小时
  }
}));

// 视图引擎
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// 路由
const authRoutes = require('./routes/auth');
const adminRoutes = require('./routes/admin');
const staffRoutes = require('./routes/staff');

app.use('/', authRoutes);
app.use('/admin', adminRoutes);
app.use('/staff', staffRoutes);

// 首页重定向
app.get('/', (req, res) => {
  if (req.session && req.session.user) {
    if (req.session.user.role === 1) {
      res.redirect('/admin');
    } else {
      res.redirect('/staff');
    }
  } else {
    res.redirect('/login');
  }
});

// 404处理
app.use((req, res) => {
  res.status(404).render('error', {
    message: '页面不存在',
    user: req.session ? req.session.user : null
  });
});

// 错误处理
app.use((err, req, res, next) => {
  console.error(err.stack);
  res.status(500).render('error', {
    message: '服务器内部错误',
    user: req.session ? req.session.user : null
  });
});

// 启动服务器
async function start() {
  try {
    await db.initDatabase();
    console.log('数据库初始化成功');

    app.listen(PORT, () => {
      console.log(`
╔════════════════════════════════════════╗
║     库存管理系统已启动                 ║
╠════════════════════════════════════════╣
║  访问地址: http://localhost:${PORT}       ║
║  默认账号: admin                       ║
║  默认密码: admin123                    ║
╚════════════════════════════════════════╝
      `);
    });
  } catch (err) {
    console.error('启动失败:', err);
    process.exit(1);
  }
}

start();
