const express = require('express');
const session = require('express-session');
const path = require('path');
const fs = require('fs');
const http = require('http');
const https = require('https');

const app = express();
const PORT = parseInt(process.env.PORT, 10) || 80;
const HTTPS_PORT = parseInt(process.env.HTTPS_PORT, 10) || 443;

// SSL 证书路径（从环境变量读取）
// SSL_KEY_PATH:  私钥文件路径（如 /etc/ssl/private/server.key）
// SSL_CERT_PATH: 证书文件路径（如 /etc/ssl/certs/server.crt）
// SSL_CA_PATH:   CA 证书链路径（可选）
const SSL_KEY_PATH = process.env.SSL_KEY_PATH || '';
const SSL_CERT_PATH = process.env.SSL_CERT_PATH || '';
const SSL_CA_PATH = process.env.SSL_CA_PATH || '';
const ENABLE_HTTPS = !!(SSL_KEY_PATH && SSL_CERT_PATH);

// 初始化数据库
const db = require('./database/init');

// 中间件
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// 启用 HTTPS 时让 Express 信任反向代理（如有），并在 cookie 上启用 secure
if (ENABLE_HTTPS) {
  app.set('trust proxy', 1);
}

// Session 配置
app.use(session({
  secret: process.env.SESSION_SECRET || 'inventory-system-secret-key-2024',
  resave: false,
  saveUninitialized: false,
  cookie: {
    maxAge: 24 * 60 * 60 * 1000, // 24小时
    secure: ENABLE_HTTPS,
    httpOnly: true
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
    if (req.session.user.role === 0 || req.session.user.role === 1) {
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

// 读取 SSL 证书
function loadSslOptions() {
  const options = {
    key: fs.readFileSync(SSL_KEY_PATH),
    cert: fs.readFileSync(SSL_CERT_PATH)
  };
  if (SSL_CA_PATH && fs.existsSync(SSL_CA_PATH)) {
    options.ca = fs.readFileSync(SSL_CA_PATH);
  }
  return options;
}

// 启动服务器
async function start() {
  try {
    await db.initDatabase();
    console.log('数据库初始化成功');

    if (ENABLE_HTTPS) {
      const sslOptions = loadSslOptions();

      // HTTPS 主服务
      https.createServer(sslOptions, app).listen(HTTPS_PORT, () => {
        console.log(`
╔════════════════════════════════════════╗
║     宣传品管理系统已启动 (HTTPS)       ║
╠════════════════════════════════════════╣
║  访问地址: https://localhost:${HTTPS_PORT}
║  默认账号: admin
║  默认密码: admin123
╚════════════════════════════════════════╝
        `);
      });

      // HTTP -> HTTPS 自动重定向
      const redirectApp = express();
      redirectApp.use((req, res) => {
        const host = (req.headers.host || '').split(':')[0];
        const target = HTTPS_PORT === 443
          ? `https://${host}${req.url}`
          : `https://${host}:${HTTPS_PORT}${req.url}`;
        res.redirect(301, target);
      });
      http.createServer(redirectApp).listen(PORT, () => {
        console.log(`HTTP 端口 ${PORT} 已启用，自动重定向到 HTTPS`);
      });
    } else {
      http.createServer(app).listen(PORT, () => {
        console.log(`
╔════════════════════════════════════════╗
║     宣传品管理系统已启动 (HTTP)        ║
╠════════════════════════════════════════╣
║  访问地址: http://localhost:${PORT}
║  默认账号: admin
║  默认密码: admin123
╚════════════════════════════════════════╝
        `);
        console.log('提示: 配置 SSL_KEY_PATH 与 SSL_CERT_PATH 环境变量可启用 HTTPS');
      });
    }
  } catch (err) {
    console.error('启动失败:', err);
    process.exit(1);
  }
}

start();
