# 库存管理系统 - 项目文档

## 项目概述
宣传品库存管理系统，支持两级权限（一级管理员、二级人员），用于管理活动宣传品的分配和领用。

## 技术栈
- **后端**: Node.js + Express 4.18
- **数据库**: SQLite (sql.js - 纯JS实现，无需编译)
- **前端**: EJS模板 + Bootstrap 5.3 + Bootstrap Icons
- **文件处理**: xlsx (Excel导入导出)
- **认证**: express-session + bcryptjs

## 项目结构
```
inventory-system/
├── app.js                    # 主入口，端口3000
├── package.json
├── database/
│   ├── init.js               # 数据库初始化和CRUD方法
│   └── inventory.db          # SQLite数据库文件
├── routes/
│   ├── auth.js               # 登录/登出路由
│   ├── admin.js              # 一级管理员路由（/admin/*）
│   └── staff.js              # 二级人员路由（/staff/*）
├── middleware/
│   └── auth.js               # isAuthenticated, isAdmin, isStaff
├── views/
│   ├── login.ejs             # 登录页
│   ├── error.ejs             # 错误页
│   ├── partials/
│   │   ├── admin-header.ejs  # 管理员导航头
│   │   ├── staff-header.ejs  # 员工导航头
│   │   └── footer.ejs        # 公共页脚
│   ├── admin/
│   │   ├── dashboard.ejs     # 管理员首页
│   │   ├── users.ejs         # 人员管理
│   │   ├── activities.ejs    # 活动列表
│   │   ├── activity-detail.ejs # 活动详情
│   │   ├── inventory.ejs     # 库存概览
│   │   └── allocations.ejs   # 配额调整
│   └── staff/
│       ├── dashboard.ejs     # 员工首页
│       ├── inventory.ejs     # 库存查看
│       ├── usage.ejs         # 领用录入
│       └── history.ejs       # 领用历史
├── public/css/
│   └── style.css             # 自定义样式
└── uploads/                  # Excel上传临时目录
```

## 数据库表结构

### users (用户表)
| 字段 | 类型 | 说明 |
|------|------|------|
| id | INTEGER PK | 自增主键 |
| username | TEXT UNIQUE | 用户名 |
| password | TEXT | bcrypt加密密码 |
| name | TEXT | 姓名 |
| role | INTEGER | 1=管理员, 2=二级人员 |
| department_id | INTEGER | 所属部门ID |
| created_at | DATETIME | 创建时间 |

### departments (部门表)
| 字段 | 类型 | 说明 |
|------|------|------|
| id | INTEGER PK | 自增主键 |
| name | TEXT UNIQUE | 部门名称 |
| created_at | DATETIME | 创建时间 |

### activities (活动表)
| 字段 | 类型 | 说明 |
|------|------|------|
| id | INTEGER PK | 自增主键 |
| name | TEXT UNIQUE | 活动名称 |
| created_at | DATETIME | 创建时间 |

### materials (宣传品表)
| 字段 | 类型 | 说明 |
|------|------|------|
| id | INTEGER PK | 自增主键 |
| activity_id | INTEGER FK | 所属活动 |
| name | TEXT | 宣传品名称 |
| unit | TEXT | 单位(个/张/本) |
| total_quantity | INTEGER | 总数量 |
| created_at | DATETIME | 创建时间 |

### department_allocations (部门配额表)
| 字段 | 类型 | 说明 |
|------|------|------|
| id | INTEGER PK | 自增主键 |
| department_id | INTEGER FK | 部门ID |
| material_id | INTEGER FK | 宣传品ID |
| allocated_quantity | INTEGER | 分配数量 |
| used_quantity | INTEGER | 已使用数量 |

### usage_records (领用记录表)
| 字段 | 类型 | 说明 |
|------|------|------|
| id | INTEGER PK | 自增主键 |
| department_id | INTEGER FK | 部门ID |
| material_id | INTEGER FK | 宣传品ID |
| quantity | INTEGER | 领用数量 |
| customer_name | TEXT | 领用客户名称 |
| remark | TEXT | 备注 |
| created_by | INTEGER FK | 录入人ID |
| created_at | DATETIME | 领用时间 |

## 核心API路由

### 认证 (routes/auth.js)
- `GET /login` - 登录页面
- `POST /login` - 登录处理
- `GET /logout` - 登出

### 管理员 (routes/admin.js)
- `GET /admin` - 仪表盘
- `GET /admin/users` - 人员列表
- `POST /admin/users/import` - 导入人员Excel
- `POST /admin/users/add` - 添加单个用户
- `POST /admin/users/delete/:id` - 删除用户
- `GET /admin/activities` - 活动列表
- `POST /admin/activities/import` - 导入活动Excel
- `GET /admin/activities/:id` - 活动详情
- `GET /admin/inventory` - 库存概览
- `GET /admin/allocations` - 配额管理页面
- `POST /admin/allocations/update` - 批量更新配额
- `GET /admin/export/inventory` - 导出库存报表
- `GET /admin/export/usage` - 导出领用明细
- `GET /admin/template/:type` - 下载Excel模板

### 员工 (routes/staff.js)
- `GET /staff` - 工作台
- `GET /staff/inventory` - 库存查看
- `GET /staff/usage` - 领用录入页面
- `POST /staff/usage` - 提交领用
- `GET /staff/history` - 领用历史
- `GET /staff/export/inventory` - 导出本部门库存
- `GET /staff/export/usage` - 导出本部门领用明细

## 数据库操作 (database/init.js)
```javascript
const db = require('./database/init');

// 初始化数据库
await db.initDatabase();

// 查询多条
db.query('SELECT * FROM users WHERE role = ?', [2]);

// 查询单条
db.get('SELECT * FROM users WHERE id = ?', [1]);

// 执行写操作（自动保存）
db.run('INSERT INTO users (...) VALUES (...)', [...]);

// 获取最后插入ID
db.getLastInsertId();
```

## Excel模板格式

### 人员名单模板
| 用户名 | 姓名 | 密码 | 部门/网点 |
|--------|------|------|-----------|
| zhangsan | 张三 | 123456 | 营业一部 |

### 活动宣传品模板
| 活动名称 | 宣传品名称 | 数量 | 单位 |
|----------|------------|------|------|
| 春季促销 | 宣传海报 | 1000 | 张 |

## 默认账号
- 管理员: admin / admin123

## 常用命令
```bash
# 启动系统
cd C:/Users/ranxi/inventory-system
npm start
# 或
node app.js

# 访问地址
http://localhost:3000
```

## 注意事项
1. sql.js是纯JS实现，每次写操作后调用saveDatabase()保存到文件
2. 密码使用bcryptjs加密，salt rounds = 10
3. Session有效期24小时
4. 文件上传使用multer，临时存储在uploads目录
5. 导出Excel使用xlsx库的buffer模式直接返回
