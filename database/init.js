const initSqlJs = require('sql.js');
const fs = require('fs');
const path = require('path');
const bcrypt = require('bcryptjs');

const DB_PATH = path.join(__dirname, 'inventory.db');

let db = null;

async function initDatabase() {
  const SQL = await initSqlJs();

  // 如果数据库文件存在，加载它；否则创建新的
  if (fs.existsSync(DB_PATH)) {
    const buffer = fs.readFileSync(DB_PATH);
    db = new SQL.Database(buffer);
    // 运行数据库迁移
    migrateDatabase();
  } else {
    db = new SQL.Database();
    createTables();
    insertDefaultAdmin();
    saveDatabase();
  }

  return db;
}

// 数据库迁移：添加新字段
function migrateDatabase() {
  // 检查 usage_records 表是否有 customer_name 字段
  const columns = query("PRAGMA table_info(usage_records)");
  const hasCustomerName = columns.some(col => col.name === 'customer_name');

  if (!hasCustomerName) {
    db.run("ALTER TABLE usage_records ADD COLUMN customer_name TEXT");
    saveDatabase();
    console.log('数据库迁移: 已添加 customer_name 字段');
  }
}

function createTables() {
  // 用户表
  db.run(`
    CREATE TABLE IF NOT EXISTS users (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT UNIQUE NOT NULL,
      password TEXT NOT NULL,
      name TEXT NOT NULL,
      role INTEGER NOT NULL DEFAULT 2,
      department_id INTEGER,
      created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )
  `);

  // 部门/网点表
  db.run(`
    CREATE TABLE IF NOT EXISTS departments (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT UNIQUE NOT NULL,
      created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )
  `);

  // 活动表
  db.run(`
    CREATE TABLE IF NOT EXISTS activities (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT UNIQUE NOT NULL,
      created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )
  `);

  // 宣传品表
  db.run(`
    CREATE TABLE IF NOT EXISTS materials (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      activity_id INTEGER NOT NULL,
      name TEXT NOT NULL,
      unit TEXT NOT NULL,
      total_quantity INTEGER NOT NULL,
      created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
      FOREIGN KEY (activity_id) REFERENCES activities(id)
    )
  `);

  // 部门宣传品配额表
  db.run(`
    CREATE TABLE IF NOT EXISTS department_allocations (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      department_id INTEGER NOT NULL,
      material_id INTEGER NOT NULL,
      allocated_quantity INTEGER DEFAULT 0,
      used_quantity INTEGER DEFAULT 0,
      FOREIGN KEY (department_id) REFERENCES departments(id),
      FOREIGN KEY (material_id) REFERENCES materials(id),
      UNIQUE(department_id, material_id)
    )
  `);

  // 领用记录表
  db.run(`
    CREATE TABLE IF NOT EXISTS usage_records (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      department_id INTEGER NOT NULL,
      material_id INTEGER NOT NULL,
      quantity INTEGER NOT NULL,
      customer_name TEXT,
      remark TEXT,
      created_by INTEGER NOT NULL,
      created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
      FOREIGN KEY (department_id) REFERENCES departments(id),
      FOREIGN KEY (material_id) REFERENCES materials(id),
      FOREIGN KEY (created_by) REFERENCES users(id)
    )
  `);
}

function insertDefaultAdmin() {
  const hashedPassword = bcrypt.hashSync('admin123', 10);
  db.run(`
    INSERT INTO users (username, password, name, role, department_id)
    VALUES ('admin', ?, '系统管理员', 1, NULL)
  `, [hashedPassword]);
}

function saveDatabase() {
  const data = db.export();
  const buffer = Buffer.from(data);
  fs.writeFileSync(DB_PATH, buffer);
}

function getDatabase() {
  return db;
}

// 通用查询方法
function query(sql, params = []) {
  const stmt = db.prepare(sql);
  stmt.bind(params);
  const results = [];
  while (stmt.step()) {
    results.push(stmt.getAsObject());
  }
  stmt.free();
  return results;
}

function run(sql, params = []) {
  db.run(sql, params);
  saveDatabase();
  return db.getRowsModified();
}

function get(sql, params = []) {
  const results = query(sql, params);
  return results.length > 0 ? results[0] : null;
}

function getLastInsertId() {
  const result = query('SELECT last_insert_rowid() as id');
  return result[0].id;
}

module.exports = {
  initDatabase,
  getDatabase,
  saveDatabase,
  query,
  run,
  get,
  getLastInsertId
};
