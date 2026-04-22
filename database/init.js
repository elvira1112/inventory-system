const initSqlJs = require('sql.js');
const fs = require('fs');
const path = require('path');
const bcrypt = require('bcryptjs');

const DB_PATH = path.join(__dirname, 'inventory.db');

let db = null;
let lastInsertId = null;

async function initDatabase() {
  const SQL = await initSqlJs();

  if (fs.existsSync(DB_PATH)) {
    const buffer = fs.readFileSync(DB_PATH);
    db = new SQL.Database(buffer);
    migrateDatabase();
  } else {
    db = new SQL.Database();
    createTables();
    insertDefaultAdmin();
    saveDatabase();
  }

  return db;
}

function createTables() {
  db.run(`
    CREATE TABLE IF NOT EXISTS users (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT UNIQUE NOT NULL,
      password TEXT NOT NULL,
      name TEXT NOT NULL,
      role INTEGER NOT NULL DEFAULT 2,
      department_id INTEGER,
      must_change_password INTEGER DEFAULT 0,
      initial_password TEXT,
      last_password_hash TEXT,
      created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS departments (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT UNIQUE NOT NULL,
      created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS activities (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT UNIQUE NOT NULL,
      department_id INTEGER,
      registration_date TEXT,
      created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )
  `);

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

  db.run(`
    CREATE TABLE IF NOT EXISTS department_allocations (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      department_id INTEGER NOT NULL,
      material_id INTEGER NOT NULL,
      allocated_quantity INTEGER DEFAULT 0,
      used_quantity INTEGER DEFAULT 0,
      recovered_quantity INTEGER DEFAULT 0,
      FOREIGN KEY (department_id) REFERENCES departments(id),
      FOREIGN KEY (material_id) REFERENCES materials(id),
      UNIQUE(department_id, material_id)
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS usage_records (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      department_id INTEGER NOT NULL,
      material_id INTEGER NOT NULL,
      quantity INTEGER NOT NULL,
      customer_name TEXT,
      remark TEXT,
      record_type TEXT DEFAULT 'usage',
      created_by INTEGER NOT NULL,
      created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
      FOREIGN KEY (department_id) REFERENCES departments(id),
      FOREIGN KEY (material_id) REFERENCES materials(id),
      FOREIGN KEY (created_by) REFERENCES users(id)
    )
  `);
}

function migrateDatabase() {
  let changed = false;

  changed = addColumnIfMissing('users', 'must_change_password', 'INTEGER DEFAULT 0') || changed;
  changed = addColumnIfMissing('users', 'initial_password', 'TEXT') || changed;
  changed = addColumnIfMissing('users', 'last_password_hash', 'TEXT') || changed;
  changed = addColumnIfMissing('activities', 'department_id', 'INTEGER') || changed;
  changed = addColumnIfMissing('activities', 'registration_date', 'TEXT') || changed;
  changed = addColumnIfMissing('department_allocations', 'recovered_quantity', 'INTEGER DEFAULT 0') || changed;
  changed = addColumnIfMissing('usage_records', 'customer_name', 'TEXT') || changed;
  changed = addColumnIfMissing('usage_records', 'record_type', "TEXT DEFAULT 'usage'") || changed;

  const admin = get("SELECT id, role FROM users WHERE username = 'admin'");
  if (admin && admin.role !== 0) {
    db.run("UPDATE users SET role = 0, must_change_password = 0 WHERE username = 'admin'");
    changed = true;
  }

  if (changed) {
    saveDatabase();
  }
}

function addColumnIfMissing(tableName, columnName, definition) {
  const columns = query(`PRAGMA table_info(${tableName})`).map(col => col.name);
  if (!columns.includes(columnName)) {
    db.run(`ALTER TABLE ${tableName} ADD COLUMN ${columnName} ${definition}`);
    return true;
  }
  return false;
}

function insertDefaultAdmin() {
  const hashedPassword = bcrypt.hashSync('admin123', 10);
  db.run(`
    INSERT INTO users (username, password, name, role, department_id, must_change_password, initial_password)
    VALUES ('admin', ?, '系统管理员', 0, NULL, 0, NULL)
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
  lastInsertId = query('SELECT last_insert_rowid() as id')[0].id;
  saveDatabase();
  return db.getRowsModified();
}

function get(sql, params = []) {
  const results = query(sql, params);
  return results.length > 0 ? results[0] : null;
}

function getLastInsertId() {
  if (lastInsertId !== null) {
    return lastInsertId;
  }
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
