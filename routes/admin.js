const express = require('express');
const router = express.Router();
const multer = require('multer');
const XLSX = require('xlsx');
const bcrypt = require('bcryptjs');
const path = require('path');
const db = require('../database/init');
const { isAuthenticated, isAdmin, isSuperAdmin } = require('../middleware/auth');

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, path.join(__dirname, '../uploads')),
  filename: (req, file, cb) => cb(null, Date.now() + '-' + file.originalname)
});
const upload = multer({ storage });

const managedDepartmentOrder = [
  '个人金融业务部',
  '公司金融业务部',
  '机构金融业务部',
  '普惠金融业务部',
  '本级业务部',
  '转塘支行',
  '文三路支行',
  '象山路小微企业专营支行',
  '象山路微企业\n专营支行',
  '银马支行'
];
const excludedAllocationDepartments = ['风险管理部', '行长室', '综合管理部'];

router.use(isAuthenticated);
router.use(isAdmin);

function isSuper(user) {
  return user.role === 0;
}

function escapeLike(value) {
  return String(value || '').replace(/[%_]/g, '\\$&');
}

function getDepartmentByName(name) {
  return db.get('SELECT id, name FROM departments WHERE name = ?', [String(name || '').trim()]);
}

function ensureDepartment(name) {
  const cleanName = String(name || '').trim();
  if (!cleanName) return null;
  let dept = getDepartmentByName(cleanName);
  if (!dept) {
    db.run('INSERT INTO departments (name) VALUES (?)', [cleanName]);
    dept = { id: db.getLastInsertId(), name: cleanName };
  }
  return dept;
}

function roleName(role) {
  if (role === 0) return '超级管理员';
  if (role === 1) return '一级人员';
  return '二级人员';
}

function formatDateTime(value) {
  if (!value) return '';
  const d = new Date(String(value).replace(' ', 'T') + '+08:00');
  if (Number.isNaN(d.getTime())) return value;
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

function normalizeExcelDate(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }
  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      const pad = n => String(n).padStart(2, '0');
      return `${parsed.y}-${pad(parsed.m)}-${pad(parsed.d)}`;
    }
  }
  return String(value).trim();
}

function userScopeWhere(user, alias = 'a') {
  if (isSuper(user)) {
    return { where: '1=1', params: [] };
  }
  return { where: `${alias}.department_id = ?`, params: [user.department_id] };
}

function getActivitiesForUser(user) {
  const scope = userScopeWhere(user, 'a');
  return db.query(`
    SELECT a.*, d.name as department_name
    FROM activities a
    LEFT JOIN departments d ON a.department_id = d.id
    WHERE ${scope.where}
    ORDER BY a.name
  `, scope.params);
}

function getReportDepartments() {
  const departments = db.query('SELECT * FROM departments');
  return departments
    .filter(dept => !excludedAllocationDepartments.includes(dept.name))
    .sort((a, b) => {
      const ia = managedDepartmentOrder.indexOf(a.name);
      const ib = managedDepartmentOrder.indexOf(b.name);
      if (ia !== -1 || ib !== -1) return (ia === -1 ? 999 : ia) - (ib === -1 ? 999 : ib);
      return a.name.localeCompare(b.name, 'zh-CN');
    });
}

function getMaterialInventoryRows(user, activityId) {
  const params = [];
  let activityFilter = '';
  if (activityId) {
    activityFilter = 'AND a.id = ?';
    params.push(activityId);
  }
  const scope = userScopeWhere(user, 'a');
  params.push(...scope.params);

  return db.query(`
    SELECT
      d.name as owner_department,
      a.id as activity_id,
      a.name as activity_name,
      m.id as material_id,
      m.name as material_name,
      m.unit,
      m.total_quantity,
      COALESCE(SUM(da.allocated_quantity), 0) as allocated,
      COALESCE(SUM(da.used_quantity), 0) as used,
      COALESCE(SUM(da.recovered_quantity), 0) as recovered,
      m.total_quantity - COALESCE(SUM(da.allocated_quantity), 0) + COALESCE(SUM(da.recovered_quantity), 0) as unallocated
    FROM materials m
    JOIN activities a ON m.activity_id = a.id
    LEFT JOIN departments d ON a.department_id = d.id
    LEFT JOIN department_allocations da ON m.id = da.material_id
    WHERE 1=1 ${activityFilter} AND ${scope.where}
    GROUP BY m.id
    ORDER BY d.name, a.name, m.name
  `, params);
}

function buildInventorySheetRows(user, activityId) {
  const materials = getMaterialInventoryRows(user, activityId);
  const departments = getReportDepartments();
  const rows = materials.map(item => {
    const row = {
      '牵头管理部门': item.owner_department || '',
      '营销活动名称': item.activity_name,
      '宣传品名称': item.material_name,
      '总量': item.total_quantity,
      '合计已分配数': item.allocated,
      '总剩余库存数': item.unallocated
    };

    const allocs = db.query(`
      SELECT d.name as department_name,
             COALESCE(da.allocated_quantity, 0) as allocated,
             COALESCE(da.used_quantity, 0) as used,
             COALESCE(da.recovered_quantity, 0) as recovered
      FROM departments d
      LEFT JOIN department_allocations da ON d.id = da.department_id AND da.material_id = ?
    `, [item.material_id]);
    const byDept = Object.fromEntries(allocs.map(a => [a.department_name, a]));

    departments.forEach(dept => {
      const alloc = byDept[dept.name] || { allocated: 0, used: 0, recovered: 0 };
      row[`${dept.name}-合计领入数量`] = alloc.allocated;
      row[`${dept.name}-已使用数量`] = alloc.used;
      row[`${dept.name}-剩余库存`] = alloc.allocated - alloc.used - alloc.recovered;
    });

    return row;
  });
  return rows;
}

function buildDetailRows(user, activityId) {
  const scope = userScopeWhere(user, 'a');
  const params = [];
  let activityFilter = '';
  if (activityId) {
    activityFilter = 'AND a.id = ?';
    params.push(activityId);
  }
  params.push(...scope.params);

  const allocationRows = db.query(`
    SELECT od.name as owner_department, a.name as activity_name, m.name as material_name,
           d.name as department_name, da.allocated_quantity as quantity,
           m.total_quantity - COALESCE(total_alloc.total_allocated, 0) + COALESCE(total_alloc.total_recovered, 0) as remaining_unallocated
    FROM department_allocations da
    JOIN materials m ON da.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    LEFT JOIN departments od ON a.department_id = od.id
    JOIN departments d ON da.department_id = d.id
    LEFT JOIN (
      SELECT material_id, SUM(allocated_quantity) as total_allocated, SUM(recovered_quantity) as total_recovered
      FROM department_allocations GROUP BY material_id
    ) total_alloc ON total_alloc.material_id = m.id
    WHERE da.allocated_quantity > 0 ${activityFilter} AND ${scope.where}
    ORDER BY od.name, a.name, m.name, d.name
  `, params);

  const rows = allocationRows.map(r => ({
    '牵头管理部门': r.owner_department || '',
    '营销活动名称': r.activity_name,
    '宣传品名称': r.material_name,
    '分配/回收标识': '分配',
    '部门/网点': r.department_name,
    '数量': r.quantity,
    '剩余未分配库存数': r.remaining_unallocated
  }));

  const recoveryRows = db.query(`
    SELECT od.name as owner_department, a.name as activity_name, m.name as material_name,
           d.name as department_name, ur.quantity
    FROM usage_records ur
    JOIN materials m ON ur.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    LEFT JOIN departments od ON a.department_id = od.id
    JOIN departments d ON ur.department_id = d.id
    WHERE ur.record_type = 'recovery' ${activityId ? 'AND a.id = ?' : ''} AND ${scope.where}
    ORDER BY ur.created_at DESC
  `, params);

  recoveryRows.forEach(r => rows.push({
    '牵头管理部门': r.owner_department || '',
    '营销活动名称': r.activity_name,
    '宣传品名称': r.material_name,
    '分配/回收标识': '回收',
    '部门/网点': r.department_name,
    '数量': r.quantity,
    '剩余未分配库存数': ''
  }));

  return rows;
}

router.get('/', (req, res) => {
  const scope = userScopeWhere(req.session.user, 'a');
  const stats = {
    departmentCount: db.get('SELECT COUNT(*) as count FROM departments').count,
    primaryUserCount: db.get('SELECT COUNT(*) as count FROM users WHERE role = 1').count,
    staffUserCount: db.get('SELECT COUNT(*) as count FROM users WHERE role = 2').count,
    activityCount: db.get(`SELECT COUNT(*) as count FROM activities a WHERE ${scope.where}`, scope.params).count,
    materialCount: db.get(`
      SELECT COUNT(*) as count FROM materials m
      JOIN activities a ON m.activity_id = a.id
      WHERE ${scope.where}
    `, scope.params).count
  };
  res.render('admin/dashboard', { user: req.session.user, stats, currentPage: 'dashboard' });
});

router.get('/users', isSuperAdmin, (req, res) => {
  const users = db.query(`
    SELECT u.*, d.name as department_name
    FROM users u
    LEFT JOIN departments d ON u.department_id = d.id
    ORDER BY u.role, u.id DESC
  `).map(u => ({ ...u, role_name: roleName(u.role) }));
  const departments = db.query('SELECT * FROM departments ORDER BY name');
  res.render('admin/users', { user: req.session.user, users, departments, currentPage: 'users' });
});

router.post('/users/import', isSuperAdmin, upload.single('file'), (req, res) => {
  try {
    if (!req.file) return res.redirect('/admin/users?error=请选择文件');
    const workbook = XLSX.readFile(req.file.path);
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
    let imported = 0;

    for (const row of data) {
      const username = row['用户名'] || row.username;
      const name = row['姓名'] || row.name;
      const password = String(row['密码'] || row.password || '123456');
      const deptName = row['部门/网点'] || row['部门'] || row.department;
      const role = Number(row['权限'] || row['角色'] || row.role || 2);
      if (!username || !name) continue;
      const dept = role === 0 ? null : ensureDepartment(deptName);
      if (role !== 0 && !dept) continue;
      if (db.get('SELECT id FROM users WHERE username = ?', [username])) continue;
      db.run(
        'INSERT INTO users (username, password, name, role, department_id, must_change_password, initial_password) VALUES (?, ?, ?, ?, ?, 1, ?)',
        [username, bcrypt.hashSync(password, 10), name, role, dept ? dept.id : null, password]
      );
      imported++;
    }
    res.redirect(`/admin/users?success=成功导入 ${imported} 名人员`);
  } catch (err) {
    res.redirect('/admin/users?error=导入失败: ' + encodeURIComponent(err.message));
  }
});

router.post('/users/add', isSuperAdmin, (req, res) => {
  const { username, name, password, department_id, role } = req.body;
  try {
    if (db.get('SELECT id FROM users WHERE username = ?', [username])) {
      return res.redirect('/admin/users?error=用户名已存在');
    }
    const rawPassword = password || '123456';
    const userRole = Number(role || 2);
    db.run(
      'INSERT INTO users (username, password, name, role, department_id, must_change_password, initial_password) VALUES (?, ?, ?, ?, ?, 1, ?)',
      [username, bcrypt.hashSync(rawPassword, 10), name, userRole, userRole === 0 ? null : department_id, rawPassword]
    );
    res.redirect('/admin/users?success=人员添加成功');
  } catch (err) {
    res.redirect('/admin/users?error=' + encodeURIComponent(err.message));
  }
});

router.post('/users/update/:id', isSuperAdmin, (req, res) => {
  const { name, role, department_id } = req.body;
  const userRole = Number(role || 2);
  db.run('UPDATE users SET name = ?, role = ?, department_id = ? WHERE id = ? AND username <> ?',
    [name, userRole, userRole === 0 ? null : department_id, req.params.id, 'admin']);
  res.redirect('/admin/users?success=人员信息已更新');
});

router.post('/users/reset-password/:id', isSuperAdmin, (req, res) => {
  const rawPassword = req.body.password || '123456';
  const hashedPassword = bcrypt.hashSync(rawPassword, 10);
  db.run(
    'UPDATE users SET last_password_hash = password, password = ?, must_change_password = 1, initial_password = ? WHERE id = ?',
    [hashedPassword, rawPassword, req.params.id]
  );
  res.redirect('/admin/users?success=密码已重置');
});

router.post('/users/delete/:id', isSuperAdmin, (req, res) => {
  db.run("DELETE FROM users WHERE id = ? AND username <> 'admin'", [req.params.id]);
  res.redirect('/admin/users?success=人员已删除');
});

router.get('/departments', isSuperAdmin, (req, res) => {
  const departments = db.query(`
    SELECT d.*, COUNT(u.id) as user_count
    FROM departments d
    LEFT JOIN users u ON d.id = u.department_id
    GROUP BY d.id
    ORDER BY d.name
  `);
  res.render('admin/departments', { user: req.session.user, departments, currentPage: 'departments' });
});

router.post('/departments/add', isSuperAdmin, (req, res) => {
  try {
    ensureDepartment(req.body.name);
    res.redirect('/admin/departments?success=部门/网点已添加');
  } catch (err) {
    res.redirect('/admin/departments?error=' + encodeURIComponent(err.message));
  }
});

router.post('/departments/import', isSuperAdmin, upload.single('file'), (req, res) => {
  try {
    if (!req.file) return res.redirect('/admin/departments?error=请选择文件');
    const workbook = XLSX.readFile(req.file.path);
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
    let imported = 0;
    for (const row of data) {
      const name = row['部门/网点'] || row['部门名称'] || row.name;
      if (name && !getDepartmentByName(name)) {
        ensureDepartment(name);
        imported++;
      }
    }
    res.redirect(`/admin/departments?success=成功导入 ${imported} 个部门/网点`);
  } catch (err) {
    res.redirect('/admin/departments?error=导入失败: ' + encodeURIComponent(err.message));
  }
});

router.post('/departments/update/:id', isSuperAdmin, (req, res) => {
  db.run('UPDATE departments SET name = ? WHERE id = ?', [req.body.name, req.params.id]);
  res.redirect('/admin/departments?success=部门/网点已更新');
});

router.post('/departments/delete/:id', isSuperAdmin, (req, res) => {
  const used = db.get('SELECT COUNT(*) as count FROM users WHERE department_id = ?', [req.params.id]).count;
  if (used > 0) return res.redirect('/admin/departments?error=该部门/网点下仍有人员，不能删除');
  db.run('DELETE FROM departments WHERE id = ?', [req.params.id]);
  res.redirect('/admin/departments?success=部门/网点已删除');
});

router.get('/activities', (req, res) => {
  const scope = userScopeWhere(req.session.user, 'a');
  const activities = db.query(`
    SELECT a.*, d.name as department_name,
           COUNT(m.id) as material_count,
           COALESCE(SUM(m.total_quantity), 0) as total_items
    FROM activities a
    LEFT JOIN departments d ON a.department_id = d.id
    LEFT JOIN materials m ON a.id = m.activity_id
    WHERE ${scope.where}
    GROUP BY a.id
    ORDER BY a.id DESC
  `, scope.params);
  res.render('admin/activities', { user: req.session.user, activities, currentPage: 'activities' });
});

router.post('/activities/import', upload.single('file'), (req, res) => {
  try {
    if (!req.file) return res.redirect('/admin/activities?error=请选择文件');
    if (!req.session.user.department_id) {
      return res.redirect('/admin/activities?error=当前账号未绑定部门/网点，不能导入活动');
    }
    const workbook = XLSX.readFile(req.file.path);
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
    let imported = 0;
    const ownerDept = db.get('SELECT * FROM departments WHERE id = ?', [req.session.user.department_id]);
    if (!ownerDept) {
      return res.redirect('/admin/activities?error=当前账号所属部门/网点不存在，不能导入活动');
    }

    for (const row of data) {
      const activityName = row['活动名称'] || row['营销活动名称'] || row.activity;
      const materialName = row['宣传品名称'] || row.material;
      const quantity = parseInt(row['数量'] || row.quantity, 10) || 0;
      const unit = '份';
      const registrationDate = normalizeExcelDate(row['登记日期'] || row.registration_date);
      if (!activityName || !materialName) continue;

      let activity = db.get('SELECT id FROM activities WHERE name = ?', [activityName]);
      if (!activity) {
        db.run('INSERT INTO activities (name, department_id, registration_date) VALUES (?, ?, ?)', [activityName, ownerDept.id, registrationDate]);
        activity = { id: db.getLastInsertId() };
      } else {
        db.run('UPDATE activities SET department_id = ?, registration_date = COALESCE(NULLIF(?, \'\'), registration_date) WHERE id = ?', [ownerDept.id, registrationDate, activity.id]);
      }

      const existingMaterial = db.get('SELECT id FROM materials WHERE activity_id = ? AND name = ?', [activity.id, materialName]);
      if (existingMaterial) {
        db.run('UPDATE materials SET total_quantity = total_quantity + ? WHERE id = ?', [quantity, existingMaterial.id]);
      } else {
        db.run('INSERT INTO materials (activity_id, name, unit, total_quantity) VALUES (?, ?, ?, ?)', [activity.id, materialName, unit, quantity]);
      }
      imported++;
    }
    res.redirect(`/admin/activities?success=成功导入 ${imported} 条记录`);
  } catch (err) {
    res.redirect('/admin/activities?error=导入失败: ' + encodeURIComponent(err.message));
  }
});

router.get('/activities/:id', (req, res) => {
  const scope = userScopeWhere(req.session.user, 'a');
  const activity = db.get(`SELECT a.*, d.name as department_name FROM activities a LEFT JOIN departments d ON a.department_id = d.id WHERE a.id = ? AND ${scope.where}`, [req.params.id, ...scope.params]);
  if (!activity) return res.redirect('/admin/activities?error=活动不存在或无权查看');
  const materials = db.query(`
    SELECT m.*, COALESCE(SUM(da.allocated_quantity), 0) as allocated,
           COALESCE(SUM(da.used_quantity), 0) as used,
           COALESCE(SUM(da.recovered_quantity), 0) as recovered
    FROM materials m
    LEFT JOIN department_allocations da ON m.id = da.material_id
    WHERE m.activity_id = ?
    GROUP BY m.id
  `, [req.params.id]);
  res.render('admin/activity-detail', { user: req.session.user, activity, materials, currentPage: 'activities' });
});

router.get('/inventory', (req, res) => {
  const selectedActivity = req.query.activity_id || '';
  const activities = getActivitiesForUser(req.session.user);
  const inventory = selectedActivity ? getMaterialInventoryRows(req.session.user, selectedActivity) : [];
  const departments = getReportDepartments();
  const departmentStocks = {};
  if (selectedActivity) {
    const allocations = db.query(`
      SELECT da.material_id, d.name as department_name, da.allocated_quantity, da.used_quantity, da.recovered_quantity
      FROM department_allocations da
      JOIN departments d ON da.department_id = d.id
      JOIN materials m ON da.material_id = m.id
      WHERE m.activity_id = ?
    `, [selectedActivity]);
    allocations.forEach(a => {
      departmentStocks[`${a.material_id}_${a.department_name}`] = {
        allocated: a.allocated_quantity || 0,
        used: a.used_quantity || 0,
        remaining: (a.allocated_quantity || 0) - (a.used_quantity || 0) - (a.recovered_quantity || 0)
      };
    });
  }
  res.render('admin/inventory', { user: req.session.user, activities, inventory, departments, departmentStocks, selectedActivity, currentPage: 'inventory' });
});

router.get('/allocations', (req, res) => {
  const activities = getActivitiesForUser(req.session.user);
  const departments = getReportDepartments();
  const activityId = req.query.activity_id;
  let materials = [];
  let allocations = [];
  if (activityId) {
    const scope = userScopeWhere(req.session.user, 'a');
    const activity = db.get(`SELECT a.id FROM activities a WHERE a.id = ? AND ${scope.where}`, [activityId, ...scope.params]);
    if (!activity) return res.redirect('/admin/inventory?error=活动不存在或无权操作');
    materials = db.query(`
      SELECT m.*, COALESCE(SUM(da.allocated_quantity), 0) as total_allocated
      FROM materials m
      LEFT JOIN department_allocations da ON m.id = da.material_id
      WHERE m.activity_id = ?
      GROUP BY m.id
    `, [activityId]);
    allocations = db.query(`
      SELECT d.id as department_id, d.name as department_name, m.id as material_id,
             COALESCE(da.allocated_quantity, 0) as allocated,
             COALESCE(da.used_quantity, 0) as used,
             COALESCE(da.recovered_quantity, 0) as recovered
      FROM departments d
      CROSS JOIN materials m
      LEFT JOIN department_allocations da ON d.id = da.department_id AND m.id = da.material_id
      WHERE m.activity_id = ?
    `, [activityId]);
  }
  res.render('admin/allocations', { user: req.session.user, departments, activities, materials, allocations, selectedActivity: activityId, currentPage: 'inventory' });
});

router.post('/allocations/update', (req, res) => {
  const { activity_id, allocations } = req.body;
  try {
    for (const key in allocations || {}) {
      const [, deptId, , matId] = key.split('_');
      const quantity = parseInt(allocations[key], 10) || 0;
      const existing = db.get('SELECT id FROM department_allocations WHERE department_id = ? AND material_id = ?', [deptId, matId]);
      if (existing) {
        db.run('UPDATE department_allocations SET allocated_quantity = ? WHERE id = ?', [quantity, existing.id]);
      } else if (quantity > 0) {
        db.run('INSERT INTO department_allocations (department_id, material_id, allocated_quantity, used_quantity, recovered_quantity) VALUES (?, ?, ?, 0, 0)', [deptId, matId, quantity]);
      }
    }
    res.redirect(`/admin/allocations?activity_id=${activity_id}&success=分配成功`);
  } catch (err) {
    res.redirect(`/admin/allocations?activity_id=${activity_id}&error=${encodeURIComponent(err.message)}`);
  }
});

router.get('/recovery', (req, res) => {
  const activityId = req.query.activity_id || '';
  const activities = getActivitiesForUser(req.session.user);
  const departments = getReportDepartments();
  let materials = [];
  if (activityId) {
    materials = db.query(`
      SELECT m.id as material_id, m.name as material_name, m.unit, d.id as department_id, d.name as department_name,
             da.allocated_quantity - da.used_quantity - da.recovered_quantity as remaining
      FROM department_allocations da
      JOIN materials m ON da.material_id = m.id
      JOIN departments d ON da.department_id = d.id
      WHERE m.activity_id = ? AND (da.allocated_quantity - da.used_quantity - da.recovered_quantity) > 0
      ORDER BY d.name, m.name
    `, [activityId]);
  }
  res.render('admin/recovery', { user: req.session.user, activities, departments, materials, selectedActivity: activityId, currentPage: 'inventory' });
});

router.post('/recovery', (req, res) => {
  const { activity_id, material_id, department_id, quantity, remark } = req.body;
  const qty = parseInt(quantity, 10) || 0;
  if (qty <= 0) return res.redirect(`/admin/recovery?activity_id=${activity_id}&error=回收数量必须大于0`);
  const allocation = db.get(`
    SELECT id, allocated_quantity - used_quantity - recovered_quantity as remaining
    FROM department_allocations WHERE department_id = ? AND material_id = ?
  `, [department_id, material_id]);
  if (!allocation || allocation.remaining < qty) {
    return res.redirect(`/admin/recovery?activity_id=${activity_id}&error=可回收库存不足`);
  }
  db.run('UPDATE department_allocations SET recovered_quantity = recovered_quantity + ? WHERE id = ?', [qty, allocation.id]);
  db.run(
    'INSERT INTO usage_records (department_id, material_id, quantity, customer_name, remark, record_type, created_by) VALUES (?, ?, ?, ?, ?, ?, ?)',
    [department_id, material_id, qty, '-', remark || '回收上交', 'recovery', req.session.user.id]
  );
  res.redirect(`/admin/recovery?activity_id=${activity_id}&success=回收成功`);
});

router.get('/export/inventory', (req, res) => {
  const data = buildInventorySheetRows(req.session.user, req.query.activity_id);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), isSuper(req.session.user) ? '超级管理员宣传品库存报表' : '一级人员宣传库存报表');
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=inventory-report.xlsx');
  res.send(buffer);
});

router.get('/export/detail', (req, res) => {
  const data = buildDetailRows(req.session.user, req.query.activity_id);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), isSuper(req.session.user) ? '超级管理员宣传品明细报表' : '一级人员宣传品明细报表');
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=material-detail-report.xlsx');
  res.send(buffer);
});

router.get('/export/usage', (req, res) => {
  const data = db.query(`
    SELECT d.name as '部门/网点', a.name as '活动名称', m.name as '宣传品名称', m.unit as '单位',
           ur.quantity as '数量', ur.customer_name as '领用客户', u.name as '录入员工',
           ur.created_at as raw_time, ur.remark as '备注',
           CASE ur.record_type WHEN 'recovery' THEN '回收' ELSE '领用' END as '记录类型'
    FROM usage_records ur
    JOIN departments d ON ur.department_id = d.id
    JOIN materials m ON ur.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    JOIN users u ON ur.created_by = u.id
    ORDER BY ur.created_at DESC
  `).map(r => {
    r['记录时间'] = formatDateTime(r.raw_time);
    delete r.raw_time;
    return r;
  });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), '领用明细');
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=usage-report.xlsx');
  res.send(buffer);
});

router.get('/template/:type', (req, res) => {
  const { type } = req.params;
  let data;
  let filename;
  if (type === 'users') {
    data = [{ '用户名': 'zhangsan', '姓名': '张三', '密码': '123456', '权限': 2, '部门/网点': '营业一部' }];
    filename = 'user-template.xlsx';
  } else if (type === 'activities') {
    data = [{ '登记日期': '2026-04-23', '活动名称': '春季促销', '宣传品名称': '宣传海报', '数量': 1000 }];
    filename = 'activity-template.xlsx';
  } else if (type === 'departments') {
    data = [{ '部门/网点': '个人金融业务部' }];
    filename = 'department-template.xlsx';
  } else {
    return res.status(400).send('未知模板类型');
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), 'Sheet1');
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename=${filename}`);
  res.send(buffer);
});

module.exports = router;
