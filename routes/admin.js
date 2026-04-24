const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const bcrypt = require('bcryptjs');
const path = require('path');
const db = require('../database/init');
const { isAuthenticated, isAdmin, isSuperAdmin } = require('../middleware/auth');
const { formatDateTime, normalizeExcelDate, passwordRuleError } = require('../utils/helpers');

const router = express.Router();

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, path.join(__dirname, '../uploads')),
  filename: (req, file, cb) => cb(null, `${Date.now()}-${file.originalname}`)
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
  '象山路微企业专营支行',
  '钱江支行'
];
const excludedAllocationDepartments = ['风险管理部', '行长室', '综合管理部'];
const comprehensiveDepartmentName = '综合管理部';

router.use(isAuthenticated);
router.use(isAdmin);
router.use((req, res, next) => {
  res.locals.query = req.query;
  next();
});
router.use((req, res, next) => {
  if (req.session.user.role === 1 && req.session.user.must_change_password && !req.path.startsWith('/password') && req.path !== '/logout') {
    return res.redirect('/admin/password?first=1');
  }
  next();
});

function isSuper(user) {
  return user.role === 0;
}

function isComprehensivePrimary(user) {
  if (!user || user.role !== 1 || !user.department_id) return false;
  const department = db.get('SELECT name FROM departments WHERE id = ?', [user.department_id]);
  return !!department && department.name === comprehensiveDepartmentName;
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
  let department = getDepartmentByName(cleanName);
  if (!department) {
    db.run('INSERT INTO departments (name) VALUES (?)', [cleanName]);
    department = { id: db.getLastInsertId(), name: cleanName };
  }
  return department;
}

function roleName(role) {
  if (role === 0) return '超级管理员';
  if (role === 1) return '一级人员';
  return '二级人员';
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
    SELECT a.*, d.name AS department_name
    FROM activities a
    LEFT JOIN departments d ON a.department_id = d.id
    WHERE ${scope.where}
    ORDER BY a.name
  `, scope.params);
}

function getFilterDepartments(selectedDepartmentId) {
  const base = db.query('SELECT * FROM departments');
  const filtered = selectedDepartmentId
    ? base.filter(item => String(item.id) === String(selectedDepartmentId))
    : base.filter(item => !excludedAllocationDepartments.includes(item.name));

  return filtered.sort((a, b) => {
    const ia = managedDepartmentOrder.indexOf(a.name);
    const ib = managedDepartmentOrder.indexOf(b.name);
    if (ia !== -1 || ib !== -1) return (ia === -1 ? 999 : ia) - (ib === -1 ? 999 : ib);
    return a.name.localeCompare(b.name, 'zh-CN');
  });
}

function getMaterialInventoryRows(user, filters = {}) {
  const scope = userScopeWhere(user, 'a');
  const params = [...scope.params];
  const where = [`${scope.where}`];

  if (filters.activityId) {
    where.push('a.id = ?');
    params.push(filters.activityId);
  }

  if (filters.departmentId && isSuper(user)) {
    where.push(`
      EXISTS (
        SELECT 1
        FROM department_allocations da_filter
        WHERE da_filter.material_id = m.id AND da_filter.department_id = ?
      )
    `);
    params.push(filters.departmentId);
  }

  return db.query(`
    SELECT
      d.name AS owner_department,
      a.id AS activity_id,
      a.name AS activity_name,
      m.id AS material_id,
      m.name AS material_name,
      m.unit,
      m.total_quantity,
      COALESCE(SUM(da.allocated_quantity - da.recovered_quantity), 0) AS allocated,
      COALESCE(SUM(da.used_quantity), 0) AS used,
      COALESCE(SUM(da.recovered_quantity), 0) AS recovered,
      m.total_quantity - COALESCE(SUM(da.allocated_quantity - da.recovered_quantity), 0) AS unallocated
    FROM materials m
    JOIN activities a ON m.activity_id = a.id
    LEFT JOIN departments d ON a.department_id = d.id
    LEFT JOIN department_allocations da ON da.material_id = m.id
    WHERE ${where.join(' AND ')}
    GROUP BY m.id
    ORDER BY d.name, a.name, m.name
  `, params);
}

function getDepartmentAllocationMap(activityId, departmentId) {
  const params = [activityId];
  let extraWhere = '';
  if (departmentId) {
    extraWhere = 'AND d.id = ?';
    params.push(departmentId);
  }

  const rows = db.query(`
    SELECT
      da.material_id,
      d.name AS department_name,
      d.id AS department_id,
      COALESCE(da.allocated_quantity, 0) - COALESCE(da.recovered_quantity, 0) AS allocated,
      COALESCE(da.used_quantity, 0) AS used,
      COALESCE(da.recovered_quantity, 0) AS recovered,
      COALESCE(da.allocated_quantity, 0) - COALESCE(da.used_quantity, 0) - COALESCE(da.recovered_quantity, 0) AS remaining
    FROM department_allocations da
    JOIN departments d ON da.department_id = d.id
    JOIN materials m ON da.material_id = m.id
    WHERE m.activity_id = ? ${extraWhere}
  `, params);

  return Object.fromEntries(rows.map(row => [`${row.material_id}_${row.department_name}`, row]));
}

function buildInventorySheetRows(user, filters = {}) {
  const materials = getMaterialInventoryRows(user, filters);
  const departments = getFilterDepartments(filters.departmentId);

  return materials.map(item => {
    const row = {
      '活动负责部门': item.owner_department || '',
      '活动名称': item.activity_name,
      '宣传品名称': item.material_name,
      '总量': item.total_quantity,
      '合计已分配': item.allocated,
      '剩余可分配': item.unallocated
    };

    const allocs = db.query(`
      SELECT
        d.name AS department_name,
        COALESCE(da.allocated_quantity, 0) - COALESCE(da.recovered_quantity, 0) AS allocated,
        COALESCE(da.used_quantity, 0) AS used,
        COALESCE(da.recovered_quantity, 0) AS recovered
      FROM departments d
      LEFT JOIN department_allocations da ON d.id = da.department_id AND da.material_id = ?
      ${filters.departmentId ? 'WHERE d.id = ?' : ''}
    `, filters.departmentId ? [item.material_id, filters.departmentId] : [item.material_id]);
    const byDepartment = Object.fromEntries(allocs.map(record => [record.department_name, record]));

    departments.forEach(department => {
      const allocation = byDepartment[department.name] || { allocated: 0, used: 0, recovered: 0 };
      row[`${department.name}-已分配`] = allocation.allocated;
      row[`${department.name}-已领用`] = allocation.used;
      row[`${department.name}-剩余库存`] = allocation.allocated - allocation.used;
    });

    return row;
  });
}

function buildDetailRows(user, filters = {}) {
  const scope = userScopeWhere(user, 'a');
  const params = [...scope.params];
  const where = [`${scope.where}`];

  if (filters.activityId) {
    where.push('a.id = ?');
    params.push(filters.activityId);
  }

  if (filters.departmentId) {
    where.push('d.id = ?');
    params.push(filters.departmentId);
  }

  const allocationRows = db.query(`
    SELECT
      od.name AS owner_department,
      a.name AS activity_name,
      m.name AS material_name,
      d.name AS department_name,
      COALESCE(da.allocated_quantity, 0) - COALESCE(da.recovered_quantity, 0) AS quantity
    FROM department_allocations da
    JOIN materials m ON da.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    LEFT JOIN departments od ON a.department_id = od.id
    JOIN departments d ON da.department_id = d.id
    WHERE ${where.join(' AND ')} AND (COALESCE(da.allocated_quantity, 0) - COALESCE(da.recovered_quantity, 0)) > 0
    ORDER BY d.name, a.name, m.name
  `, params).map(row => ({
    '活动负责部门': row.owner_department || '',
    '部门/网点': row.department_name,
    '活动名称': row.activity_name,
    '宣传品名称': row.material_name,
    '类型': '分配',
    '数量': row.quantity,
    '时间': '',
    '备注': ''
  }));

  const recoveryRows = db.query(`
    SELECT
      od.name AS owner_department,
      a.name AS activity_name,
      m.name AS material_name,
      d.name AS department_name,
      ur.quantity,
      ur.created_at,
      ur.remark
    FROM usage_records ur
    JOIN materials m ON ur.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    LEFT JOIN departments od ON a.department_id = od.id
    JOIN departments d ON ur.department_id = d.id
    WHERE ${where.join(' AND ')} AND ur.record_type = 'recovery'
    ORDER BY d.name, a.name, m.name, ur.created_at DESC
  `, params).map(row => ({
    '活动负责部门': row.owner_department || '',
    '部门/网点': row.department_name,
    '活动名称': row.activity_name,
    '宣传品名称': row.material_name,
    '类型': '回收',
    '数量': row.quantity,
    '时间': formatDateTime(row.created_at),
    '备注': row.remark || ''
  }));

  return [...allocationRows, ...recoveryRows].sort((a, b) => {
    return `${a['部门/网点']}${a['活动名称']}${a['宣传品名称']}${a['类型']}${a['时间']}`.localeCompare(
      `${b['部门/网点']}${b['活动名称']}${b['宣传品名称']}${b['类型']}${b['时间']}`,
      'zh-CN'
    );
  });
}

function buildMergedSheet(rows, sheetName, mergeColumns = []) {
  const headers = rows.length > 0 ? Object.keys(rows[0]) : [];
  const data = [headers, ...rows.map(row => headers.map(key => row[key] ?? ''))];
  const sheet = XLSX.utils.aoa_to_sheet(data);

  sheet['!cols'] = headers.map(header => ({
    wch: Math.max(header.length + 4, ...rows.map(row => String(row[header] ?? '').length + 2), 12)
  }));

  if (mergeColumns.length > 0 && rows.length > 1) {
    const merges = [];
    mergeColumns.forEach(columnIndex => {
      let start = 1;
      for (let rowIndex = 2; rowIndex <= data.length; rowIndex++) {
        const current = rowIndex < data.length ? data[rowIndex][columnIndex] : null;
        const previous = data[rowIndex - 1][columnIndex];
        const previousKeys = mergeColumns
          .filter(index => index < columnIndex)
          .map(index => data[rowIndex - 1][index])
          .join('|');
        const currentKeys = rowIndex < data.length
          ? mergeColumns.filter(index => index < columnIndex).map(index => data[rowIndex][index]).join('|')
          : '';

        if (current !== previous || previousKeys !== currentKeys) {
          if (rowIndex - start > 1 && previous !== '') {
            merges.push({
              s: { r: start, c: columnIndex },
              e: { r: rowIndex - 1, c: columnIndex }
            });
          }
          start = rowIndex;
        }
      }
    });
    sheet['!merges'] = merges;
  }

  return { sheet, sheetName };
}

router.get('/', (req, res) => {
  const scope = userScopeWhere(req.session.user, 'a');
  const stats = {
    departmentCount: db.get('SELECT COUNT(*) AS count FROM departments').count,
    primaryUserCount: db.get('SELECT COUNT(*) AS count FROM users WHERE role = 1').count,
    staffUserCount: db.get('SELECT COUNT(*) AS count FROM users WHERE role = 2').count,
    activityCount: db.get(`SELECT COUNT(*) AS count FROM activities a WHERE ${scope.where}`, scope.params).count,
    materialCount: db.get(`
      SELECT COUNT(*) AS count
      FROM materials m
      JOIN activities a ON m.activity_id = a.id
      WHERE ${scope.where}
    `, scope.params).count
  };

  res.render('admin/dashboard', {
    user: req.session.user,
    stats,
    canDeleteActivities: isComprehensivePrimary(req.session.user),
    currentPage: 'dashboard'
  });
});

router.get('/password', (req, res) => {
  if (req.session.user.role !== 1) {
    return res.redirect('/admin');
  }

  res.render('admin/password', {
    user: req.session.user,
    first: !!req.query.first,
    error: req.query.error || null,
    success: req.query.success || null,
    currentPage: 'password'
  });
});

router.post('/password', (req, res) => {
  if (req.session.user.role !== 1) {
    return res.redirect('/admin');
  }

  const { old_password, new_password, confirm_password } = req.body;
  const user = db.get('SELECT * FROM users WHERE id = ?', [req.session.user.id]);

  if (!user) return res.redirect('/logout');
  if (!bcrypt.compareSync(old_password || '', user.password)) {
    return res.redirect('/admin/password?error=' + encodeURIComponent('原密码不正确'));
  }
  if (new_password !== confirm_password) {
    return res.redirect('/admin/password?error=' + encodeURIComponent('两次输入的新密码不一致'));
  }

  const ruleError = passwordRuleError(new_password, user);
  if (ruleError) {
    return res.redirect('/admin/password?error=' + encodeURIComponent(ruleError));
  }

  const hashedPassword = bcrypt.hashSync(new_password, 10);
  db.run(
    'UPDATE users SET last_password_hash = password, password = ?, must_change_password = 0, initial_password = NULL WHERE id = ?',
    [hashedPassword, user.id]
  );
  req.session.user.must_change_password = 0;
  res.redirect('/admin/password?success=' + encodeURIComponent('密码修改成功'));
});

router.get('/users', isSuperAdmin, (req, res) => {
  const users = db.query(`
    SELECT u.*, d.name AS department_name
    FROM users u
    LEFT JOIN departments d ON u.department_id = d.id
    ORDER BY u.role, u.id DESC
  `).map(user => ({ ...user, role_name: roleName(user.role) }));
  const departments = db.query('SELECT * FROM departments ORDER BY name');

  res.render('admin/users', {
    user: req.session.user,
    users,
    departments,
    currentPage: 'users'
  });
});

router.post('/users/import', isSuperAdmin, upload.single('file'), (req, res) => {
  try {
    if (!req.file) return res.redirect('/admin/users?error=' + encodeURIComponent('请选择文件'));
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

      const department = role === 0 ? null : ensureDepartment(deptName);
      if (role !== 0 && !department) continue;
      if (db.get('SELECT id FROM users WHERE username = ?', [username])) continue;

      db.run(
        'INSERT INTO users (username, password, name, role, department_id, must_change_password, initial_password) VALUES (?, ?, ?, ?, ?, 1, ?)',
        [username, bcrypt.hashSync(password, 10), name, role, department ? department.id : null, password]
      );
      imported++;
    }

    res.redirect(`/admin/users?success=${encodeURIComponent(`成功导入 ${imported} 名人员`)}`);
  } catch (err) {
    res.redirect('/admin/users?error=' + encodeURIComponent(`导入失败: ${err.message}`));
  }
});

router.post('/users/add', isSuperAdmin, (req, res) => {
  const { username, name, password, department_id, role } = req.body;

  try {
    if (db.get('SELECT id FROM users WHERE username = ?', [username])) {
      return res.redirect('/admin/users?error=' + encodeURIComponent('用户名已存在'));
    }

    const rawPassword = password || '123456';
    const userRole = Number(role || 2);
    db.run(
      'INSERT INTO users (username, password, name, role, department_id, must_change_password, initial_password) VALUES (?, ?, ?, ?, ?, 1, ?)',
      [username, bcrypt.hashSync(rawPassword, 10), name, userRole, userRole === 0 ? null : department_id, rawPassword]
    );
    res.redirect('/admin/users?success=' + encodeURIComponent('人员添加成功'));
  } catch (err) {
    res.redirect('/admin/users?error=' + encodeURIComponent(err.message));
  }
});

router.post('/users/update/:id', isSuperAdmin, (req, res) => {
  const { name, role, department_id } = req.body;
  const userRole = Number(role || 2);
  db.run(
    'UPDATE users SET name = ?, role = ?, department_id = ? WHERE id = ? AND username <> ?',
    [name, userRole, userRole === 0 ? null : department_id, req.params.id, 'admin']
  );
  res.redirect('/admin/users?success=' + encodeURIComponent('人员信息已更新'));
});

router.post('/users/reset-password/:id', isSuperAdmin, (req, res) => {
  const rawPassword = req.body.password || '123456';
  const hashedPassword = bcrypt.hashSync(rawPassword, 10);
  db.run(
    'UPDATE users SET last_password_hash = password, password = ?, must_change_password = 1, initial_password = ? WHERE id = ?',
    [hashedPassword, rawPassword, req.params.id]
  );
  res.redirect('/admin/users?success=' + encodeURIComponent('密码已重置'));
});

router.post('/users/delete/:id', isSuperAdmin, (req, res) => {
  db.run("DELETE FROM users WHERE id = ? AND username <> 'admin'", [req.params.id]);
  res.redirect('/admin/users?success=' + encodeURIComponent('人员已删除'));
});

router.get('/departments', isSuperAdmin, (req, res) => {
  const departments = db.query(`
    SELECT d.*, COUNT(u.id) AS user_count
    FROM departments d
    LEFT JOIN users u ON d.id = u.department_id
    GROUP BY d.id
    ORDER BY d.name
  `);

  res.render('admin/departments', {
    user: req.session.user,
    departments,
    currentPage: 'departments'
  });
});

router.post('/departments/add', isSuperAdmin, (req, res) => {
  try {
    ensureDepartment(req.body.name);
    res.redirect('/admin/departments?success=' + encodeURIComponent('部门/网点已添加'));
  } catch (err) {
    res.redirect('/admin/departments?error=' + encodeURIComponent(err.message));
  }
});

router.post('/departments/import', isSuperAdmin, upload.single('file'), (req, res) => {
  try {
    if (!req.file) return res.redirect('/admin/departments?error=' + encodeURIComponent('请选择文件'));
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

    res.redirect(`/admin/departments?success=${encodeURIComponent(`成功导入 ${imported} 个部门/网点`)}`);
  } catch (err) {
    res.redirect('/admin/departments?error=' + encodeURIComponent(`导入失败: ${err.message}`));
  }
});

router.post('/departments/update/:id', isSuperAdmin, (req, res) => {
  db.run('UPDATE departments SET name = ? WHERE id = ?', [req.body.name, req.params.id]);
  res.redirect('/admin/departments?success=' + encodeURIComponent('部门/网点已更新'));
});

router.post('/departments/delete/:id', isSuperAdmin, (req, res) => {
  const used = db.get('SELECT COUNT(*) AS count FROM users WHERE department_id = ?', [req.params.id]).count;
  if (used > 0) {
    return res.redirect('/admin/departments?error=' + encodeURIComponent('该部门/网点下仍有人员，不能删除'));
  }
  db.run('DELETE FROM departments WHERE id = ?', [req.params.id]);
  res.redirect('/admin/departments?success=' + encodeURIComponent('部门/网点已删除'));
});

router.get('/activities', (req, res) => {
  const scope = userScopeWhere(req.session.user, 'a');
  const activities = db.query(`
    SELECT
      a.*,
      d.name AS department_name,
      COUNT(m.id) AS material_count,
      COALESCE(SUM(m.total_quantity), 0) AS total_items,
      COALESCE(SUM(CASE WHEN da.allocated_quantity > 0 THEN 1 ELSE 0 END), 0) AS distributed_count
    FROM activities a
    LEFT JOIN departments d ON a.department_id = d.id
    LEFT JOIN materials m ON a.id = m.activity_id
    LEFT JOIN department_allocations da ON da.material_id = m.id
    WHERE ${scope.where}
    GROUP BY a.id
    ORDER BY a.id DESC
  `, scope.params).map(activity => ({
    ...activity,
    can_delete: isComprehensivePrimary(req.session.user) && Number(activity.distributed_count || 0) === 0
  }));

  res.render('admin/activities', {
    user: req.session.user,
    activities,
    canDeleteActivities: isComprehensivePrimary(req.session.user),
    currentPage: 'activities'
  });
});

router.post('/activities/import', upload.single('file'), (req, res) => {
  try {
    if (!req.file) return res.redirect('/admin/activities?error=' + encodeURIComponent('请选择文件'));
    if (!req.session.user.department_id) {
      return res.redirect('/admin/activities?error=' + encodeURIComponent('当前账号未绑定部门/网点，不能导入活动'));
    }

    const workbook = XLSX.readFile(req.file.path);
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
    const ownerDepartment = db.get('SELECT * FROM departments WHERE id = ?', [req.session.user.department_id]);
    let imported = 0;

    if (!ownerDepartment) {
      return res.redirect('/admin/activities?error=' + encodeURIComponent('当前账号所属部门/网点不存在，不能导入活动'));
    }

    for (const row of data) {
      const activityName = row['活动名称'] || row['营销活动名称'] || row.activity;
      const materialName = row['宣传品名称'] || row.material;
      const quantity = parseInt(row['数量'] || row.quantity, 10) || 0;
      const unit = row['单位'] || row.unit || '份';
      const registrationDate = normalizeExcelDate(row['登记日期'] || row.registration_date, XLSX);
      if (!activityName || !materialName) continue;

      let activity = db.get('SELECT id FROM activities WHERE name = ?', [activityName]);
      if (!activity) {
        db.run(
          'INSERT INTO activities (name, department_id, registration_date) VALUES (?, ?, ?)',
          [activityName, ownerDepartment.id, registrationDate]
        );
        activity = { id: db.getLastInsertId() };
      } else {
        db.run(
          'UPDATE activities SET department_id = ?, registration_date = COALESCE(NULLIF(?, \'\'), registration_date) WHERE id = ?',
          [ownerDepartment.id, registrationDate, activity.id]
        );
      }

      const existingMaterial = db.get('SELECT id FROM materials WHERE activity_id = ? AND name = ?', [activity.id, materialName]);
      if (existingMaterial) {
        db.run('UPDATE materials SET total_quantity = total_quantity + ? WHERE id = ?', [quantity, existingMaterial.id]);
      } else {
        db.run('INSERT INTO materials (activity_id, name, unit, total_quantity) VALUES (?, ?, ?, ?)', [activity.id, materialName, unit, quantity]);
      }
      imported++;
    }

    res.redirect(`/admin/activities?success=${encodeURIComponent(`成功导入 ${imported} 条记录`)}`);
  } catch (err) {
    res.redirect('/admin/activities?error=' + encodeURIComponent(`导入失败: ${err.message}`));
  }
});

router.post('/activities/delete/:id', (req, res) => {
  if (!isComprehensivePrimary(req.session.user)) {
    return res.redirect('/admin/activities?error=' + encodeURIComponent('只有综合管理部的一级人员可删除活动'));
  }

  const activity = db.get('SELECT * FROM activities WHERE id = ?', [req.params.id]);
  if (!activity) {
    return res.redirect('/admin/activities?error=' + encodeURIComponent('活动不存在'));
  }

  const distributed = db.get(`
    SELECT COUNT(*) AS count
    FROM department_allocations da
    JOIN materials m ON da.material_id = m.id
    WHERE m.activity_id = ? AND da.allocated_quantity > 0
  `, [req.params.id]);

  if (distributed.count > 0) {
    return res.redirect('/admin/activities?error=' + encodeURIComponent('活动下已有宣传品派发，不能删除'));
  }

  db.run('DELETE FROM materials WHERE activity_id = ?', [req.params.id]);
  db.run('DELETE FROM activities WHERE id = ?', [req.params.id]);
  res.redirect('/admin/activities?success=' + encodeURIComponent('活动已删除'));
});

router.get('/activities/:id', (req, res) => {
  const scope = userScopeWhere(req.session.user, 'a');
  const activity = db.get(`
    SELECT a.*, d.name AS department_name
    FROM activities a
    LEFT JOIN departments d ON a.department_id = d.id
    WHERE a.id = ? AND ${scope.where}
  `, [req.params.id, ...scope.params]);

  if (!activity) {
    return res.redirect('/admin/activities?error=' + encodeURIComponent('活动不存在或无权查看'));
  }

  const materials = db.query(`
    SELECT
      m.*,
      COALESCE(SUM(da.allocated_quantity - da.recovered_quantity), 0) AS allocated,
      COALESCE(SUM(da.used_quantity), 0) AS used,
      COALESCE(SUM(da.recovered_quantity), 0) AS recovered
    FROM materials m
    LEFT JOIN department_allocations da ON m.id = da.material_id
    WHERE m.activity_id = ?
    GROUP BY m.id
  `, [req.params.id]);

  res.render('admin/activity-detail', {
    user: req.session.user,
    activity,
    materials,
    currentPage: 'activities'
  });
});

router.get('/inventory', (req, res) => {
  const selectedActivity = req.query.activity_id || '';
  const selectedDepartment = req.query.department_id || '';
  const activities = isSuper(req.session.user)
    ? db.query(`
        SELECT a.id, a.name, d.name AS department_name
        FROM activities a
        LEFT JOIN departments d ON a.department_id = d.id
        ORDER BY a.name
      `)
    : getActivitiesForUser(req.session.user);
  const departments = getFilterDepartments(selectedDepartment);
  const showData = isSuper(req.session.user) || !!selectedActivity;
  const inventory = showData ? getMaterialInventoryRows(req.session.user, { activityId: selectedActivity, departmentId: selectedDepartment }) : [];
  const detailRows = showData ? buildDetailRows(req.session.user, { activityId: selectedActivity, departmentId: selectedDepartment }) : [];
  const departmentStocks = selectedActivity ? getDepartmentAllocationMap(selectedActivity, selectedDepartment) : {};

  res.render('admin/inventory', {
    user: req.session.user,
    activities,
    allDepartments: db.query('SELECT * FROM departments ORDER BY name'),
    inventory,
    detailRows,
    departments,
    departmentStocks,
    selectedActivity,
    selectedDepartment,
    currentPage: 'inventory'
  });
});

router.get('/allocations', (req, res) => {
  const activities = getActivitiesForUser(req.session.user);
  const departments = getFilterDepartments();
  const activityId = req.query.activity_id || '';
  let materials = [];
  let allocations = [];

  if (activityId) {
    const scope = userScopeWhere(req.session.user, 'a');
    const activity = db.get(`SELECT a.id FROM activities a WHERE a.id = ? AND ${scope.where}`, [activityId, ...scope.params]);
    if (!activity) {
      return res.redirect('/admin/inventory?error=' + encodeURIComponent('活动不存在或无权操作'));
    }

    materials = db.query(`
      SELECT
        m.*,
        COALESCE(SUM(da.allocated_quantity - da.recovered_quantity), 0) AS total_allocated
      FROM materials m
      LEFT JOIN department_allocations da ON m.id = da.material_id
      WHERE m.activity_id = ?
      GROUP BY m.id
    `, [activityId]).map(material => ({
      ...material,
      remaining_allocable: material.total_quantity - material.total_allocated
    }));

    allocations = db.query(`
      SELECT
        d.id AS department_id,
        d.name AS department_name,
        m.id AS material_id,
        COALESCE(da.allocated_quantity, 0) - COALESCE(da.recovered_quantity, 0) AS allocated,
        COALESCE(da.used_quantity, 0) AS used,
        COALESCE(da.recovered_quantity, 0) AS recovered,
        COALESCE(da.allocated_quantity, 0) - COALESCE(da.used_quantity, 0) - COALESCE(da.recovered_quantity, 0) AS remaining
      FROM departments d
      CROSS JOIN materials m
      LEFT JOIN department_allocations da ON d.id = da.department_id AND m.id = da.material_id
      WHERE m.activity_id = ? AND d.name NOT IN (${excludedAllocationDepartments.map(() => '?').join(', ')})
      ORDER BY d.name, m.name
    `, [activityId, ...excludedAllocationDepartments]);
  }

  res.render('admin/allocations', {
    user: req.session.user,
    departments,
    activities,
    materials,
    allocations,
    selectedActivity: activityId,
    currentPage: 'inventory'
  });
});

router.post('/allocations/update', (req, res) => {
  const { activity_id, allocations = {} } = req.body;

  try {
    const materials = db.query('SELECT id, total_quantity, name FROM materials WHERE activity_id = ?', [activity_id]);
    const materialSummary = Object.fromEntries(materials.map(material => [material.id, { ...material, total: 0 }]));
    const desiredValues = {};

    for (const [key, value] of Object.entries(allocations)) {
      const match = key.match(/^dept_(\d+)_mat_(\d+)$/);
      if (!match) continue;
      const departmentId = Number(match[1]);
      const materialId = Number(match[2]);
      const desiredAllocated = Math.max(parseInt(value, 10) || 0, 0);
      const existing = db.get('SELECT * FROM department_allocations WHERE department_id = ? AND material_id = ?', [departmentId, materialId]);
      const used = existing ? existing.used_quantity || 0 : 0;
      const recovered = existing ? existing.recovered_quantity || 0 : 0;

      if (desiredAllocated < used) {
        const material = materialSummary[materialId];
        return res.redirect(`/admin/allocations?activity_id=${activity_id}&error=${encodeURIComponent(`${material ? material.name : '宣传品'}的分配数不能小于已领用数`)}`);
      }

      desiredValues[key] = { departmentId, materialId, desiredAllocated, existing, recovered };
      if (materialSummary[materialId]) {
        materialSummary[materialId].total += desiredAllocated;
      }
    }

    const overflowMaterial = Object.values(materialSummary).find(material => material.total > material.total_quantity);
    if (overflowMaterial) {
      return res.redirect(`/admin/allocations?activity_id=${activity_id}&error=${encodeURIComponent(`${overflowMaterial.name}分配数量超过剩余可分配数`)}`);
    }

    Object.values(desiredValues).forEach(item => {
      const grossAllocated = item.desiredAllocated + item.recovered;
      if (item.existing) {
        db.run('UPDATE department_allocations SET allocated_quantity = ? WHERE id = ?', [grossAllocated, item.existing.id]);
      } else if (grossAllocated > 0) {
        db.run(
          'INSERT INTO department_allocations (department_id, material_id, allocated_quantity, used_quantity, recovered_quantity) VALUES (?, ?, ?, 0, 0)',
          [item.departmentId, item.materialId, grossAllocated]
        );
      }
    });

    res.redirect(`/admin/allocations?activity_id=${activity_id}&success=${encodeURIComponent('分配成功')}`);
  } catch (err) {
    res.redirect(`/admin/allocations?activity_id=${activity_id}&error=${encodeURIComponent(err.message)}`);
  }
});

router.get('/recovery', (req, res) => {
  const activityId = req.query.activity_id || '';
  const activities = getActivitiesForUser(req.session.user);
  const departments = getFilterDepartments();
  let materials = [];

  if (activityId) {
    materials = db.query(`
      SELECT
        m.id AS material_id,
        m.name AS material_name,
        m.unit,
        d.id AS department_id,
        d.name AS department_name,
        da.allocated_quantity - da.used_quantity - da.recovered_quantity AS remaining
      FROM department_allocations da
      JOIN materials m ON da.material_id = m.id
      JOIN departments d ON da.department_id = d.id
      WHERE m.activity_id = ? AND (da.allocated_quantity - da.used_quantity - da.recovered_quantity) > 0
      ORDER BY d.name, m.name
    `, [activityId]);
  }

  res.render('admin/recovery', {
    user: req.session.user,
    activities,
    departments,
    materials,
    selectedActivity: activityId,
    currentPage: 'inventory'
  });
});

router.post('/recovery', (req, res) => {
  const { activity_id, material_id, department_id, quantity, remark } = req.body;
  const qty = parseInt(quantity, 10) || 0;

  if (qty <= 0) {
    return res.redirect(`/admin/recovery?activity_id=${activity_id}&error=${encodeURIComponent('回收数量必须大于0')}`);
  }

  const allocation = db.get(`
    SELECT id, allocated_quantity - used_quantity - recovered_quantity AS remaining
    FROM department_allocations
    WHERE department_id = ? AND material_id = ?
  `, [department_id, material_id]);

  if (!allocation || allocation.remaining < qty) {
    return res.redirect(`/admin/recovery?activity_id=${activity_id}&error=${encodeURIComponent('可回收库存不足')}`);
  }

  db.run('UPDATE department_allocations SET recovered_quantity = recovered_quantity + ? WHERE id = ?', [qty, allocation.id]);
  db.run(
    'INSERT INTO usage_records (department_id, material_id, quantity, customer_name, remark, record_type, created_by) VALUES (?, ?, ?, ?, ?, ?, ?)',
    [department_id, material_id, qty, '-', remark || '回收上交', 'recovery', req.session.user.id]
  );
  res.redirect(`/admin/recovery?activity_id=${activity_id}&success=${encodeURIComponent('回收成功')}`);
});

router.get('/export/inventory', (req, res) => {
  const filters = {
    activityId: req.query.activity_id || '',
    departmentId: req.query.department_id || ''
  };
  const data = buildInventorySheetRows(req.session.user, filters);
  const wb = XLSX.utils.book_new();
  const { sheet, sheetName } = buildMergedSheet(data, '库存报表', [0, 1]);
  XLSX.utils.book_append_sheet(wb, sheet, sheetName);
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=inventory-report.xlsx');
  res.send(buffer);
});

router.get('/export/detail', (req, res) => {
  const filters = {
    activityId: req.query.activity_id || '',
    departmentId: req.query.department_id || ''
  };
  const data = buildDetailRows(req.session.user, filters);
  const wb = XLSX.utils.book_new();
  const { sheet, sheetName } = buildMergedSheet(data, '明细报表', [0, 1, 2]);
  XLSX.utils.book_append_sheet(wb, sheet, sheetName);
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=material-detail-report.xlsx');
  res.send(buffer);
});

router.get('/export/usage', (req, res) => {
  const data = db.query(`
    SELECT
      d.name AS '部门/网点',
      a.name AS '活动名称',
      m.name AS '宣传品名称',
      m.unit AS '单位',
      ur.quantity AS '数量',
      ur.customer_name AS '领用客户',
      u.name AS '录入员工',
      ur.created_at AS raw_time,
      ur.remark AS '备注',
      CASE ur.record_type WHEN 'recovery' THEN '回收' ELSE '领用' END AS '记录类型'
    FROM usage_records ur
    JOIN departments d ON ur.department_id = d.id
    JOIN materials m ON ur.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    JOIN users u ON ur.created_by = u.id
    ORDER BY ur.created_at DESC
  `).map(row => ({
    ...row,
    时间: formatDateTime(row.raw_time),
    raw_time: undefined
  })).map(({ raw_time, ...rest }) => rest);

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
    data = [{ 用户名: 'zhangsan', 姓名: '张三', 密码: '123456', 权限: 2, '部门/网点': '营业一部' }];
    filename = 'user-template.xlsx';
  } else if (type === 'activities') {
    data = [{ 登记日期: '2026-04-23', 活动名称: '春季促销', 宣传品名称: '宣传海报', 数量: 1000, 单位: '张' }];
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
