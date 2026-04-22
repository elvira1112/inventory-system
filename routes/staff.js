const express = require('express');
const router = express.Router();
const XLSX = require('xlsx');
const bcrypt = require('bcryptjs');
const db = require('../database/init');
const { isAuthenticated } = require('../middleware/auth');

router.use(isAuthenticated);
router.use((req, res, next) => {
  if (req.session.user.role !== 2) {
    return res.redirect('/admin');
  }
  next();
});

router.use((req, res, next) => {
  if (req.session.user.must_change_password && !req.path.startsWith('/password') && req.path !== '/logout') {
    return res.redirect('/staff/password?first=1');
  }
  next();
});

function passwordRuleError(password, user) {
  if (!password || password.length < 6) return '密码长度至少6位';
  if (/^(\d)\1{3,}$/.test(password)) return '密码不能是4位以上重复数字';
  for (let i = 0; i <= password.length - 3; i++) {
    const part = password.slice(i, i + 3);
    if (/^\d{3}$/.test(part)) {
      const nums = part.split('').map(Number);
      if (nums[1] === nums[0] + 1 && nums[2] === nums[1] + 1) return '密码不能包含3位以上连续数字';
    }
  }
  if (user && bcrypt.compareSync(password, user.password)) return '新密码不能和上一次密码相同';
  return null;
}

function formatDateTime(value) {
  if (!value) return '';
  const d = new Date(String(value).replace(' ', 'T') + '+08:00');
  if (Number.isNaN(d.getTime())) return value;
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

function getStaffActivities(departmentId) {
  return db.query(`
    SELECT DISTINCT a.id, a.name
    FROM department_allocations da
    JOIN materials m ON da.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    WHERE da.department_id = ?
    ORDER BY a.name
  `, [departmentId]);
}

function getStaffInventory(departmentId, activityId) {
  return db.query(`
    SELECT
      d.name as department_name,
      a.id as activity_id,
      a.name as activity_name,
      m.id as material_id,
      m.name as material_name,
      m.unit,
      COALESCE(da.allocated_quantity, 0) as allocated,
      COALESCE(da.used_quantity, 0) as used,
      COALESCE(da.recovered_quantity, 0) as recovered,
      COALESCE(da.allocated_quantity, 0) - COALESCE(da.used_quantity, 0) - COALESCE(da.recovered_quantity, 0) as remaining
    FROM department_allocations da
    JOIN departments d ON da.department_id = d.id
    JOIN materials m ON da.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    WHERE da.department_id = ? ${activityId ? 'AND a.id = ?' : ''}
    ORDER BY a.name, m.name
  `, activityId ? [departmentId, activityId] : [departmentId]);
}

router.get('/', (req, res) => {
  const departmentId = req.session.user.department_id;
  const department = db.get('SELECT * FROM departments WHERE id = ?', [departmentId]);
  const inventory = getStaffInventory(departmentId).slice(0, 8);
  const recentUsage = db.query(`
    SELECT a.name as activity_name, m.name as material_name, m.unit, ur.quantity,
           ur.customer_name, ur.remark, ur.created_at, ur.record_type
    FROM usage_records ur
    JOIN materials m ON ur.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    WHERE ur.department_id = ?
    ORDER BY ur.created_at DESC
    LIMIT 10
  `, [departmentId]).map(r => ({ ...r, created_at: formatDateTime(r.created_at) }));

  res.render('staff/dashboard', { user: req.session.user, department, inventory, recentUsage, currentPage: 'dashboard' });
});

router.get('/password', (req, res) => {
  res.render('staff/password', {
    user: req.session.user,
    first: !!req.query.first,
    error: req.query.error || null,
    success: req.query.success || null,
    currentPage: 'password'
  });
});

router.post('/password', (req, res) => {
  const { old_password, new_password, confirm_password } = req.body;
  const user = db.get('SELECT * FROM users WHERE id = ?', [req.session.user.id]);
  if (!user) return res.redirect('/logout');
  if (!bcrypt.compareSync(old_password || '', user.password)) {
    return res.redirect('/staff/password?error=原密码不正确');
  }
  if (new_password !== confirm_password) {
    return res.redirect('/staff/password?error=两次输入的新密码不一致');
  }
  const ruleError = passwordRuleError(new_password, user);
  if (ruleError) return res.redirect('/staff/password?error=' + encodeURIComponent(ruleError));
  const hashedPassword = bcrypt.hashSync(new_password, 10);
  db.run('UPDATE users SET last_password_hash = password, password = ?, must_change_password = 0, initial_password = NULL WHERE id = ?', [hashedPassword, user.id]);
  req.session.user.must_change_password = 0;
  res.redirect('/staff/password?success=密码修改成功');
});

router.get('/inventory', (req, res) => {
  const departmentId = req.session.user.department_id;
  const selectedActivity = req.query.activity_id || '';
  const activities = getStaffActivities(departmentId);
  const inventory = selectedActivity ? getStaffInventory(departmentId, selectedActivity) : [];
  res.render('staff/inventory', { user: req.session.user, activities, inventory, selectedActivity, currentPage: 'inventory' });
});

router.get('/usage', (req, res) => {
  const departmentId = req.session.user.department_id;
  const materials = db.query(`
    SELECT a.name as activity_name, m.id as material_id, m.name as material_name, m.unit,
           COALESCE(da.allocated_quantity, 0) - COALESCE(da.used_quantity, 0) - COALESCE(da.recovered_quantity, 0) as remaining
    FROM department_allocations da
    JOIN materials m ON da.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    WHERE da.department_id = ?
      AND (COALESCE(da.allocated_quantity, 0) - COALESCE(da.used_quantity, 0) - COALESCE(da.recovered_quantity, 0)) > 0
    ORDER BY a.name, m.name
  `, [departmentId]);

  res.render('staff/usage', { user: req.session.user, materials, currentPage: 'usage' });
});

router.post('/usage', (req, res) => {
  const { material_id, quantity, customer_name, remark } = req.body;
  const departmentId = req.session.user.department_id;
  const userId = req.session.user.id;
  const qty = parseInt(quantity, 10) || 0;

  if (qty <= 0) return res.redirect('/staff/usage?error=领用数量必须大于0');
  if (!customer_name || !customer_name.trim()) return res.redirect('/staff/usage?error=请填写领用客户名称');

  try {
    const allocation = db.get(`
      SELECT id, allocated_quantity - used_quantity - recovered_quantity as remaining
      FROM department_allocations
      WHERE department_id = ? AND material_id = ?
    `, [departmentId, material_id]);

    if (!allocation || allocation.remaining < qty) return res.redirect('/staff/usage?error=库存不足');

    db.run(
      'INSERT INTO usage_records (department_id, material_id, quantity, customer_name, remark, record_type, created_by) VALUES (?, ?, ?, ?, ?, ?, ?)',
      [departmentId, material_id, qty, customer_name.trim(), remark || '', 'usage', userId]
    );
    db.run('UPDATE department_allocations SET used_quantity = used_quantity + ? WHERE id = ?', [qty, allocation.id]);
    res.redirect('/staff/usage?success=领用记录添加成功');
  } catch (err) {
    res.redirect('/staff/usage?error=' + encodeURIComponent(err.message));
  }
});

router.get('/history', (req, res) => {
  const departmentId = req.session.user.department_id;
  const mode = req.query.mode === 'mine' ? 'mine' : 'department';
  const params = [departmentId];
  let userFilter = '';
  if (mode === 'mine') {
    userFilter = 'AND ur.created_by = ?';
    params.push(req.session.user.id);
  }

  const records = db.query(`
    SELECT a.name as activity_name, m.name as material_name, m.unit, ur.quantity,
           ur.customer_name, ur.remark, u.name as created_by_name, ur.created_at, ur.record_type
    FROM usage_records ur
    JOIN materials m ON ur.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    JOIN users u ON ur.created_by = u.id
    WHERE ur.department_id = ? ${userFilter}
    ORDER BY ur.created_at DESC
  `, params).map(r => ({ ...r, created_at: formatDateTime(r.created_at) }));

  res.render('staff/history', { user: req.session.user, records, mode, currentPage: 'history' });
});

router.get('/export/inventory', (req, res) => {
  const departmentId = req.session.user.department_id;
  const department = db.get('SELECT name FROM departments WHERE id = ?', [departmentId]);
  const selectedActivity = req.query.activity_id || '';
  const rows = getStaffInventory(departmentId, selectedActivity).map(item => ({
    '部门/网点': item.department_name,
    '营销活动名称': item.activity_name,
    '宣传品名称': item.material_name,
    '分发入库总量': item.allocated,
    '已使用数量': item.used,
    '已回收上交数量': item.recovered,
    '剩余库存数量': item.remaining
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), '二级人员宣传库存报表');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(buildStaffDetailRows(departmentId)), '二级人员宣传品明细报表');
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  const filename = encodeURIComponent(`${department.name}-宣传品库存及明细台账.xlsx`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename=${filename}`);
  res.send(buffer);
});

router.get('/export/usage', (req, res) => {
  const departmentId = req.session.user.department_id;
  const mineOnly = req.query.mode === 'mine';
  const params = [departmentId];
  let userFilter = '';
  if (mineOnly) {
    userFilter = 'AND ur.created_by = ?';
    params.push(req.session.user.id);
  }
  const department = db.get('SELECT name FROM departments WHERE id = ?', [departmentId]);
  const data = db.query(`
    SELECT d.name as '部门/网点', a.name as '活动名称', m.name as '宣传品名称', m.unit as '单位',
           ur.quantity as '领用数量', ur.customer_name as '领用客户', u.name as '录入员工',
           ur.created_at as raw_time, ur.remark as '备注'
    FROM usage_records ur
    JOIN departments d ON ur.department_id = d.id
    JOIN materials m ON ur.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    JOIN users u ON ur.created_by = u.id
    WHERE ur.department_id = ? AND ur.record_type = 'usage' ${userFilter}
    ORDER BY ur.created_at DESC
  `, params).map(r => {
    r['领用时间'] = formatDateTime(r.raw_time);
    delete r.raw_time;
    return r;
  });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), '二级人员宣传品明细报表');
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  const filename = encodeURIComponent(`${department.name}-${mineOnly ? '本人' : '本部门'}领用明细.xlsx`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename=${filename}`);
  res.send(buffer);
});

function buildStaffDetailRows(departmentId) {
  return db.query(`
    SELECT d.name as department_name, a.name as activity_name, m.name as material_name, m.unit,
           ur.quantity, ur.customer_name, u.name as created_by_name, ur.created_at, ur.remark
    FROM usage_records ur
    JOIN departments d ON ur.department_id = d.id
    JOIN materials m ON ur.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    JOIN users u ON ur.created_by = u.id
    WHERE ur.department_id = ? AND ur.record_type = 'usage'
    ORDER BY ur.created_at DESC
  `, [departmentId]).map((r, index) => ({
    '序号': index + 1,
    '部门/网点': r.department_name,
    '活动名称': r.activity_name,
    '宣传品名称': r.material_name,
    '单位': r.unit,
    '领用数量': r.quantity,
    '领用客户': r.customer_name || '-',
    '录入员工': r.created_by_name || '-',
    '领用时间': formatDateTime(r.created_at),
    '备注': r.remark || ''
  }));
}

module.exports = router;
