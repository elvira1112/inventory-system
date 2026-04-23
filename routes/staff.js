const express = require('express');
const XLSX = require('xlsx');
const bcrypt = require('bcryptjs');
const db = require('../database/init');
const { isAuthenticated } = require('../middleware/auth');
const { formatDateTime, maskCustomerName, passwordRuleError } = require('../utils/helpers');

const router = express.Router();

router.use(isAuthenticated);
router.use((req, res, next) => {
  res.locals.query = req.query;
  next();
});
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
      d.name AS department_name,
      a.id AS activity_id,
      a.name AS activity_name,
      m.id AS material_id,
      m.name AS material_name,
      m.unit,
      COALESCE(da.allocated_quantity, 0) - COALESCE(da.recovered_quantity, 0) AS allocated,
      COALESCE(da.used_quantity, 0) AS used,
      COALESCE(da.recovered_quantity, 0) AS recovered,
      COALESCE(da.allocated_quantity, 0) - COALESCE(da.used_quantity, 0) - COALESCE(da.recovered_quantity, 0) AS remaining
    FROM department_allocations da
    JOIN departments d ON da.department_id = d.id
    JOIN materials m ON da.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    WHERE da.department_id = ? ${activityId ? 'AND a.id = ?' : ''}
    ORDER BY a.name, m.name
  `, activityId ? [departmentId, activityId] : [departmentId]);
}

function buildStaffDetailRows(departmentId, activityId) {
  return db.query(`
    SELECT
      d.name AS department_name,
      a.name AS activity_name,
      m.name AS material_name,
      m.unit,
      ur.quantity,
      ur.customer_name,
      u.name AS created_by_name,
      ur.created_at,
      ur.remark
    FROM usage_records ur
    JOIN departments d ON ur.department_id = d.id
    JOIN materials m ON ur.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    JOIN users u ON ur.created_by = u.id
    WHERE ur.department_id = ? AND ur.record_type = 'usage' ${activityId ? 'AND a.id = ?' : ''}
    ORDER BY a.name, m.name, ur.created_at DESC
  `, activityId ? [departmentId, activityId] : [departmentId]).map((row, index) => ({
    '序号': index + 1,
    '部门/网点': row.department_name,
    '活动名称': row.activity_name,
    '宣传品名称': row.material_name,
    '单位': row.unit,
    '领用数量': row.quantity,
    '领用客户': row.customer_name || '-',
    '录入员工': row.created_by_name || '-',
    '时间': formatDateTime(row.created_at),
    '备注': row.remark || ''
  }));
}

function buildMergedSheet(rows, sheetName, mergeColumns = []) {
  const headers = rows.length > 0 ? Object.keys(rows[0]) : [];
  const data = [headers, ...rows.map(row => headers.map(key => row[key] ?? ''))];
  const sheet = XLSX.utils.aoa_to_sheet(data);

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
  const departmentId = req.session.user.department_id;
  const department = db.get('SELECT * FROM departments WHERE id = ?', [departmentId]);
  const inventory = getStaffInventory(departmentId).slice(0, 8);
  const recentUsage = db.query(`
    SELECT a.name AS activity_name, m.name AS material_name, m.unit, ur.quantity,
           ur.customer_name, ur.remark, ur.created_at, ur.record_type
    FROM usage_records ur
    JOIN materials m ON ur.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    WHERE ur.department_id = ?
    ORDER BY ur.created_at DESC
    LIMIT 10
  `, [departmentId]).map(row => ({
    ...row,
    customer_name_masked: maskCustomerName(row.customer_name),
    created_at: formatDateTime(row.created_at)
  }));

  res.render('staff/dashboard', {
    user: req.session.user,
    department,
    inventory,
    recentUsage,
    currentPage: 'dashboard'
  });
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
    return res.redirect('/staff/password?error=' + encodeURIComponent('原密码不正确'));
  }
  if (new_password !== confirm_password) {
    return res.redirect('/staff/password?error=' + encodeURIComponent('两次输入的新密码不一致'));
  }

  const ruleError = passwordRuleError(new_password, user);
  if (ruleError) {
    return res.redirect('/staff/password?error=' + encodeURIComponent(ruleError));
  }

  const hashedPassword = bcrypt.hashSync(new_password, 10);
  db.run(
    'UPDATE users SET last_password_hash = password, password = ?, must_change_password = 0, initial_password = NULL WHERE id = ?',
    [hashedPassword, user.id]
  );
  req.session.user.must_change_password = 0;
  res.redirect('/staff/password?success=' + encodeURIComponent('密码修改成功'));
});

router.get('/inventory', (req, res) => {
  const departmentId = req.session.user.department_id;
  const selectedActivity = req.query.activity_id || '';
  const activities = getStaffActivities(departmentId);
  const inventory = selectedActivity ? getStaffInventory(departmentId, selectedActivity) : [];

  res.render('staff/inventory', {
    user: req.session.user,
    activities,
    inventory,
    selectedActivity,
    currentPage: 'inventory'
  });
});

router.get('/usage', (req, res) => {
  const departmentId = req.session.user.department_id;
  const materials = db.query(`
    SELECT
      a.name AS activity_name,
      m.id AS material_id,
      m.name AS material_name,
      m.unit,
      COALESCE(da.allocated_quantity, 0) - COALESCE(da.used_quantity, 0) - COALESCE(da.recovered_quantity, 0) AS remaining
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

  if (qty <= 0) {
    return res.redirect('/staff/usage?error=' + encodeURIComponent('领用数量必须大于0'));
  }
  if (!customer_name || !customer_name.trim()) {
    return res.redirect('/staff/usage?error=' + encodeURIComponent('请填写领用客户名称'));
  }

  try {
    const allocation = db.get(`
      SELECT id, allocated_quantity - used_quantity - recovered_quantity AS remaining
      FROM department_allocations
      WHERE department_id = ? AND material_id = ?
    `, [departmentId, material_id]);

    if (!allocation || allocation.remaining < qty) {
      return res.redirect('/staff/usage?error=' + encodeURIComponent('库存不足'));
    }

    db.run(
      'INSERT INTO usage_records (department_id, material_id, quantity, customer_name, remark, record_type, created_by) VALUES (?, ?, ?, ?, ?, ?, ?)',
      [departmentId, material_id, qty, customer_name.trim(), remark || '', 'usage', userId]
    );
    db.run('UPDATE department_allocations SET used_quantity = used_quantity + ? WHERE id = ?', [qty, allocation.id]);
    res.redirect('/staff/usage?success=' + encodeURIComponent('领用记录添加成功'));
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
    SELECT a.name AS activity_name, m.name AS material_name, m.unit, ur.quantity,
           ur.customer_name, ur.remark, u.name AS created_by_name, ur.created_at, ur.record_type
    FROM usage_records ur
    JOIN materials m ON ur.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    JOIN users u ON ur.created_by = u.id
    WHERE ur.department_id = ? ${userFilter}
    ORDER BY ur.created_at DESC
  `, params).map(row => ({
    ...row,
    customer_name_masked: maskCustomerName(row.customer_name),
    created_at: formatDateTime(row.created_at)
  }));

  res.render('staff/history', {
    user: req.session.user,
    records,
    mode,
    currentPage: 'history'
  });
});

router.get('/export/inventory', (req, res) => {
  const departmentId = req.session.user.department_id;
  const department = db.get('SELECT name FROM departments WHERE id = ?', [departmentId]);
  const selectedActivity = req.query.activity_id || '';
  const rows = getStaffInventory(departmentId, selectedActivity).map(item => ({
    '部门/网点': item.department_name,
    '活动名称': item.activity_name,
    '宣传品名称': item.material_name,
    '当前已分配': item.allocated,
    '已领用': item.used,
    '已回收': item.recovered,
    '剩余库存': item.remaining
  }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), '库存报表');
  const detailSheet = buildMergedSheet(buildStaffDetailRows(departmentId, selectedActivity), '明细报表', [1, 2]);
  XLSX.utils.book_append_sheet(wb, detailSheet.sheet, detailSheet.sheetName);

  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  const filename = encodeURIComponent(`${department.name}-库存报表.xlsx`);
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
    SELECT d.name AS '部门/网点', a.name AS '活动名称', m.name AS '宣传品名称', m.unit AS '单位',
           ur.quantity AS '领用数量', ur.customer_name AS '领用客户', u.name AS '录入员工',
           ur.created_at AS raw_time, ur.remark AS '备注'
    FROM usage_records ur
    JOIN departments d ON ur.department_id = d.id
    JOIN materials m ON ur.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    JOIN users u ON ur.created_by = u.id
    WHERE ur.department_id = ? AND ur.record_type = 'usage' ${userFilter}
    ORDER BY ur.created_at DESC
  `, params).map(row => ({
    ...row,
    时间: formatDateTime(row.raw_time),
    raw_time: undefined
  })).map(({ raw_time, ...rest }) => rest);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), '明细报表');
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  const filename = encodeURIComponent(`${department.name}-${mineOnly ? '本人' : '本部门'}明细报表.xlsx`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename=${filename}`);
  res.send(buffer);
});

module.exports = router;
