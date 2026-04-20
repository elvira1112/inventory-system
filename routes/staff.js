const express = require('express');
const router = express.Router();
const XLSX = require('xlsx');
const db = require('../database/init');
const { isAuthenticated } = require('../middleware/auth');

// 应用中间件 - 确保是二级人员
router.use(isAuthenticated);
router.use((req, res, next) => {
  if (req.session.user.role !== 2) {
    return res.redirect('/admin');
  }
  next();
});

// 员工仪表盘
router.get('/', (req, res) => {
  const departmentId = req.session.user.department_id;

  // 获取部门信息
  const department = db.get('SELECT * FROM departments WHERE id = ?', [departmentId]);

  // 获取本部门库存概览
  const inventory = db.query(`
    SELECT
      a.name as activity_name,
      m.name as material_name,
      m.unit,
      COALESCE(da.allocated_quantity, 0) as allocated,
      COALESCE(da.used_quantity, 0) as used,
      COALESCE(da.allocated_quantity, 0) - COALESCE(da.used_quantity, 0) as remaining
    FROM department_allocations da
    JOIN materials m ON da.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    WHERE da.department_id = ?
    ORDER BY a.name, m.name
  `, [departmentId]);

  // 获取最近领用记录
  const recentUsage = db.query(`
    SELECT
      a.name as activity_name,
      m.name as material_name,
      m.unit,
      ur.quantity,
      ur.remark,
      ur.created_at
    FROM usage_records ur
    JOIN materials m ON ur.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    WHERE ur.department_id = ?
    ORDER BY ur.created_at DESC
    LIMIT 10
  `, [departmentId]);

  res.render('staff/dashboard', {
    user: req.session.user,
    department,
    inventory,
    recentUsage
  });
});

// 库存查看页面
router.get('/inventory', (req, res) => {
  const departmentId = req.session.user.department_id;

  const inventory = db.query(`
    SELECT
      a.id as activity_id,
      a.name as activity_name,
      m.id as material_id,
      m.name as material_name,
      m.unit,
      COALESCE(da.allocated_quantity, 0) as allocated,
      COALESCE(da.used_quantity, 0) as used,
      COALESCE(da.allocated_quantity, 0) - COALESCE(da.used_quantity, 0) as remaining
    FROM department_allocations da
    JOIN materials m ON da.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    WHERE da.department_id = ?
    ORDER BY a.name, m.name
  `, [departmentId]);

  // 按活动分组
  const groupedInventory = {};
  for (const item of inventory) {
    if (!groupedInventory[item.activity_name]) {
      groupedInventory[item.activity_name] = [];
    }
    groupedInventory[item.activity_name].push(item);
  }

  res.render('staff/inventory', {
    user: req.session.user,
    groupedInventory
  });
});

// 领用录入页面
router.get('/usage', (req, res) => {
  const departmentId = req.session.user.department_id;

  // 获取可领用的宣传品（有剩余配额的）
  const materials = db.query(`
    SELECT
      a.name as activity_name,
      m.id as material_id,
      m.name as material_name,
      m.unit,
      COALESCE(da.allocated_quantity, 0) - COALESCE(da.used_quantity, 0) as remaining
    FROM department_allocations da
    JOIN materials m ON da.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    WHERE da.department_id = ?
      AND (COALESCE(da.allocated_quantity, 0) - COALESCE(da.used_quantity, 0)) > 0
    ORDER BY a.name, m.name
  `, [departmentId]);

  res.render('staff/usage', {
    user: req.session.user,
    materials
  });
});

// 提交领用记录
router.post('/usage', (req, res) => {
  const { material_id, quantity, customer_name, remark } = req.body;
  const departmentId = req.session.user.department_id;
  const userId = req.session.user.id;
  const qty = parseInt(quantity) || 0;

  if (qty <= 0) {
    return res.redirect('/staff/usage?error=领用数量必须大于0');
  }

  if (!customer_name || !customer_name.trim()) {
    return res.redirect('/staff/usage?error=请填写领用客户名称');
  }

  try {
    // 检查剩余配额
    const allocation = db.get(`
      SELECT
        da.id,
        da.allocated_quantity,
        da.used_quantity,
        da.allocated_quantity - da.used_quantity as remaining
      FROM department_allocations da
      WHERE da.department_id = ? AND da.material_id = ?
    `, [departmentId, material_id]);

    if (!allocation || allocation.remaining < qty) {
      return res.redirect('/staff/usage?error=库存不足');
    }

    // 创建领用记录
    db.run(
      'INSERT INTO usage_records (department_id, material_id, quantity, customer_name, remark, created_by) VALUES (?, ?, ?, ?, ?, ?)',
      [departmentId, material_id, qty, customer_name.trim(), remark || '', userId]
    );

    // 更新已使用数量
    db.run(
      'UPDATE department_allocations SET used_quantity = used_quantity + ? WHERE id = ?',
      [qty, allocation.id]
    );

    res.redirect('/staff/usage?success=领用记录添加成功');
  } catch (err) {
    res.redirect('/staff/usage?error=' + err.message);
  }
});

// 领用历史
router.get('/history', (req, res) => {
  const departmentId = req.session.user.department_id;

  const records = db.query(`
    SELECT
      a.name as activity_name,
      m.name as material_name,
      m.unit,
      ur.quantity,
      ur.customer_name,
      ur.remark,
      u.name as created_by_name,
      ur.created_at
    FROM usage_records ur
    JOIN materials m ON ur.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    JOIN users u ON ur.created_by = u.id
    WHERE ur.department_id = ?
    ORDER BY ur.created_at DESC
  `, [departmentId]);

  res.render('staff/history', {
    user: req.session.user,
    records
  });
});

// 导出本部门库存报表
router.get('/export/inventory', (req, res) => {
  const departmentId = req.session.user.department_id;
  const department = db.get('SELECT name FROM departments WHERE id = ?', [departmentId]);

  const data = db.query(`
    SELECT
      a.name as '活动名称',
      m.name as '宣传品名称',
      m.unit as '单位',
      COALESCE(da.allocated_quantity, 0) as '分配数量',
      COALESCE(da.used_quantity, 0) as '已使用',
      COALESCE(da.allocated_quantity, 0) - COALESCE(da.used_quantity, 0) as '剩余数量'
    FROM department_allocations da
    JOIN materials m ON da.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    WHERE da.department_id = ?
    ORDER BY a.name, m.name
  `, [departmentId]);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, '库存报表');

  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  const filename = encodeURIComponent(`${department.name}-库存报表.xlsx`);

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename=${filename}`);
  res.send(buffer);
});

// 导出本部门领用明细
router.get('/export/usage', (req, res) => {
  const departmentId = req.session.user.department_id;
  const department = db.get('SELECT name FROM departments WHERE id = ?', [departmentId]);

  const data = db.query(`
    SELECT
      a.name as '活动名称',
      m.name as '宣传品名称',
      m.unit as '单位',
      ur.quantity as '领用数量',
      ur.customer_name as '领用客户',
      u.name as '录入员工',
      ur.created_at as '领用时间',
      ur.remark as '备注'
    FROM usage_records ur
    JOIN materials m ON ur.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    JOIN users u ON ur.created_by = u.id
    WHERE ur.department_id = ?
    ORDER BY ur.created_at DESC
  `, [departmentId]);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, '领用明细');

  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  const filename = encodeURIComponent(`${department.name}-领用明细.xlsx`);

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename=${filename}`);
  res.send(buffer);
});

module.exports = router;
