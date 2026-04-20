const express = require('express');
const router = express.Router();
const multer = require('multer');
const XLSX = require('xlsx');
const bcrypt = require('bcryptjs');
const path = require('path');
const db = require('../database/init');
const { isAuthenticated, isAdmin } = require('../middleware/auth');

// 文件上传配置
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, path.join(__dirname, '../uploads'));
  },
  filename: (req, file, cb) => {
    cb(null, Date.now() + '-' + file.originalname);
  }
});
const upload = multer({ storage });

// 应用中间件
router.use(isAuthenticated);
router.use(isAdmin);

// 管理员仪表盘
router.get('/', (req, res) => {
  const stats = {
    departmentCount: db.get('SELECT COUNT(*) as count FROM departments').count,
    userCount: db.get('SELECT COUNT(*) as count FROM users WHERE role = 2').count,
    activityCount: db.get('SELECT COUNT(*) as count FROM activities').count,
    materialCount: db.get('SELECT COUNT(*) as count FROM materials').count
  };
  res.render('admin/dashboard', { user: req.session.user, stats, title: '管理员仪表盘' });
});

// ============ 人员管理 ============

// 人员列表页面
router.get('/users', (req, res) => {
  const users = db.query(`
    SELECT u.*, d.name as department_name
    FROM users u
    LEFT JOIN departments d ON u.department_id = d.id
    WHERE u.role = 2
    ORDER BY u.id DESC
  `);
  const departments = db.query('SELECT * FROM departments ORDER BY name');
  res.render('admin/users', { user: req.session.user, users, departments });
});

// 导入人员名单
router.post('/users/import', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.redirect('/admin/users?error=请选择文件');
    }

    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    let imported = 0;
    let errors = [];

    for (const row of data) {
      const username = row['用户名'] || row['username'];
      const name = row['姓名'] || row['name'];
      const password = row['密码'] || row['password'] || '123456';
      const deptName = row['部门/网点'] || row['部门'] || row['department'];

      if (!username || !name || !deptName) {
        errors.push(`行数据不完整: ${JSON.stringify(row)}`);
        continue;
      }

      // 确保部门存在
      let dept = db.get('SELECT id FROM departments WHERE name = ?', [deptName]);
      if (!dept) {
        db.run('INSERT INTO departments (name) VALUES (?)', [deptName]);
        dept = { id: db.getLastInsertId() };
      }

      // 检查用户是否存在
      const existingUser = db.get('SELECT id FROM users WHERE username = ?', [username]);
      if (existingUser) {
        errors.push(`用户已存在: ${username}`);
        continue;
      }

      // 创建用户
      const hashedPassword = bcrypt.hashSync(password.toString(), 10);
      db.run(
        'INSERT INTO users (username, password, name, role, department_id) VALUES (?, ?, ?, 2, ?)',
        [username, hashedPassword, name, dept.id]
      );
      imported++;
    }

    res.redirect(`/admin/users?success=成功导入 ${imported} 名用户${errors.length > 0 ? '，部分失败' : ''}`);
  } catch (err) {
    console.error(err);
    res.redirect('/admin/users?error=导入失败: ' + err.message);
  }
});

// 添加单个用户
router.post('/users/add', (req, res) => {
  const { username, name, password, department_id } = req.body;

  try {
    const existingUser = db.get('SELECT id FROM users WHERE username = ?', [username]);
    if (existingUser) {
      return res.redirect('/admin/users?error=用户名已存在');
    }

    const hashedPassword = bcrypt.hashSync(password || '123456', 10);
    db.run(
      'INSERT INTO users (username, password, name, role, department_id) VALUES (?, ?, ?, 2, ?)',
      [username, hashedPassword, name, department_id]
    );

    res.redirect('/admin/users?success=用户添加成功');
  } catch (err) {
    res.redirect('/admin/users?error=' + err.message);
  }
});

// 删除用户
router.post('/users/delete/:id', (req, res) => {
  db.run('DELETE FROM users WHERE id = ? AND role = 2', [req.params.id]);
  res.redirect('/admin/users?success=用户删除成功');
});

// ============ 活动和宣传品管理 ============

// 活动列表页面
router.get('/activities', (req, res) => {
  const activities = db.query(`
    SELECT a.*,
           COUNT(m.id) as material_count,
           COALESCE(SUM(m.total_quantity), 0) as total_items
    FROM activities a
    LEFT JOIN materials m ON a.id = m.activity_id
    GROUP BY a.id
    ORDER BY a.id DESC
  `);
  res.render('admin/activities', { user: req.session.user, activities });
});

// 导入活动和宣传品
router.post('/activities/import', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.redirect('/admin/activities?error=请选择文件');
    }

    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    let imported = 0;

    for (const row of data) {
      const activityName = row['活动名称'] || row['activity'];
      const materialName = row['宣传品名称'] || row['material'];
      const quantity = parseInt(row['数量'] || row['quantity']) || 0;
      const unit = row['单位'] || row['unit'] || '个';

      if (!activityName || !materialName) {
        continue;
      }

      // 确保活动存在
      let activity = db.get('SELECT id FROM activities WHERE name = ?', [activityName]);
      if (!activity) {
        db.run('INSERT INTO activities (name) VALUES (?)', [activityName]);
        activity = { id: db.getLastInsertId() };
      }

      // 检查宣传品是否已存在
      const existingMaterial = db.get(
        'SELECT id FROM materials WHERE activity_id = ? AND name = ?',
        [activity.id, materialName]
      );

      if (existingMaterial) {
        // 更新数量
        db.run(
          'UPDATE materials SET total_quantity = total_quantity + ? WHERE id = ?',
          [quantity, existingMaterial.id]
        );
      } else {
        // 创建新宣传品
        db.run(
          'INSERT INTO materials (activity_id, name, unit, total_quantity) VALUES (?, ?, ?, ?)',
          [activity.id, materialName, unit, quantity]
        );
      }
      imported++;
    }

    res.redirect(`/admin/activities?success=成功导入 ${imported} 条记录`);
  } catch (err) {
    console.error(err);
    res.redirect('/admin/activities?error=导入失败: ' + err.message);
  }
});

// 查看活动详情（宣传品列表）
router.get('/activities/:id', (req, res) => {
  const activity = db.get('SELECT * FROM activities WHERE id = ?', [req.params.id]);
  if (!activity) {
    return res.redirect('/admin/activities?error=活动不存在');
  }

  const materials = db.query(`
    SELECT m.*,
           COALESCE(SUM(da.allocated_quantity), 0) as allocated,
           COALESCE(SUM(da.used_quantity), 0) as used
    FROM materials m
    LEFT JOIN department_allocations da ON m.id = da.material_id
    WHERE m.activity_id = ?
    GROUP BY m.id
  `, [req.params.id]);

  res.render('admin/activity-detail', { user: req.session.user, activity, materials });
});

// ============ 库存管理 ============

// 库存概览
router.get('/inventory', (req, res) => {
  const inventory = db.query(`
    SELECT
      a.name as activity_name,
      m.id as material_id,
      m.name as material_name,
      m.unit,
      m.total_quantity,
      COALESCE(SUM(da.allocated_quantity), 0) as allocated,
      COALESCE(SUM(da.used_quantity), 0) as used,
      m.total_quantity - COALESCE(SUM(da.allocated_quantity), 0) as unallocated
    FROM materials m
    JOIN activities a ON m.activity_id = a.id
    LEFT JOIN department_allocations da ON m.id = da.material_id
    GROUP BY m.id
    ORDER BY a.name, m.name
  `);

  res.render('admin/inventory', { user: req.session.user, inventory });
});

// ============ 配额调整 ============

// 配额管理页面
router.get('/allocations', (req, res) => {
  const departments = db.query('SELECT * FROM departments ORDER BY name');
  const activities = db.query('SELECT * FROM activities ORDER BY name');

  // 获取当前选择的活动
  const activityId = req.query.activity_id;
  let materials = [];
  let allocations = [];

  if (activityId) {
    // 获取宣传品及其库存统计（总数、已分配总数）
    materials = db.query(`
      SELECT
        m.*,
        COALESCE(SUM(da.allocated_quantity), 0) as total_allocated
      FROM materials m
      LEFT JOIN department_allocations da ON m.id = da.material_id
      WHERE m.activity_id = ?
      GROUP BY m.id
    `, [activityId]);

    // 获取每个部门对每个宣传品的分配情况
    allocations = db.query(`
      SELECT
        d.id as department_id,
        d.name as department_name,
        m.id as material_id,
        m.name as material_name,
        m.unit,
        COALESCE(da.allocated_quantity, 0) as allocated,
        COALESCE(da.used_quantity, 0) as used
      FROM departments d
      CROSS JOIN materials m
      LEFT JOIN department_allocations da ON d.id = da.department_id AND m.id = da.material_id
      WHERE m.activity_id = ?
      ORDER BY d.name, m.name
    `, [activityId]);
  }

  res.render('admin/allocations', {
    user: req.session.user,
    departments,
    activities,
    materials,
    allocations,
    selectedActivity: activityId
  });
});

// 批量更新配额
router.post('/allocations/update', (req, res) => {
  const { activity_id, allocations } = req.body;

  try {
    // allocations格式: { "dept_1_mat_2": "100", ... }
    for (const key in allocations) {
      const [_, deptId, __, matId] = key.split('_');
      const quantity = parseInt(allocations[key]) || 0;

      // 使用 UPSERT 逻辑
      const existing = db.get(
        'SELECT id FROM department_allocations WHERE department_id = ? AND material_id = ?',
        [deptId, matId]
      );

      if (existing) {
        db.run(
          'UPDATE department_allocations SET allocated_quantity = ? WHERE id = ?',
          [quantity, existing.id]
        );
      } else if (quantity > 0) {
        db.run(
          'INSERT INTO department_allocations (department_id, material_id, allocated_quantity, used_quantity) VALUES (?, ?, ?, 0)',
          [deptId, matId, quantity]
        );
      }
    }

    res.redirect(`/admin/allocations?activity_id=${activity_id}&success=配额更新成功`);
  } catch (err) {
    res.redirect(`/admin/allocations?activity_id=${activity_id}&error=${err.message}`);
  }
});

// ============ 报表导出 ============

// 导出库存报表
router.get('/export/inventory', (req, res) => {
  const data = db.query(`
    SELECT
      a.name as '活动名称',
      m.name as '宣传品名称',
      m.unit as '单位',
      m.total_quantity as '总数量',
      COALESCE(SUM(da.allocated_quantity), 0) as '已分配',
      COALESCE(SUM(da.used_quantity), 0) as '已使用',
      m.total_quantity - COALESCE(SUM(da.allocated_quantity), 0) as '未分配'
    FROM materials m
    JOIN activities a ON m.activity_id = a.id
    LEFT JOIN department_allocations da ON m.id = da.material_id
    GROUP BY m.id
    ORDER BY a.name, m.name
  `);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, '库存报表');

  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=inventory-report.xlsx');
  res.send(buffer);
});

// 导出领用明细报表
router.get('/export/usage', (req, res) => {
  const data = db.query(`
    SELECT
      d.name as '部门/网点',
      a.name as '活动名称',
      m.name as '宣传品名称',
      m.unit as '单位',
      ur.quantity as '领用数量',
      ur.customer_name as '领用客户',
      u.name as '录入员工',
      ur.created_at as '领用时间',
      ur.remark as '备注'
    FROM usage_records ur
    JOIN departments d ON ur.department_id = d.id
    JOIN materials m ON ur.material_id = m.id
    JOIN activities a ON m.activity_id = a.id
    JOIN users u ON ur.created_by = u.id
    ORDER BY ur.created_at DESC
  `);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, '领用明细');

  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=usage-report.xlsx');
  res.send(buffer);
});

// 下载导入模板
router.get('/template/:type', (req, res) => {
  const { type } = req.params;
  let data, filename;

  if (type === 'users') {
    data = [{ '用户名': 'zhangsan', '姓名': '张三', '密码': '123456', '部门/网点': '营业一部' }];
    filename = 'user-template.xlsx';
  } else if (type === 'activities') {
    data = [{ '活动名称': '春季促销', '宣传品名称': '宣传海报', '数量': 1000, '单位': '张' }];
    filename = 'activity-template.xlsx';
  } else {
    return res.status(400).send('未知模板类型');
  }

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename=${filename}`);
  res.send(buffer);
});

module.exports = router;
