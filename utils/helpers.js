const bcrypt = require('bcryptjs');

function formatDateTime(value) {
  if (!value) return '';
  const text = String(value);
  const normalized = text.includes('T') ? text : text.replace(' ', 'T');
  const date = new Date(`${normalized}+08:00`);
  if (Number.isNaN(date.getTime())) return text;
  const pad = n => String(n).padStart(2, '0');
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
}

function formatDate(value) {
  if (!value) return '';
  const text = String(value);
  const normalized = text.includes('T') ? text : text.replace(' ', 'T');
  const date = new Date(`${normalized}+08:00`);
  if (Number.isNaN(date.getTime())) return text;
  const pad = n => String(n).padStart(2, '0');
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

function normalizeExcelDate(value, XLSX) {
  if (!value) return '';
  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }
  if (typeof value === 'number' && XLSX && XLSX.SSF) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      const pad = n => String(n).padStart(2, '0');
      return `${parsed.y}-${pad(parsed.m)}-${pad(parsed.d)}`;
    }
  }
  return String(value).trim();
}

function passwordRuleError(password, user) {
  if (!password || password.length < 6) {
    return '密码长度至少6位';
  }
  if (!/[a-zA-Z]/.test(password) || !/[0-9]/.test(password)) {
    return '密码必须同时包含数字和字母';
  }
  if (/^(\d)\1{3,}$/.test(password)) {
    return '密码不能是4位及以上重复数字';
  }
  for (let i = 0; i <= password.length - 3; i++) {
    const part = password.slice(i, i + 3);
    if (/^\d{3}$/.test(part)) {
      const nums = part.split('').map(Number);
      if (nums[1] === nums[0] + 1 && nums[2] === nums[1] + 1) {
        return '密码不能包含3位及以上连续数字';
      }
    }
  }
  if (user && bcrypt.compareSync(password, user.password)) {
    return '新密码不能和上一密码相同';
  }
  return null;
}

function maskCustomerName(value) {
  const name = String(value || '').trim();
  if (!name) return '';
  const chars = Array.from(name);
  const allChinese = chars.every(char => /[\u4e00-\u9fa5]/.test(char));
  if (!allChinese) return name;
  if (chars.length === 2) {
    return `${chars[0]}*`;
  }
  if (chars.length >= 3) {
    return `${chars[0]}${'*'.repeat(chars.length - 2)}${chars[chars.length - 1]}`;
  }
  return name;
}

/**
 * 敏感信息脱敏函数
 * 识别并脱敏：19位银行卡号、18位身份证号、11位手机号、12位/8位电话号码
 * 保留首尾部分，中间用 * 替代
 */
function maskSensitive(value) {
  if (!value) return value;
  var str = String(value);

  // 身份证号（末位可能为X）：保留前3 + 后4
  str = str.replace(/\b(\d{3})\d{11}([\dXx])(\d{3})\b/gi, '$1***********$2$3');

  // 处理纯数字段
  str = str.replace(/\d+/g, function(match) {
    var len = match.length;
    // 19位银行卡号/账号：前4 + *********** + 后4
    if (len === 19) return match.substring(0, 4) + '***********' + match.substring(15);
    // 18位纯数字身份证号：前3 + *********** + 后4
    if (len === 18) return match.substring(0, 3) + '***********' + match.substring(14);
    // 12位电话号码：前4 + **** + 后4
    if (len === 12) return match.substring(0, 4) + '****' + match.substring(8);
    // 11位手机号（1开头）：前3 + **** + 后4
    if (len === 11 && match[0] === '1') return match.substring(0, 3) + '****' + match.substring(7);
    // 8位电话号码：前2 + **** + 后2
    if (len === 8) return match.substring(0, 2) + '****' + match.substring(6);
    return match;
  });

  return str;
}

/**
 * 格式化活动名称显示，若已删除则追加删除原因
 */
function formatActivityDisplay(name, deleteReason) {
  if (!deleteReason) return name;
  return `${name} 【已删除：${deleteReason}】`;
}

/**
 * 获取北京时间（UTC+8）的日期字符串 YYYY-MM-DD
 */
function beijingDateString(date = new Date()) {
  const beijing = new Date(date.getTime() + 8 * 60 * 60 * 1000);
  const pad = n => String(n).padStart(2, '0');
  return `${beijing.getUTCFullYear()}-${pad(beijing.getUTCMonth() + 1)}-${pad(beijing.getUTCDate())}`;
}

/**
 * 生成指定月份（默认当月）的自然日日期列表，格式 YYYY-MM-DD
 */
function buildMonthDateOptions(date = new Date()) {
  const beijing = new Date(date.getTime() + 8 * 60 * 60 * 1000);
  const year = beijing.getUTCFullYear();
  const month = beijing.getUTCMonth();
  const days = new Date(Date.UTC(year, month + 1, 0)).getUTCDate();
  const pad = n => String(n).padStart(2, '0');
  const list = [];
  for (let d = 1; d <= days; d++) {
    list.push(`${year}-${pad(month + 1)}-${pad(d)}`);
  }
  return list;
}

/**
 * 校验日期字符串是否为合法的 YYYY-MM-DD
 */
function isValidDateString(value) {
  return typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value);
}

/**
 * 生成跨月日期分组（默认最近 12 个月，含当月，倒序），
 * 每个分组包含 month(YYYY-MM)、label(YYYY年MM月)、dates(日期数组，正序)
 * 当月只生成到今天（不允许未来日期）
 */
function buildDateGroups(monthCount = 12, date = new Date()) {
  const beijing = new Date(date.getTime() + 8 * 60 * 60 * 1000);
  const todayYear = beijing.getUTCFullYear();
  const todayMonth = beijing.getUTCMonth();
  const todayDay = beijing.getUTCDate();
  const pad = n => String(n).padStart(2, '0');
  const groups = [];

  for (let offset = 0; offset < monthCount; offset++) {
    const cursor = new Date(Date.UTC(todayYear, todayMonth - offset, 1));
    const year = cursor.getUTCFullYear();
    const month = cursor.getUTCMonth();
    const isCurrentMonth = offset === 0;
    const days = isCurrentMonth ? todayDay : new Date(Date.UTC(year, month + 1, 0)).getUTCDate();

    const dates = [];
    for (let d = 1; d <= days; d++) {
      dates.push(`${year}-${pad(month + 1)}-${pad(d)}`);
    }

    groups.push({
      month: `${year}-${pad(month + 1)}`,
      label: `${year}年${pad(month + 1)}月`,
      dates
    });
  }

  return groups;
}

module.exports = {
  formatDateTime,
  formatDate,
  normalizeExcelDate,
  passwordRuleError,
  maskCustomerName,
  maskSensitive,
  formatActivityDisplay,
  beijingDateString,
  buildMonthDateOptions,
  buildDateGroups,
  isValidDateString
};
