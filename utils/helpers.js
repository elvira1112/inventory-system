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

module.exports = {
  formatDateTime,
  normalizeExcelDate,
  passwordRuleError,
  maskCustomerName
};
