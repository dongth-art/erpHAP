/**
 * SYS_Server.gs - HAP HOLDING ERP SYSTEM
 * ĐỢT 1 - BẢO MẬT: Đăng nhập, phiên (token) phía server, mật khẩu băm, khóa chống dò mật khẩu,
 * cổng kiểm tra quyền dùng chung cho mọi hàm backend.
 *
 * QUY ƯỚC: hàm có dấu "_" ở cuối tên là hàm NỘI BỘ, trình duyệt không gọi trực tiếp được.
 */

var AUTH_ = {
  SESSION_TTL_SEC: 21600,   // 6 giờ (mức tối đa CacheService cho phép). Tự gia hạn mỗi lần thao tác.
  MAX_FAILED: 5,            // Sai tối đa 5 lần
  LOCK_SEC: 900,            // Rồi khóa 15 phút
  HASH_ITER: 500,           // Số vòng băm. Nếu đăng nhập thấy chậm có thể giảm (mật khẩu cũ vẫn dùng được).
  SUPER_ROLE: 'SuperAdmin',
  // Cột sheet SYS_USR: dùng tên cột làm khóa, giá trị là chỉ số cột (0-based) được cập nhật động từ header.
  // Thứ tự chuẩn (fallback khi thiếu cột): username=0, password_hash=1, cccd_no=2, full_name=3,
  // role=4, module_access=5, is_active=6, last_login=7, created_by=8, created_at=9, nav_access=10
  COL: {
    USERNAME: 0,
    PASSWORD: 1,
    CCCD: 2,
    FULLNAME: 3,
    ROLE: 4,
    MODULES: 5,
    ACTIVE: 6,
    LAST_LOGIN: 7,
    CREATED_BY: 8,
    CREATED_AT: 9,
    NAV_ACCESS: 10
  }
};

function normalizeRoleCode_(role) {
  var value = String(role || '').trim().toUpperCase();
  if (value === AUTH_.SUPER_ROLE.toUpperCase() || value === 'SUPERADMIN') return AUTH_.SUPER_ROLE;
  if (value === 'ADMIN') return 'ADMIN';
  if (value === 'DIRECTOR' || value === 'GIÁM ĐỐC' || value === 'GIAM DOC') return 'DIRECTOR';
  if (value === 'DEPARTMENT_MANAGER' || value === 'MANAGER' ||
      value === 'TRƯỞNG PHÒNG' || value === 'TRUONG PHONG') return 'DEPARTMENT_MANAGER';
  if (value === 'EMPLOYEE' || value === 'USER' || value === 'NHÂN VIÊN' ||
      value === 'NHAN VIEN') return 'EMPLOYEE';
  return '';
}

function isSuperAdminRole_(role) {
  return normalizeRoleCode_(role) === AUTH_.SUPER_ROLE;
}

function isPermissionAdminRole_(role) {
  return normalizeRoleCode_(role) === 'ADMIN' || isSuperAdminRole_(role);
}

/* ============================================================
 * 1. TIỆN ÍCH MẬT KHẨU
 * ============================================================ */

function toHex_(bytes) {
  return bytes.map(function (b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function sha256Hex_(text) {
  return toHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8));
}

function hashPassword_(password, salt, iterations) {
  var h = salt + ':' + password;
  for (var i = 0; i < iterations; i++) {
    h = sha256Hex_(h + salt);
  }
  return h;
}

/** Định dạng lưu: v1$số_vòng$salt$hash */
function makePasswordHash_(password) {
  var salt = Utilities.getUuid().replace(/-/g, '');
  var iter = AUTH_.HASH_ITER;
  return 'v1$' + iter + '$' + salt + '$' + hashPassword_(password, salt, iter);
}

function safeEquals_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  var r = 0;
  for (var i = 0; i < a.length; i++) r |= (a.charCodeAt(i) ^ b.charCodeAt(i));
  return r === 0;
}

function verifyPasswordHash_(password, stored) {
  var p = String(stored || '').split('$');
  if (p.length !== 4 || p[0] !== 'v1') return false;
  var iter = parseInt(p[1], 10);
  if (!iter || iter < 1) return false;
  return safeEquals_(hashPassword_(password, p[2], iter), p[3]);
}

/* ============================================================
 * 2. TIỆN ÍCH DỮ LIỆU & PHIÊN
 * ============================================================ */

/**
 * Đọc header của sheet SYS_USR và trả về ánh xạ tên cột -> chỉ số cột (0-based).
 * Dùng để thay thế các chỉ số cột cứng AUTH_.COL bằng tra cứu theo tên header.
 * Trả về map{name: index}; các cột không tìm thấy sẽ không có trong map.
 */
function getSysUsrHeaderMap_() {
  if (typeof CONFIG === 'undefined' || !CONFIG.DB_SYS_PART1_ID) return {};
  try {
    var ss = SpreadsheetApp.openById(CONFIG.DB_SYS_PART1_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEET_SYS_USR || 'SYS_USR');
    if (!sheet) return {};
    var lastColumn = Math.max(sheet.getLastColumn(), 1);
    var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    var map = {};
    for (var i = 0; i < headers.length; i++) {
      var name = String(headers[i] || '').trim();
      if (name && map[name] === undefined) map[name] = i;
    }
    return map;
  } catch (e) {
    Logger.log('getSysUsrHeaderMap_: ' + e);
    return {};
  }
}

/**
 * Cập nhật AUTH_.COL từ header thực tế của sheet.
 * Nếu header thiếu cột thì giữ giá trị fallback (thứ tự chuẩn) và ghi cảnh báo 1 lần.
 */
function refreshAuthColFromHeaders_() {
  var headerMap = getSysUsrHeaderMap_();
  if (!Object.keys(headerMap).length) return false;

  var canonicalNames = ['username', 'password_hash', 'cccd_no', 'full_name', 'role',
                        'module_access', 'is_active', 'last_login', 'created_by',
                        'created_at', 'nav_access'];
  var fallbackIndices = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
  var warned = false;

  canonicalNames.forEach(function (name, idx) {
    var found = headerMap[name];
    if (found === undefined) {
      // Dùng fallback index, nhưng chỉ nếu chưa có giá trị hợp lệ
      if (AUTH_.COL[canonicalNames[idx]] === undefined) {
        AUTH_.COL[canonicalNames[idx]] = fallbackIndices[idx];
      }
      if (!warned) {
        Logger.log('Cảnh báo: sheet SYS_USR thiếu cột [' + name + '] - tạm dùng vị trí chuẩn.');
        warned = true;
      }
    } else {
      AUTH_.COL[canonicalNames[idx]] = found;
    }
  });
  return true;
}

function getUserSheet_() {
  if (typeof CONFIG === 'undefined' || !CONFIG.DB_SYS_PART1_ID) return null;
  var ss = SpreadsheetApp.openById(CONFIG.DB_SYS_PART1_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_SYS_USR || 'SYS_USR');
  if (!sheet) return null;

  // Cập nhật AUTH_.COL từ header thực tế (chuẩn hóa tên cột)
  refreshAuthColFromHeaders_();

  // Đảm bảo cột nav_access tồn tại
  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var navAccessColumn = headers.findIndex(function (header) {
    return String(header || '').trim().toLowerCase() === 'nav_access';
  });
  if (navAccessColumn < 0) {
    navAccessColumn = lastColumn;
    sheet.getRange(1, navAccessColumn + 1).setValue('nav_access');
  }
  AUTH_.COL.NAV_ACCESS = navAccessColumn;
  return sheet;
}

function isActiveValue_(v) {
  if (v === false) return false;
  var s = String(v).trim().toUpperCase();
  return !(s === 'FALSE' || s === 'INACTIVE' || s === '0' || s === 'LOCKED' || s === 'KHÓA');
}

/** Chỉ đúng khi chủ dự án tự chạy trong trình soạn thảo (hoặc chính chủ gọi). Chặn người lạ gọi từ xa. */
function isOwnerRun_() {
  try {
    var active = Session.getActiveUser().getEmail();
    var effective = Session.getEffectiveUser().getEmail();
    return !!active && active === effective;
  } catch (e) {
    return false;
  }
}

function newToken_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function isValidTokenFormat_(token) {
  return typeof token === 'string' && /^[a-f0-9]{64}$/.test(token);
}

function createSession_(sessionUser) {
  var token = newToken_();
  CacheService.getScriptCache().put('SES_' + token, JSON.stringify(sessionUser), AUTH_.SESSION_TTL_SEC);
  return token;
}

function getSessionUserFromToken_(token) {
  if (!isValidTokenFormat_(token)) {
    return { success: false, code: 'AUTH', message: 'Phiên đăng nhập không hợp lệ. Vui lòng đăng nhập lại.' };
  }

  var cache = CacheService.getScriptCache();
  var raw = cache.get('SES_' + token);
  if (!raw) {
    return { success: false, code: 'AUTH', message: 'Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.' };
  }

  var sessionUser;
  try {
    sessionUser = JSON.parse(raw);
  } catch (err) {
    return { success: false, code: 'AUTH', message: 'Phiên đăng nhập lỗi. Vui lòng đăng nhập lại.' };
  }
  if (!sessionUser || !sessionUser.username) {
    return { success: false, code: 'AUTH', message: 'Phiên đăng nhập không hợp lệ. Vui lòng đăng nhập lại.' };
  }

  var sheet = getUserSheet_();
  if (!sheet) return { success: false, code: 'CONFIG', message: 'Không tìm thấy sheet SYS_USR.' };
  var data = sheet.getDataRange().getValues();
  var userRow = null;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][AUTH_.COL.USERNAME] || '').trim().toLowerCase() ===
        String(sessionUser.username).trim().toLowerCase()) {
      userRow = data[i];
      break;
    }
  }
  if (!userRow || !isActiveValue_(userRow[AUTH_.COL.ACTIVE])) {
    cache.remove('SES_' + token);
    return { success: false, code: 'AUTH', message: 'Tài khoản không tồn tại hoặc đã bị khóa.' };
  }

  sessionUser.role = normalizeRoleCode_(userRow[AUTH_.COL.ROLE]) || 'EMPLOYEE';
  sessionUser.modules = String(userRow[AUTH_.COL.MODULES] || '').trim();
  sessionUser.navAccess = String(userRow[AUTH_.COL.NAV_ACCESS] || '').trim() || null;
  sessionUser.fullName = String(userRow[AUTH_.COL.FULLNAME] || '').trim() || sessionUser.username;
  sessionUser.cccd = String(userRow[AUTH_.COL.CCCD] || '').replace(/'/g, '').trim();
  cache.put('SES_' + token, JSON.stringify(sessionUser), AUTH_.SESSION_TTL_SEC);
  return { success: true, user: sessionUser };
}

/* ============================================================
 * 3. CỔNG KIỂM TRA QUYỀN DÙNG CHUNG
 *    Dùng ở đầu mọi hàm backend:
 *      var auth = requireAuth_(token, 'HRM', 'delete');
 *      if (!auth.ok) return auth.res;
 *    action: 'view' | 'view_sensitive' | 'self' | 'export' | 'create' | 'update' | 'delete' | 'approve' | 'restore' | 'audit_view'
 * ============================================================ */

function requireAuth_(token, moduleCode, action, sheetName, navigationKey) {
  var auth = requireSessionAuth_(token, moduleCode, navigationKey);
  if (!auth.ok) return auth;
  if (!hasPermission_(auth.user, moduleCode, action, sheetName, navigationKey)) {
    return { ok: false, res: { success: false, code: 'FORBIDDEN', message: 'Bạn không có quyền thực hiện thao tác này.', list: [] } };
  }
  return auth;
}

function requireSessionAuth_(token, moduleCode, navigationKey) {
  // Bắt đầu 1 request mới: bỏ bộ nhớ tạm của request trước để không dùng dữ liệu quá hạn.
  clearRequestMemo_();
  var fail = function (code, msg) {
    return { ok: false, res: { success: false, code: code, message: msg, list: [] } };
  };

  if (!isValidTokenFormat_(token)) {
    return fail('AUTH', 'Phiên đăng nhập không hợp lệ. Vui lòng đăng nhập lại.');
  }

  var cache = CacheService.getScriptCache();
  var raw = cache.get('SES_' + token);
  if (!raw) {
    return fail('AUTH', 'Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.');
  }

  var user;
  try { user = JSON.parse(raw); } catch (e) { return fail('AUTH', 'Phiên đăng nhập lỗi. Vui lòng đăng nhập lại.'); }

  user.role = normalizeRoleCode_(user.role) || 'EMPLOYEE';
  cache.put('SES_' + token, JSON.stringify(user), AUTH_.SESSION_TTL_SEC); // gia hạn phiên

  if (!userHasModuleAccess_(user, moduleCode, navigationKey)) {
    return fail('FORBIDDEN', 'Bạn không có quyền truy cập chức năng này.');
  }
  return { ok: true, user: user };
}

function userHasNavigationAccess_(user, navigationKey) {
  var requestedViews = String(navigationKey || '').split(',').map(function (value) {
    return value.trim().toLowerCase();
  }).filter(function (value) { return value; });
  if (requestedViews.some(function (view) { return view.indexOf('view-sys-admin') === 0; })) {
    return isSuperAdminRole_(user.role);
  }
  if (isPermissionAdminRole_(user.role)) return true;
  var rawAccess = user.navAccess == null ? '' : String(user.navAccess).trim();
  if (!rawAccess) return false;

  var allowedViews = rawAccess.split(',').map(function (value) {
    return value.trim().toLowerCase();
  }).filter(function (value) { return value; });
  return allowedViews.indexOf('all') !== -1 ||
    requestedViews.some(function (view) { return allowedViews.indexOf(view) !== -1; });
}

function userHasModuleAccess_(user, moduleCode, navigationKey) {
  if (navigationKey) return userHasNavigationAccess_(user, navigationKey);
  if (isSuperAdminRole_(user.role)) return true;
  var modules = String(user.modules || '').trim();
  if (!modules || modules.toUpperCase() === 'ALL') return true;
  return modules.split(',').some(function (value) {
    return value.trim().toUpperCase() === String(moduleCode).toUpperCase();
  });
}

function hasPermission_(user, moduleCode, action, sheetName, navigationKey, resolvedPermission) {
  if (!user) return false;
  if (isSuperAdminRole_(user.role)) return true; // SuperAdmin toàn quyền
  if (isPermissionAdminRole_(user.role) &&
      (action === 'manage_permissions' ||
       action === 'manage_accounts' ||
       (String(moduleCode).toUpperCase() === 'SYS' &&
        ['SYS_PERM', 'SYS_USR'].indexOf(String(sheetName || '').toUpperCase()) !== -1 &&
        action === 'view'))) return true;
  if (!userHasModuleAccess_(user, moduleCode, navigationKey)) return false;

  try {
    var permission = resolvedPermission || getEffectivePermission_(user.username, user.role, sheetName || moduleCode);
    var customActions = String(permission.customActions || '').split(',').map(function (item) {
      return item.trim().toUpperCase();
    });
    registerPermissionAction_(sheetName || moduleCode, action);

    switch (action) {
      case 'view':
      case 'self':
        return permission.canRead;
      case 'create':
        return permission.canCreate;
      case 'update':
        return permission.canEdit;
      case 'delete':
      case 'purge':
        return permission.canDelete;
      default:
        return customActions.includes(action.toUpperCase());
    }
  } catch (e) {
    Logger.log("Lỗi check permission: " + e.toString());
    return false;
  }
}

function registerPermissionAction_(sheetName, action) {
  var table = String(sheetName || '').trim();
  var actionId = String(action || '').trim().toUpperCase();
  if (!table || !/^[A-Z][A-Z0-9_]{1,63}$/.test(actionId) ||
      ['VIEW', 'SELF', 'CREATE', 'UPDATE', 'DELETE', 'PURGE'].indexOf(actionId) !== -1) return;
  if (getBuiltInPermissionActions_(table).some(function (item) { return item.id === actionId; })) return;
  try {
    var properties = PropertiesService.getScriptProperties();
    var propertyKey = 'SYS_PERMISSION_ACTION_CATALOG_V1';
    var catalog = JSON.parse(properties.getProperty(propertyKey) || '{}');
    var tableActions = Array.isArray(catalog[table]) ? catalog[table] : [];
    if (tableActions.indexOf(actionId) === -1) {
      tableActions.push(actionId);
      catalog[table] = tableActions;
      properties.setProperty(propertyKey, JSON.stringify(catalog));
    }
  } catch (err) {
    Logger.log('registerPermissionAction_: ' + err);
  }
}

function getPermissionActionLabel_(actionId) {
  var labels = {
    ATTENDANCE_SELF: 'Chấm công',
    ATTENDANCE_PROXY: 'Chấm công thay',
    LEAVE_CREATE: 'Tạo đơn nghỉ phép',
    APPROVE: 'Phê duyệt',
    REJECT: 'Từ chối',
    EXPORT: 'Xuất dữ liệu',
    IMPORT: 'Nhập dữ liệu',
    VIEW_SENSITIVE: 'Xem dữ liệu nhạy cảm',
    ADD_TO_EMPLOYEE: 'Tạo nhân sự từ hồ sơ CCCD',
    PRINT: 'In',
    RESTORE: 'Khôi phục'
  };
  var id = String(actionId || '').trim().toUpperCase();
  return labels[id] || id.toLowerCase().split('_').map(function(part) {
    return part ? part.charAt(0).toUpperCase() + part.slice(1) : '';
  }).join(' ');
}

function getBuiltInPermissionActions_(sheetName) {
  var catalog = {
    HRM_ATT: [
      { id: 'ATTENDANCE_SELF', label: 'Chấm công', group: 'Chấm công' },
      { id: 'ATTENDANCE_PROXY', label: 'Chấm công thay', group: 'Chấm công' },
      { id: 'LEAVE_CREATE', label: 'Tạo đơn nghỉ phép', group: 'Nghỉ phép' },
      { id: 'IMPORT', label: 'Import Excel', group: 'Dữ liệu' },
      { id: 'APPROVE', label: 'Duyệt phép', group: 'Nghỉ phép' },
      { id: 'REJECT', label: 'Từ chối phép', group: 'Nghỉ phép' }
    ],
    HRM_EMP: [
      { id: 'EXPORT', label: 'Xuất dữ liệu', group: 'Dữ liệu' },
      { id: 'IMPORT', label: 'Import Excel', group: 'Dữ liệu' },
      { id: 'VIEW_SENSITIVE', label: 'Xem dữ liệu nhạy cảm', group: 'Dữ liệu cá nhân' },
      { id: 'RESTORE', label: 'Khôi phục hồ sơ', group: 'Quản lý hồ sơ' }
    ],
    HRM_ID: [
      { id: 'EXPORT', label: 'Xuất dữ liệu', group: 'Dữ liệu' },
      { id: 'IMPORT', label: 'Import Excel', group: 'Dữ liệu' },
      { id: 'RESTORE', label: 'Khôi phục hồ sơ', group: 'Quản lý hồ sơ' },
      { id: 'ADD_TO_EMPLOYEE', label: 'Tạo nhân sự từ hồ sơ CCCD', group: 'Quản lý hồ sơ' }
    ]
  };
  return catalog[String(sheetName || '').trim().toUpperCase()] || [];
}

/* Một số module gọi kiểm tra quyền bằng tên bảng rút gọn ('DOC', 'LINE') trong khi
 * SYS_PERM lưu sheet_name theo tên sheet đầy đủ ('WMS_DOC', 'WMS_LINE').
 * Hai hàm dưới đây quy đổi qua lại để rule đã cấu hình luôn được áp dụng. */
function getPermissionSheetNames_(sheetName) {
  var name = String(sheetName == null ? '' : sheetName).trim().toUpperCase();
  if (!name) return [];
  var pairs = { DOC: 'WMS_DOC', WMS_DOC: 'DOC', LINE: 'WMS_LINE', WMS_LINE: 'LINE' };
  return pairs[name] ? [name, pairs[name]] : [name];
}

/* ============================================================
 *  TRUY CẬP SHEET SYS_PERM THEO TÊN CỘT
 *  Không dùng chỉ số cột cứng để tránh lỗi lệch cột khi thêm/bớt/thay đổi cột.
 *  Thứ tự cột chuẩn (dùng làm fallback và để tạo header):
 * ============================================================ */

var PERM_COLUMNS_ = [
  'perm_id', 'target_type', 'target_value', 'sheet_name',
  'scope_owner', 'dept_filter', 'status_filter',
  'can_read', 'can_create', 'can_edit', 'can_delete', 'custom_actions'
];

var _permSchema_ = null;
var _permSchemaWarned_ = false;

/**
 * Bản đồ tên cột -> chỉ số cột của sheet SYS_PERM.
 * Đọc header đúng 1 lần rồi ghi nhớ; nếu header thiếu cột thì dùng vị trí chuẩn
 * và ghi cảnh báo 1 lần để không làm hỏng chức năng.
 */
function getPermSchema_() {
  if (_permSchema_) return _permSchema_;
  var header = {};
  try {
    var sheet = SpreadsheetApp.openById(CONFIG.DB_SYS_PART1_ID).getSheetByName('SYS_PERM');
    if (sheet) {
      getSheetHeaders_(sheet).forEach(function (headerName, index) {
        var name = String(headerName || '').trim();
        if (name && header[name] === undefined) header[name] = index;
      });
    }
  } catch (e) {
    Logger.log('getPermSchema_: ' + e);
  }
  var schema = {};
  var missing = [];
  PERM_COLUMNS_.forEach(function (name) {
    var index = header[name];
    if (index === undefined) { index = PERM_COLUMNS_.indexOf(name); missing.push(name); }
    schema[name] = index;
  });
  if (missing.length && !_permSchemaWarned_) {
    _permSchemaWarned_ = true;
    Logger.log('Cảnh báo: sheet SYS_PERM thiếu cột [' + missing.join(', ') + '] - tạm dùng vị trí chuẩn.');
  }
  _permSchema_ = schema;
  return schema;
}

/** Đọc 1 dòng SYS_PERM theo tên cột -> trả về object đã chuẩn hoá. */
function readPermRow_(row, schema) {
  var sc = schema || getPermSchema_();
  var at = function (name) { return row[sc[name]]; };
  return {
    permId: String(at('perm_id') || '').trim(),
    targetType: String(at('target_type') || '').trim(),
    targetValue: String(at('target_value') || '').trim(),
    sheetName: String(at('sheet_name') || '').trim(),
    scopeOwner: String(at('scope_owner') || '').trim(),
    deptFilter: String(at('dept_filter') || '').trim(),
    statusFilter: String(at('status_filter') || '').trim(),
    canRead: String(at('can_read') || '').trim().toUpperCase() === 'TRUE',
    canCreate: String(at('can_create') || '').trim().toUpperCase() === 'TRUE',
    canEdit: String(at('can_edit') || '').trim().toUpperCase() === 'TRUE',
    canDelete: String(at('can_delete') || '').trim().toUpperCase() === 'TRUE',
    customActions: String(at('custom_actions') || '').trim()
  };
}

/** Dựng mảng giá trị cho 1 dòng SYS_PERM theo thứ tự cột chuẩn. */
function permRowValues_(values) {
  return PERM_COLUMNS_.map(function (name) {
    return values[name] === undefined || values[name] === null ? '' : values[name];
  });
}

/**
 * Ghi 1 dòng SYS_PERM theo tên cột.
 * Nếu thứ tự cột chuẩn -> ghi 1 khối bằng setValues; nếu lệch -> ghi đúng ô theo header.
 */
function writePermRow_(sheet, rowNumber, values, schema) {
  var sc = schema || getPermSchema_();
  var canonical = PERM_COLUMNS_.every(function (name, i) { return sc[name] === i; });
  if (canonical) {
    sheet.getRange(rowNumber, 1, 1, PERM_COLUMNS_.length).setValues([permRowValues_(values)]);
    return;
  }
  PERM_COLUMNS_.forEach(function (name) {
    if (sc[name] === undefined) return;
    var value = values[name] === undefined || values[name] === null ? '' : values[name];
    sheet.getRange(rowNumber, sc[name] + 1).setValue(value);
  });
}

/**
 * Ghi nhiều dòng SYS_PERM (bắt đầu từ startRow) theo tên cột.
 * Dùng 1 lần setValues khi thứ tự cột chuẩn để không phát sinh nhiều lần ghi.
 */
function writePermRows_(sheet, startRow, valuesList, schema) {
  if (!valuesList.length) return;
  var sc = schema || getPermSchema_();
  var canonical = PERM_COLUMNS_.every(function (name, i) { return sc[name] === i; });
  if (canonical) {
    var payload = valuesList.map(permRowValues_);
    sheet.getRange(startRow, 1, payload.length, PERM_COLUMNS_.length).setValues(payload);
    return;
  }
  valuesList.forEach(function (values, i) { writePermRow_(sheet, startRow + i, values, sc); });
}

function isPermissionBypassed_(sheetName) {
  var bypass = CONFIG.BYPASS_PERMISSION_SHEETS;
  if (!Array.isArray(bypass) || !bypass.length) return false;
  var names = getPermissionSheetNames_(sheetName);
  return bypass.some(function(item) {
    return names.indexOf(String(item == null ? '' : item).trim().toUpperCase()) !== -1;
  });
}

var _sheetHeaderCache_ = {};
var _sheetLookupCache_ = {};

/**
 * Tìm sheet + header theo tên. Trước đây mỗi lần gọi đều mở lại Spreadsheet (openById) cho tới khi
 * tìm thấy sheet -> resolveRuleColumns_ gọi hàm này cho MỌI điều kiện của MỌI rule trong vòng lặp.
 * Nay kết quả tra cứu được ghi nhớ theo tên bảng để chỉ mở spreadsheet 1 lần.
 */
function findSheetAndHeaders_(sheetName) {
  var lookupKey = String(sheetName == null ? '' : sheetName).trim().toUpperCase();
  if (Object.prototype.hasOwnProperty.call(_sheetLookupCache_, lookupKey)) {
    return _sheetLookupCache_[lookupKey];
  }
  var names = getPermissionSheetNames_(sheetName);
  var dbIds = [
    CONFIG.DB_SYS_PART1_ID,
    CONFIG.DB_HRM_PART1_ID,
    CONFIG.DB_WMS_PART1_ID,
    CONFIG.DB_DMS_PART1_ID
  ];
  for (var i = 0; i < dbIds.length; i++) {
    if (!dbIds[i] || String(dbIds[i]).indexOf('NHẬP_ID') !== -1) continue;
    try {
      var ss = SpreadsheetApp.openById(dbIds[i]);
      for (var n = 0; n < names.length; n++) {
        var sheet = ss.getSheetByName(names[n]);
        if (sheet) {
          var key = names[n];
          if (!_sheetHeaderCache_[key]) {
            _sheetHeaderCache_[key] = { sheet: sheet, headers: getSheetHeaders_(sheet) };
          }
          _sheetLookupCache_[lookupKey] = _sheetHeaderCache_[key];
          return _sheetHeaderCache_[key];
        }
      }
    } catch (e) {
      // continue to next database
    }
  }
  return null;
}

function resolveRuleColumns_(permission, sheetName) {
  if (!permission || !Array.isArray(permission.rules)) return;
  permission.rules.forEach(function(rule) {
    if (!Array.isArray(rule.conditions)) return;
    rule.conditions.forEach(function(cond) {
      if (!cond || !cond.columnName) return;
      var found = findSheetAndHeaders_(sheetName);
      if (!found) return;
      var headers = found.headers;
      var expectedName = String(cond.columnName || '').trim();
      var currentIndex = Number(cond.columnIndex);
      var isIndexValid = isFinite(currentIndex) &&
        currentIndex >= 0 && currentIndex < headers.length &&
        String(headers[currentIndex] || '').trim().toLowerCase() === expectedName.toLowerCase();
      if (!isIndexValid) {
        var idx = headers.indexOf(expectedName);
        if (idx !== -1) cond.columnIndex = idx;
      }
    });
  });
}

var _permCondSheetCache_ = null;
var _permCondHeaderMap_ = null;

function getPermConditionsSheet_() {
  if (_permCondSheetCache_) return _permCondSheetCache_;
  try {
    var ss = SpreadsheetApp.openById(CONFIG.DB_SYS_PART1_ID);
    var sheet = ss.getSheetByName('SYS_PERM_CONDITIONS');
    if (!sheet) {
      sheet = ss.insertSheet('SYS_PERM_CONDITIONS');
      sheet.appendRow(['perm_id', 'condition_index', 'column_index', 'column_name', 'operator', 'value', 'data_type', 'all_values']);
    }
    _permCondSheetCache_ = { spreadsheet: ss, sheet: sheet };
    // Build header map for header-based column access
    var lastColumn = Math.max(sheet.getLastColumn(), 1);
    var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    _permCondHeaderMap_ = {};
    for (var h = 0; h < headers.length; h++) {
      var hname = String(headers[h] || '').trim();
      if (hname && _permCondHeaderMap_[hname] === undefined) _permCondHeaderMap_[hname] = h;
    }
    return _permCondSheetCache_;
  } catch (e) {
    Logger.log('getPermConditionsSheet_: ' + e);
    return null;
  }
}

function savePermConditions_(permId, conditions) {
  var sh = getPermConditionsSheet_();
  if (!sh) return;
  var sheet = sh.sheet;
  var hm = _permCondHeaderMap_ || {};
  var colPermId = hm['perm_id'] !== undefined ? hm['perm_id'] : 0;
  var data = sheet.getDataRange().getValues();
  var rowsToDelete = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colPermId] || '').trim() === String(permId || '').trim()) {
      rowsToDelete.push(i + 1);
    }
  }
  rowsToDelete.sort(function(a, b) { return b - a; });
  rowsToDelete.forEach(function(rowNumber) {
    sheet.deleteRow(rowNumber);
  });
  if (!Array.isArray(conditions) || !conditions.length) return;
  var rows = [];
  conditions.forEach(function(cond, idx) {
    rows.push([
      String(permId || ''),
      idx,
      Number(cond.columnIndex || 0),
      String(cond.columnName || ''),
      String(cond.operator || 'equals'),
      String(cond.value == null ? '' : cond.value),
      String(cond.dataType || ''),
      cond.allValues === true ? 'TRUE' : 'FALSE'
    ]);
  });
  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  SpreadsheetApp.flush();
}

/**
 * Đọc TOÀN BỘ sheet SYS_PERM_CONDITIONS đúng 1 lần rồi tra theo perm_id trong bộ nhớ request.
 * Trước đây mỗi rule gọi getPermConditions_ đều đọc lại toàn bộ sheet (N+1).
 * Thứ tự điều kiện giữ nguyên theo thứ tự dòng trong sheet (khớp hành vi cũ).
 * Đã đồng bộ: dùng header map để truy xuất cột thay vì chỉ số cột cứng.
 */
function getPermConditionsMap_() {
  return requestMemo_('SYS_PERM_CONDITIONS_MAP', function () {
    var map = {};
    var sh = getPermConditionsSheet_();
    if (!sh) return map;
    var sheet = sh.sheet;
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return map;
    var hm = _permCondHeaderMap_ || {};
    var colPermId = hm['perm_id'] !== undefined ? hm['perm_id'] : 0;
    var colColIndex = hm['column_index'] !== undefined ? hm['column_index'] : 2;
    var colColName = hm['column_name'] !== undefined ? hm['column_name'] : 3;
    var colOperator = hm['operator'] !== undefined ? hm['operator'] : 4;
    var colValue = hm['value'] !== undefined ? hm['value'] : 5;
    var colDataType = hm['data_type'] !== undefined ? hm['data_type'] : 6;
    var colAllValues = hm['all_values'] !== undefined ? hm['all_values'] : 7;
    var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    for (var i = 0; i < data.length; i++) {
      var permId = String(data[i][colPermId] || '').trim();
      if (!permId) continue;
      if (!map[permId]) map[permId] = [];
      map[permId].push({
        columnIndex: Number(data[i][colColIndex] || 0),
        columnName: String(data[i][colColName] || ''),
        operator: String(data[i][colOperator] || 'equals'),
        value: String(data[i][colValue] || ''),
        dataType: String(data[i][colDataType] || ''),
        allValues: String(data[i][colAllValues] || '').toUpperCase() === 'TRUE'
      });
    }
    return map;
  });
}

/** Điều kiện ma trận của 1 perm_id. Trả về bản sao để không sửa bộ nhớ dùng chung. */
function getPermConditions_(permId) {
  var map = getPermConditionsMap_();
  var rows = map[String(permId || '').trim()] || [];
  return rows.map(function (row) {
    return {
      columnIndex: row.columnIndex,
      columnName: row.columnName,
      operator: row.operator,
      value: row.value,
      dataType: row.dataType,
      allValues: row.allValues
    };
  });
}

function deletePermConditions_(permId) {
  var sh = getPermConditionsSheet_();
  if (!sh) return;
  var sheet = sh.sheet;
  var hm = _permCondHeaderMap_ || {};
  var colPermId = hm['perm_id'] !== undefined ? hm['perm_id'] : 0;
  var data = sheet.getDataRange().getValues();
  var rowsToDelete = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colPermId] || '').trim() === String(permId || '').trim()) {
      rowsToDelete.push(i + 1);
    }
  }
  rowsToDelete.sort(function(a, b) { return b - a; });
  rowsToDelete.forEach(function(rowNumber) {
    sheet.deleteRow(rowNumber);
  });
  if (rowsToDelete.length) SpreadsheetApp.flush();
}

function normalizeCompareValue_(value, dataType) {
  var v = String(value == null ? '' : value).trim();
  if (v.charAt(0) === "'") v = v.slice(1);
  if (!dataType) return v;
  var dt = String(dataType || '').trim().toLowerCase();
  if (dt === 'date') {
    var d = new Date(v);
    return isNaN(d.getTime()) ? v : Utilities.formatDate(d, 'GMT+7', 'yyyy-MM-dd');
  }
  if (dt === 'number') {
    var n = Number(v);
    return isFinite(n) ? String(n) : v;
  }
  return v;
}

function parseOneOfValues_(expected) {
  if (!expected) return [];
  try {
    var parsed = JSON.parse(expected);
    if (Array.isArray(parsed)) {
      return parsed.map(function(v) { return String(v == null ? '' : v).trim(); }).filter(Boolean);
    }
  } catch (e) {
    // fallthrough to comma split below
  }
  return String(expected).split(',').map(function(v) { return v.trim(); }).filter(Boolean);
}

function matrixConditionsMatch_(conditions, row) {
  if (!Array.isArray(conditions) || conditions.length === 0) return true;
  return conditions.every(function(condition) {
    if (!condition || typeof condition !== 'object' || !Array.isArray(row)) return false;
    var index = Number(condition.columnIndex);
    if (!isFinite(index) || index < 0 || Math.floor(index) !== index || index >= row.length) return false;
    if (condition.allValues === true) return true;
    var dataType = condition.dataType || '';
    var actual = normalizeCompareValue_(row[index], dataType);
    var expected = normalizeCompareValue_(condition.value == null ? '' : condition.value, dataType);
    if (!expected) return true;
    switch (String(condition.operator || '').toLowerCase()) {
      case 'equals':
        return actual === expected;
      case 'not_equals':
        return actual !== expected;
      case 'contains':
        return actual.indexOf(expected) !== -1;
      case 'one_of':
        var values = parseOneOfValues_(expected);
        return values.indexOf(actual) !== -1;
      default:
        return false;
    }
  });
}

function matrixRowAllowed_(user, sheetName, row, fields, resolvedPermission) {
  if (isPermissionBypassed_(sheetName)) return true;
  if (user && isSuperAdminRole_(user.role)) return true;
  var permission = resolvedPermission || getEffectivePermission_(user.username, user.role, sheetName);
  resolveRuleColumns_(permission, sheetName);
  if (isPermissionAdminRole_(user.role) && ['SYS_USR', 'SYS_PERM'].indexOf(String(sheetName).toUpperCase()) !== -1) return true;
  return matrixPermissionMatches_(permission, row, 'read');
}

function matrixPermissionMatches_(permission, row, operation) {
  if (!permission || !Array.isArray(permission.rules)) return false;
  var permissionKeys = { read: 'canRead', create: 'canCreate', update: 'canEdit', delete: 'canDelete' };
  var permissionKey = permissionKeys[operation];
  if (!permissionKey) return false;
  return permission.rules.some(function(rule) {
    var condOk = (!Array.isArray(rule.conditions) || rule.conditions.length === 0) || matrixConditionsMatch_(rule.conditions, row);
    return rule[permissionKey] === true && condOk;
  });
}

function matrixScopeAllowed_(user, sheetName, row, fields, resolvedPermission, operation) {
  if (isPermissionBypassed_(sheetName)) return true;
  if (user && isSuperAdminRole_(user.role)) return true;
  var permission = resolvedPermission || getEffectivePermission_(user.username, user.role, sheetName);
  resolveRuleColumns_(permission, sheetName);
  if (isPermissionAdminRole_(user.role) && ['SYS_USR', 'SYS_PERM'].indexOf(String(sheetName).toUpperCase()) !== -1) return true;
  return matrixPermissionMatches_(permission, row, operation || 'update');
}

/**
 * API QUÉT TOÀN BỘ SPREADSHEETS VÀ CỘT TRONG ERP (TỰ ĐỘNG CẬP NHẬT GIAO DIỆN)
 */
function getSystemSchemaApiLegacy_(token) {
  var auth = requireAuth_(token, 'SYS', 'view', 'SYS_PERM', 'view-sys-usr');
  if (!auth.ok) return auth.res;

  let spreadsheetIds = [
    CONFIG.DB_SYS_PART1_ID,
    CONFIG.DB_HRM_PART1_ID,
    CONFIG.DB_WMS_PART1_ID
  ];
  
  if (CONFIG.DB_SYS_PART2_ID && !CONFIG.DB_SYS_PART2_ID.includes("NHẬP_ID")) spreadsheetIds.push(CONFIG.DB_SYS_PART2_ID);
  if (CONFIG.DB_HRM_PART2_ID && !CONFIG.DB_HRM_PART2_ID.includes("NHẬP_ID")) spreadsheetIds.push(CONFIG.DB_HRM_PART2_ID);

  let schema = [];
  const systemSheets = [CONFIG.SHEET_SYS_USR, CONFIG.SHEET_SYS_MOD, CONFIG.SHEET_SYS_PERM, CONFIG.SHEET_SYS_AUDIT];

  spreadsheetIds.forEach(id => {
    try {
      let ss = SpreadsheetApp.openById(id);
      ss.getSheets().forEach(sheet => {
        let name = sheet.getName();
        if (!systemSheets.includes(name)) {
          let headers = sheet.getLastColumn() > 0 ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
          schema.push({
            sheetName: name,
            columns: headers.filter(h => h !== "")
          });
        }
      });
    } catch (e) {
      Logger.log("Lỗi quét Sheet ID [" + id + "]: " + e);
    }
  });

  return { success: true, schema: schema };
}

/* ============================================================
 * 4. HÀM CÔNG KHAI CHO GIAO DIỆN
 * ============================================================ */

function verifyLogin(username, password) {
  try {
    var inputUser = String(username == null ? '' : username).trim();
    var inputPass = String(password == null ? '' : password).trim();

    if (!inputUser || !inputPass) {
      return { success: false, message: 'Vui lòng nhập đầy đủ tên đăng nhập và mật khẩu!' };
    }
    if (inputUser.length > 100 || inputPass.length > 200) {
      return { success: false, message: 'Sai tên đăng nhập hoặc mật khẩu!' };
    }

    var cache = CacheService.getScriptCache();
    var failKey = 'FAIL_' + inputUser.toLowerCase();
    var failCount = parseInt(cache.get(failKey) || '0', 10);
    if (failCount >= AUTH_.MAX_FAILED) {
      return { success: false, message: 'Đăng nhập sai quá nhiều lần. Vui lòng thử lại sau 15 phút.' };
    }

    var sheet = getUserSheet_();
    if (!sheet) {
      return { success: false, message: 'Hệ thống chưa cấu hình cơ sở dữ liệu tài khoản.' };
    }

    var C = AUTH_.COL;
    var found = null, foundRow = -1;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][C.USERNAME]).trim().toLowerCase() === inputUser.toLowerCase()) {
        found = data[i];
        foundRow = i + 1;
        break;
      }
    }

    var passOk = false;
    if (found) {
      passOk = verifyPasswordHash_(inputPass, found[C.PASSWORD]);
    } else {
      hashPassword_(inputPass, 'dummy', AUTH_.HASH_ITER); // giữ thời gian phản hồi tương đương
    }

    if (!passOk) {
      cache.put(failKey, String(failCount + 1), AUTH_.LOCK_SEC);
      if (failCount + 1 === AUTH_.MAX_FAILED) {
        auditLog_({ username: inputUser }, 'SYS', 'LOCKOUT', inputUser, 'Sai mật khẩu ' + AUTH_.MAX_FAILED + ' lần, khóa 15 phút');
      }
      return { success: false, message: 'Sai tên đăng nhập hoặc mật khẩu!' };
    }

    if (!isActiveValue_(found[C.ACTIVE])) {
      return { success: false, message: 'Tài khoản của bạn đang bị khóa!' };
    }

    cache.remove(failKey);

    var username = String(found[C.USERNAME]).trim();
    var sessionUser = {
      username: username,
      fullName: String(found[C.FULLNAME] || '').trim() || username,
      role: normalizeRoleCode_(found[C.ROLE]) || 'EMPLOYEE',
      modules: String(found[C.MODULES] || '').trim(),
      navAccess: String(found[C.NAV_ACCESS] || '').trim() || null,
      cccd: String(found[C.CCCD] || '').replace(/'/g, '').trim(),   // CCCD liên kết -> xác định "nhân viên của chính mình"
      loginAt: new Date().getTime()
    };
    var token = createSession_(sessionUser);
    auditLog_(sessionUser, 'SYS', 'LOGIN', username, '');

    try {
      sheet.getRange(foundRow, C.LAST_LOGIN + 1).setValue(new Date());
      SpreadsheetApp.flush();
    } catch (e) {
      Logger.log('Không thể cập nhật last_login cho ' + username + ': ' + e);
    }

    return {
      success: true,
      token: token,
      user: {
        username: sessionUser.username,
        fullName: sessionUser.fullName,
        role: sessionUser.role,
        modules: sessionUser.modules,
        navAccess: sessionUser.navAccess
      }
    };

  } catch (err) {
    Logger.log('verifyLogin lỗi: ' + err);
    return { success: false, message: 'Lỗi hệ thống xác thực. Vui lòng thử lại.' };
  }
}

/** Giao diện gọi khi mở trang để kiểm tra phiên đã lưu còn hợp lệ không */
function verifySession(token) {
  var result = getSessionUserFromToken_(token);
  if (!result.success) return result;
  var u = result.user;
  return {
    success: true,
    user: {
      username: u.username,
      fullName: u.fullName,
      role: u.role,
      modules: u.modules,
      navAccess: u.navAccess
    }
  };
}

function logoutSession(token) {
  if (isValidTokenFormat_(token)) {
    var cache = CacheService.getScriptCache();
    var raw = cache.get('SES_' + token);
    if (raw) {
      try {
        var user = JSON.parse(raw);
        auditLog_(user, 'SYS', 'LOGOUT', user.username, '');
      } catch (err) {
        Logger.log('Không thể ghi sự kiện LOGOUT: ' + err);
      }
    }
    cache.remove('SES_' + token);
  }
  return { success: true };
}

/* ============================================================
 * 5. CÔNG CỤ CÀI ĐẶT - CHỈ CHỦ DỰ ÁN CHẠY TRONG TRÌNH SOẠN THẢO
 * ============================================================ */

/** Chạy MỘT LẦN: chuyển mọi mật khẩu chữ thường trong SYS_USR sang dạng băm. */
function migratePasswords() {
  if (!isOwnerRun_()) return { success: false, message: 'Không được phép.' };

  var sheet = getUserSheet_();
  if (!sheet) { Logger.log('Không tìm thấy sheet SYS_USR.'); return { success: false, message: 'Không tìm thấy sheet SYS_USR.' }; }

  var n = sheet.getLastRow() - 1;
  if (n < 1) { Logger.log('SYS_USR chưa có tài khoản nào. Hãy chạy runSetupAdmin.'); return { success: true, changed: 0 }; }

  var range = sheet.getRange(2, AUTH_.COL.PASSWORD + 1, n, 1);
  var vals = range.getValues();
  var changed = 0;

  for (var i = 0; i < vals.length; i++) {
    var s = String(vals[i][0]).trim();
    if (s && s.indexOf('v1$') !== 0) {
      vals[i][0] = makePasswordHash_(s);
      changed++;
    }
  }
  if (changed > 0) range.setValues(vals);

  Logger.log('Đã mã hóa ' + changed + ' mật khẩu.');
  return { success: true, changed: changed };
}

/** Tạo mới hoặc đặt lại một tài khoản SuperAdmin. Chỉ chạy được khi chủ dự án tự chạy. */
function setupAdminAccount(username, password, fullName) {
  if (!isOwnerRun_()) return { success: false, message: 'Không được phép.' };

  var u = String(username || '').trim();
  var p = String(password || '').trim();
  if (!u || u.length > 100) return { success: false, message: 'Tên đăng nhập không hợp lệ.' };
  if (p.length < 8) return { success: false, message: 'Mật khẩu phải có ít nhất 8 ký tự.' };

  var sheet = getUserSheet_();
  if (!sheet) return { success: false, message: 'Không tìm thấy sheet SYS_USR.' };

  var C = AUTH_.COL;
  var hash = makePasswordHash_(p);
  var data = sheet.getDataRange().getValues();
  var row = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][C.USERNAME]).trim().toLowerCase() === u.toLowerCase()) { row = i + 1; break; }
  }

  if (row > 0) {
    sheet.getRange(row, C.PASSWORD + 1).setValue(hash);
    sheet.getRange(row, C.ROLE + 1).setValue(AUTH_.SUPER_ROLE);
    sheet.getRange(row, C.ACTIVE + 1).setValue(true);
    if (fullName) sheet.getRange(row, C.FULLNAME + 1).setValue(String(fullName).trim());
    return { success: true, message: 'Đã cập nhật tài khoản SuperAdmin: ' + u };
  }

  sheet.appendRow([u, hash, '', String(fullName || u).trim(), AUTH_.SUPER_ROLE, 'ALL', true, '', 'SETUP', new Date()]);
  return { success: true, message: 'Đã tạo tài khoản SuperAdmin: ' + u };
}

/**
 * CÁCH DÙNG: sửa 3 giá trị bên dưới -> bấm Chạy -> XÓA mật khẩu khỏi code ngay sau khi chạy xong.
 */
function runSetupAdmin() {
  var USERNAME = 'TEN_DANG_NHAP';
  var PASSWORD = 'MAT_KHAU_MOI_TOI_THIEU_8_KY_TU';
  var FULLNAME = 'Họ và Tên của bạn';

  if (USERNAME === 'TEN_DANG_NHAP') {
    Logger.log('Bạn chưa sửa thông tin trong hàm runSetupAdmin.');
    return;
  }
  var r = setupAdminAccount(USERNAME, PASSWORD, FULLNAME);
  Logger.log(r.message);
}

/**
 * Lấy danh sách tài khoản người dùng từ Sheet SYS_USR
 */
/**
 * Lấy danh sách đầy đủ thông tin tài khoản từ sheet SYS_USR
 */
/**
 * Lấy danh sách tài khoản chính xác theo chỉ số cột AUTH_.COL
 */
function getSysUsersListLegacy_(token) {
  var auth = requireAuth_(token, 'SYS', 'view', 'SYS_USR', 'view-sys-usr');
  if (!auth.ok) return auth.res;
  try {
    var sheet = getUserSheet_();
    if (!sheet) return { success: false, message: 'Chưa cấu hình bảng dữ liệu SYS_USR' };

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, list: [] };

    var C = AUTH_.COL;
    var list = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var uName = String(row[C.USERNAME] || '').trim();
      var fName = String(row[C.FULLNAME] || '').trim();
      var cccd = String(row[C.CCCD] || '').replace(/'/g, '').trim();
      var rRole = normalizeRoleCode_(row[C.ROLE]) || 'EMPLOYEE';
      var mods = String(row[C.MODULES] || '').trim() || 'ALL';
      var active = isActiveValue_(row[C.ACTIVE]);
      var creator = String(row[C.CREATED_BY] || '').trim();

      if (uName && uName.toLowerCase() !== 'username') {
        list.push({
          username: uName,
          full_name: fName,
          cccd_no: cccd,
          role: rRole,
          module_access: mods,
          nav_access: String(row[C.NAV_ACCESS] || '').trim(),
          is_active: active,
          status: active ? 'ACTIVE' : 'LOCKED',
          last_login: row[C.LAST_LOGIN] ? fmtDateTime_(row[C.LAST_LOGIN]) : '',
          created_by: creator ? creator : 'admin',
          created_at: row[C.CREATED_AT] ? fmtDateTime_(row[C.CREATED_AT]) : ''
        });
      }
    }

    return { success: true, list: list };
  } catch (err) {
    return { success: false, message: 'Lỗi nạp danh sách SYS_USR: ' + err.toString() };
  }
}

/**
 * Ghi trực tiếp xuống Google Sheet SYS_USR chuẩn theo vị trí cột
 */
function saveSysUserLegacy_(payload) {
  try {
    var sheet = getUserSheet_();
    if (!sheet) return { success: false, message: 'Không tìm thấy sheet SYS_USR' };

    var uName = String(payload.username || '').trim();
    var mode = payload.mode || 'CREATE';

    if (!uName) return { success: false, message: 'Tên đăng nhập không được để trống!' };

    var C = AUTH_.COL;
    var data = sheet.getDataRange().getValues();
    var targetRow = -1;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][C.USERNAME]).trim().toLowerCase() === uName.toLowerCase()) {
        targetRow = i + 1;
        break;
      }
    }

    if (mode === 'CREATE') {
      if (targetRow > 0) {
        return { success: false, message: 'Tên đăng nhập [' + uName + '] đã tồn tại trong hệ thống!' };
      }
      if (!payload.password) {
        return { success: false, message: 'Vui lòng nhập mật khẩu cho tài khoản mới!' };
      }

      var passHash = makePasswordHash_(payload.password);
      
      // Tạo mảng đủ 10 cột theo đúng AUTH_.COL
      var newRow = new Array(10);
      newRow[C.USERNAME] = uName;
      newRow[C.PASSWORD] = passHash;
      newRow[C.CCCD] = String(payload.cccd_no || '').trim();
      newRow[C.FULLNAME] = String(payload.full_name || '').trim();
      newRow[C.ROLE] = requestedRole;
      newRow[C.MODULES] = String(payload.module_access || 'ALL').trim();
      newRow[C.ACTIVE] = (payload.status === 'ACTIVE' || payload.is_active === true);
      newRow[C.LAST_LOGIN] = '';
      newRow[C.CREATED_BY] = auth.user.username;
      newRow[C.CREATED_AT] = new Date();

      sheet.appendRow(newRow);
      SpreadsheetApp.flush(); // Ép Google Sheet lưu dữ liệu ngay lập tức
      return { success: true, message: 'Tạo thành công tài khoản [' + uName + ']!' };

    } else {
      // MODE UPDATE (SỬA TÀI KHOẢN)
      if (targetRow <= 0) {
        return { success: false, message: 'Không tìm thấy tài khoản [' + uName + '] để cập nhật!' };
      }

      sheet.getRange(targetRow, C.CCCD + 1).setValue(String(payload.cccd_no || '').trim());
      sheet.getRange(targetRow, C.FULLNAME + 1).setValue(String(payload.full_name || '').trim());
      sheet.getRange(targetRow, C.ROLE + 1).setValue(normalizeRoleCode_(payload.role) || 'EMPLOYEE');
      sheet.getRange(targetRow, C.MODULES + 1).setValue(String(payload.module_access || 'ALL').trim());
      sheet.getRange(targetRow, C.ACTIVE + 1).setValue(payload.status === 'ACTIVE' || payload.is_active === true);

      // Nếu có nhập mật khẩu mới thì cập nhật băm mật khẩu
      if (payload.password && String(payload.password).trim() !== '') {
        var newPassHash = makePasswordHash_(payload.password);
        sheet.getRange(targetRow, C.PASSWORD + 1).setValue(newPassHash);
      }

      SpreadsheetApp.flush(); // Ép Google Sheet lưu dữ liệu ngay lập tức
      return { success: true, message: 'Đã cập nhật thông tin tài khoản [' + uName + '] thành công!' };
    }

  } catch (err) {
    return { success: false, message: 'Lỗi lưu tài khoản: ' + err.toString() };
  }
}

/**
 * Hàm Wrapper lưu tài khoản - Bỏ qua vướng mắc Token nếu dùng tài khoản Quản trị
 */
function saveSysUserApi(token, payload) {
  return saveSysUser(token, payload);
}

function getSysUsersListApi(token) {
  return getSysUsersList(token);
}

// Khai báo alias để router gọi tên nào cũng chạy thành công
function getSysUsersListApi(token) {
  return getSysUsersList(token);
}

/* ============================================================
 * 6. MA TRẬN PHÂN QUYỀN (MATRIX AUTHORIZATION - DYNAMIC RBAC/ABAC)
 * ============================================================ */

function migrateLegacyRoleCodes_(auth) {
  if (!isPermissionAdminRole_(auth.user.role)) return { success: true, convertedCount: 0 };

  var migrationProperty = 'SYS_USR_ROLE_CODES_MIGRATION_V2';
  var properties = PropertiesService.getScriptProperties();
  if (properties.getProperty(migrationProperty) === 'DONE') {
    return { success: true, convertedCount: 0 };
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { success: false, message: 'Không thể khóa dữ liệu để chuyển role MANAGER. Vui lòng tải lại sau.' };
  }
  try {
    var usersSheet = getUserSheet_();
    if (!usersSheet) return { success: false, message: 'Không tìm thấy bảng dữ liệu SYS_USR.' };
    var C = AUTH_.COL;
    var lastUserRow = usersSheet.getLastRow();
    var roleChanges = [];
    if (lastUserRow > 1) {
      var userRoles = usersSheet.getRange(2, C.ROLE + 1, lastUserRow - 1, 1).getValues();
      for (var i = 0; i < userRoles.length; i++) {
        var oldRole = String(userRoles[i][0] || '').trim();
        var normalizedRole = normalizeRoleCode_(oldRole);
        if (normalizedRole && normalizedRole !== oldRole) {
          roleChanges.push({ row: i + 2, value: normalizedRole });
        }
      }
    }

    roleChanges.forEach(function(change) {
      usersSheet.getRange(change.row, C.ROLE + 1).setValue(change.value);
    });
    if (roleChanges.length) SpreadsheetApp.flush();
    for (var j = 0; j < roleChanges.length; j++) {
      if (String(usersSheet.getRange(roleChanges[j].row, C.ROLE + 1).getValue()).trim() !== roleChanges[j].value) {
        return { success: false, message: 'Không xác minh được ánh xạ role cũ. Vui lòng rà soát SYS_USR.' };
      }
    }
    if (roleChanges.length) {
      auditLog_(auth.user, 'SYS', 'MIGRATE_ROLE_CODES', 'SYS_USR', 'converted=' + roleChanges.length);
    }
    properties.setProperty(migrationProperty, 'DONE');
    return { success: true, convertedCount: roleChanges.length };
  } catch (err) {
    Logger.log('migrateLegacyRoleCodes_: ' + err);
    return { success: false, message: 'Không thể ánh xạ role cũ: ' + String(err && err.message ? err.message : err) };
  } finally {
    lock.releaseLock();
  }
}

function toSysUserListItem_(row) {
  var C = AUTH_.COL;
  var isActive = isActiveValue_(row[C.ACTIVE]);
  return {
    username: String(row[C.USERNAME] || '').trim(),
    full_name: String(row[C.FULLNAME] || '').trim(),
    cccd_no: String(row[C.CCCD] || '').replace(/'/g, '').trim(),
    role: normalizeRoleCode_(row[C.ROLE]) || 'EMPLOYEE',
    nav_access: String(row[C.NAV_ACCESS] || '').trim(),
    is_active: isActive,
    status: isActive ? 'ACTIVE' : 'LOCKED',
    last_login: row[C.LAST_LOGIN] ? fmtDateTime_(row[C.LAST_LOGIN]) : '',
    created_by: String(row[C.CREATED_BY] || '').trim(),
    created_at: row[C.CREATED_AT] ? fmtDateTime_(row[C.CREATED_AT]) : ''
  };
}

function writeSysUserCccdAsText_(sheet, rowNumber, cccd) {
  var cccdCell = sheet.getRange(rowNumber, AUTH_.COL.CCCD + 1);
  cccdCell.setNumberFormat('@');
  cccdCell.setValue(String(cccd || '').trim());
}

/**
 * Lấy danh sách tất cả người dùng từ sheet SYS_USR cho giao diện Phân quyền
 */
function getSysUsersList(token) {
  var auth = requireAuth_(token, 'SYS', 'view', 'SYS_USR', 'view-sys-usr');
  if (!auth.ok) return auth.res;
  var migration = migrateLegacyRoleCodes_(auth);
  if (!migration.success) return migration;
  try {
    var sheet = getUserSheet_();
    if (!sheet) return { success: false, message: 'Chưa cấu hình Sheet SYS_USR' };

    var C = AUTH_.COL;
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, list: [], migratedRoleCount: migration.convertedCount };
    var rowCount = lastRow - 1;
    var usernames = sheet.getRange(2, C.USERNAME + 1, rowCount, 1).getValues();
    var details = sheet.getRange(2, C.CCCD + 1, rowCount, C.NAV_ACCESS - C.CCCD + 1).getValues();
    var list = [];

    // ADMIN/SuperAdmin không cần đọc SYS_PERM (đã bypass ở matrixRowAllowed_) -> tránh đọc thừa.
    var isAdminRole = isPermissionAdminRole_(auth.user.role);
    var dataPermission = isAdminRole ? null : getEffectivePermission_(auth.user.username, auth.user.role, 'SYS_USR');

    for (var i = 0; i < rowCount; i++) {
      var row = [];
      row[C.USERNAME] = usernames[i][0];
      for (var column = C.CCCD; column <= C.NAV_ACCESS; column++) {
        row[column] = details[i][column - C.CCCD];
      }
      var uName = String(row[C.USERNAME] || '').trim();

      if (uName && uName.toLowerCase() !== 'username' &&
          (isAdminRole ||
           matrixRowAllowed_(auth.user, 'SYS_USR', row, { owner: C.CREATED_BY, status: C.ACTIVE }, dataPermission))) {
        list.push(toSysUserListItem_(row));
      }
    }

    return { success: true, list: list, migratedRoleCount: migration.convertedCount };
  } catch (err) {
    return { success: false, message: 'Lỗi nạp SYS_USR: ' + err.toString() };
  }
}

function getSysUserLoginHistory(token, targetUsername) {
  var auth = requireAuth_(token, 'SYS', 'manage_accounts', 'SYS_USR', 'view-sys-usr');
  if (!auth.ok) return auth.res;

  var username = String(targetUsername || '').trim();
  if (!username) return fail_('Thiếu tên tài khoản cần xem lịch sử.');

  try {
    var usersSheet = getUserSheet_();
    if (!usersSheet) return fail_('Không tìm thấy sheet SYS_USR.');
    var users = usersSheet.getDataRange().getValues();
    var C = AUTH_.COL;
    var targetRow = null;
    for (var i = 1; i < users.length; i++) {
      if (String(users[i][C.USERNAME] || '').trim().toLowerCase() === username.toLowerCase()) {
        targetRow = users[i];
        break;
      }
    }
    if (!targetRow) return fail_('Không tìm thấy tài khoản cần xem lịch sử.');

    var targetRole = String(targetRow[C.ROLE] || '').trim().toUpperCase();
    if (!isSuperAdminRole_(auth.user.role) &&
        (targetRole === 'ADMIN' || targetRole === AUTH_.SUPER_ROLE.toUpperCase())) {
      return { success: false, code: 'FORBIDDEN', message: 'Chỉ SuperAdmin được xem lịch sử đăng nhập của tài khoản ADMIN/SuperAdmin.' };
    }

    var auditSpreadsheet = SpreadsheetApp.openById(CONFIG.DB_SYS_PART1_ID);
    var auditSheet = auditSpreadsheet.getSheetByName(CONFIG.SHEET_SYS_AUDIT || 'SYS_AUDIT');
    if (!auditSheet || auditSheet.getLastRow() <= 1) return { success: true, list: [] };

    var lastRow = auditSheet.getLastRow();
    var scanRows = Math.min(lastRow - 1, 5000);
    var auditRows = auditSheet.getRange(lastRow - scanRows + 1, 1, scanRows, 6).getValues();
    var events = [];
    for (var j = auditRows.length - 1; j >= 0 && events.length < 10; j--) {
      var row = auditRows[j];
      var action = String(row[3] || '').trim().toUpperCase();
      if (String(row[1] || '').trim().toLowerCase() !== username.toLowerCase() ||
          (action !== 'LOGIN' && action !== 'LOGOUT')) continue;
      events.push({
        time: fmtDateTime_(row[0]),
        action: action,
        ip_address: '',
        user_agent: ''
      });
    }
    return { success: true, list: events };
  } catch (err) {
    Logger.log('getSysUserLoginHistory: ' + err);
    return { success: false, message: 'Không tải được lịch sử đăng nhập: ' + err.message, list: [] };
  }
}

function getSysUsersListApi(token) {
  return getSysUsersList(token);
}

/**
 * Lấy quy tắc phân quyền hiệu lực của một User/Role đối với 1 Sheet cụ thể
 */
function getEffectivePermission_(username, role, sheetName) {
  var roleCode = normalizeRoleCode_(role) || String(role || '').trim().toUpperCase();

  // CacheService (5 phút) theo user + sheet, giảm đọc SYS_PERM mỗi request.
  var cacheKey = 'PERM_EFF|' + String(username || '').trim().toLowerCase() + '|' + String(sheetName || '').trim().toUpperCase();
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get(cacheKey);
    if (cached) {
      var parsed = JSON.parse(cached);
      if (parsed && parsed._role === roleCode) return parsed.value;
    }
  } catch (e) {}
  var result = isSuperAdminRole_(roleCode) ? {
      canRead: true, canCreate: true, canEdit: true, canDelete: true,
      rules: [{ conditions: [], canRead: true, canCreate: true, canEdit: true, canDelete: true }],
      customActions: "APPROVE,EXPORT,PRINT,RESTORE"
    } : requestMemo_('PERM_EFF|' + String(username).trim().toLowerCase() + '|' + roleCode + '|' + String(sheetName).trim().toUpperCase(), function () {
      return computeEffectivePermission_(username, role, roleCode, sheetName);
    });
  try {
    var cache = CacheService.getScriptCache();
    cache.put(cacheKey, JSON.stringify({ _role: roleCode, value: result }), 300);
  } catch (e) {}
  return result;
}

function computeEffectivePermission_(username, role, roleCode, sheetName) {
  var defaultPerm = { canRead: false, canCreate: false, canEdit: false, canDelete: false, rules: [], customActions: "" };

  try {
    var ss = SpreadsheetApp.openById(CONFIG.DB_SYS_PART1_ID);
    var sheet = ss.getSheetByName('SYS_PERM');
    if (!sheet) return defaultPerm;

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return defaultPerm;
    var targetNames = getPermissionSheetNames_(sheetName);
    var schema = getPermSchema_();
    var effective = { canRead: false, canCreate: false, canEdit: false, canDelete: false, rules: [], customActions: [] };
    for (var i = 1; i < data.length; i++) {
      var r = readPermRow_(data[i], schema);
      if (r.targetType.toUpperCase() === 'USER' &&
          r.targetValue.toLowerCase() === String(username).trim().toLowerCase() &&
          targetNames.indexOf(r.sheetName.toUpperCase()) !== -1) {
        var actions = r.customActions.split(',').map(function(action) { return action.trim().toUpperCase(); });
        actions.forEach(function(action) {
          if (action && effective.customActions.indexOf(action) === -1) effective.customActions.push(action);
        });
        var permId = r.permId;
        var conditions = getPermConditions_(permId);
        if (!Array.isArray(conditions) || conditions.length === 0) continue;
        if (conditions.length > 12) throw new Error('Cấu hình rule ma trận không hợp lệ (quá 12 điều kiện).');
        var rule = {
          conditions: conditions,
          canRead: r.canRead,
          canCreate: r.canCreate,
          canEdit: r.canEdit,
          canDelete: r.canDelete
        };
        effective.rules.push(rule);
        effective.canRead = effective.canRead || rule.canRead;
        effective.canCreate = effective.canCreate || rule.canCreate;
        effective.canEdit = effective.canEdit || rule.canEdit;
        effective.canDelete = effective.canDelete || rule.canDelete;
      }
    }
    effective.customActions = effective.customActions.join(',');
    return effective;
  } catch (err) {
    Logger.log("Lỗi đọc SYS_PERM: " + err.toString());
    return defaultPerm;
  }
}

/**
 * Quét toàn bộ Schema các Sheet & Cột trong hệ thống để phục vụ giao diện Phân quyền
 */
function getDistinctColumnValuesApi(token, sheetName, columnIndex) {
  var auth = requireAuth_(token, 'SYS', 'view', 'SYS_PERM', 'view-sys-usr');
  if (!auth.ok) return auth.res;
  try {
    Logger.log('getDistinctColumnValuesApi: sheetName=' + sheetName + ', columnIndex=' + columnIndex);
    if (!isFinite(columnIndex) || columnIndex < 0) {
      return { success: false, message: 'Chỉ số cột không hợp lệ.' };
    }
    var dbIds = [CONFIG.DB_SYS_PART1_ID, CONFIG.DB_HRM_PART1_ID, CONFIG.DB_WMS_PART1_ID, CONFIG.DB_DMS_PART1_ID];
    var found = null;
    var lastError = null;
    var foundDbId = null;
    for (var i = 0; i < dbIds.length; i++) {
      var dbId = dbIds[i];
      if (!dbId || dbId.indexOf('NHẬP_ID') !== -1) continue;
      if (found) break;
      Logger.log('getDistinctColumnValuesApi: thử mở DB #' + i + ' id=' + dbId);
      try {
        var ss = SpreadsheetApp.openById(dbId);
        var sh = ss.getSheetByName(sheetName);
        if (sh) {
          found = sh;
          foundDbId = dbId;
          Logger.log('getDistinctColumnValuesApi: TÌM THẤY sheet "' + sheetName + '" trong DB id=' + dbId);
        } else {
          Logger.log('getDistinctColumnValuesApi: KHÔNG TÌM THẤY sheet "' + sheetName + '" trong DB id=' + dbId);
        }
      } catch (e) {
        lastError = e;
        Logger.log('getDistinctColumnValuesApi: LỖI mở DB id=' + dbId + ': ' + e.toString());
        continue;
      }
    }
    if (!found) {
      if (lastError) {
        Logger.log('getDistinctColumnValuesApi: KHÔNG TÌM THẤY sheet, lỗi cuối=' + lastError.toString());
        return { success: false, message: lastError.toString() };
      }
      Logger.log('getDistinctColumnValuesApi: KHÔNG TÌM THẤY sheet "' + sheetName + '" trong tất cả DB.');
      return { success: false, message: 'Không tìm thấy sheet ' + sheetName };
    }
    var lastRow = found.getLastRow();
    var lastColumn = found.getLastColumn();
    Logger.log('getDistinctColumnValuesApi: sheet="' + sheetName + "' lastRow=" + lastRow + ' lastColumn=' + lastColumn + ' foundDbId=' + foundDbId);
    if (lastRow < 2) return { success: true, values: [] };
    if (columnIndex >= lastColumn) {
      return { success: false, message: 'Cột #' + (columnIndex + 1) + ' vượt quá số cột của sheet (' + lastColumn + ').' };
    }
    var vals = found.getRange(2, columnIndex + 1, lastRow - 1, 1).getValues();
    var set = {};
    vals.forEach(function(r) { var v = String(r[0] || '').trim(); if (v) set[v] = true; });
    var resultValues = Object.keys(set).sort();
    Logger.log('getDistinctColumnValuesApi: cột #' + columnIndex + ' trả về ' + resultValues.length + ' giá trị: ' + JSON.stringify(resultValues.slice(0, 10)));
    return { success: true, values: resultValues };
  } catch (err) {
    Logger.log('getDistinctColumnValuesApi: NGOẠI LỆ: ' + err.toString());
    return { success: false, message: err.toString() };
  }
}
function getSystemSchemaApi(token) {
  var auth = requireAuth_(token, 'SYS', 'view', 'SYS_PERM', 'view-sys-usr');
  if (!auth.ok) return auth.res;
  try {
    var schema = [];
    var registeredActions = JSON.parse(
      PropertiesService.getScriptProperties().getProperty('SYS_PERMISSION_ACTION_CATALOG_V1') || '{}'
    );
    var configuredActionsBySheet = {};
    var permissionSheet = SpreadsheetApp.openById(CONFIG.DB_SYS_PART1_ID).getSheetByName('SYS_PERM');
    if (permissionSheet && permissionSheet.getLastRow() > 1) {
      var permissionRows = permissionSheet.getDataRange().getValues();
      var permSchema = getPermSchema_();
      permissionRows.slice(1).forEach(function(row) {
        var perm = readPermRow_(row, permSchema);
        var tableName = perm.sheetName;
        perm.customActions.split(',').forEach(function(action) {
          var actionId = action.trim().toUpperCase();
          if (!tableName || !actionId) return;
          if (!configuredActionsBySheet[tableName]) configuredActionsBySheet[tableName] = [];
          if (configuredActionsBySheet[tableName].indexOf(actionId) === -1) {
            configuredActionsBySheet[tableName].push(actionId);
          }
        });
      });
    }
    
    // List các DB Spreadsheet ID cần quét (giữ nguyên thứ tự để getDistinctColumnValuesApi tìm trùng sheet)
    var dbIds = [CONFIG.DB_SYS_PART1_ID, CONFIG.DB_HRM_PART1_ID, CONFIG.DB_WMS_PART1_ID, CONFIG.DB_DMS_PART1_ID];

    dbIds.forEach(function(dbId) {
      if (!dbId || dbId.indexOf('NHẬP_ID') !== -1) return;
      try {
        var ss = SpreadsheetApp.openById(dbId);
        var sheets = ss.getSheets();
        sheets.forEach(function(sh) {
          var name = sh.getName();
          if (CONFIG.BYPASS_PERMISSION_SHEETS && CONFIG.BYPASS_PERMISSION_SHEETS.indexOf(String(name).toUpperCase()) !== -1) return;
          if (name.indexOf('SYS_') === 0 || name.indexOf('HRM_') === 0 ||
             name.indexOf('WMS_') === 0 || name === CONFIG.SHEET_WMS_WAREHOUSE ||
             name === CONFIG.SHEET_WMS_LOCATION || name === CONFIG.SHEET_WMS_SKU) {
            try {
              var headers = getSheetHeaders_(sh);
            } catch (e) {
              Logger.log('getSystemSchemaApi: lỗi đọc header sheet ' + name + ': ' + e);
              return;
            }
            var builtInActions = getBuiltInPermissionActions_(name);
            var actionIds = builtInActions.map(function(action) { return action.id; })
              .concat(Array.isArray(registeredActions[name]) ? registeredActions[name] : [])
              .concat(configuredActionsBySheet[name] || [])
              .filter(function(action, index, all) { return all.indexOf(action) === index; });
            schema.push({
              sheetName: name,
              columns: headers.map(function(header, index) {
                return { index: index, name: String(header || '').trim() };
              }).filter(function(column) { return column.name; }),
              actions: actionIds.map(function(actionId) {
                var builtInAction = builtInActions.find(function(action) { return action.id === actionId; });
                return builtInAction || { id: actionId, label: getPermissionActionLabel_(actionId), group: 'Hành động khác' };
              })
            });
          }
        });
      } catch (e) {
        Logger.log('getSystemSchemaApi: lỗi mở database ' + dbId + ': ' + e);
      }
    });

    return { success: true, schema: schema };
  } catch (err) {
    return { success: false, message: 'Lỗi quét Dynamic Schema: ' + err.toString() };
  }
}

/**
 * Lưu dòng quy tắc Phân quyền mới/cập nhật vào sheet SYS_PERM
 */
function getSystemTableHeaders_(sheetName) {
  var dbIds = [CONFIG.DB_SYS_PART1_ID, CONFIG.DB_HRM_PART1_ID, CONFIG.DB_WMS_PART1_ID, CONFIG.DB_DMS_PART1_ID];
  for (var i = 0; i < dbIds.length; i++) {
    if (!dbIds[i] || String(dbIds[i]).indexOf('NHẬP_ID') !== -1) continue;
    try {
      var dataSheet = SpreadsheetApp.openById(dbIds[i]).getSheetByName(sheetName);
      if (dataSheet) {
        return getSheetHeaders_(dataSheet);
      }
    } catch (e) {
      continue;
    }
  }
  return null;
}

function getKnownPermissionActions_(sheetName, permissionSheet) {
  var catalog = JSON.parse(
    PropertiesService.getScriptProperties().getProperty('SYS_PERMISSION_ACTION_CATALOG_V1') || '{}'
  );
  var actions = getBuiltInPermissionActions_(sheetName).map(function(action) { return action.id; });
  (Array.isArray(catalog[sheetName]) ? catalog[sheetName] : []).forEach(function(actionId) {
    if (actions.indexOf(actionId) === -1) actions.push(actionId);
  });
  if (permissionSheet && permissionSheet.getLastRow() > 1) {
    var sheetNames = getPermissionSheetNames_(sheetName);
    var permSchema = getPermSchema_();
    permissionSheet.getDataRange().getValues().slice(1).forEach(function(row) {
      var perm = readPermRow_(row, permSchema);
      if (perm.targetType.toUpperCase() !== 'USER' ||
          sheetNames.indexOf(perm.sheetName.toUpperCase()) === -1) return;
      perm.customActions.split(',').forEach(function(action) {
        var actionId = action.trim().toUpperCase();
        if (actionId && actions.indexOf(actionId) === -1) actions.push(actionId);
      });
    });
  }
  return actions;
}

function savePermissionRuleApi(token, rule) {
  var auth = requireAuth_(token, 'SYS', 'manage_permissions', 'SYS_PERM', 'view-sys-usr');
  if (!auth.ok || !isPermissionAdminRole_(auth.user.role)) {
    return auth.ok ? { success: false, code: 'FORBIDDEN', message: 'Chỉ ADMIN hoặc SuperAdmin được cấu hình ma trận phân quyền.' } : auth.res;
  }
  // P5: khóa toàn hệ thống để 2 người không ghi đè rule của cùng 1 bảng.
  return withLock_(function () {
    return doSavePermissionRule_(rule);
  });
}

function doSavePermissionRule_(rule) {
  try {
    if (!rule || !rule.sheetName || !rule.targetValue ||
        String(rule.targetType || '').trim().toUpperCase() !== 'USER') {
      return { success: false, message: 'Dữ liệu cấu hình phân quyền không hợp lệ!' };
    }
    var targetType = 'USER';
    var targetValue = String(rule.targetValue).trim();

    var ss = SpreadsheetApp.openById(CONFIG.DB_SYS_PART1_ID);
    var sheet = ss.getSheetByName('SYS_PERM');
    
    // Tự động khởi tạo Sheet SYS_PERM nếu chưa có (header theo thứ tự chuẩn)
    if (!sheet) {
      sheet = ss.insertSheet('SYS_PERM');
      sheet.appendRow(PERM_COLUMNS_.slice());
      _permSchema_ = null; // đọc lại header vừa tạo
    }

    if (!Array.isArray(rule.rules) || rule.rules.length > 20) {
      return { success: false, message: 'Danh sách rule không hợp lệ (tối đa 20 rule).' };
    }
    var headers = getSystemTableHeaders_(String(rule.sheetName).trim());
    if (!headers) return { success: false, message: 'Không tìm thấy bảng dữ liệu đã chọn.' };
    var validatedRules = [];
    for (var ruleIndex = 0; ruleIndex < rule.rules.length; ruleIndex++) {
      var requestedRule = rule.rules[ruleIndex] || {};
      if (!Array.isArray(requestedRule.conditions) || requestedRule.conditions.length < 1 ||
          requestedRule.conditions.length > 12) {
        return { success: false, message: 'Rule số ' + (ruleIndex + 1) + ' cần có từ 1 đến 12 điều kiện.' };
      }
      var conditions = [];
      for (var conditionIndex = 0; conditionIndex < requestedRule.conditions.length; conditionIndex++) {
        var condition = requestedRule.conditions[conditionIndex] || {};
        var columnIndex = Number(condition.columnIndex);
        var operator = String(condition.operator || '').trim().toLowerCase();
        var conditionValue = String(condition.value == null ? '' : condition.value).trim();
        var allValues = condition.allValues === true;
        if (!isFinite(columnIndex) || columnIndex < 0 || Math.floor(columnIndex) !== columnIndex ||
            columnIndex >= headers.length || !headers[columnIndex] ||
            String(condition.columnName || '').trim().toLowerCase() !== headers[columnIndex].toLowerCase() ||
            ['equals', 'not_equals', 'contains', 'one_of'].indexOf(operator) === -1 ||
            (!allValues && !conditionValue) || conditionValue.length > 500) {
          return { success: false, message: 'Điều kiện ' + (conditionIndex + 1) + ' của rule ' +
            (ruleIndex + 1) + ' không hợp lệ hoặc cột đã thay đổi.' };
        }
        conditions.push({
          columnIndex: columnIndex,
          columnName: headers[columnIndex],
          operator: operator,
          value: allValues ? '' : conditionValue,
          allValues: allValues,
          dataType: String(condition.dataType || '').trim()
        });
      }
      validatedRules.push({
        conditions: conditions,
        canRead: requestedRule.canRead === true,
        canCreate: requestedRule.canCreate === true,
        canEdit: requestedRule.canEdit === true,
        canDelete: requestedRule.canDelete === true
      });
    }
    var requestedActions = String(rule.customActions || '').split(',').map(function(action) {
      return action.trim().toUpperCase();
    }).filter(function(action, index, actions) { return action && actions.indexOf(action) === index; });
    var knownActions = getKnownPermissionActions_(String(rule.sheetName).trim(), sheet);
    if (requestedActions.some(function(action) {
      return !/^[A-Z][A-Z0-9_]{1,63}$/.test(action) || knownActions.indexOf(action) === -1;
    })) {
      return { success: false, message: 'Danh sách hành động bổ sung chứa hành động chưa được đăng ký cho bảng dữ liệu này.' };
    }

    var permSchema = getPermSchema_();
    var data = sheet.getDataRange().getValues();
    var existingRows = [];
    var targetSheetNames = getPermissionSheetNames_(rule.sheetName);
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (String(row[permSchema['target_type']] || '').toUpperCase() === targetType &&
          String(row[permSchema['target_value']] || '').trim().toLowerCase() === targetValue.toLowerCase() &&
          targetSheetNames.indexOf(String(row[permSchema['sheet_name']] || '').trim().toUpperCase()) !== -1) {
        existingRows.push(i + 1);
      }
    }

    var rowsToWrite = validatedRules.map(function(config, index) {
      config.permId = 'P' + new Date().getTime() + '_' + index;
      return [
        config.permId,
        targetType,
        targetValue,
        String(rule.sheetName).trim(),
        'FALSE',
        '',
        'ALL',
        config.canRead ? 'TRUE' : 'FALSE',
        config.canCreate ? 'TRUE' : 'FALSE',
        config.canEdit ? 'TRUE' : 'FALSE',
        config.canDelete ? 'TRUE' : 'FALSE',
        requestedActions.join(',')
      ];
    });
    if (!rowsToWrite.length && requestedActions.length) {
      var permIdA = 'P' + new Date().getTime() + '_A';
      rowsToWrite.push([
        permIdA, targetType, targetValue, String(rule.sheetName).trim(),
        'FALSE', '', 'ALL', 'FALSE', 'FALSE', 'FALSE', 'FALSE', requestedActions.join(',')
      ]);
    }
    var deleteRanges = [];
    existingRows.sort(function(a, b) { return a - b; }).forEach(function(rowNumber) {
      var lastRange = deleteRanges[deleteRanges.length - 1];
      if (lastRange && rowNumber === lastRange.end + 1) lastRange.end = rowNumber;
      else deleteRanges.push({ start: rowNumber, end: rowNumber });
    });
    deleteRanges.reverse().forEach(function(range) {
      sheet.deleteRows(range.start, range.end - range.start + 1);
    });
    if (rowsToWrite.length) {
      var startRow = sheet.getLastRow() + 1;
       var cols = PERM_COLUMNS_.filter(function (name) { return permSchema[name] !== undefined; });
      var payload = rowsToWrite.map(function (rowArr) {
        var values = {};
        PERM_COLUMNS_.forEach(function (name, i) { values[name] = rowArr[i]; });
        var arr = new Array(sheet.getLastColumn());
        cols.forEach(function (name) {
           arr[permSchema[name]] = values[name] === undefined || values[name] === null ? '' : values[name];
        });
        return arr;
      });
      sheet.getRange(startRow, 1, payload.length, payload[0].length).setValues(payload);
    }
    SpreadsheetApp.flush();

    // Lưu điều kiện chi tiết vào sheet SYS_PERM_CONDITIONS (không còn JSON trong ô)
    validatedRules.forEach(function(config) {
      if (config.permId) {
        savePermConditions_(config.permId, config.conditions || []);
      }
    });

    var savedRows = sheet.getDataRange().getValues();
    var savedNames = getPermissionSheetNames_(rule.sheetName);
    var savedCount = savedRows.slice(1).filter(function(row) {
      return String(row[permSchema['target_type']]).trim().toUpperCase() === targetType &&
        String(row[permSchema['target_value']]).trim().toLowerCase() === targetValue.toLowerCase() &&
        savedNames.indexOf(String(row[permSchema['sheet_name']]).trim().toUpperCase()) !== -1;
    }).length;
    if (savedCount !== rowsToWrite.length) {
      return { success: false, message: 'Google Sheets chưa xác nhận đủ rule ma trận đã ghi vào SYS_PERM.' };
    }

    invalidatePermissionCache_(targetValue, String(rule.sheetName).trim());

    return { success: true, message: 'Đã lưu cấu hình phân quyền thành công!' };
  } catch (err) {
    return { success: false, message: 'Lỗi lưu cấu hình phân quyền: ' + String(err && err.message ? err.message : err) };
  }
}

/** Xóa cache phân quyền của một user trên các sheet liên quan. */
function invalidatePermissionCache_(targetValue, sheetName) {
  try {
    var cache = CacheService.getScriptCache();
    var username = String(targetValue || '').trim().toLowerCase();
    var names = sheetName ? getPermissionSheetNames_(sheetName) : [];
    var keys = [];
    if (names.length) {
      names.forEach(function(name) {
        keys.push('PERM_' + username + '_' + String(name).trim().toUpperCase());
        keys.push('PERM_EFF|' + username + '|' + String(name).trim().toUpperCase());
      });
    } else {
      // Nếu không biết sheet, xóa toàn bộ cache PERM của user này (ít chính xác nhưng an toàn)
      // CacheService không hỗ trợ delete theo prefix, nên bỏ qua.
    }
    if (keys.length) cache.removeAll(keys);
  } catch (e) {
    // ignore
  }
  clearRequestMemo_();
}

/**
 * Danh sách tên bảng có thể đang có cache phân quyền (dùng để xóa cache khi đổi vai trò tài khoản).
 * Gồm danh sách cố định của toàn hệ thống + các sheet_name xuất hiện thực tế trong SYS_PERM.
 */
function getKnownPermissionSheetNames_() {
  return requestMemo_('SYS_PERM_SHEET_NAMES', function () {
    var names = {};
    var staticNames = [
      'SYS_USR', 'SYS_PERM', 'HRM_ID', 'HRM_EMP', 'HRM_ATT',
      'WAREHOUSE', 'LOCATION', 'SKU', 'DOC', 'WMS_DOC', 'LINE', 'WMS_LINE',
      'WMS_POST', 'WMS_LOT', 'WMS_QC', 'WMS_LEDGER', 'WMS_BAL',
      'WMS_WF', 'WMS_WF_STEP', 'WMS_WF_RUN', 'WMS_WF_TASK', 'WMS_WF_LOG',
      'WMS_PROD', 'WMS_GENEALOGY', 'WMS_ATTACH', 'WMS_USER_MAP',
      'DMS_DOCS', 'DMS_VERSIONS', 'DMS_COMMENTS', 'DMS_READ_ACK'
    ];
    staticNames.forEach(function(name) { names[name] = true; });
    try {
      var sheet = SpreadsheetApp.openById(CONFIG.DB_SYS_PART1_ID).getSheetByName('SYS_PERM');
      if (sheet && sheet.getLastRow() > 1) {
        var rows = sheet.getRange(2, 4, sheet.getLastRow() - 1, 1).getValues();
        rows.forEach(function(row) {
          var name = String(row[0] || '').trim().toUpperCase();
          if (name) names[name] = true;
        });
      }
    } catch (e) {
      // ignore - dùng danh sách cố định
    }
    return Object.keys(names);
  });
}

/**
 * Xóa cache phân quyền của 1 user trên MỌI bảng.
 * Dùng khi đổi vai trò / trạng thái tài khoản vì cache PERM được tính theo vai trò lúc tạo.
 */
function invalidateAllPermissionCache_(username) {
  var uname = String(username || '').trim().toLowerCase();
  if (!uname) return;
  try {
    var keys = [];
    getKnownPermissionSheetNames_().forEach(function(name) {
      getPermissionSheetNames_(name).forEach(function(alias) {
        var key = 'PERM_' + uname + '_' + String(alias).trim().toUpperCase();
        if (keys.indexOf(key) === -1) keys.push(key);
      });
    });
    var cache = CacheService.getScriptCache();
    for (var i = 0; i < keys.length; i += 50) {
      cache.removeAll(keys.slice(i, i + 50));
    }
  } catch (e) {
    // ignore
  }
  clearRequestMemo_();
}

function savePermissionRulesBatchApi(token, entries) {
  if (!Array.isArray(entries) || !entries.length || entries.length > 100) {
    return { success: false, message: 'Danh sách bảng cần lưu không hợp lệ.' };
  }
  var targetValue = String(entries[0] && entries[0].targetValue || '').trim();
  if (!targetValue) return { success: false, message: 'Thiếu tài khoản cần cấu hình.' };
  var seenSheets = {};
  var permissionSheet = SpreadsheetApp.openById(CONFIG.DB_SYS_PART1_ID).getSheetByName('SYS_PERM');
  for (var validateIndex = 0; validateIndex < entries.length; validateIndex++) {
    var candidate = entries[validateIndex] || {};
    var sheetName = String(candidate.sheetName || '').trim();
    var normalizedSheetName = sheetName.toUpperCase();
    if (!sheetName || seenSheets[normalizedSheetName] ||
        String(candidate.targetValue || '').trim().toLowerCase() !== targetValue.toLowerCase() ||
        !Array.isArray(candidate.rules) || candidate.rules.length > 20) {
      return { success: false, message: 'Bảng hoặc danh sách quy tắc trong lô lưu không hợp lệ.' };
    }
    seenSheets[normalizedSheetName] = true;
    var headers = getSystemTableHeaders_(sheetName);
    if (!headers) return { success: false, message: 'Không tìm thấy bảng dữ liệu ' + sheetName + '.' };
    for (var ruleIndex = 0; ruleIndex < candidate.rules.length; ruleIndex++) {
      var requestedRule = candidate.rules[ruleIndex] || {};
      if (!Array.isArray(requestedRule.conditions) || requestedRule.conditions.length < 1 ||
          requestedRule.conditions.length > 12) {
        return { success: false, message: 'Quy tắc ' + (ruleIndex + 1) + ' của bảng ' + sheetName + ' không hợp lệ.' };
      }
      for (var conditionIndex = 0; conditionIndex < requestedRule.conditions.length; conditionIndex++) {
        var condition = requestedRule.conditions[conditionIndex] || {};
        var columnIndex = Number(condition.columnIndex);
        var operator = String(condition.operator || '').trim().toLowerCase();
        var value = String(condition.value == null ? '' : condition.value).trim();
        if (!isFinite(columnIndex) || columnIndex < 0 || Math.floor(columnIndex) !== columnIndex ||
            columnIndex >= headers.length || !headers[columnIndex] ||
            String(condition.columnName || '').trim().toLowerCase() !== headers[columnIndex].toLowerCase() ||
            ['equals', 'not_equals', 'contains', 'one_of'].indexOf(operator) === -1 ||
            (condition.allValues !== true && !value) || value.length > 500) {
          return {
            success: false,
            message: 'Điều kiện ' + (conditionIndex + 1) + ' của quy tắc ' +
              (ruleIndex + 1) + ' trên bảng ' + sheetName + ' không hợp lệ.'
          };
        }
      }
    }
    var actions = String(candidate.customActions || '').split(',').map(function(action) {
      return action.trim().toUpperCase();
    }).filter(function(action, index, list) { return action && list.indexOf(action) === index; });
    var knownActions = getKnownPermissionActions_(sheetName, permissionSheet);
    if (actions.some(function(action) { return knownActions.indexOf(action) === -1; })) {
      return { success: false, message: 'Bảng ' + sheetName + ' có quyền chức năng chưa được đăng ký.' };
    }
  }

  var savedSheets = [];
  for (var i = 0; i < entries.length; i++) {
    var entry = entries[i] || {};
    var result = savePermissionRuleApi(token, {
      targetType: 'USER',
      targetValue: targetValue,
      sheetName: entry.sheetName,
      rules: entry.rules,
      customActions: entry.customActions
    });
    if (!result || !result.success) {
      return {
        success: false,
        message: (result && result.message ? result.message : 'Không lưu được cấu hình.') +
          (savedSheets.length ? ' Đã lưu các bảng: ' + savedSheets.join(', ') + '.' : ''),
        savedSheets: savedSheets,
        failedSheet: String(entry.sheetName || '').trim()
      };
    }
    savedSheets.push(String(entry.sheetName || '').trim());
  }
  return {
    success: true,
    message: 'Đã lưu cấu hình cho ' + savedSheets.length + ' bảng dữ liệu.',
    savedSheets: savedSheets
  };
}

function permissionRuleDeleteKey_(rule) {
  var conditions = Array.isArray(rule && rule.conditions) ? rule.conditions.map(function(condition) {
    return {
      columnIndex: Number(condition.columnIndex),
      columnName: String(condition.columnName || ''),
      operator: String(condition.operator || '').trim().toLowerCase(),
      value: String(condition.value == null ? '' : condition.value),
      allValues: condition.allValues === true
    };
  }) : [];
  return JSON.stringify({
    conditions: conditions,
    canRead: rule && rule.canRead === true,
    canCreate: rule && rule.canCreate === true,
    canEdit: rule && rule.canEdit === true,
    canDelete: rule && rule.canDelete === true
  });
}

function deletePermissionRulesApi(token, targetValue, deletions) {
  targetValue = String(targetValue || '').trim();
  if (!targetValue || !Array.isArray(deletions) || !deletions.length || deletions.length > 100) {
    return { success: false, message: 'Danh sách quyền cần xóa không hợp lệ.' };
  }
  var auth = requireAuth_(token, 'SYS', 'manage_permissions', 'SYS_PERM', 'view-sys-usr');
  if (!auth.ok || !isPermissionAdminRole_(auth.user.role)) {
    return auth.ok
      ? { success: false, code: 'FORBIDDEN', message: 'Chỉ ADMIN hoặc SuperAdmin được xóa quyền ma trận.' }
      : auth.res;
  }

  // P5: khóa toàn hệ thống để tránh xóa nhầm dòng khi có thao tác ghi song song.
  return withLock_(function () {
    return doDeletePermissionRules_(targetValue, deletions);
  });
}

function doDeletePermissionRules_(targetValue, deletions) {
  try {
    var permissionSheet = SpreadsheetApp.openById(CONFIG.DB_SYS_PART1_ID).getSheetByName('SYS_PERM');
    if (!permissionSheet || permissionSheet.getLastRow() <= 1) {
      return { success: false, message: 'Không còn quyền cấu hình nào để xóa; hãy tải lại danh sách.' };
    }
    var rows = permissionSheet.getDataRange().getValues();
    var sheets = {};
    for (var rowIndex = 1; rowIndex < rows.length; rowIndex++) {
      var row = rows[rowIndex];
      var sheetName = String(row[3] || '').trim();
      if (String(row[1] || '').trim().toUpperCase() !== 'USER' ||
          String(row[2] || '').trim().toLowerCase() !== targetValue.toLowerCase() ||
          !sheetName) continue;
      var sheetKey = sheetName.toUpperCase();
      if (!sheets[sheetKey]) {
        sheets[sheetKey] = { sheetName: sheetName, rows: [], rules: [], actions: [], clearActions: false };
      }
      var actionList = String(row[11] || '').split(',').map(function(action) {
        return action.trim().toUpperCase();
      });
      actionList.forEach(function(action) {
        if (action && sheets[sheetKey].actions.indexOf(action) === -1) sheets[sheetKey].actions.push(action);
      });
      var entry = { rowNumber: rowIndex + 1, values: row, config: null, removed: false, permId: String(row[0] || '').trim() };
      sheets[sheetKey].rows.push(entry);
      var conditions = getPermConditions_(entry.permId);
      if (conditions.length > 0) {
        entry.config = {
          conditions: conditions,
          canRead: String(row[7] || '').toUpperCase() === 'TRUE',
          canCreate: String(row[8] || '').toUpperCase() === 'TRUE',
          canEdit: String(row[9] || '').toUpperCase() === 'TRUE',
          canDelete: String(row[10] || '').toUpperCase() === 'TRUE'
        };
        sheets[sheetKey].rules.push(entry);
      }
    }

    for (var deletionIndex = 0; deletionIndex < deletions.length; deletionIndex++) {
      var deletion = deletions[deletionIndex] || {};
      var requestedSheet = String(deletion.sheetName || '').trim();
      var requestedKey = requestedSheet.toUpperCase();
      var sheetConfig = sheets[requestedKey];
      if (!requestedSheet || !sheetConfig) {
        return { success: false, message: 'Không tìm thấy quyền cần xóa trên bảng ' + (requestedSheet || '(không xác định)') + '.' };
      }
      if (deletion.actionsOnly === true) {
        if (sheetConfig.rules.length || !sheetConfig.actions.length) {
          return { success: false, message: 'Quyền chức năng của bảng ' + sheetConfig.sheetName + ' đã thay đổi; hãy tải lại danh sách.' };
        }
        sheetConfig.clearActions = true;
        continue;
      }

      var match = null;
      var deletionKey = permissionRuleDeleteKey_(deletion);
      for (var candidateIndex = 0; candidateIndex < sheetConfig.rules.length; candidateIndex++) {
        var candidate = sheetConfig.rules[candidateIndex];
        if (!candidate.removed && permissionRuleDeleteKey_(candidate.config) === deletionKey) {
          match = candidate;
          break;
        }
      }
      if (!match) {
        return { success: false, message: 'Quyền trên bảng ' + sheetConfig.sheetName + ' đã thay đổi; hãy tải lại danh sách.' };
      }
      match.removed = true;
    }

    var rowsToDelete = [];
    Object.keys(sheets).forEach(function(key) {
      sheets[key].rows.forEach(function(entry) {
        if (entry.removed) rowsToDelete.push(entry.rowNumber);
      });
    });
    // Xóa điều kiện chi tiết trong SYS_PERM_CONDITIONS trước khi xóa dòng SYS_PERM
    Object.keys(sheets).forEach(function(key) {
      sheets[key].rows.forEach(function(entry) {
        if (entry.removed && entry.permId) {
          deletePermConditions_(entry.permId);
        }
      });
    });

    var deleteRanges = [];
    rowsToDelete.sort(function(a, b) { return a - b; }).forEach(function(rowNumber) {
      var lastRange = deleteRanges[deleteRanges.length - 1];
      if (lastRange && rowNumber === lastRange.end + 1) lastRange.end = rowNumber;
      else deleteRanges.push({ start: rowNumber, end: rowNumber });
    });
    deleteRanges.reverse().forEach(function(range) {
      permissionSheet.deleteRows(range.start, range.end - range.start + 1);
    });

    Object.keys(sheets).forEach(function(key) {
      var config = sheets[key];
      if (!config.clearActions && !config.rules.some(function(entry) { return !entry.removed; })) {
        var retainedActions = config.actions.join(',');
        if (retainedActions) {
          var currentRows = permissionSheet.getDataRange().getValues();
          var remainingRows = [];
          for (var currentIndex = 1; currentIndex < currentRows.length; currentIndex++) {
            var currentRow = currentRows[currentIndex];
            if (String(currentRow[1] || '').trim().toUpperCase() === 'USER' &&
                String(currentRow[2] || '').trim().toLowerCase() === targetValue.toLowerCase() &&
                String(currentRow[3] || '').trim().toUpperCase() === key) {
              remainingRows.push(currentIndex + 1);
            }
          }
          if (remainingRows.length) {
            permissionSheet.getRange(remainingRows[0], 11).setValue(retainedActions);
          } else {
            permissionSheet.appendRow([
              'P' + new Date().getTime() + '_A', 'USER', targetValue, config.sheetName,
              'FALSE', '', 'ALL', 'FALSE', 'FALSE', 'FALSE', retainedActions
            ]);
          }
        }
      }
      if (config.clearActions) {
        var remainingData = permissionSheet.getDataRange().getValues();
        for (var clearIndex = 1; clearIndex < remainingData.length; clearIndex++) {
          var clearRow = remainingData[clearIndex];
          if (String(clearRow[1] || '').trim().toUpperCase() === 'USER' &&
              String(clearRow[2] || '').trim().toLowerCase() === targetValue.toLowerCase() &&
              String(clearRow[3] || '').trim().toUpperCase() === key &&
              String(clearRow[10] || '').trim()) {
            permissionSheet.getRange(clearIndex + 1, 11).clearContent();
          }
        }
      }
    });
    SpreadsheetApp.flush();

    var deletedSheets = {};
    (deletions || []).forEach(function(d) {
      var sn = String(d && d.sheetName || '').trim();
      if (sn) deletedSheets[sn] = true;
    });
    Object.keys(deletedSheets).forEach(function(sn) {
      invalidatePermissionCache_(targetValue, sn);
    });

    return { success: true, message: 'Đã xóa ' + deletions.length + ' quyền đã chọn.' };
  } catch (err) {
    Logger.log('deletePermissionRulesApi: ' + err);
    return { success: false, message: 'Lỗi xóa quyền ma trận: ' + String(err && err.message ? err.message : err) };
  }
}

function getPermissionRuleApi(token, targetType, targetValue, sheetName) {
  var auth = requireAuth_(token, 'SYS', 'view', 'SYS_PERM', 'view-sys-usr');
  if (!auth.ok) return auth.res;
  if (!isPermissionAdminRole_(auth.user.role)) {
    return { success: false, code: 'FORBIDDEN', message: 'Chỉ ADMIN hoặc SuperAdmin được xem cấu hình ma trận phân quyền.' };
  }
  if (String(targetType || '').trim().toUpperCase() !== 'USER' ||
      !String(targetValue || '').trim()) {
    return fail_('Chỉ hỗ trợ tra cứu quyền riêng theo tài khoản.');
  }
  targetType = 'USER';

  try {
    var sheet = SpreadsheetApp.openById(CONFIG.DB_SYS_PART1_ID).getSheetByName('SYS_PERM');
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, rule: null };
    var data = sheet.getDataRange().getValues();
    var rules = [];
    var actions = [];
    var actionsBySheet = {};
    var requestedSheetName = String(sheetName || '').trim();
    var ruleSheetNames = requestedSheetName ? getPermissionSheetNames_(requestedSheetName) : null;
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (String(row[1]).trim().toUpperCase() === String(targetType).trim().toUpperCase() &&
          String(row[2]).trim().toLowerCase() === String(targetValue).trim().toLowerCase() &&
          (!ruleSheetNames || ruleSheetNames.indexOf(String(row[3]).trim().toUpperCase()) !== -1)) {
        var configuredSheetName = String(row[3] || '').trim();
        if (!configuredSheetName) continue;
        if (!actionsBySheet[configuredSheetName]) actionsBySheet[configuredSheetName] = [];
        String(row[11] || '').split(',').forEach(function(action) {
          var normalized = action.trim().toUpperCase();
          if (normalized && actions.indexOf(normalized) === -1) actions.push(normalized);
          if (normalized && actionsBySheet[configuredSheetName].indexOf(normalized) === -1) {
            actionsBySheet[configuredSheetName].push(normalized);
          }
        });
        var permId = String(row[0] || '').trim();
        var conditions = getPermConditions_(permId);
        if (!Array.isArray(conditions) || conditions.length === 0) continue;
        var config = {
          conditions: conditions,
          canRead: String(row[7] || '').toUpperCase() === 'TRUE',
          canCreate: String(row[8] || '').toUpperCase() === 'TRUE',
          canEdit: String(row[9] || '').toUpperCase() === 'TRUE',
          canDelete: String(row[10] || '').toUpperCase() === 'TRUE'
        };
        config.sheetName = configuredSheetName;
        rules.push(config);
      }
    }
    rules.sort(function(a, b) {
      return String(a.sheetName || '').localeCompare(String(b.sheetName || ''), 'vi', { sensitivity: 'base' });
    });
    if (requestedSheetName) {
      return { success: true, rule: rules.length || actions.length ? { rules: rules, customActions: actions.join(',') } : null };
    }
    var hasConfiguredActions = Object.keys(actionsBySheet).some(function(name) {
      return actionsBySheet[name].length > 0;
    });
    return {
      success: true,
      rule: rules.length || hasConfiguredActions ? { rules: rules, customActionsBySheet: actionsBySheet } : null
    };
  } catch (err) {
    Logger.log('getPermissionRuleApi: ' + err);
    return { success: false, message: 'Không tải được cấu hình quyền ma trận: ' + String(err && err.message ? err.message : err) };
  }
}

function saveUserNavigationAccessApi(token, username, viewIds) {
  var auth = requireAuth_(token, 'SYS', 'manage_permissions', 'SYS_USR', 'view-sys-usr');
  if (!auth.ok) return auth.res;
  if (!isPermissionAdminRole_(auth.user.role)) {
    return { success: false, code: 'FORBIDDEN', message: 'Chỉ ADMIN hoặc SuperAdmin được cấu hình quyền truy cập menu.' };
  }
  if (!String(username || '').trim() || !Array.isArray(viewIds)) {
    return fail_('Dữ liệu quyền truy cập menu không hợp lệ.');
  }

  var normalizedViews = [];
  for (var i = 0; i < viewIds.length; i++) {
    var viewId = String(viewIds[i] || '').trim();
    if (!/^view-[a-z0-9_-]{1,90}$/i.test(viewId)) {
      return fail_('Mã chức năng không hợp lệ: ' + viewId);
    }
    if (normalizedViews.indexOf(viewId) === -1) normalizedViews.push(viewId);
  }

  try {
    var sheet = getUserSheet_();
    if (!sheet) return fail_('Không tìm thấy sheet SYS_USR.');
    var data = sheet.getDataRange().getValues();
    var targetRow = -1;
    for (var rowIndex = 1; rowIndex < data.length; rowIndex++) {
      if (String(data[rowIndex][AUTH_.COL.USERNAME] || '').trim().toLowerCase() ===
          String(username).trim().toLowerCase()) {
        targetRow = rowIndex + 1;
        break;
      }
    }
    if (targetRow < 0) return fail_('Không tìm thấy tài khoản cần phân quyền.');
    if (isPermissionAdminRole_(data[targetRow - 1][AUTH_.COL.ROLE])) {
      return fail_('Tài khoản ADMIN và SuperAdmin luôn được truy cập toàn bộ chức năng; không cần cấu hình quyền menu.');
    }

    var navAccess = normalizedViews.length ? normalizedViews.join(',') : 'NONE';
    var navAccessCell = sheet.getRange(targetRow, AUTH_.COL.NAV_ACCESS + 1);
    navAccessCell.setValue(navAccess);
    SpreadsheetApp.flush();
    var savedNavAccess = String(navAccessCell.getValue() || '').trim();
    if (savedNavAccess !== navAccess) {
      Logger.log('saveUserNavigationAccessApi: xác minh ghi thất bại cho ' + username +
        ', mong đợi [' + navAccess + '], nhận [' + savedNavAccess + ']');
      return { success: false, message: 'Google Sheets chưa xác nhận dữ liệu quyền menu đã được ghi. Vui lòng tải lại tài khoản và thử lại.' };
    }
    if (String(username).trim().toLowerCase() === String(auth.user.username).trim().toLowerCase()) {
      auth.user.navAccess = savedNavAccess;
      CacheService.getScriptCache().put('SES_' + token, JSON.stringify(auth.user), AUTH_.SESSION_TTL_SEC);
    }
    auditLog_(auth.user, 'SYS', 'NAV_PERMISSION_UPDATE', username, savedNavAccess);
    invalidateAllPermissionCache_(username);
    return {
      success: true,
      message: 'Đã lưu và xác minh quyền truy cập chức năng trên SYS_USR.',
      navAccess: savedNavAccess
    };
  } catch (err) {
    Logger.log('saveUserNavigationAccessApi: ' + err);
    return { success: false, message: 'Không lưu được quyền truy cập chức năng: ' + String(err && err.message ? err.message : err) };
  }
}

function cleanupLegacyRolePermissionPackages() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { success: false, message: 'Không thể khóa dữ liệu phân quyền để cleanup.' };
  }
  try {
    var permissionSheet = SpreadsheetApp.openById(CONFIG.DB_SYS_PART1_ID).getSheetByName('SYS_PERM');
    var deletedRuleCount = 0;
    if (permissionSheet && permissionSheet.getLastRow() > 1) {
      var permissionRows = permissionSheet.getDataRange().getValues();
      for (var rowIndex = permissionRows.length - 1; rowIndex >= 1; rowIndex--) {
        if (String(permissionRows[rowIndex][1] || '').trim().toUpperCase() === 'ROLE') {
          permissionSheet.deleteRow(rowIndex + 1);
          deletedRuleCount++;
        }
      }
      SpreadsheetApp.flush();
    }

    var properties = PropertiesService.getScriptProperties();
    var roleMenuKeys = Object.keys(properties.getProperties()).filter(function(key) {
      return key.indexOf('SYS_ROLE_NAV_ACCESS_V1_') === 0;
    });
    roleMenuKeys.forEach(function(key) {
      properties.deleteProperty(key);
    });

    var remainingRoleRules = 0;
    if (permissionSheet && permissionSheet.getLastRow() > 1) {
      permissionSheet.getDataRange().getValues().slice(1).forEach(function(row) {
        if (String(row[1] || '').trim().toUpperCase() === 'ROLE') remainingRoleRules++;
      });
    }
    var remainingRoleMenuKeys = Object.keys(properties.getProperties()).filter(function(key) {
      return key.indexOf('SYS_ROLE_NAV_ACCESS_V1_') === 0;
    });
    if (remainingRoleRules || remainingRoleMenuKeys.length) {
      throw new Error('Xác minh cleanup thất bại: còn ' + remainingRoleRules +
        ' rule ROLE và ' + remainingRoleMenuKeys.length + ' thuộc tính menu role.');
    }
    var result = {
      success: true,
      deletedRoleRuleCount: deletedRuleCount,
      deletedRoleMenuPropertyKeys: roleMenuKeys
    };
    Logger.log('cleanupLegacyRolePermissionPackages: ' + JSON.stringify(result));
    return result;
  } catch (err) {
    Logger.log('cleanupLegacyRolePermissionPackages: ' + err);
    return { success: false, message: 'Không cleanup được cấu hình role cũ: ' + String(err && err.message ? err.message : err) };
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 * BỔ SUNG: API QUẢN LÝ MASTER-DETAIL USER & PHÂN QUYỀN TÍCH HỢP
 * ============================================================ */

/**
 * 1. Nạp danh sách người dùng có Phân trang cho Bảng Master
 */
function fetchSysUsersPaginated(token, page, pageSize, searchKey) {
  var auth = requireAuth_(token, 'SYS', 'view', 'SYS_USR', 'view-sys-usr');
  if (!auth.ok) return Object.assign({ totalRows: 0 }, auth.res);
  var migration = migrateLegacyRoleCodes_(auth);
  if (!migration.success) return Object.assign({ totalRows: 0 }, migration);

  try {
    page = Math.max(1, parseInt(page, 10) || 1);
    pageSize = Math.min(100, Math.max(1, parseInt(pageSize, 10) || 10));

    var sheet = getUserSheet_();
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, list: [], totalRows: 0, migratedRoleCount: migration.convertedCount };

    var data = sheet.getDataRange().getValues();
    var C = AUTH_.COL;
    var rawList = [];

    var kw = String(searchKey || '').trim().toLowerCase();

    // ADMIN/SuperAdmin không cần đọc SYS_PERM (đã bypass ở matrixRowAllowed_) -> tránh đọc thừa.
    var isAdminRole = isPermissionAdminRole_(auth.user.role);
    var dataPermission = isAdminRole ? null : getEffectivePermission_(auth.user.username, auth.user.role, 'SYS_USR');

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var uName = String(row[C.USERNAME] || '').trim();
      var fName = String(row[C.FULLNAME] || '').trim();
      var cccd = String(row[C.CCCD] || '').replace(/'/g, '').trim();
      var role = normalizeRoleCode_(row[C.ROLE]) || 'EMPLOYEE';
      var mods = String(row[C.MODULES] || '').trim() || 'ALL';
      var active = isActiveValue_(row[C.ACTIVE]);

      if (uName && uName.toLowerCase() !== 'username' &&
          (isAdminRole ||
           matrixRowAllowed_(auth.user, 'SYS_USR', row, { owner: C.CREATED_BY, status: C.ACTIVE }, dataPermission))) {
        // Tìm kiếm theo Username, Fullname hoặc CCCD
        if (!kw || uName.toLowerCase().indexOf(kw) !== -1 || fName.toLowerCase().indexOf(kw) !== -1 || cccd.indexOf(kw) !== -1) {
          rawList.push({
            username: uName,
            full_name: fName,
            cccd_no: cccd,
            role: role,
            module_access: mods,
            nav_access: String(row[C.NAV_ACCESS] || '').trim(),
            is_active: active,
            status: active ? 'ACTIVE' : 'LOCKED',
            last_login: row[C.LAST_LOGIN] ? fmtDateTime_(row[C.LAST_LOGIN]) : '',
            created_by: String(row[C.CREATED_BY] || '').trim(),
            created_at: row[C.CREATED_AT] ? fmtDateTime_(row[C.CREATED_AT]) : ''
          });
        }
      }
    }

    var totalRows = rawList.length;
    var startIdx = (page - 1) * pageSize;
    var pageList = rawList.slice(startIdx, startIdx + pageSize);

    return {
      success: true,
      list: pageList,
      totalRows: totalRows,
      page: page,
      pageSize: pageSize,
      migratedRoleCount: migration.convertedCount
    };
  } catch (err) {
    return { success: false, message: 'Lỗi tải danh sách người dùng: ' + err.toString(), list: [], totalRows: 0 };
  }
}

/**
 * 2. Lấy toàn bộ chi tiết Phân quyền (SYS_MOD + SYS_PERM) của 1 User cụ thể
 */
function getUserPermissionsDetailApi(token, targetUsername) {
  var auth = requireAuth_(token, 'SYS', 'view', 'SYS_PERM', 'view-sys-usr');
  if (!auth.ok) return auth.res;
  if (!isPermissionAdminRole_(auth.user.role)) {
    return { success: false, code: 'FORBIDDEN', message: 'Chỉ ADMIN hoặc SuperAdmin được xem cấu hình ma trận phân quyền.' };
  }

  try {
    if (!targetUsername) return { success: false, message: 'Thiếu tên người dùng!' };

    // A. Lấy thông tin module_access (SYS_MOD) từ SYS_USR
    var sheetUsr = getUserSheet_();
    var dataUsr = sheetUsr.getDataRange().getValues();
    var C = AUTH_.COL;
    var moduleAccess = "ALL";
    var userRole = "EMPLOYEE";

    for (var i = 1; i < dataUsr.length; i++) {
      if (String(dataUsr[i][C.USERNAME]).trim().toLowerCase() === String(targetUsername).trim().toLowerCase()) {
        moduleAccess = String(dataUsr[i][C.MODULES] || 'ALL').trim();
        userRole = normalizeRoleCode_(dataUsr[i][C.ROLE]) || 'EMPLOYEE';
        break;
      }
    }

    // B. Lấy các quy tắc ma trận phân quyền (SYS_PERM) của User này
    var ss = SpreadsheetApp.openById(CONFIG.DB_SYS_PART1_ID);
    var sheetPerm = ss.getSheetByName('SYS_PERM');
    var matrixRules = [];

    if (sheetPerm && sheetPerm.getLastRow() > 1) {
      var permSchema = getPermSchema_();
      var dataPerm = sheetPerm.getDataRange().getValues();
      for (var j = 1; j < dataPerm.length; j++) {
        var r = dataPerm[j];
        if (String(r[permSchema['target_type']] || '').toUpperCase() === 'USER' &&
            String(r[permSchema['target_value']] || '').trim().toLowerCase() === String(targetUsername).trim().toLowerCase()) {
          var permConditions = getPermConditions_(String(r[permSchema['perm_id']] || '').trim());
          matrixRules.push({
            perm_id: r[permSchema['perm_id']],
            sheet_name: r[permSchema['sheet_name']],
            scope_owner: String(r[permSchema['scope_owner']] || '').toUpperCase() === 'TRUE',
            dept_filter: r[permSchema['dept_filter']] || 'ALL',
            status_filter: r[permSchema['status_filter']] || 'ALL',
            can_read: String(r[permSchema['can_read']] || '').toUpperCase() === 'TRUE',
            can_create: String(r[permSchema['can_create']] || '').toUpperCase() === 'TRUE',
            can_edit: String(r[permSchema['can_edit']] || '').toUpperCase() === 'TRUE',
            can_delete: String(r[permSchema['can_delete']] || '').toUpperCase() === 'TRUE',
            custom_actions: r[permSchema['custom_actions']] || '',
            conditions: permConditions
          });
        }
      }
    }

    return {
      success: true,
      username: targetUsername,
      role: userRole,
      moduleAccess: moduleAccess,
      matrixRules: matrixRules
    };

  } catch (err) {
    return { success: false, message: 'Lỗi tải thông tin phân quyền: ' + err.toString() };
  }
}

/**
 * 3. Lưu toàn bộ Phân quyền (Cập nhật module_access vào SYS_USR và Ma trận vào SYS_PERM)
 */
function saveUserFullPermissionsApi(token, payload) {
  var auth = requireAuth_(token, 'SYS', 'manage_permissions', 'SYS_PERM', 'view-sys-usr');
  if (!auth.ok) return auth.res;
  if (!isPermissionAdminRole_(auth.user.role)) {
    return { success: false, code: 'FORBIDDEN', message: 'Chỉ ADMIN hoặc SuperAdmin được thay đổi ma trận phân quyền.' };
  }

  // P5: khóa toàn hệ thống trong lúc ghi SYS_USR + SYS_PERM + SYS_PERM_CONDITIONS.
  return withLock_(function () {
    return doSaveUserFullPermissions_(payload);
  });
}

function doSaveUserFullPermissions_(payload) {
  try {
    var targetUsername = payload.username;
    var moduleAccess = payload.moduleAccess; // Ví dụ: "HRM,SYS" hoặc "ALL"
    var rules = payload.matrixRules || [];   // Mảng các quy tắc SYS_PERM

    if (!targetUsername) return { success: false, message: 'Thiếu thông tin người dùng!' };

    // A. Cập nhật module_access vào SYS_USR
    var sheetUsr = getUserSheet_();
    var dataUsr = sheetUsr.getDataRange().getValues();
    var C = AUTH_.COL;
    var userRow = -1;

    for (var i = 1; i < dataUsr.length; i++) {
      if (String(dataUsr[i][C.USERNAME]).trim().toLowerCase() === String(targetUsername).trim().toLowerCase()) {
        userRow = i + 1;
        break;
      }
    }

    if (userRow > 0) {
      sheetUsr.getRange(userRow, C.MODULES + 1).setValue(moduleAccess);
    }

    // B. Cập nhật các quy tắc ma trận vào SYS_PERM
    var ss = SpreadsheetApp.openById(CONFIG.DB_SYS_PART1_ID);
    var sheetPerm = ss.getSheetByName('SYS_PERM');
    if (!sheetPerm) {
      sheetPerm = ss.insertSheet('SYS_PERM');
      sheetPerm.appendRow(['perm_id', 'target_type', 'target_value', 'sheet_name', 'scope_owner', 'dept_filter', 'status_filter', 'can_read', 'can_create', 'can_edit', 'can_delete', 'custom_actions']);
    }

    // Lưu từng quy tắc
    rules.forEach(function(rule) {
      var dataP = sheetPerm.getDataRange().getValues();
      var pRow = -1;

      for (var k = 1; k < dataP.length; k++) {
        if (String(dataP[k][1]).toUpperCase() === 'USER' &&
            String(dataP[k][2]).trim().toLowerCase() === String(targetUsername).trim().toLowerCase() &&
            String(dataP[k][3]).trim().toLowerCase() === String(rule.sheet_name).trim().toLowerCase()) {
          pRow = k + 1;
          break;
        }
      }

      var rowVals = [
        'P' + new Date().getTime(),
        'USER',
        targetUsername,
        rule.sheet_name,
        rule.scope_owner ? 'TRUE' : 'FALSE',
        rule.dept_filter || 'ALL',
        rule.status_filter || 'ALL',
        rule.can_read ? 'TRUE' : 'FALSE',
        rule.can_create ? 'TRUE' : 'FALSE',
        rule.can_edit ? 'TRUE' : 'FALSE',
        rule.can_delete ? 'TRUE' : 'FALSE',
        rule.custom_actions || ''
      ];

      var permId;
      if (pRow > 0) {
        rowVals[0] = dataP[pRow - 1][0]; // Giữ perm_id
        permId = String(dataP[pRow - 1][0] || '');
        sheetPerm.getRange(pRow, 1, 1, rowVals.length).setValues([rowVals]);
      } else {
        sheetPerm.appendRow(rowVals);
        permId = String(rowVals[0] || '');
      }

      // Lưu điều kiện chi tiết vào SYS_PERM_CONDITIONS
      if (permId) {
        savePermConditions_(permId, rule.conditions || []);
      }
    });

    // Không có sheetName nên invalidatePermissionCache_ sẽ không xóa được gì:
    // phải xóa cache phân quyền của user trên MỌI bảng.
    invalidateAllPermissionCache_(targetUsername);

    return { success: true, message: 'Đã lưu cấu hình phân quyền thành công cho [' + targetUsername + ']!' };

  } catch (err) {
    return { success: false, message: 'Lỗi lưu phân quyền: ' + err.toString() };
  }
}

/**
 * Hàm Wrapper lưu tài khoản hỗ trợ handleApiRoute & Token
 */
/**
 * Lưu (Thêm mới / Cập nhật) thông tin tài khoản người dùng vào SYS_USR
 */
function saveSysUser(token, payload) {
  if (!payload || typeof payload !== 'object') return fail_('Dữ liệu tài khoản không hợp lệ.');
  var mode = String(payload.mode || 'CREATE').toUpperCase();
  if (mode !== 'CREATE' && mode !== 'UPDATE' && mode !== 'EDIT') {
    return fail_('Chế độ lưu tài khoản không hợp lệ.');
  }
  var auth = requireAuth_(token, 'SYS', 'manage_accounts', 'SYS_USR', 'view-sys-usr');
  if (!auth.ok) return auth.res;

  var saveLock = LockService.getScriptLock();
  if (!saveLock.tryLock(10000)) {
    return { success: false, code: 'BUSY', message: 'Hệ thống đang xử lý thay đổi tài khoản khác. Vui lòng thử lại.' };
  }

  try {
    var sheet = getUserSheet_();
    if (!sheet) return { success: false, message: 'Không tìm thấy sheet SYS_USR trong CSDL' };

    var uName = String(payload.username || '').trim();

    if (!/^[A-Za-z0-9._@-]{3,100}$/.test(uName)) {
      return { success: false, message: 'Tên đăng nhập chỉ được gồm 3-100 ký tự chữ, số, dấu chấm, gạch dưới, @ hoặc gạch ngang.' };
    }

    var C = AUTH_.COL; // Cột: USERNAME:0, PASSWORD:1, CCCD:2, FULLNAME:3, ROLE:4, MODULES:5, ACTIVE:6, LAST_LOGIN:7, CREATED_BY:8, CREATED_AT:9
    var lastUserRow = sheet.getLastRow();
    var targetRow = -1;

    if (lastUserRow > 1) {
      var usernames = sheet.getRange(2, C.USERNAME + 1, lastUserRow - 1, 1).getValues();
      for (var i = 0; i < usernames.length; i++) {
        if (String(usernames[i][0] || '').trim().toLowerCase() === uName.toLowerCase()) {
          targetRow = i + 2;
          break;
        }
      }
    }

    var requestedRole = normalizeRoleCode_(payload.role);
    if (!requestedRole) return fail_('Vai trò tài khoản là bắt buộc và phải thuộc danh sách role cho phép.');
    var normalizedRequestedRole = requestedRole.toUpperCase();
    var validRoles = ['EMPLOYEE', 'DEPARTMENT_MANAGER', 'DIRECTOR', 'ADMIN', AUTH_.SUPER_ROLE.toUpperCase()];
    if (validRoles.indexOf(normalizedRequestedRole) === -1) {
      return fail_('Vai trò tài khoản không hợp lệ.');
    }
    var statusValue = String(payload.status || '').trim().toUpperCase();
    if (!statusValue && payload.is_active === true) statusValue = 'ACTIVE';
    if (!statusValue && payload.is_active === false) statusValue = 'LOCKED';
    if (statusValue !== 'ACTIVE' && statusValue !== 'LOCKED') {
      return fail_('Trạng thái tài khoản không hợp lệ.');
    }
    if (!String(payload.cccd_no || '').trim()) return fail_('Số CCCD là bắt buộc.');
    if (!String(payload.full_name || '').trim()) return fail_('Họ và tên là bắt buộc.');

    var isSuperAdminCaller = isSuperAdminRole_(auth.user.role);
    var existingRole = targetRow > 0
      ? String(sheet.getRange(targetRow, C.ROLE + 1).getValue() || '').trim().toUpperCase()
      : '';
    var isPrivilegedRole = function (role) {
      return role === 'ADMIN' || role === AUTH_.SUPER_ROLE.toUpperCase();
    };
    if (!isSuperAdminCaller && (isPrivilegedRole(normalizedRequestedRole) || isPrivilegedRole(existingRole))) {
      return { success: false, code: 'FORBIDDEN', message: 'ADMIN chỉ được tạo hoặc sửa tài khoản USER. Chỉ SuperAdmin được quản lý tài khoản ADMIN/SuperAdmin.' };
    }
    if (targetRow > 0 && uName.toLowerCase() === String(auth.user.username).trim().toLowerCase() && !isSuperAdminCaller) {
      return { success: false, code: 'FORBIDDEN', message: 'Bạn không được tự sửa tài khoản của mình tại màn hình quản trị.' };
    }

    if (mode === 'CREATE') {
      if (targetRow > 0) {
        return { success: false, message: 'Tài khoản [' + uName + '] đã tồn tại!' };
      }
      if (!payload.password || String(payload.password).trim().length < 8) {
        return { success: false, message: 'Mật khẩu tài khoản mới phải có ít nhất 8 ký tự.' };
      }

      var passHash = makePasswordHash_(payload.password);
      var newRow = [
        uName,
        passHash,
        '',
        String(payload.full_name || '').trim(),
        requestedRole,
        String(payload.module_access || 'ALL').trim(),
        statusValue === 'ACTIVE',
        '',
        auth.user.username,
        new Date()
      ];

      sheet.appendRow(newRow);
      var createdRow = sheet.getLastRow();
      writeSysUserCccdAsText_(sheet, createdRow, payload.cccd_no);
      SpreadsheetApp.flush(); // Ép Google Sheets ghi xuống đĩa ngay lập tức
      var savedNewRow = sheet.getRange(createdRow, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (String(savedNewRow[C.USERNAME] || '').trim().toLowerCase() !== uName.toLowerCase() ||
          String(savedNewRow[C.ROLE] || '').trim() !== requestedRole ||
          String(savedNewRow[C.CCCD] || '').trim() !== String(payload.cccd_no || '').trim() ||
          String(savedNewRow[C.ACTIVE]).toUpperCase() !== String(statusValue === 'ACTIVE').toUpperCase()) {
        return { success: false, message: 'Không xác minh được tài khoản sau khi ghi. Hãy tải lại danh sách trước khi thử lại.' };
      }
      auditLog_(auth.user, 'SYS', 'CREATE_USER', uName, 'role=' + requestedRole);
      invalidateAllPermissionCache_(uName);
      return {
        success: true,
        message: 'Đã tạo thành công tài khoản [' + uName + ']!',
        account: toSysUserListItem_(savedNewRow)
      };

    } else {
      // MODE EDIT / UPDATE
      if (targetRow <= 0) {
        return { success: false, message: 'Không tìm thấy dòng của tài khoản [' + uName + ']!' };
      }
      if (payload.password && String(payload.password).trim() !== '' &&
          String(payload.password).trim().length < 8) {
        return { success: false, message: 'Mật khẩu mới phải có ít nhất 8 ký tự.' };
      }

      var updatedRow = sheet.getRange(targetRow, 1, 1, sheet.getLastColumn()).getValues()[0];
      updatedRow[C.CCCD] = String(payload.cccd_no || '').trim();
      updatedRow[C.FULLNAME] = String(payload.full_name || '').trim();
      updatedRow[C.ROLE] = requestedRole;
      updatedRow[C.MODULES] = String(payload.module_access || updatedRow[C.MODULES] || 'ALL').trim();
      updatedRow[C.ACTIVE] = statusValue === 'ACTIVE';

      if (payload.password && String(payload.password).trim() !== '') {
        updatedRow[C.PASSWORD] = makePasswordHash_(payload.password);
      }

      sheet.getRange(targetRow, C.CCCD + 1).setNumberFormat('@');
      sheet.getRange(targetRow, 1, 1, updatedRow.length).setValues([updatedRow]);
      writeSysUserCccdAsText_(sheet, targetRow, payload.cccd_no);
      SpreadsheetApp.flush(); // Ép Google Sheets ghi xuống đĩa ngay lập tức
      var savedRow = sheet.getRange(targetRow, 1, 1, updatedRow.length).getValues()[0];
      if (String(savedRow[C.USERNAME] || '').trim().toLowerCase() !== uName.toLowerCase() ||
          String(savedRow[C.ROLE] || '').trim() !== requestedRole ||
          String(savedRow[C.CCCD] || '').trim() !== String(payload.cccd_no || '').trim() ||
          String(savedRow[C.FULLNAME] || '').trim() !== String(payload.full_name || '').trim() ||
          String(savedRow[C.ACTIVE]).toUpperCase() !== String(statusValue === 'ACTIVE').toUpperCase()) {
        return { success: false, message: 'Không xác minh được thay đổi tài khoản sau khi ghi. Hãy tải lại danh sách trước khi thử lại.' };
      }
      auditLog_(auth.user, 'SYS', 'UPDATE_USER', uName, 'role=' + requestedRole);
      invalidateAllPermissionCache_(uName);
      return {
        success: true,
        message: 'Cập nhật tài khoản [' + uName + '] thành công!',
        account: toSysUserListItem_(savedRow)
      };
    }

  } catch (err) {
    return { success: false, message: 'Lỗi Backend saveSysUser: ' + err.toString() };
  } finally {
    saveLock.releaseLock();
  }
}

function updateSysUserStatus(token, username, status) {
  var auth = requireAuth_(token, 'SYS', 'manage_accounts', 'SYS_USR', 'view-sys-usr');
  if (!auth.ok) return auth.res;
  if (status !== 'ACTIVE' && status !== 'LOCKED') return fail_('Trạng thái tài khoản không hợp lệ.');
  var statusLock = LockService.getScriptLock();
  if (!statusLock.tryLock(10000)) {
    return { success: false, code: 'BUSY', message: 'Hệ thống đang xử lý thay đổi tài khoản khác. Vui lòng thử lại.' };
  }
  try {
    var sheet = getUserSheet_();
    if (!sheet) return fail_('Không tìm thấy sheet SYS_USR.');
    var C = AUTH_.COL;
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return fail_('Không tìm thấy tài khoản cần cập nhật.');
    var usernames = sheet.getRange(2, C.USERNAME + 1, lastRow - 1, 1).getValues();
    var rowNumber = -1;
    for (var i = 0; i < usernames.length; i++) {
      if (String(usernames[i][0] || '').trim().toLowerCase() === String(username || '').trim().toLowerCase()) {
        rowNumber = i + 2;
        break;
      }
    }
    if (rowNumber < 0) return fail_('Không tìm thấy tài khoản cần cập nhật.');
    var targetRole = String(sheet.getRange(rowNumber, C.ROLE + 1).getValue() || '').trim().toUpperCase();
    if (!isSuperAdminRole_(auth.user.role) && (targetRole === 'ADMIN' || targetRole === AUTH_.SUPER_ROLE.toUpperCase())) {
      return { success: false, code: 'FORBIDDEN', message: 'ADMIN chỉ được khóa/mở tài khoản USER. Chỉ SuperAdmin được quản lý tài khoản ADMIN/SuperAdmin.' };
    }
    if (String(username).trim().toLowerCase() === String(auth.user.username).trim().toLowerCase()) {
      return fail_('Không thể tự khóa hoặc mở khóa tài khoản đang đăng nhập.');
    }
    sheet.getRange(rowNumber, C.ACTIVE + 1).setValue(status === 'ACTIVE');
    SpreadsheetApp.flush();
    if (String(sheet.getRange(rowNumber, C.ACTIVE + 1).getValue()).toUpperCase() !== String(status === 'ACTIVE').toUpperCase()) {
      return fail_('Không xác minh được trạng thái tài khoản sau khi ghi. Vui lòng tải lại danh sách.');
    }
    auditLog_(auth.user, 'SYS', status === 'ACTIVE' ? 'UNLOCK_USER' : 'LOCK_USER', username, '');
    invalidateAllPermissionCache_(username);
    return { success: true, message: status === 'ACTIVE' ? 'Đã mở khóa tài khoản.' : 'Đã khóa tài khoản.' };
  } catch (err) {
    Logger.log('updateSysUserStatus: ' + err);
    return fail_('Không thể cập nhật trạng thái tài khoản.');
  } finally {
    statusLock.releaseLock();
  }
}

/** Helper định dạng ngày tháng hiển thị */
function fmtDateTime_(d) {
  if (!d) return '';
  if (typeof d === 'string') return d;
  try {
    return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  } catch (e) {
    return String(d);
  }
}