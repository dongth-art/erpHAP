/**
 * HRM_EMP_Server.gs - HAP HOLDING ERP SYSTEM (Hồ sơ Nhân viên - 34 cột)
 * ĐỢT 2a: mọi hàm công khai nhận "token" ở tham số đầu tiên và kiểm tra quyền qua requireAuth_.
 * ĐỢT 2b: xóa mềm (Thùng rác), chặn trùng mã NV / CCCD, ghi nhật ký thao tác.
 * Đọc/ghi theo HEADER (tên cột), không phụ thuộc thứ tự cột.
 */

/* ============================================================
 * CẤU TRÚC 34 CỘT (đúng thứ tự trong sheet HRM_EMP)
 * ============================================================ */

function empCols_() {
  return [
    'cccd_no', 'emp_id', 'full_name', 'dob', 'gender', 'curr_addr',
    'avatar_url', 'phone', 'email', 'company_code', 'unit_code', 'dept_code',
    'job_title', 'work_location', 'manager_id', 'join_date', 'contract_type',
    'contract_no', 'contract_start_date', 'contract_end_date', 'contract_status', 'work_status',
    'resign_date', 'salary_grade', 'base_salary', 'social_insurance_no', 'tax_code',
    'bank_account_no', 'bank_name', 'bank_account_holder', 'academic_level',
    'created_by', 'created_at', 'updated_at'
  ];
}

/** Cột P... của HRM_EMP: sau 34 cột dữ liệu -> cột 35 = is_deleted, 36 = deleted_by, 37 = deleted_at */
function empFlagCol_() { return 35; }

/** Các cột luôn lưu dạng CHỮ (mã định danh/mã danh mục). Thiếu cột nào ở đây, Sheets sẽ tự đổi giá trị toàn số thành number. */
function empTextCols_() {
  return ['cccd_no', 'emp_id', 'phone', 'manager_id', 'company_code', 'unit_code', 'dept_code',
          'contract_no', 'social_insurance_no', 'tax_code', 'bank_account_no'];
}

function empSet_(arr) {
  var o = {};
  arr.forEach(function (k) { o[k] = true; });
  return o;
}

function getHRMEMPHeaderMap_() {
  try {
    var sheet = getEMPSheet_(false);
    if (!sheet) return {};
    if (typeof getHeaderMap_ === 'function') {
      return getHeaderMap_(sheet);
    }
    if (typeof getSheetHeaders_ === 'function') {
      var headers = getSheetHeaders_(sheet);
      var map = {};
      for (var i = 0; i < headers.length; i++) {
        var h = String(headers[i] || '').trim();
        if (h) map[h] = i;
      }
      return map;
    }
    return {};
  } catch (e) {
    return {};
  }
}

function hreColIdx_(name) {
  try {
    var map = getHRMEMPHeaderMap_();
    if (map && map[name] !== undefined && map[name] !== null) {
      return map[name];
    }
    var cols = empCols_();
    var idx = cols.indexOf(name);
    if (idx >= 0) return idx;
  } catch (e) {}
  return -1;
}

function HRMEMPColumns_() { return empCols_(); }

function getEMPConfig_() {
  var c = (typeof CONFIG !== 'undefined') ? CONFIG : {};
  return {
    dbId: c.DB_HRM_PART1_ID || '',
    sheetEmp: c.SHEET_HRM_EMP || 'HRM_EMP'
  };
}

function getEMPSheet_(createIfMissing) {
  var cfg = getEMPConfig_();
  if (!cfg.dbId) throw new Error('Chưa cấu hình DB_HRM_PART1_ID');
  var ss = SpreadsheetApp.openById(cfg.dbId);
  var sheet = ss.getSheetByName(cfg.sheetEmp);
  if (!sheet && createIfMissing) {
    sheet = ss.insertSheet(cfg.sheetEmp);
    sheet.appendRow(empCols_());
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** Dòng sheet (mảng 34 ô) -> đối tượng */
function empRowToObj_(row) {
  var map = getHRMEMPHeaderMap_();
  var cols = empCols_();
  var dateKeys = empSet_(['dob', 'join_date', 'contract_start_date', 'contract_end_date', 'resign_date']);
  var idKeys = empSet_(empTextCols_());
  var dtKeys = empSet_(['created_at', 'updated_at']);
  var o = {};
  for (var i = 0; i < cols.length; i++) {
    var k = cols[i];
    var idx = map[k] !== undefined ? map[k] : i;
    var v = row[idx];
    if (idKeys[k]) o[k] = cleanId_(v);
    else if (dateKeys[k]) o[k] = fmtDate_(v);
    else if (dtKeys[k]) o[k] = fmtDateTime_(v);
    else if (k === 'base_salary') o[k] = v || 0;
    else o[k] = v || '';
  }
  return o;
}

/** Ẩn các trường nhạy cảm (lương, thuế, BHXH, ngân hàng) nếu người dùng không có quyền */
function empMaskSensitive_(o) {
  ['salary_grade', 'base_salary', 'social_insurance_no', 'tax_code',
   'bank_account_no', 'bank_name', 'bank_account_holder'].forEach(function (k) {
    o[k] = (k === 'base_salary') ? 0 : '';
  });
  return o;
}

/* ============================================================
 * TRA CỨU NỘI BỘ (dùng cho Chấm công, kiểm tra trùng...)
 * ============================================================ */

/** Đọc toàn bộ cột + cờ xóa của toàn bảng trong 1 lần */
function empScan_(sheet) {
  var out = { n: 0, rows: [], flags: [] };
  if (!sheet || sheet.getLastRow() <= 1) return out;
  out.n = sheet.getLastRow() - 1;
  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  out.rows = sheet.getRange(2, 1, out.n, lastColumn).getValues();
  out.flags = readDeletedFlags_(sheet, empFlagCol_(), 2, out.n);
  return out;
}

/** Tìm nhân viên ĐANG HOẠT ĐỘNG theo cột (0 = cccd_no, 1 = emp_id) */
function empFindByColumn_(colIndex0, value) {
  var sheet = getEMPSheet_(false);
  if (!sheet || sheet.getLastRow() <= 1) return null;
  var target = cleanId_(value).toLowerCase();
  if (!target) return null;

  var n = sheet.getLastRow() - 1;
  var col = sheet.getRange(2, colIndex0 + 1, n, 1).getValues();
  var flags = readDeletedFlags_(sheet, empFlagCol_(), 2, n);
  for (var i = 0; i < col.length; i++) {
    if (!flags[i] && cleanId_(col[i][0]).toLowerCase() === target) {
      var lastColumn = Math.max(sheet.getLastColumn(), 1);
      return empRowToObj_(sheet.getRange(i + 2, 1, 1, lastColumn).getValues()[0]);
    }
  }
  return null;
}

function empFindByEmpId_(empId) { return empFindByColumn_(1, empId); }
function empFindByCccd_(cccd) { return empFindByColumn_(0, cccd); }

/** Bản đồ mã nhân viên (chữ thường) -> họ tên. Dùng để hiển thị tên người duyệt. */
function empNameMap_() {
  var map = {};
  var sheet = getEMPSheet_(false);
  if (!sheet || sheet.getLastRow() <= 1) return map;
  var empIdIdx = hreColIdx_('emp_id');
  var fullNameIdx = hreColIdx_('full_name');
  var n = sheet.getLastRow() - 1;
  var lastColumn = Math.max(sheet.getLastColumn(), empIdIdx + 1, fullNameIdx + 1);
  var d = sheet.getRange(2, 1, n, lastColumn).getValues();
  d.forEach(function (r) {
    var k = cleanId_(r[empIdIdx]);
    if (k) map[k.toLowerCase()] = String(r[fullNameIdx] || '');
  });
  return map;
}

/**
 * Chuỗi quản lý của một nhân viên (theo cột manager_id).
 * Cố ý tính CẢ nhân viên đã xóa mềm để chuỗi duyệt không bị đứt (đơn không tự động được duyệt khi quản lý bị xóa).
 */
function getApprovalHierarchyForEmployee_(empId) {
  try {
    var sheet = getEMPSheet_(false);
    if (!sheet || sheet.getLastRow() <= 1) return { manager_id: '', hierarchy: [] };

    var empIdIdx = hreColIdx_('emp_id');
    var managerIdx = hreColIdx_('manager_id');
    var n = sheet.getLastRow() - 1;
    var lastColumn = Math.max(sheet.getLastColumn(), Math.max(empIdIdx, managerIdx) + 1);
    var data = sheet.getRange(2, 1, n, lastColumn).getValues();
    var dict = {};
    for (var i = 0; i < data.length; i++) {
      var eId = cleanId_(data[i][empIdIdx]).toLowerCase();
      if (eId) dict[eId] = cleanId_(data[i][managerIdx]);
    }

    var direct = dict[cleanId_(empId).toLowerCase()] || '';
    var hierarchy = [], visited = {}, curr = direct;
    while (curr && !visited[curr.toLowerCase()]) {
      visited[curr.toLowerCase()] = true;
      hierarchy.push(curr);
      curr = dict[curr.toLowerCase()] || '';
    }
    return { manager_id: direct, hierarchy: hierarchy };
  } catch (err) {
    Logger.log('getApprovalHierarchyForEmployee_: ' + err);
    return { manager_id: '', hierarchy: [] };
  }
}

/* ============================================================
 * ĐỌC DỮ LIỆU (chỉ nhân viên đang hoạt động)
 * ============================================================ */

function fetchHRMEMPListPaginated(token, page, pageSize) {
  var auth = requireAuth_(token, 'HRM', 'view', 'HRM_EMP', 'view-hrm-emp,view-hrm-org,view-hrm-att');
  if (!auth.ok) return Object.assign({ totalRows: 0 }, auth.res);
  var dataPermission = getEffectivePermission_(auth.user.username, auth.user.role, 'HRM_EMP');
  var permissions = {
    create: hasPermission_(auth.user, 'HRM', 'create', 'HRM_EMP', undefined, dataPermission),
    update: hasPermission_(auth.user, 'HRM', 'update', 'HRM_EMP', undefined, dataPermission),
    delete: hasPermission_(auth.user, 'HRM', 'delete', 'HRM_EMP', undefined, dataPermission),
    export: hasPermission_(auth.user, 'HRM', 'export', 'HRM_EMP', undefined, dataPermission),
    import: hasPermission_(auth.user, 'HRM', 'import', 'HRM_EMP', undefined, dataPermission),
    restore: hasPermission_(auth.user, 'HRM', 'restore', 'HRM_EMP', undefined, dataPermission),
    viewSensitive: hasPermission_(auth.user, 'HRM', 'view_sensitive', 'HRM_EMP', undefined, dataPermission)
  };

  try {
    page = Math.max(1, parseInt(page, 10) || 1);
    pageSize = Math.min(3000, Math.max(1, parseInt(pageSize, 10) || 20));

    var sheet = getEMPSheet_(false);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, list: [], totalRows: 0, permissions: permissions };

    var n = sheet.getLastRow() - 1;
    var flags = readDeletedFlags_(sheet, empFlagCol_(), 2, n);
    var lastColumn = Math.max(sheet.getLastColumn(), 1);
    var data = sheet.getRange(2, 1, n, lastColumn).getValues();
    var empMap = getHRMEMPHeaderMap_();
    var empIdIdx = (empMap['emp_id'] !== undefined) ? empMap['emp_id'] : 1;
    var ownerIdx = (empMap['created_by'] !== undefined) ? empMap['created_by'] : 31;
    var deptIdx = (empMap['dept_code'] !== undefined) ? empMap['dept_code'] : 11;
    var statusIdx = (empMap['work_status'] !== undefined) ? empMap['work_status'] : 21;
    var active = [];
    for (var i = 0; i < n; i++) {
      if (!flags[i] && cleanId_(data[i][empIdIdx]) &&
          matrixRowAllowed_(auth.user, 'HRM_EMP', data[i], { owner: ownerIdx, dept: deptIdx, status: statusIdx }, dataPermission)) active.push(i + 2);
    }

    var totalRows = active.length;
    var pageRows = active.slice((page - 1) * pageSize, page * pageSize);
    if (pageRows.length === 0) return { success: true, list: [], totalRows: totalRows, permissions: permissions };

    var minR = pageRows[0];
    var showSensitive = permissions.viewSensitive;

    var empMap = getHRMEMPHeaderMap_();
    var empIdIdx = (empMap['emp_id'] !== undefined) ? empMap['emp_id'] : 1;
    var list = [];
    pageRows.forEach(function (r) {
      var row = data[r - 2];
      if (cleanId_(row[empIdIdx]) !== '') {
        var o = empRowToObj_(row);
        list.push(showSensitive ? o : empMaskSensitive_(o));
      }
    });
    return { success: true, list: list, totalRows: totalRows, permissions: permissions };
  } catch (err) {
    Logger.log('fetchHRMEMPListPaginated: ' + err);
    return { success: false, message: 'Lỗi tải dữ liệu nhân viên', list: [], totalRows: 0 };
  }
}

function fetchAllHRMEMPRecords(token) {
  var auth = requireAuth_(token, 'HRM', 'export', 'HRM_EMP', 'view-hrm-emp');
  if (!auth.ok) return auth.res;

  try {
    var dataPermission = getEffectivePermission_(auth.user.username, auth.user.role, 'HRM_EMP');
    var sheet = getEMPSheet_(false);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, list: [] };

    var n = sheet.getLastRow() - 1;
    var lastColumn = Math.max(sheet.getLastColumn(), 1);
    var data = sheet.getRange(2, 1, n, lastColumn).getValues();
    var flags = readDeletedFlags_(sheet, empFlagCol_(), 2, n);
    var showSensitive = hasPermission_(auth.user, 'HRM', 'view_sensitive', 'HRM_EMP', undefined, dataPermission);
    var empMap = getHRMEMPHeaderMap_();
    var empIdIdx = (empMap['emp_id'] !== undefined) ? empMap['emp_id'] : 1;
    var ownerIdx = (empMap['created_by'] !== undefined) ? empMap['created_by'] : 31;
    var deptIdx = (empMap['dept_code'] !== undefined) ? empMap['dept_code'] : 11;
    var statusIdx = (empMap['work_status'] !== undefined) ? empMap['work_status'] : 21;
    var list = [];
    for (var i = 0; i < data.length; i++) {
      if (!flags[i] && cleanId_(data[i][empIdIdx]) !== '' &&
          matrixRowAllowed_(auth.user, 'HRM_EMP', data[i], { owner: ownerIdx, dept: deptIdx, status: statusIdx }, dataPermission)) {
        var o = empRowToObj_(data[i]);
        list.push(showSensitive ? o : empMaskSensitive_(o));
      }
    }
    auditLog_(auth.user, 'HRM_EMP', 'EXPORT', '', 'Xuất ' + list.length + ' hồ sơ nhân viên');
    return { success: true, list: list };
  } catch (err) {
    Logger.log('fetchAllHRMEMPRecords: ' + err);
    return { success: false, message: 'Lỗi tải dữ liệu xuất Excel', list: [] };
  }
}

/* ============================================================
 * GHI DỮ LIỆU
 * ============================================================ */

function empTextVal_(val) {
  var s = cleanId_(val);
  return s ? "'" + s : '';
}

/** Đối tượng nhập liệu -> mảng 34 ô */
function empObjToRow_(r, createdBy, createdAt, updatedAt) {
  var cccdIdx = hreColIdx_('cccd_no');
  var empIdIdx = hreColIdx_('emp_id');
  var fullNameIdx = hreColIdx_('full_name');
  var dobIdx = hreColIdx_('dob');
  var genderIdx = hreColIdx_('gender');
  var currAddrIdx = hreColIdx_('curr_addr');
  var avatarIdx = hreColIdx_('avatar_url');
  var phoneIdx = hreColIdx_('phone');
  var emailIdx = hreColIdx_('email');
  var companyIdx = hreColIdx_('company_code');
  var unitIdx = hreColIdx_('unit_code');
  var deptIdx = hreColIdx_('dept_code');
  var jobIdx = hreColIdx_('job_title');
  var workLocIdx = hreColIdx_('work_location');
  var managerIdx = hreColIdx_('manager_id');
  var joinDateIdx = hreColIdx_('join_date');
  var contractTypeIdx = hreColIdx_('contract_type');
  var contractNoIdx = hreColIdx_('contract_no');
  var contractStartIdx = hreColIdx_('contract_start_date');
  var contractEndIdx = hreColIdx_('contract_end_date');
  var contractStatusIdx = hreColIdx_('contract_status');
  var workStatusIdx = hreColIdx_('work_status');
  var resignDateIdx = hreColIdx_('resign_date');
  var salaryGradeIdx = hreColIdx_('salary_grade');
  var baseSalaryIdx = hreColIdx_('base_salary');
  var siIdx = hreColIdx_('social_insurance_no');
  var taxIdx = hreColIdx_('tax_code');
  var bankAccIdx = hreColIdx_('bank_account_no');
  var bankNameIdx = hreColIdx_('bank_name');
  var bankHolderIdx = hreColIdx_('bank_account_holder');
  var acadIdx = hreColIdx_('academic_level');
  var createdByIdx = hreColIdx_('created_by');
  var createdAtIdx = hreColIdx_('created_at');
  var updatedAtIdx = hreColIdx_('updated_at');
  var textIds = empSet_(empTextCols_());
  var pass = empSet_(['dob', 'join_date', 'contract_start_date', 'contract_end_date', 'resign_date']);
  var defaults = { gender: 'Nam', contract_status: 'Hiệu lực', work_status: 'Đang làm việc' };
  var cols = empCols_();
  var row = [];

  for (var i = 0; i < 31; i++) {
    var k = cols[i], v = r[k];
    if (textIds[k]) row.push(empTextVal_(v));
    else if (pass[k]) row.push(v || '');
    else if (k === 'base_salary') row.push(parseFloat(v) || 0);
    else row.push(safeText_(v) || defaults[k] || '');
  }
  var map = getHRMEMPHeaderMap_();
  var createdByIdx = (map['created_by'] !== undefined) ? map['created_by'] : 31;
  var createdAtIdx = (map['created_at'] !== undefined) ? map['created_at'] : 32;
  var updatedAtIdx = (map['updated_at'] !== undefined) ? map['updated_at'] : 33;
  var len = Math.max(row.length, createdByIdx + 1, createdAtIdx + 1, updatedAtIdx + 1);
  while (row.length < len) row.push('');
  row[createdByIdx] = createdBy;
  row[createdAtIdx] = createdAt;
  row[updatedAtIdx] = updatedAt;
  return row;
}

function saveHRMEMPRecord(token, record) {
  var auth = requireSessionAuth_(token, 'HRM', 'view-hrm-emp');
  if (!auth.ok) return auth.res;

  try {
    if (!record || !cleanId_(record.emp_id)) return fail_('Thiếu Mã nhân viên.');
    if (!record.full_name || !String(record.full_name).trim()) return fail_('Thiếu Họ và Tên.');

    return withLock_(function () {
      var sheet = getEMPSheet_(true);
      var empId = cleanId_(record.emp_id);
      var cccd = cleanId_(record.cccd_no);
      var now = new Date();
      var username = auth.user.username;

      var scan = empScan_(sheet);
      var empMap2 = getHRMEMPHeaderMap_();
      var empIdIdx2 = (empMap2['emp_id'] !== undefined) ? empMap2['emp_id'] : 1;
      var cccdIdx2 = (empMap2['cccd_no'] !== undefined) ? empMap2['cccd_no'] : 0;
      var rowIdx = -1, rowDeleted = false, cccdOwner = '';
      for (var i = 0; i < scan.n; i++) {
        var rid = cleanId_(scan.rows[i][empIdIdx2]);
        if (rid === empId) {
          if (!scan.flags[i]) { rowIdx = i + 2; rowDeleted = false; break; }
          if (rowIdx < 0) { rowIdx = i + 2; rowDeleted = true; }
        } else if (cccd && !scan.flags[i] && cleanId_(scan.rows[i][cccdIdx2]) === cccd) {
          cccdOwner = rid;   // số CCCD này đã thuộc nhân viên khác đang hoạt động
        }
      }
      if (cccdOwner) return fail_('Số CCCD ' + cccd + ' đã thuộc nhân viên có mã ' + cccdOwner + '. Mỗi người chỉ có một hồ sơ nhân viên.');

      if (rowIdx > 0) {
        var lastColumn = Math.max(sheet.getLastColumn(), 1);
        var old = sheet.getRange(rowIdx, 1, 1, lastColumn).getValues()[0];
        var empMap = getHRMEMPHeaderMap_();
        var createdByIdx = (empMap['created_by'] !== undefined) ? empMap['created_by'] : 31;
        var createdAtIdx = (empMap['created_at'] !== undefined) ? empMap['created_at'] : 32;
        var ownerIdx = (empMap['created_by'] !== undefined) ? empMap['created_by'] : 31;
        var deptIdx = (empMap['dept_code'] !== undefined) ? empMap['dept_code'] : 11;
        var statusIdx = (empMap['work_status'] !== undefined) ? empMap['work_status'] : 21;
        if (!hasPermission_(auth.user, 'HRM', 'update', 'HRM_EMP') ||
            !matrixScopeAllowed_(auth.user, 'HRM_EMP', old, { owner: ownerIdx, dept: deptIdx, status: statusIdx }, null, 'update')) {
          return fail_('Bạn không có quyền sửa hồ sơ ngoài phạm vi được phân quyền.', { code: 'FORBIDDEN' });
        }
        if (rowDeleted && !hasPermission_(auth.user, 'HRM', 'restore', 'HRM_EMP')) {
          return fail_('Mã nhân viên này đang nằm trong Thùng rác. Chỉ SuperAdmin được khôi phục.', { code: 'FORBIDDEN' });
        }
        var row = empObjToRow_(record, old[createdByIdx] || username, old[createdAtIdx] || now, now);
        if (!matrixScopeAllowed_(auth.user, 'HRM_EMP', row, { owner: ownerIdx, dept: deptIdx, status: statusIdx }, null, 'update')) {
          return fail_('Thông tin mới nằm ngoài phạm vi được phân quyền.', { code: 'FORBIDDEN' });
        }
        var changed = diffFields_(old, row, empCols_().slice(0, 31));
        sheet.getRange(rowIdx, 1, 1, lastColumn).setValues([row]);
        if (rowDeleted) restoreRows_(sheet, [rowIdx], empFlagCol_());
        auditLog_(auth.user, 'HRM_EMP', rowDeleted ? 'RESTORE_UPDATE' : 'UPDATE', empId, changed ? 'Đổi: ' + changed : 'Không đổi nội dung');
        return { success: true, message: 'Cập nhật hồ sơ nhân viên thành công!' };
      }

      var empMap = getHRMEMPHeaderMap_();
      var ownerIdx = (empMap['created_by'] !== undefined) ? empMap['created_by'] : 31;
      var deptIdx = (empMap['dept_code'] !== undefined) ? empMap['dept_code'] : 11;
      var statusIdx = (empMap['work_status'] !== undefined) ? empMap['work_status'] : 21;
      var newRow = empObjToRow_(record, username, now, now);
      if (!hasPermission_(auth.user, 'HRM', 'create', 'HRM_EMP') ||
          !matrixScopeAllowed_(auth.user, 'HRM_EMP', newRow, { owner: ownerIdx, dept: deptIdx, status: statusIdx }, null, 'create')) {
        return fail_('Thông tin mới nằm ngoài phạm vi được phân quyền.', { code: 'FORBIDDEN' });
      }
      var lastColumn = Math.max(sheet.getLastColumn(), 1);
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, lastColumn).setValues([newRow]);
      auditLog_(auth.user, 'HRM_EMP', 'CREATE', empId, String(record.full_name).trim());
      return { success: true, message: 'Thêm mới nhân viên thành công!' };
    });
  } catch (err) {
    Logger.log('saveHRMEMPRecord: ' + err);
    return fail_('Không thể lưu hồ sơ nhân viên.');
  }
}

function saveBulkHRMEMPRecords(token, recordsList) {
  var auth = requireAuth_(token, 'HRM', 'import', 'HRM_EMP', 'view-hrm-emp');
  if (!auth.ok) return auth.res;

  try {
    if (!Array.isArray(recordsList) || recordsList.length === 0) return fail_('Không có dữ liệu hợp lệ để nhập.');
    if (recordsList.length > 2000) return fail_('Mỗi lần chỉ nhập tối đa 2000 dòng.');

    return withLock_(function () {
      var sheet = getEMPSheet_(true);
      var now = new Date();

      // Mã NV đã có (cả trong Thùng rác, để không dùng lại) và CCCD của nhân viên đang hoạt động
      var usedIds = {}, activeCccd = {};
      var scan = empScan_(sheet);
      var empMap2 = getHRMEMPHeaderMap_();
      var empIdIdx2 = (empMap2['emp_id'] !== undefined) ? empMap2['emp_id'] : 1;
      var cccdIdx2 = (empMap2['cccd_no'] !== undefined) ? empMap2['cccd_no'] : 0;
      for (var i = 0; i < scan.n; i++) {
        var rid = cleanId_(scan.rows[i][empIdIdx2]);
        if (rid) usedIds[rid.toLowerCase()] = true;
        var rc = cleanId_(scan.rows[i][cccdIdx2]);
        if (rc && !scan.flags[i]) activeCccd[rc] = rid;
      }

      var rows = [], dupId = [], dupCccd = [], outOfScope = [], invalid = 0;
      var empMap = getHRMEMPHeaderMap_();
      var ownerIdx = (empMap['created_by'] !== undefined) ? empMap['created_by'] : 31;
      var deptIdx = (empMap['dept_code'] !== undefined) ? empMap['dept_code'] : 11;
      var statusIdx = (empMap['work_status'] !== undefined) ? empMap['work_status'] : 21;
      recordsList.forEach(function (r) {
        var id = r ? cleanId_(r.emp_id) : '';
        if (!id || !r.full_name || !String(r.full_name).trim()) { invalid++; return; }
        var c = cleanId_(r.cccd_no);
        if (usedIds[id.toLowerCase()]) { dupId.push(id); return; }
        if (c && activeCccd[c]) { dupCccd.push(c); return; }
        var newRow = empObjToRow_(r, auth.user.username, now, now);
        if (!matrixScopeAllowed_(auth.user, 'HRM_EMP', newRow, { owner: ownerIdx, dept: deptIdx, status: statusIdx }, null, 'create')) {
          outOfScope.push(id);
          return;
        }
        usedIds[id.toLowerCase()] = true;
        if (c) activeCccd[c] = id;
        rows.push(newRow);
      });

      var parts = [];
      if (dupId.length) parts.push(dupId.length + ' trùng mã NV: ' + sampleList_(dupId));
      if (dupCccd.length) parts.push(dupCccd.length + ' trùng CCCD: ' + sampleList_(dupCccd));
      if (outOfScope.length) parts.push(outOfScope.length + ' ngoài phạm vi phân quyền: ' + sampleList_(outOfScope));
      if (invalid) parts.push(invalid + ' thiếu mã NV/họ tên');
      var skippedMsg = parts.length ? ' Bỏ qua ' + (dupId.length + dupCccd.length + outOfScope.length + invalid) + ' dòng (' + parts.join('; ') + ').' : '';

      if (rows.length === 0) return fail_('Không có nhân viên mới để nhập.' + skippedMsg);

      var lastColumn = Math.max(sheet.getLastColumn(), 1);
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, lastColumn).setValues(rows);
      var empIdIdx = hreColIdx_('emp_id');
      auditLog_(auth.user, 'HRM_EMP', 'IMPORT', sampleList_(rows.map(function (x) { return cleanId_(x[empIdIdx]); }), 5), 'Nhập ' + rows.length + ' nhân viên, bỏ qua ' + (recordsList.length - rows.length));
      return { success: true, message: 'Đã nhập thành công ' + rows.length + ' nhân viên.' + skippedMsg };
    });
  } catch (err) {
    Logger.log('saveBulkHRMEMPRecords: ' + err);
    return fail_('Không thể nhập dữ liệu từ Excel.');
  }
}

/* ============================================================
 * XÓA MỀM
 * ============================================================ */

function deleteBulkHRMEMPRecords(token, empIdList) {
  var auth = requireAuth_(token, 'HRM', 'delete', 'HRM_EMP', 'view-hrm-emp');
  if (!auth.ok) return auth.res;

  try {
    if (!Array.isArray(empIdList) || empIdList.length === 0) return fail_('Danh sách mã nhân viên cần xóa trống.');

    return withLock_(function () {
      var sheet = getEMPSheet_(false);
      if (!sheet || sheet.getLastRow() <= 1) return fail_('Bảng dữ liệu đang trống.');

      var targets = {};
      empIdList.forEach(function (x) { var k = cleanId_(x); if (k) targets[k.toLowerCase()] = true; });

      var scan = empScan_(sheet);
      var dataPermission = getEffectivePermission_(auth.user.username, auth.user.role, 'HRM_EMP');
      var empMap = getHRMEMPHeaderMap_();
      var empIdIdx = (empMap['emp_id'] !== undefined) ? empMap['emp_id'] : 1;
      var managerIdx = (empMap['manager_id'] !== undefined) ? empMap['manager_id'] : 14;
      var ownerIdx = (empMap['created_by'] !== undefined) ? empMap['created_by'] : 31;
      var deptIdx = (empMap['dept_code'] !== undefined) ? empMap['dept_code'] : 11;
      var statusIdx = (empMap['work_status'] !== undefined) ? empMap['work_status'] : 21;
      var rowsToDelete = [], ids = [];
      for (var i = 0; i < scan.n; i++) {
        var rid = cleanId_(scan.rows[i][empIdIdx]);
        if (!scan.flags[i] && rid && targets[rid.toLowerCase()] &&
            matrixScopeAllowed_(auth.user, 'HRM_EMP', scan.rows[i], { owner: ownerIdx, dept: deptIdx, status: statusIdx }, dataPermission, 'delete')) {
          rowsToDelete.push(i + 2); ids.push(rid);
        }
      }
      if (rowsToDelete.length === 0) return fail_('Không tìm thấy nhân viên phù hợp để xóa.');

      // Cảnh báo: nhân viên bị xóa đang là quản lý của người khác (chuỗi duyệt phép)
      var idSet = {}; ids.forEach(function (x) { idSet[x.toLowerCase()] = true; });
      var orphaned = 0;
      var empIdIdx2 = (empMap['emp_id'] !== undefined) ? empMap['emp_id'] : 1;
      var managerIdx2 = (empMap['manager_id'] !== undefined) ? empMap['manager_id'] : 14;
      for (var j = 0; j < scan.n; j++) {
        if (scan.flags[j]) continue;
        var rid2 = cleanId_(scan.rows[j][empIdIdx2]).toLowerCase();
        var mgr = cleanId_(scan.rows[j][managerIdx2]).toLowerCase();
        if (mgr && idSet[mgr] && !idSet[rid2]) orphaned++;
      }

      var cnt = softDeleteRows_(sheet, rowsToDelete, empFlagCol_(), auth.user.username);
      auditLog_(auth.user, 'HRM_EMP', 'DELETE', sampleList_(ids, 20), 'Chuyển ' + cnt + ' nhân viên vào Thùng rác' + (orphaned ? '; ' + orphaned + ' NV đang có người bị xóa làm quản lý' : ''));
      return {
        success: true,
        message: 'Đã chuyển ' + cnt + ' hồ sơ nhân viên vào Thùng rác (SuperAdmin có thể khôi phục).' +
          (orphaned ? ' Lưu ý: ' + orphaned + ' nhân viên đang có người này làm quản lý, hãy cập nhật lại Mã Quản Lý cho họ.' : '')
      };
    });
  } catch (err) {
    Logger.log('deleteBulkHRMEMPRecords: ' + err);
    return fail_('Không thể xóa dữ liệu.');
  }
}

function deleteHRMEMPRecord(token, empId) {
  return deleteBulkHRMEMPRecords(token, [empId]);
}

/**
 * CHẠY MỘT LẦN (chủ dự án, trong trình soạn thảo): chuyển các cột mã (manager_id, dept_code...) của dữ liệu đã có
 * từ dạng số về dạng CHỮ. Lưu ý: số 0 đứng đầu đã mất từ trước (nếu có) thì không khôi phục được.
 */
function normalizeEmpTextColumns() {
  if (!isOwnerRun_()) return { success: false, message: 'Không được phép.' };

  return withLock_(function () {
    var sheet = getEMPSheet_(false);
    if (!sheet || sheet.getLastRow() <= 1) { Logger.log('Bảng nhân viên trống.'); return { success: true, changed: 0 }; }

    var cols = empCols_(), n = sheet.getLastRow() - 1, changed = 0, detail = [];
    empTextCols_().forEach(function (key) {
      var c = cols.indexOf(key) + 1;
      var rng = sheet.getRange(2, c, n, 1);
      var vals = rng.getValues(), cnt = 0;
      for (var i = 0; i < vals.length; i++) {
        var v = vals[i][0];
        if (typeof v === 'number') cnt++;
        vals[i][0] = (v === '' || v == null) ? '' : cleanId_(v);
      }
      rng.setNumberFormat('@');   // định dạng cột là Văn bản
      rng.setValues(vals);
      if (cnt) detail.push(key + ': ' + cnt);
      changed += cnt;
    });
    Logger.log('Đã chuyển ' + changed + ' ô từ số sang chữ. ' + (detail.length ? detail.join('; ') : ''));
    return { success: true, changed: changed };
  });
}