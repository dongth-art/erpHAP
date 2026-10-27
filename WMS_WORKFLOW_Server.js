var WMS_WORKFLOW_HEADERS_ = {
  WF:['workflow_version_id', 'workflow_id', 'wf_code', 'wf_name', 'doc_type',
    'warehouse_scope_id', 'version_no', 'status', 'is_active',
    'created_by_username', 'created_by_emp_id', 'created_at',
    'updated_by_username', 'updated_at', 'activated_by_username',
    'activated_at', 'change_reason'],
  STEP:['workflow_step_id', 'workflow_version_id', 'step_no', 'step_name',
    'approver_emp_id', 'approver_username', 'approver_name_snapshot',
    'can_self_approve', 'is_stock_posting_step', 'created_at'],
  RUN:['workflow_run_id', 'doc_id', 'workflow_version_id', 'run_no',
    'run_status', 'current_step_no', 'started_by_username',
    'started_by_emp_id', 'started_at', 'completed_at', 'previous_run_id',
    'return_reason', 'created_at'],
  TASK:['task_id', 'workflow_run_id', 'workflow_step_id', 'step_no',
    'step_name_snapshot', 'approver_emp_id', 'approver_username',
    'approver_name_snapshot', 'can_self_approve', 'is_stock_posting_step',
    'task_status', 'acted_at', 'decision'],
  LOG:['action_id', 'workflow_run_id', 'task_id', 'doc_id', 'action_type',
    'comment', 'actor_username', 'actor_emp_id', 'actor_name_snapshot',
    'action_at', 'from_status', 'to_status'],
  USER_MAP:['mapping_id', 'username', 'emp_id', 'is_primary', 'status',
    'created_by', 'created_at', 'updated_by', 'updated_at']
};

var WMS_WORKFLOW_NAV_ = 'view-wms-wf';
var WMS_WORKFLOW_INBOX_NAV_ = 'view-wms-wf-inbox';

var WF_CACHE_TTL_SEC = 120;
var _wfCache_ = null;
var _wfCacheTime_ = 0;

function getWorkflowCache_() {
  var now = Date.now();
  if (_wfCache_ && now - _wfCacheTime_ < WF_CACHE_TTL_SEC * 1000) {
    return _wfCache_;
  }
  var cache = CacheService.getScriptCache();
  var raw = cache.get('WMS_WF_CACHE');
  if (raw) {
    try {
      _wfCache_ = JSON.parse(raw);
      _wfCacheTime_ = now;
      return _wfCache_;
    } catch (e) {
      cache.remove('WMS_WF_CACHE');
    }
  }
  return null;
}

function setWorkflowCache_(wfRows, stepRows) {
  _wfCache_ = { wf: wfRows, step: stepRows };
  _wfCacheTime_ = Date.now();
  try {
    CacheService.getScriptCache().put('WMS_WF_CACHE', JSON.stringify({ wf: wfRows, step: stepRows }), WF_CACHE_TTL_SEC);
  } catch (e) { /* cache ignore */ }
}

function invalidateWorkflowCache_() {
  _wfCache_ = null;
  _wfCacheTime_ = 0;
  try { CacheService.getScriptCache().remove('WMS_WF_CACHE'); } catch (e) { /* ignore */ }
}

function wmsWorkflowAuth_(token, tableName, action, navigationKey) {
  var auth = wmsCanRead_(token, tableName, navigationKey || WMS_WORKFLOW_NAV_);
  if (!auth.ok) return auth;
  var allowed = action === 'create' ? auth.canCreate
    : action === 'update' ? auth.canEdit
    : action === 'delete' ? auth.canDelete : true;
  if (!allowed) {
    return {
      ok:false,
      result:fail_('Bạn không có quyền ' +
        (action === 'create' ? 'tạo' : action === 'update' ? 'cập nhật' : 'xóa') +
        ' dữ liệu ' + tableName + '.', { code:'FORBIDDEN' })
    };
  }
  return auth;
}

/**
 * Tối ưu: Dùng lại instance spreadsheet đã mở để tránh gọi openById() nhiều lần
 */
/**
 * Tối ưu Bước 1: Dùng lại instance spreadsheet đã mở để tránh gọi openById() lặp lại
 */
/**
 * Tối ưu: Dùng lại instance spreadsheet đã mở để tránh gọi openById() lặp lại
 */
function wmsWorkflowGetSheets_(existingSpreadsheet) {
  var spreadsheet = existingSpreadsheet || SpreadsheetApp.openById(CONFIG.DB_WMS_PART1_ID);
  return {
    spreadsheet: spreadsheet,
    wf: wmsGetSheet_('SHEET_WMS_WF', WMS_WORKFLOW_HEADERS_.WF, spreadsheet),
    step: wmsGetSheet_('SHEET_WMS_WF_STEP', WMS_WORKFLOW_HEADERS_.STEP, spreadsheet),
    run: wmsGetSheet_('SHEET_WMS_WF_RUN', WMS_WORKFLOW_HEADERS_.RUN, spreadsheet),
    task: wmsGetSheet_('SHEET_WMS_WF_TASK', WMS_WORKFLOW_HEADERS_.TASK, spreadsheet),
    log: wmsGetSheet_('SHEET_WMS_WF_LOG', WMS_WORKFLOW_HEADERS_.LOG, spreadsheet),
    docs: wmsDocGetSheet_(spreadsheet),
    lines: wmsGetSheet_('SHEET_WMS_LINE', WMS_LINE_HEADERS_, spreadsheet)
  };
}

/**
 * Helper: Đọc nhanh theo cột dùng TextFinder thay vì load toàn bộ Sheet
 */
function wmsWorkflowReadByColumnFast_(source, headerName, expectedValue) {
  var colIndex = source.headers.indexOf(headerName);
  if (colIndex < 0) return [];
  var lastRow = source.sheet.getLastRow();
  if (lastRow < 2) return [];

  var finder = source.sheet.getRange(2, colIndex + 1, lastRow - 1, 1)
    .createTextFinder(String(expectedValue || ''))
    .matchEntireCell(true)
    .matchCase(false);
  var ranges = finder.findAll();
  if (!ranges.length) return [];

  var results = [];
  ranges.forEach(function(range) {
    var rIndex = range.getRowIndex();
    var rowVal = source.sheet.getRange(rIndex, 1, 1, source.headers.length).getValues()[0];
    results.push({ index: rIndex - 2, row: rowVal });
  });
  return results;
}

function wmsWorkflowGetSaveSheets_() {
  var spreadsheet = SpreadsheetApp.openById(CONFIG.DB_WMS_PART1_ID);
  return {
    wf:wmsGetSheet_('SHEET_WMS_WF', WMS_WORKFLOW_HEADERS_.WF, spreadsheet),
    step:wmsGetSheet_('SHEET_WMS_WF_STEP', WMS_WORKFLOW_HEADERS_.STEP, spreadsheet)
  };
}

function wmsWorkflowReadRows_(source) {
  var lastRow = source.sheet.getLastRow();
  return lastRow > 1
    ? source.sheet.getRange(2, 1, lastRow - 1, source.headers.length).getValues()
    : [];
}

/**
 * Đọc các dòng trong sheet có giá trị cột columnName khớp expectedValue.
 * Dùng TextFinder để chỉ quét cột cần tìm, không đọc toàn bộ sheet vào bộ nhớ.
 * Trả về [{index, row}, ...] tương thích với wmsWorkflowReadRows_.
 */
function wmsWorkflowReadByColumnFast_(source, columnName, expectedValue) {
  var colIndex = source.headers.indexOf(columnName);
  if (colIndex < 0) throw new Error('Sheet thiếu cột ' + columnName + '.');
  var lastRow = source.sheet.getLastRow();
  if (lastRow < 2) return [];
  var textFinder = source.sheet.getRange(2, colIndex + 1, lastRow - 1, 1).createTextFinder(String(expectedValue || ''))
    .matchEntireCell(true)
    .matchCase(false);
  var ranges = textFinder.findAll();
  if (!ranges.length) return [];
  var minRow = ranges[0].getRowIndex();
  var maxRow = ranges[ranges.length - 1].getRowIndex();
  var allRows = source.sheet.getRange(minRow, 1, maxRow - minRow + 1, source.headers.length).getValues();
  var rangeSet = {};
  ranges.forEach(function(range) { rangeSet[range.getRowIndex()] = true; });
  var results = [];
  allRows.forEach(function(row, offset) {
    var actualRow = minRow + offset;
    if (rangeSet[actualRow]) {
      results.push({ index: actualRow - 2, row: row });
    }
  });
  return results;
}

function wmsWorkflowReadSelectedRows_(source, selectedIndexes) {
  if (!selectedIndexes.length) return [];
  selectedIndexes.sort(function(a, b) { return a - b; });
  var result = [];
  var start = selectedIndexes[0], previous = start;
  var readBlock = function(first, last) {
    Array.prototype.push.apply(result,
      source.sheet.getRange(first + 2, 1, last - first + 1, source.headers.length).getValues());
  };
  for (var i = 1; i < selectedIndexes.length; i++) {
    if (selectedIndexes[i] !== previous + 1) {
      readBlock(start, previous);
      start = selectedIndexes[i];
    }
    previous = selectedIndexes[i];
  }
  readBlock(start, previous);
  return result;
}

function wmsWorkflowReadChangedRows_(source, since, updatedHeader) {
  var updatedIndex = source.headers.indexOf(updatedHeader);
  if (updatedIndex < 0) throw new Error('Sheet thiếu cột ' + updatedHeader + '.');
  var lastRow = source.sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = source.sheet.getRange(2, updatedIndex + 1, lastRow - 1, 1).getValues();
  var selected = [];
  values.forEach(function(row, index) {
    var time = new Date(row[0]).getTime();
    if (isFinite(time) && time >= since) selected.push(index);
  });
  return wmsWorkflowReadSelectedRows_(source, selected);
}

function wmsWorkflowReadRowsForIds_(source, idHeader, expectedIds) {
  var idIndex = source.headers.indexOf(idHeader);
  if (idIndex < 0) throw new Error('Sheet thiếu cột ' + idHeader + '.');
  var ids = {};
  expectedIds.forEach(function(id) { ids[String(id || '')] = true; });
  var lastRow = source.sheet.getLastRow();
  if (lastRow < 2 || !Object.keys(ids).length) return [];
  var values = source.sheet.getRange(2, idIndex + 1, lastRow - 1, 1).getValues();
  var selected = [];
  values.forEach(function(row, index) {
    if (ids[String(row[0] || '')]) selected.push(index);
  });
  return wmsWorkflowReadSelectedRows_(source, selected);
}

function wmsWorkflowReadEmployeeRows_(expectedIds) {
  if (!CONFIG.DB_HRM_PART1_ID) throw new Error('Chưa cấu hình DB_HRM_PART1_ID.');
  var spreadsheet = SpreadsheetApp.openById(CONFIG.DB_HRM_PART1_ID);
  var sheet = spreadsheet.getSheetByName(CONFIG.SHEET_HRM_EMP || 'HRM_EMP');
  if (!sheet) throw new Error('Không tìm thấy sheet HRM_EMP.');
  var width = sheet.getLastColumn();
  if (width < 3 || sheet.getLastRow() < 2) return { headers:[], rows:[] };
  var headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(function(value) {
    return String(value || '').trim();
  });
  var missing = ['emp_id', 'full_name', 'work_status'].filter(function(header) {
    return headers.indexOf(header) < 0;
  });
  if (missing.length) throw new Error('HRM_EMP thiếu cột: ' + missing.join(', ') + '.');
  var rowCount = sheet.getLastRow() - 1;
  if (Array.isArray(expectedIds)) {
    var targets = {};
    expectedIds.forEach(function(id) {
      var normalizedId = String(id || '').replace(/^'/, '').trim();
      if (normalizedId) targets[normalizedId] = true;
    });
    var idIndex = headers.indexOf('emp_id');
    var idValues = rowCount
      ? sheet.getRange(2, idIndex + 1, rowCount, 1).getValues()
      : [];
    var selected = [];
    idValues.forEach(function(row, index) {
      var empId = String(row[0] || '').replace(/^'/, '').trim();
      if (targets[empId]) selected.push(index);
    });
    return {
      headers:headers,
      rows:wmsWorkflowReadSelectedRows_({ sheet:sheet, headers:headers }, selected)
    };
  }
  return {
    headers:headers,
    rows:rowCount ? sheet.getRange(2, 1, rowCount, width).getValues() : []
  };
}

function wmsWorkflowEmployeeIndex_(employeeData, employeeAuth) {
  var empIndex = employeeData.headers.indexOf('emp_id');
  var nameIndex = employeeData.headers.indexOf('full_name');
  var statusIndex = employeeData.headers.indexOf('work_status');
  var deletedIndex = employeeData.headers.indexOf('is_deleted');
  var employees = {};
  employeeData.rows.forEach(function(row) {
    if (employeeAuth && !matrixRowAllowed_(employeeAuth.user, 'HRM_EMP', row, null, employeeAuth.permission)) return;
    var empId = String(row[empIndex] || '').replace(/^'/, '').trim();
    var workStatus = String(row[statusIndex] || '').trim().toLowerCase();
    if (!empId || (deletedIndex >= 0 && wmsIsDeleted_(row[deletedIndex])) ||
        ['đang làm việc', 'active', 'working'].indexOf(workStatus) < 0) return;
    if (employees[empId]) throw new Error('Mã nhân viên bị trùng trong HRM_EMP: ' + empId + '.');
    employees[empId] = {
      emp_id:empId,
      full_name:String(row[nameIndex] || '').trim()
    };
  });
  return employees;
}

function wmsWorkflowReadAccount_(username) {
  if (!CONFIG.DB_SYS_PART1_ID) throw new Error('Chưa cấu hình DB_SYS_PART1_ID.');
  var spreadsheet = SpreadsheetApp.openById(CONFIG.DB_SYS_PART1_ID);
  var sheet = spreadsheet.getSheetByName(CONFIG.SHEET_SYS_USR || 'SYS_USR');
  if (!sheet || sheet.getLastRow() < 2) return null;
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  var target = String(username || '').trim().toLowerCase();
  var matches = rows.filter(function(row) {
    return String(row[0] || '').trim().toLowerCase() === target && isActiveValue_(row[6]);
  });
  if (matches.length > 1) throw new Error('Username bị trùng trong SYS_USR.');
  return matches.length ? { username:String(matches[0][0]).trim(), full_name:String(matches[0][3] || '').trim() } : null;
}

function wmsWorkflowReadAccounts_() {
  if (!CONFIG.DB_SYS_PART1_ID) throw new Error('Chưa cấu hình DB_SYS_PART1_ID.');
  var spreadsheet = SpreadsheetApp.openById(CONFIG.DB_SYS_PART1_ID);
  var sheet = spreadsheet.getSheetByName(CONFIG.SHEET_SYS_USR || 'SYS_USR');
  if (!sheet || sheet.getLastRow() < 2) return [];
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  var accounts = {};
  rows.forEach(function(row) {
    if (!isActiveValue_(row[6])) return;
    var username = String(row[0] || '').trim();
    if (!username) return;
    var key = username.toLowerCase();
    if (accounts[key]) throw new Error('Username bị trùng trong SYS_USR: ' + username + '.');
    accounts[key] = { username:username, full_name:String(row[3] || '').trim() || username };
  });
  return Object.keys(accounts).map(function(key) { return accounts[key]; })
    .sort(function(a, b) { return a.username.localeCompare(b.username); });
}

function wmsWorkflowReadRequestedAccounts_(usernames) {
  if (!CONFIG.DB_SYS_PART1_ID) throw new Error('Chưa cấu hình DB_SYS_PART1_ID.');
  var targets = {};
  (usernames || []).forEach(function(username) {
    var key = String(username || '').trim().toLowerCase();
    if (key) targets[key] = true;
  });
  var spreadsheet = SpreadsheetApp.openById(CONFIG.DB_SYS_PART1_ID);
  var sheet = spreadsheet.getSheetByName(CONFIG.SHEET_SYS_USR || 'SYS_USR');
  if (!sheet || sheet.getLastRow() < 2 || !Object.keys(targets).length) return {};
  var rowCount = sheet.getLastRow() - 1;
  var usernamesInSheet = sheet.getRange(2, 1, rowCount, 1).getValues();
  var selected = [];
  usernamesInSheet.forEach(function(row, index) {
    var key = String(row[0] || '').trim().toLowerCase();
    if (targets[key]) selected.push(index);
  });

  var activeValues = [];
  if (selected.length) {
    selected.sort(function(a, b) { return a - b; });
    var start = selected[0], previous = start;
    var readBlock = function(first, last) {
      Array.prototype.push.apply(activeValues,
        sheet.getRange(first + 2, 7, last - first + 1, 1).getValues());
    };
    for (var i = 1; i < selected.length; i++) {
      if (selected[i] !== previous + 1) {
        readBlock(start, previous);
        start = selected[i];
      }
      previous = selected[i];
    }
    readBlock(start, previous);
  }

  var accounts = {};
  selected.forEach(function(rowIndex, index) {
    if (!isActiveValue_(activeValues[index][0])) return;
    var username = String(usernamesInSheet[rowIndex][0] || '').trim();
    var key = username.toLowerCase();
    if (accounts[key]) throw new Error('Username bị trùng trong SYS_USR: ' + username + '.');
    accounts[key] = { username:username };
  });
  return accounts;
}

function wmsWorkflowAccountIndex_() {
  var accounts = wmsWorkflowReadAccounts_();
  var index = {};
  accounts.forEach(function(account) { index[account.username.toLowerCase()] = account; });
  return index;
}

function wmsWorkflowGetSetup(token, since) {
  var wfAuth = wmsWorkflowAuth_(token, 'WMS_WF', 'read');
  if (!wfAuth.ok) return wfAuth.result;
  var stepAuth = wmsWorkflowAuth_(token, 'WMS_WF_STEP', 'read');
  if (!stepAuth.ok) return stepAuth.result;
  var cursor = Number(since) || 0;
  var incremental = cursor > 0;
  try {
    var sheets = wmsWorkflowGetSheets_(typeof docSpreadsheet !== 'undefined' ? docSpreadsheet : null);
    var employees = {};
    var warehouses = [];
    var accounts = [];
    var employeeAccess = true;
    if (!incremental) {
      var employeeAuth = wmsCanRead_(token, 'HRM_EMP', WMS_WORKFLOW_NAV_);
      if (employeeAuth.ok) {
        employees = wmsWorkflowEmployeeIndex_(wmsWorkflowReadEmployeeRows_(), employeeAuth);
      } else if (employeeAuth.result && employeeAuth.result.code === 'FORBIDDEN') {
        employeeAccess = false;
      } else {
        return employeeAuth.result;
      }
    }
    var workflowRows, workflowSteps;
    if (!incremental) {
      var cached = getWorkflowCache_();
      if (cached) {
        workflowRows = cached.wf;
        workflowSteps = cached.step;
      } else {
        workflowRows = wmsWorkflowReadRows_(sheets.wf);
        workflowSteps = wmsWorkflowReadRows_(sheets.step);
        setWorkflowCache_(workflowRows, workflowSteps);
      }
    } else {
      workflowRows = wmsWorkflowReadChangedRows_(sheets.wf, cursor, 'updated_at');
      var changedVersionIds = workflowRows.map(function(row) {
        return String(row[sheets.wf.headers.indexOf('workflow_version_id')] || '');
      });
      workflowSteps = wmsWorkflowReadRowsForIds_(sheets.step, 'workflow_version_id', changedVersionIds);
    }
    var changedVersionIds = workflowRows.map(function(row) {
      return String(row[sheets.wf.headers.indexOf('workflow_version_id')] || '');
    });
    var workflows = workflowRows.filter(function(row) {
      return matrixRowAllowed_(wfAuth.user, 'WMS_WF', row, null, wfAuth.permission);
    }).map(function(row) {
      var versionId = String(row[sheets.wf.headers.indexOf('workflow_version_id')] || '');
      var steps = workflowSteps.filter(function(step) {
        return String(step[sheets.step.headers.indexOf('workflow_version_id')] || '') === versionId &&
          matrixRowAllowed_(stepAuth.user, 'WMS_WF_STEP', step, null, stepAuth.permission);
      }).map(function(step) {
        return wmsRowsAsObjects_(sheets.step.headers, [step])[0];
      }).sort(function(a, b) { return Number(a.step_no) - Number(b.step_no); });
      var item = wmsRowsAsObjects_(sheets.wf.headers, [row])[0];
      item.steps = steps;
      return item;
    });
    if (!incremental) {
      var warehouseAuth = wmsCanRead_(token, 'WAREHOUSE', WMS_WORKFLOW_NAV_);
      if (!warehouseAuth.ok) return warehouseAuth.result;
      var warehouseSource = wmsGetSheet_('SHEET_WMS_WAREHOUSE', [
        'warehouse_id', 'warehouse_code', 'warehouse_name', 'is_active', 'is_deleted'
      ]);
      warehouses = wmsWorkflowReadRows_(warehouseSource).filter(function(row) {
        return wmsBoolean_(row[warehouseSource.headers.indexOf('is_active')]) &&
          !wmsIsDeleted_(row[warehouseSource.headers.indexOf('is_deleted')]) &&
          matrixRowAllowed_(warehouseAuth.user, 'WAREHOUSE', row, null, warehouseAuth.permission);
      }).map(function(row) {
        return {
          warehouse_id:row[warehouseSource.headers.indexOf('warehouse_id')],
          warehouse_code:row[warehouseSource.headers.indexOf('warehouse_code')],
          warehouse_name:row[warehouseSource.headers.indexOf('warehouse_name')]
        };
      });
      accounts = wmsWorkflowReadAccounts_();
    }
    var response = {
      success:true,
      workflows:workflows,
      incremental:incremental,
      cursor:Date.now(),
      permissions:{
        createWorkflow:wfAuth.canCreate && stepAuth.canCreate,
        updateWorkflow:wfAuth.canEdit,
        deleteWorkflow:wfAuth.canDelete
      }
    };
    if (!incremental) {
      response.employees = Object.keys(employees).map(function(empId) { return employees[empId]; });
      response.employeeAccess = employeeAccess;
      response.accounts = accounts;
      response.warehouses = warehouses;
    }
    return response;
  } catch (err) {
    Logger.log('wmsWorkflowGetSetup: ' + err);
    return { success:false, message:'Không tải được cấu hình luồng duyệt: ' + err.message };
  }
}

function wmsWorkflowSaveMapping(token, record) {
  var auth = wmsWorkflowAuth_(token, 'WMS_USER_MAP', record && record.mapping_id ? 'update' : 'create');
  if (!auth.ok) return auth.result;
  try {
    var source = wmsGetSheet_('SHEET_WMS_USER_MAP', WMS_WORKFLOW_HEADERS_.USER_MAP);
    var username = wmsText_(record && record.username, 100);
    var empId = wmsText_(record && record.emp_id, 100);
    if (!username || !empId) return fail_('Hãy chọn nhân viên và nhập username.');
    var employeeIndex = wmsWorkflowEmployeeIndex_(
      wmsWorkflowReadEmployeeRows_([empId]));
    if (!employeeIndex[empId]) return fail_('Nhân viên không tồn tại hoặc không còn làm việc.');
    var account = wmsWorkflowReadAccount_(username);
    if (!account) return fail_('Username không tồn tại hoặc tài khoản đã khóa.');
    var rows = wmsWorkflowReadRows_(source);
    var idIndex = source.headers.indexOf('mapping_id');
    var statusIndex = source.headers.indexOf('status');
    var usernameIndex = source.headers.indexOf('username');
    var empIndex = source.headers.indexOf('emp_id');
    var mappingId = String(record.mapping_id || '').trim();
    var existing = mappingId ? rows.map(function(row, index) {
      return { row:row, index:index };
    }).filter(function(entry) {
      return String(entry.row[idIndex] || '') === mappingId;
    }) : [];
    if (existing.length > 1) throw new Error('Mã mapping bị trùng trong WMS_USER_MAP.');
    if (mappingId && !existing.length) return fail_('Không tìm thấy mapping cần cập nhật.');
    var duplicate = rows.some(function(row, index) {
      if (existing.length && index === existing[0].index) return false;
      return String(row[statusIndex] || '').toUpperCase() === WMS_USER_MAP_STATUS_.ACTIVE &&
        (String(row[usernameIndex] || '').trim().toLowerCase() === username.toLowerCase() ||
         String(row[empIndex] || '').replace(/^'/, '').trim().toLowerCase() === empId.toLowerCase());
    });
    if (duplicate) return fail_('Username hoặc mã nhân viên đã được liên kết với một mapping đang hoạt động.');
    var now = new Date();
    var values = {
      mapping_id:mappingId || 'MAP-' + Utilities.getUuid().replace(/-/g, '').substring(0, 16).toUpperCase(),
      username:account.username,
      emp_id:empId,
      is_primary:true,
      status:WMS_USER_MAP_STATUS_.ACTIVE,
      created_by:existing.length ? existing[0].row[source.headers.indexOf('created_by')] : auth.user.username,
      created_at:existing.length ? existing[0].row[source.headers.indexOf('created_at')] : now,
      updated_by:auth.user.username,
      updated_at:now
    };
    var newRow = wmsRowForWrite_(source.headers, values);
    if (!matrixScopeAllowed_(auth.user, 'WMS_USER_MAP', newRow, null, auth.permission,
        existing.length ? 'update' : 'create')) {
      return fail_('Mapping nằm ngoài phạm vi dữ liệu được phân quyền.', { code:'FORBIDDEN' });
    }
    withLock_(function() {
      if (existing.length) {
        source.sheet.getRange(existing[0].index + 2, 1, 1, source.headers.length).setValues([newRow]);
      } else {
        source.sheet.getRange(source.sheet.getLastRow() + 1, 1, 1, source.headers.length).setValues([newRow]);
      }
    });
    wmsAudit_(auth.user, 'WMS_USER_MAP', existing.length ? 'UPDATE' : 'CREATE',
      values.mapping_id, username + ' → ' + empId);
    return { success:true, message:'Đã lưu liên kết tài khoản và nhân viên.' };
  } catch (err) {
    Logger.log('wmsWorkflowSaveMapping: ' + err);
    return { success:false, message:'Không lưu được mapping tài khoản: ' + err.message };
  }
}

function wmsWorkflowSetMappingStatus(token, mappingId, status) {
  var auth = wmsWorkflowAuth_(token, 'WMS_USER_MAP', 'update');
  if (!auth.ok) return auth.result;
  status = String(status || '').toUpperCase();
  if ([WMS_USER_MAP_STATUS_.ACTIVE, WMS_USER_MAP_STATUS_.INACTIVE].indexOf(status) < 0) {
    return fail_('Trạng thái mapping không hợp lệ.');
  }
  try {
    var source = wmsGetSheet_('SHEET_WMS_USER_MAP', WMS_WORKFLOW_HEADERS_.USER_MAP);
    return withLock_(function() {
      var rows = wmsWorkflowReadRows_(source);
      var index = rows.findIndex(function(row) {
        return String(row[source.headers.indexOf('mapping_id')] || '') === String(mappingId || '');
      });
      if (index < 0) return fail_('Không tìm thấy mapping.');
      var oldRow = rows[index];
      if (!matrixScopeAllowed_(auth.user, 'WMS_USER_MAP', oldRow, null, auth.permission, 'update')) {
        return fail_('Mapping nằm ngoài phạm vi dữ liệu được phân quyền.', { code:'FORBIDDEN' });
      }
      if (status === WMS_USER_MAP_STATUS_.ACTIVE) {
        var duplicate = rows.some(function(row, rowIndex) {
          return rowIndex !== index &&
            String(row[source.headers.indexOf('status')] || '').toUpperCase() === WMS_USER_MAP_STATUS_.ACTIVE &&
            (String(row[source.headers.indexOf('username')] || '').toLowerCase() ===
               String(oldRow[source.headers.indexOf('username')] || '').toLowerCase() ||
             String(row[source.headers.indexOf('emp_id')] || '').replace(/^'/, '').toLowerCase() ===
               String(oldRow[source.headers.indexOf('emp_id')] || '').replace(/^'/, '').toLowerCase());
        });
        if (duplicate) return fail_('Username hoặc nhân viên đã có mapping đang hoạt động khác.');
      }
      var changed = oldRow.slice();
      changed[source.headers.indexOf('status')] = status;
      changed[source.headers.indexOf('updated_by')] = auth.user.username;
      changed[source.headers.indexOf('updated_at')] = new Date();
      source.sheet.getRange(index + 2, 1, 1, source.headers.length).setValues([changed]);
      wmsAudit_(auth.user, 'WMS_USER_MAP', status === 'ACTIVE' ? 'ACTIVATE' : 'DEACTIVATE',
        mappingId, String(oldRow[source.headers.indexOf('username')] || ''));
      return { success:true, message:status === 'ACTIVE' ? 'Đã kích hoạt mapping.' : 'Đã ngừng mapping.' };
    });
  } catch (err) {
    Logger.log('wmsWorkflowSetMappingStatus: ' + err);
    return { success:false, message:'Không đổi được trạng thái mapping: ' + err.message };
  }
}

function wmsWorkflowSave(token, record) {
  var startedAt = Date.now();
  var perf = {
    step_count:Array.isArray(record && record.steps) ? record.steps.length : 0,
    warehouse_scoped:!!(record && record.warehouse_scope_id),
    lock:{}
  };
  var result;
  try {
    result = wmsWorkflowSaveImpl_(token, record, perf);
    return result;
  } finally {
    perf.total_ms = Date.now() - startedAt;
    perf.outcome = result && result.success ? 'success' : 'failure';
    Logger.log('[WMS_PERF] ' + JSON.stringify(perf));
  }
}

function wmsWorkflowSaveImpl_(token, record, perf) {
  var phaseStartedAt = Date.now();
  var wfAuth = wmsWorkflowAuth_(token, 'WMS_WF', 'create');
  if (!wfAuth.ok) {
    perf.auth_ms = Date.now() - phaseStartedAt;
    return wfAuth.result;
  }
  var stepAuth = wmsWorkflowAuth_(token, 'WMS_WF_STEP', 'create');
  if (!stepAuth.ok) {
    perf.auth_ms = Date.now() - phaseStartedAt;
    return stepAuth.result;
  }
  perf.auth_ms = Date.now() - phaseStartedAt;
  if (!record || typeof record !== 'object') return fail_('Dữ liệu workflow không hợp lệ.');
  try {
    phaseStartedAt = Date.now();
    var sheets = wmsWorkflowGetSaveSheets_();
    perf.workflow_sheet_setup_ms = Date.now() - phaseStartedAt;
    var docType = String(record.doc_type || '').toUpperCase();
    if ([WMS_DOC_TYPES_.PURCHASE_RECEIPT, WMS_DOC_TYPES_.SALES_ISSUE].indexOf(docType) < 0) {
      return fail_('Loại chứng từ không hợp lệ.');
    }
    var wfCode = wmsText_(record.wf_code, 30).toUpperCase();
    var wfName = wmsText_(record.wf_name, 150);
    var warehouseScopeId = wmsText_(record.warehouse_scope_id, 100);
    var changeReason = wmsText_(record.change_reason, 500);
    if (!/^[A-Z0-9_-]{2,30}$/.test(wfCode)) return fail_('Mã workflow chỉ gồm A-Z, 0-9, dấu gạch ngang hoặc gạch dưới.');
    if (!wfName) return fail_('Tên workflow là bắt buộc.');
    if (!Array.isArray(record.steps) || !record.steps.length || record.steps.length > 20) {
      return fail_('Workflow cần từ 1 đến 20 bước duyệt.');
    }
    phaseStartedAt = Date.now();
    var requestedEmployeeIds = record.steps.map(function(step) {
      return step && step.approver_emp_id;
    });
    var requestedUsernames = record.steps.map(function(step) {
      return step && step.approver_username;
    });
    var employeeIndex = wmsWorkflowEmployeeIndex_(
      wmsWorkflowReadEmployeeRows_(requestedEmployeeIds));
    perf.employee_lookup_ms = Date.now() - phaseStartedAt;
    phaseStartedAt = Date.now();
    var accountIndex = wmsWorkflowReadRequestedAccounts_(requestedUsernames);
    perf.account_lookup_ms = Date.now() - phaseStartedAt;
    phaseStartedAt = Date.now();
    var normalizedSteps = record.steps.map(function(step, index) {
      var username = String(step && step.approver_username || '').trim();
      var empId = String(step && step.approver_emp_id || '').replace(/^'/, '').trim();
      var person = employeeIndex[empId];
      var account = accountIndex[username.toLowerCase()];
      var stepName = wmsText_(step && step.step_name, 100);
      if (!person) throw new Error('Bước ' + (index + 1) + ': hãy chọn nhân viên đang làm việc.');
      if (!account) throw new Error('Bước ' + (index + 1) + ': username không tồn tại hoặc tài khoản đã khóa.');
      if (!stepName) throw new Error('Bước ' + (index + 1) + ': tên bước là bắt buộc.');
      return {
        workflow_step_id:'WFS-' + Utilities.getUuid().replace(/-/g, '').substring(0, 16).toUpperCase(),
        step_no:index + 1,
        step_name:stepName,
        approver_emp_id:person.emp_id,
        approver_username:account.username,
        approver_name_snapshot:person.full_name,
        can_self_approve:wmsBoolean_(step.can_self_approve),
         is_stock_posting_step:wmsBoolean_(step.is_stock_posting_step),
        created_at:new Date()
      };
    });
    perf.step_validation_ms = Date.now() - phaseStartedAt;
    if (warehouseScopeId) {
      phaseStartedAt = Date.now();
      var warehouseAuth = wmsCanRead_(token, 'WAREHOUSE', WMS_WORKFLOW_NAV_);
      if (!warehouseAuth.ok) return warehouseAuth.result;
      var warehouses = wmsGetSheet_('SHEET_WMS_WAREHOUSE', [
        'warehouse_id', 'warehouse_code', 'warehouse_name', 'is_deleted'
      ]);
      var warehouseRow = wmsWorkflowReadRows_(warehouses).find(function(row) {
        return String(row[warehouses.headers.indexOf('warehouse_id')] || '') === warehouseScopeId &&
          !wmsIsDeleted_(row[warehouses.headers.indexOf('is_deleted')]) &&
          wmsBoolean_(row[warehouses.headers.indexOf('is_active')]) &&
          matrixRowAllowed_(warehouseAuth.user, 'WAREHOUSE', row, null, warehouseAuth.permission);
      });
      if (!warehouseRow) return fail_('Kho áp dụng không hợp lệ hoặc ngoài phạm vi.');
      perf.warehouse_lookup_ms = Date.now() - phaseStartedAt;
    }
    var result = withLock_(function() {
      phaseStartedAt = Date.now();
      var wfRows = wmsWorkflowReadRows_(sheets.wf);
      perf.workflow_read_ms = Date.now() - phaseStartedAt;
      var matching = wfRows.filter(function(row) {
        return String(row[sheets.wf.headers.indexOf('wf_code')] || '').toUpperCase() === wfCode &&
          String(row[sheets.wf.headers.indexOf('doc_type')] || '') === docType &&
          String(row[sheets.wf.headers.indexOf('warehouse_scope_id')] || '') === warehouseScopeId;
      });
      var workflowId = '';
      var versionRows = matching;
      var sourceWorkflowId = String(record.source_workflow_id || '').trim();
      var sourceVersionId = String(record.source_workflow_version_id || '').trim();
      if (sourceWorkflowId || sourceVersionId) {
        var sourceVersion = wfRows.find(function(row) {
          return String(row[sheets.wf.headers.indexOf('workflow_version_id')] || '') === sourceVersionId &&
            String(row[sheets.wf.headers.indexOf('workflow_id')] || '') === sourceWorkflowId;
        });
        if (!sourceWorkflowId || !sourceVersionId || !sourceVersion ||
            String(sourceVersion[sheets.wf.headers.indexOf('status')] || '') === WMS_WORKFLOW_STATUS_.ARCHIVED ||
            !matrixRowAllowed_(wfAuth.user, 'WMS_WF', sourceVersion, null, wfAuth.permission) ||
            !matrixScopeAllowed_(wfAuth.user, 'WMS_WF', sourceVersion, null, wfAuth.permission, 'create')) {
          return fail_('Workflow gốc không tồn tại hoặc nằm ngoài phạm vi được phân quyền.',
            { code:'FORBIDDEN' });
        }
        workflowId = sourceWorkflowId;
        versionRows = wfRows.filter(function(row) {
          return String(row[sheets.wf.headers.indexOf('workflow_id')] || '') === workflowId;
        });
      } else {
        workflowId = matching.length
          ? String(matching[0][sheets.wf.headers.indexOf('workflow_id')] || '')
          : 'WF-' + Utilities.getUuid().replace(/-/g, '').substring(0, 16).toUpperCase();
      }
      var versionNo = versionRows.reduce(function(max, row) {
        return Math.max(max, Number(row[sheets.wf.headers.indexOf('version_no')]) || 0);
      }, 0) + 1;
      var versionId = 'WFV-' + Utilities.getUuid().replace(/-/g, '').substring(0, 16).toUpperCase();
      var now = new Date();
      var header = wmsRowForWrite_(sheets.wf.headers, {
        workflow_version_id:versionId,
        workflow_id:workflowId,
        wf_code:wfCode,
        wf_name:wfName,
        doc_type:docType,
        warehouse_scope_id:warehouseScopeId,
        version_no:versionNo,
        status:WMS_WORKFLOW_STATUS_.DRAFT,
        is_active:false,
        created_by_username:wfAuth.user.username,
        created_by_emp_id:'',
        created_at:now,
        updated_by_username:wfAuth.user.username,
        updated_at:now,
        activated_by_username:'',
        activated_at:'',
        change_reason:changeReason
      });
      if (!matrixScopeAllowed_(wfAuth.user, 'WMS_WF', header, null, wfAuth.permission, 'create')) {
        return fail_('Workflow nằm ngoài phạm vi được phân quyền.', { code:'FORBIDDEN' });
      }
      var stepRows = normalizedSteps.map(function(step) {
        var row = wmsRowForWrite_(sheets.step.headers,
          Object.assign({}, step, { workflow_version_id:versionId }));
        if (!matrixScopeAllowed_(stepAuth.user, 'WMS_WF_STEP', row, null, stepAuth.permission, 'create')) {
          throw wmsDocScopeError_('Một hoặc nhiều bước nằm ngoài phạm vi được phân quyền.');
        }
        return row;
      });
      phaseStartedAt = Date.now();
      var headerRow = sheets.wf.sheet.getLastRow() + 1;
      sheets.wf.sheet.getRange(headerRow, 1, 1, sheets.wf.headers.length).setValues([header]);
      try {
        sheets.step.sheet.getRange(sheets.step.sheet.getLastRow() + 1, 1,
          stepRows.length, sheets.step.headers.length).setValues(stepRows);
      } catch (writeError) {
        try {
          var archived = header.slice();
          archived[sheets.wf.headers.indexOf('status')] = WMS_WORKFLOW_STATUS_.ARCHIVED;
          sheets.wf.sheet.getRange(headerRow, 1, 1, sheets.wf.headers.length).setValues([archived]);
        } catch (rollbackError) {
          throw new Error(writeError.message + ' Workflow nháp cũng không thể đóng an toàn: ' + rollbackError.message);
        }
        throw writeError;
      }
      invalidateWorkflowCache_();
      perf.workflow_write_ms = Date.now() - phaseStartedAt;
      var workflow = wmsRowsAsObjects_(sheets.wf.headers, [header])[0];
      workflow.steps = stepRows.map(function(row) {
        return wmsRowsAsObjects_(sheets.step.headers, [row])[0];
      });
      return { success:true, message:'Đã tạo phiên bản workflow nháp.',
        workflow_version_id:versionId, wf_code:wfCode, version_no:versionNo, workflow:workflow };
    }, perf.lock);
    if (!result || !result.success) return result;
    phaseStartedAt = Date.now();
    wmsAudit_(wfAuth.user, 'WMS_WF', 'CREATE_DRAFT', result.workflow_version_id,
      result.wf_code + ' v' + result.version_no);
    perf.audit_ms = Date.now() - phaseStartedAt;
    return result;
  } catch (err) {
    Logger.log('wmsWorkflowSave: ' + err);
    var failure = { success:false, message:'Không lưu được workflow: ' + err.message };
    if (err.code) failure.code = err.code;
    return failure;
  }
}

function wmsWorkflowActivate(token, workflowVersionId) {
  var auth = wmsWorkflowAuth_(token, 'WMS_WF', 'update');
  if (!auth.ok) return auth.result;
  try {
    var wf = wmsGetSheet_('SHEET_WMS_WF', WMS_WORKFLOW_HEADERS_.WF);
    var step = wmsGetSheet_('SHEET_WMS_WF_STEP', WMS_WORKFLOW_HEADERS_.STEP);
    var stepAuth = wmsWorkflowAuth_(token, 'WMS_WF_STEP', 'read');
    if (!stepAuth.ok) return stepAuth.result;
    return withLock_(function() {
      var rows = wmsWorkflowReadRows_(wf);
      var idIndex = wf.headers.indexOf('workflow_version_id');
      var targetIndex = rows.findIndex(function(row) {
        return String(row[idIndex] || '') === String(workflowVersionId || '');
      });
      if (targetIndex < 0) return fail_('Không tìm thấy phiên bản workflow.');
      var target = rows[targetIndex];
      if (!matrixScopeAllowed_(auth.user, 'WMS_WF', target, null, auth.permission, 'update')) {
        return fail_('Workflow nằm ngoài phạm vi được phân quyền.', { code:'FORBIDDEN' });
      }
      if ([WMS_WORKFLOW_STATUS_.DRAFT, WMS_WORKFLOW_STATUS_.INACTIVE].indexOf(
          String(target[wf.headers.indexOf('status')] || '')) < 0) {
        return fail_('Chỉ có thể kích hoạt workflow nháp hoặc đang khóa.');
      }
      var versionId = String(target[idIndex] || '');
      var steps = wmsWorkflowReadRows_(step).filter(function(row) {
        return String(row[step.headers.indexOf('workflow_version_id')] || '') === versionId &&
          matrixRowAllowed_(stepAuth.user, 'WMS_WF_STEP', row, null, stepAuth.permission);
      }).sort(function(a, b) {
        return Number(a[step.headers.indexOf('step_no')]) - Number(b[step.headers.indexOf('step_no')]);
      });
      if (!steps.length) return fail_('Workflow chưa có bước duyệt.');
      var approverIds = steps.map(function(row) {
        return row[step.headers.indexOf('approver_emp_id')];
      });
      var employees = wmsWorkflowEmployeeIndex_(wmsWorkflowReadEmployeeRows_(approverIds));
      var accounts = wmsWorkflowAccountIndex_();
      for (var i = 0; i < steps.length; i++) {
        var username = String(steps[i][step.headers.indexOf('approver_username')] || '').toLowerCase();
        var configuredEmpId = String(steps[i][step.headers.indexOf('approver_emp_id')] || '')
          .replace(/^'/, '').trim();
        if (!employees[configuredEmpId]) {
          return fail_('Nhân viên người duyệt tại bước ' + (i + 1) + ' không còn làm việc.');
        }
        if (!accounts[username]) {
          return fail_('Tài khoản người duyệt tại bước ' + (i + 1) + ' không còn hoạt động.');
        }
      }
      var docType = wf.headers.indexOf('doc_type');
      var scope = wf.headers.indexOf('warehouse_scope_id');
      var status = wf.headers.indexOf('status');
      var active = wf.headers.indexOf('is_active');
      var activatedBy = wf.headers.indexOf('activated_by_username');
      var activatedAt = wf.headers.indexOf('activated_at');
      var updatedBy = wf.headers.indexOf('updated_by_username');
      var updatedAt = wf.headers.indexOf('updated_at');
      var changes = [];
      var activated = target.slice();
      activated[status] = WMS_WORKFLOW_STATUS_.ACTIVE;
      activated[active] = true;
      activated[activatedBy] = auth.user.username;
      activated[activatedAt] = new Date();
      activated[updatedBy] = auth.user.username;
      activated[updatedAt] = new Date();
      changes.push({ index:targetIndex, old:target.slice(), row:activated });
      var originals = changes.map(function(item) { return { index:item.index, row:item.old }; });
      try {
        wmsWorkflowWriteRows_(wf, changes);
      } catch (writeError) {
        try {
          wmsWorkflowWriteRows_(wf, originals);
        } catch (rollbackError) {
          throw new Error(writeError.message + ' Rollback kích hoạt thất bại: ' + rollbackError.message);
        }
        throw writeError;
      }
      invalidateWorkflowCache_();
      wmsAudit_(auth.user, 'WMS_WF', 'ACTIVATE', versionId,
        String(target[wf.headers.indexOf('wf_code')] || '') + ' v' +
        String(target[wf.headers.indexOf('version_no')] || ''));
      return { success:true, message:'Đã kích hoạt workflow.' };
    });
  } catch (err) {
    Logger.log('wmsWorkflowActivate: ' + err);
    var result = { success:false, message:'Không kích hoạt được workflow: ' + err.message };
    if (err.code) result.code = err.code;
    return result;
  }
}

function wmsWorkflowSetStatus(token, workflowVersionId, action) {
  action = String(action || '').toUpperCase();
  var permission = action === 'DELETE'
    ? wmsWorkflowAuth_(token, 'WMS_WF', 'delete')
    : wmsWorkflowAuth_(token, 'WMS_WF', 'update');
  if (!permission.ok) return permission.result;
  if (['LOCK', 'DELETE'].indexOf(action) < 0) return fail_('Thao tác workflow không hợp lệ.');
  try {
    var sheets = wmsWorkflowGetSheets_();
    return withLock_(function() {
      var rows = wmsWorkflowReadRows_(sheets.wf);
      var versionIndex = sheets.wf.headers.indexOf('workflow_version_id');
      var targetIndex = rows.findIndex(function(row) {
        return String(row[versionIndex] || '') === String(workflowVersionId || '');
      });
      if (targetIndex < 0) return fail_('Không tìm thấy phiên bản workflow.');
      var target = rows[targetIndex];
      if (!matrixScopeAllowed_(permission.user, 'WMS_WF', target, null, permission.permission,
          action === 'DELETE' ? 'delete' : 'update')) {
        return fail_('Workflow nằm ngoài phạm vi được phân quyền.', { code:'FORBIDDEN' });
      }
      var workflowId = String(target[sheets.wf.headers.indexOf('workflow_id')] || '');
      var versionRows = rows.map(function(row, index) {
        return { row:row, index:index };
      }).filter(function(item) {
        return String(item.row[sheets.wf.headers.indexOf('workflow_id')] || '') === workflowId;
      });
      if (action === 'LOCK') {
        if (String(target[sheets.wf.headers.indexOf('status')] || '') !== WMS_WORKFLOW_STATUS_.ACTIVE) {
          return fail_('Chỉ có thể khóa phiên bản workflow đang hoạt động.');
        }
        var locked = target.slice();
        locked[sheets.wf.headers.indexOf('status')] = WMS_WORKFLOW_STATUS_.INACTIVE;
        locked[sheets.wf.headers.indexOf('is_active')] = false;
        locked[sheets.wf.headers.indexOf('updated_by_username')] = permission.user.username;
        locked[sheets.wf.headers.indexOf('updated_at')] = new Date();
        sheets.wf.sheet.getRange(targetIndex + 2, 1, 1, sheets.wf.headers.length).setValues([locked]);
        invalidateWorkflowCache_();
        wmsAudit_(permission.user, 'WMS_WF', 'LOCK', workflowVersionId, workflowId);
        return { success:true, message:'Đã khóa workflow; các lượt duyệt đang chạy không bị ảnh hưởng.' };
      }
      var versionIds = {};
      versionRows.forEach(function(item) {
        if (!matrixScopeAllowed_(permission.user, 'WMS_WF', item.row, null, permission.permission, 'delete')) {
          throw wmsDocScopeError_('Một phiên bản workflow nằm ngoài phạm vi xóa.');
        }
        versionIds[String(item.row[versionIndex] || '')] = true;
      });
      var runVersionIndex = sheets.run.headers.indexOf('workflow_version_id');
      var versionIdList = Object.keys(versionIds);
      var inUse = versionIdList.some(function(vid) {
        return wmsWorkflowReadByColumnFast_(sheets.run, 'workflow_version_id', vid).length > 0;
      });
      if (inUse) return fail_('Workflow đã có lịch sử sử dụng; chỉ được khóa thay vì xóa.');
      var now = new Date();
      var archivedRows = versionRows.map(function(item) {
        var archived = item.row.slice();
        archived[sheets.wf.headers.indexOf('status')] = WMS_WORKFLOW_STATUS_.ARCHIVED;
        archived[sheets.wf.headers.indexOf('is_active')] = false;
        archived[sheets.wf.headers.indexOf('updated_by_username')] = permission.user.username;
        archived[sheets.wf.headers.indexOf('updated_at')] = now;
        return { index:item.index, row:archived };
      });
      try {
        wmsWorkflowWriteRows_(sheets.wf, archivedRows);
      } catch (writeError) {
        try {
          wmsWorkflowWriteRows_(sheets.wf, versionRows.map(function(item) {
            return { index:item.index, row:item.row.slice() };
          }));
        } catch (rollbackError) {
          throw new Error(writeError.message + ' Rollback xóa mềm thất bại: ' + rollbackError.message);
        }
        throw writeError;
      }
      invalidateWorkflowCache_();
      wmsAudit_(permission.user, 'WMS_WF', 'ARCHIVE', workflowId, 'No workflow runs');
      return { success:true, message:'Đã xóa mềm workflow chưa từng được sử dụng.' };
    });
  } catch (err) {
    Logger.log('wmsWorkflowSetStatus: ' + err);
    var result = { success:false, message:'Không cập nhật được workflow: ' + err.message };
    if (err.code) result.code = err.code;
    return result;
  }
}

function wmsWorkflowWriteRows_(source, changes) {
  changes.forEach(function(item) {
    source.sheet.getRange(item.index + 2, 1, 1, source.headers.length).setValues([item.row]);
  });
}

function wmsWorkflowListAvailable_(token, docType, warehouseId, navigationKey, inheritedDocAuth) {
  var auth = inheritedDocAuth && inheritedDocAuth.ok && inheritedDocAuth.user
    ? {
      ok:true,
      user:inheritedDocAuth.user,
      permission:{ rules:[{ canRead:true, conditions:[] }] }
    }
    : wmsWorkflowAuth_(token, 'WMS_WF', 'read', navigationKey);
  if (!auth.ok) return auth.result;
  var source = wmsGetSheet_('SHEET_WMS_WF', WMS_WORKFLOW_HEADERS_.WF);
  var headers = source.headers;
  var rows = wmsWorkflowReadRows_(source);
  var type = String(docType || '').toUpperCase();
  var scope = String(warehouseId || '');
  var workflows = rows.filter(function(row) {
    var rowScope = String(row[headers.indexOf('warehouse_scope_id')] || '');
    return String(row[headers.indexOf('doc_type')] || '').toUpperCase() === type &&
      String(row[headers.indexOf('status')] || '') === WMS_WORKFLOW_STATUS_.ACTIVE &&
      wmsBoolean_(row[headers.indexOf('is_active')]) &&
      (!scope || !rowScope || rowScope === scope) &&
      matrixRowAllowed_(auth.user, 'WMS_WF', row, null, auth.permission, 'read');
  }).map(function(row) {
    return {
      workflow_version_id:row[headers.indexOf('workflow_version_id')],
      wf_code:row[headers.indexOf('wf_code')],
      wf_name:row[headers.indexOf('wf_name')],
      doc_type:row[headers.indexOf('doc_type')],
      warehouse_scope_id:row[headers.indexOf('warehouse_scope_id')],
      version_no:row[headers.indexOf('version_no')]
    };
  });
  workflows.sort(function(a, b) {
    return String(a.wf_code || '').localeCompare(String(b.wf_code || ''));
  });
  return { success:true, workflows:workflows };
}

function wmsWorkflowSubmit(token, docId, selectedWorkflowVersionId) {
  try {
    var sheets = wmsWorkflowGetSheets_();
    return withLock_(function() {
      var matches = wmsDocReadByValue_(sheets.docs, 'doc_id', docId);
      if (matches.length !== 1) return fail_('Không tìm thấy phiếu cần gửi duyệt.');
      var doc = matches[0].row;
      var docType = String(doc[sheets.docs.headers.indexOf('doc_type')] || '');
      var auth = wmsDocAuth_(token, docType, 'update');
      if (!auth.ok) return auth.result;
      if (!matrixScopeAllowed_(auth.docAuth.user, 'DOC', doc, null, auth.docAuth.permission, 'update')) {
        return fail_('Phiếu nằm ngoài phạm vi gửi duyệt.', { code:'FORBIDDEN' });
      }
      if (wmsIsDeleted_(doc[sheets.docs.headers.indexOf('is_deleted')])) {
        return fail_('Phiếu đã được đánh dấu xóa.');
      }
      var statusIndex = sheets.docs.headers.indexOf('doc_status');
      var requesterIndex = sheets.docs.headers.indexOf('requested_by_username');
      if ([WMS_DOC_STATUS_.DRAFT, WMS_DOC_STATUS_.RETURNED].indexOf(String(doc[statusIndex] || '')) < 0 ||
          String(doc[sheets.docs.headers.indexOf('posting_status')] || '') !== WMS_POSTING_STATUS_.NOT_POSTED) {
        return fail_('Chỉ được gửi duyệt phiếu nháp hoặc phiếu đã trả về và chưa ghi sổ.');
      }
      if (String(doc[requesterIndex] || '').toLowerCase() !== String(auth.docAuth.user.username).toLowerCase()) {
        return fail_('Chỉ người đề nghị mới được gửi phiếu vào luồng duyệt.', { code:'FORBIDDEN' });
      }
      var lineRows = wmsDocReadByValue_(sheets.lines, 'doc_id', docId).filter(function(entry) {
        return !wmsIsDeleted_(entry.row[sheets.lines.headers.indexOf('is_deleted')]);
      });
      if (!lineRows.length) return fail_('Phiếu cần có ít nhất một dòng hàng trước khi gửi duyệt.');
      if (lineRows.some(function(entry) {
        return !matrixScopeAllowed_(auth.lineAuth.user, 'LINE', entry.row, null,
          auth.lineAuth.permission, 'read');
      })) return fail_('Một hoặc nhiều dòng phiếu nằm ngoài phạm vi dữ liệu được phân quyền.', { code:'FORBIDDEN' });

      var warehouseId = doc[sheets.docs.headers.indexOf('warehouse_id')];
      var savedWorkflowId = String(doc[sheets.docs.headers.indexOf('selected_workflow_version_id')] || '');
      var requestedWorkflowId = String(selectedWorkflowVersionId || savedWorkflowId).trim();
      if (!requestedWorkflowId) {
        return fail_('Vui lòng chọn workflow đang hoạt động trước khi gửi phiếu duyệt.');
      }
      var selectableWorkflows = wmsWorkflowListAvailable_(token, docType, warehouseId,
        wmsDocNavigationKey_(docType), auth.docAuth);
      if (!selectableWorkflows.success) return selectableWorkflows;
      var workflowOption = selectableWorkflows.workflows.find(function(item) {
        return String(item.workflow_version_id) === requestedWorkflowId;
      });
      if (!workflowOption) {
        return fail_('Workflow đã chọn không còn hoạt động hoặc không áp dụng cho loại phiếu/kho này. Hãy chọn workflow khác.');
      }
      var workflow = wmsWorkflowReadRows_(sheets.wf).find(function(row) {
        return String(row[sheets.wf.headers.indexOf('workflow_version_id')] || '') === requestedWorkflowId;
      });
      if (!workflow) return fail_('Không tìm thấy phiên bản workflow đã chọn.');
      var versionId = String(workflow[sheets.wf.headers.indexOf('workflow_version_id')] || '');
      var stepRows = wmsWorkflowReadRows_(sheets.step).filter(function(row) {
        return String(row[sheets.step.headers.indexOf('workflow_version_id')] || '') === versionId;
      }).sort(function(a, b) {
        return Number(a[sheets.step.headers.indexOf('step_no')]) -
          Number(b[sheets.step.headers.indexOf('step_no')]);
      });
      if (!stepRows.length) return fail_('Workflow đang hoạt động không có bước duyệt.');
      var approverIds = stepRows.map(function(row) {
        return row[sheets.step.headers.indexOf('approver_emp_id')];
      });
      var employees = wmsWorkflowEmployeeIndex_(wmsWorkflowReadEmployeeRows_(approverIds));
      var accounts = wmsWorkflowAccountIndex_();
      for (var i = 0; i < stepRows.length; i++) {
        var configuredUsername = String(stepRows[i][sheets.step.headers.indexOf('approver_username')] || '');
        var configuredEmpId = String(stepRows[i][sheets.step.headers.indexOf('approver_emp_id')] || '')
          .replace(/^'/, '').trim();
        if (!employees[configuredEmpId]) {
          return fail_('Nhân viên người duyệt tại bước ' + (i + 1) + ' không còn làm việc.');
        }
        if (!configuredUsername || !accounts[configuredUsername.toLowerCase()]) {
          return fail_('Tài khoản người duyệt tại bước ' + (i + 1) + ' không còn hoạt động.');
        }
      }
      var runs = wmsWorkflowReadByColumnFast_(sheets.run, 'doc_id', docId).map(function(match) {
        return match.row;
      });
      var previous = runs.length ? runs[runs.length - 1] : null;
      var previousId = previous ? String(previous[sheets.run.headers.indexOf('workflow_run_id')] || '') : '';
      var runId = 'WFR-' + Utilities.getUuid().replace(/-/g, '').substring(0, 16).toUpperCase();
      var now = new Date();
      var run = wmsRowForWrite_(sheets.run.headers, {
        workflow_run_id:runId,
        doc_id:docId,
        workflow_version_id:versionId,
        run_no:runs.length + 1,
        run_status:WMS_WORKFLOW_RUN_STATUS_.PENDING,
        current_step_no:1,
        started_by_username:auth.docAuth.user.username,
        started_by_emp_id:auth.docAuth.user.emp_id || '',
        started_at:now,
        completed_at:'',
        previous_run_id:previousId,
        return_reason:'',
        created_at:now
      });
      var task = wmsWorkflowTaskRow_(sheets.task.headers, runId, stepRows[0], sheets.step.headers);
      var requesterActor = Object.assign({}, auth.docAuth.user, {
        emp_id:auth.docAuth.user.emp_id || '',
        fullName:auth.docAuth.user.fullName || auth.docAuth.user.username
      });
      var log = wmsWorkflowLogRow_(sheets.log.headers, runId, '', docId,
        WMS_WORKFLOW_ACTIONS_.SUBMIT, '', requesterActor, 'DRAFT', 'PENDING', now);
      var oldDoc = doc.slice();
      doc[sheets.docs.headers.indexOf('workflow_run_id')] = runId;
      doc[sheets.docs.headers.indexOf('selected_workflow_version_id')] = versionId;
      doc[sheets.docs.headers.indexOf('requested_by_emp_id')] = auth.docAuth.user.emp_id || '';
      doc[sheets.docs.headers.indexOf('requested_by_name_snapshot')] =
        auth.docAuth.user.fullName || auth.docAuth.user.username;
      doc[sheets.docs.headers.indexOf('approval_status')] = WMS_APPROVAL_STATUS_.PENDING;
      doc[statusIndex] = WMS_DOC_STATUS_.IN_APPROVAL;
      doc[sheets.docs.headers.indexOf('submitted_at')] = now;
      doc[sheets.docs.headers.indexOf('updated_by')] = auth.docAuth.user.username;
      doc[sheets.docs.headers.indexOf('updated_at')] = now;
      var runRow = sheets.run.sheet.getLastRow() + 1;
      var taskRow = sheets.task.sheet.getLastRow() + 1;
      var logRow = sheets.log.sheet.getLastRow() + 1;
      var runWritten = false, taskWritten = false, logWritten = false, docWritten = false;
      try {
        sheets.run.sheet.getRange(runRow, 1, 1, sheets.run.headers.length).setValues([run]);
        runWritten = true;
        sheets.task.sheet.getRange(taskRow, 1, 1, sheets.task.headers.length).setValues([task]);
        taskWritten = true;
        sheets.docs.sheet.getRange(matches[0].index + 2, 1, 1, sheets.docs.headers.length).setValues([doc]);
        docWritten = true;
        sheets.log.sheet.getRange(logRow, 1, 1, sheets.log.headers.length).setValues([log]);
        logWritten = true;
      } catch (writeError) {
        var rollbackErrors = [];
        if (docWritten) {
          try { sheets.docs.sheet.getRange(matches[0].index + 2, 1, 1, sheets.docs.headers.length).setValues([oldDoc]); }
          catch (error) { rollbackErrors.push('khôi phục phiếu: ' + error.message); }
        }
        if (taskWritten) {
          try {
            var skipped = task.slice();
            skipped[sheets.task.headers.indexOf('task_status')] = WMS_WORKFLOW_TASK_STATUS_.SKIPPED;
            sheets.task.sheet.getRange(taskRow, 1, 1, sheets.task.headers.length).setValues([skipped]);
          } catch (error) { rollbackErrors.push('đóng task phát sinh: ' + error.message); }
        }
        if (runWritten) {
          try {
            var superseded = run.slice();
            superseded[sheets.run.headers.indexOf('run_status')] = WMS_WORKFLOW_RUN_STATUS_.SUPERSEDED;
            sheets.run.sheet.getRange(runRow, 1, 1, sheets.run.headers.length).setValues([superseded]);
          } catch (error) { rollbackErrors.push('đóng lần chạy phát sinh: ' + error.message); }
        }
        if (logWritten) {
          try {
            var cancelledLog = log.slice();
            cancelledLog[sheets.log.headers.indexOf('action_type')] = WMS_WORKFLOW_ACTIONS_.CANCEL;
            cancelledLog[sheets.log.headers.indexOf('comment')] = 'Rollback do lỗi gửi duyệt.';
            sheets.log.sheet.getRange(logRow, 1, 1, sheets.log.headers.length).setValues([cancelledLog]);
          } catch (error) { rollbackErrors.push('đóng log phát sinh: ' + error.message); }
        }
        throw new Error(writeError.message + (rollbackErrors.length ? ' Rollback lỗi: ' + rollbackErrors.join('; ') : ''));
      }
      wmsAudit_(auth.docAuth.user, 'DOC', 'SUBMIT_FOR_APPROVAL', docId, runId);
      return { success:true, message:'Đã gửi phiếu duyệt theo ' +
        String(workflow[sheets.wf.headers.indexOf('wf_code')] || '') + '.',
        workflow_run_id:runId };
    });
  } catch (err) {
    Logger.log('wmsWorkflowSubmit: ' + err);
    return { success:false, message:'Không gửi được phiếu duyệt: ' + err.message };
  }
}

function wmsWorkflowTaskRow_(headers, runId, step, stepHeaders) {
  var value = function(header) { return step[stepHeaders.indexOf(header)]; };
  return wmsRowForWrite_(headers, {
    task_id:'WFT-' + Utilities.getUuid().replace(/-/g, '').substring(0, 16).toUpperCase(),
    workflow_run_id:runId,
    workflow_step_id:value('workflow_step_id'),
    step_no:value('step_no'),
    step_name_snapshot:value('step_name'),
    approver_emp_id:value('approver_emp_id'),
    approver_username:value('approver_username'),
    approver_name_snapshot:value('approver_name_snapshot'),
    can_self_approve:value('can_self_approve'),
     is_stock_posting_step:value('is_stock_posting_step'),
    task_status:WMS_WORKFLOW_TASK_STATUS_.PENDING,
    acted_at:'',
    decision:''
  });
}

function wmsWorkflowLogRow_(headers, runId, taskId, docId, action, comment, actor, from, to, now) {
  return wmsRowForWrite_(headers, {
    action_id:'WFL-' + Utilities.getUuid().replace(/-/g, '').substring(0, 16).toUpperCase(),
    workflow_run_id:runId,
    task_id:taskId,
    doc_id:docId,
    action_type:action,
    comment:comment,
    actor_username:actor.username,
    actor_emp_id:actor.emp_id || '',
    actor_name_snapshot:actor.fullName || actor.username,
    action_at:now,
    from_status:from,
    to_status:to
  });
}

function wmsWorkflowMyTasks(token) {
  var taskAuth = wmsWorkflowAuth_(token, 'WMS_WF_TASK', 'read', WMS_WORKFLOW_INBOX_NAV_);
  if (!taskAuth.ok) return taskAuth.result;
  var docAuth = wmsCanRead_(token, 'DOC', WMS_WORKFLOW_INBOX_NAV_);
  if (!docAuth.ok) return docAuth.result;
  try {
    var sheets = wmsWorkflowGetSheets_();
    var username = String(taskAuth.user.username || '').toLowerCase();

    // Chỉ đọc task của user hiện tại (không đọc toàn bộ sheet task)
    var taskMatches = wmsWorkflowReadByColumnFast_(sheets.task, 'approver_username', username);

    // Lọc task_status = PENDING và collect runId
    var taskRows = [];
    var runIds = {};
    taskMatches.forEach(function(match) {
      var task = match.row;
      if (String(task[sheets.task.headers.indexOf('task_status')] || '') !== WMS_WORKFLOW_TASK_STATUS_.PENDING) return;
      taskRows.push(match);
      var runId = String(task[sheets.task.headers.indexOf('workflow_run_id')] || '');
      if (runId) runIds[runId] = true;
    });

    // Đọc run theo từng runId (không đọc toàn bộ sheet run)
    var runMap = {};
    Object.keys(runIds).forEach(function(runId) {
      var runMatches = wmsWorkflowReadByColumnFast_(sheets.run, 'workflow_run_id', runId);
      runMatches.forEach(function(rm) {
        runMap[String(rm.row[sheets.run.headers.indexOf('workflow_run_id')] || '')] = rm.row;
      });
    });

    // Collect docId từ các run, đọc doc theo từng docId
    var docIds = {};
    Object.keys(runMap).forEach(function(runId) {
      var run = runMap[runId];
      var docId = String(run[sheets.run.headers.indexOf('doc_id')] || '');
      if (docId) docIds[docId] = true;
    });
    var docMap = {};
    Object.keys(docIds).forEach(function(docId) {
      var docMatches = wmsDocReadByValue_(sheets.docs, 'doc_id', docId);
      docMatches.forEach(function(dm) {
        docMap[String(dm.row[sheets.docs.headers.indexOf('doc_id')] || '')] = dm.row;
      });
    });

    var list = taskRows.filter(function(match) {
      var task = match.row;
      var runId = String(task[sheets.task.headers.indexOf('workflow_run_id')] || '');
      var run = runMap[runId];
      if (!run || String(run[sheets.run.headers.indexOf('run_status')] || '') !== WMS_WORKFLOW_RUN_STATUS_.PENDING ||
          Number(run[sheets.run.headers.indexOf('current_step_no')]) !== Number(task[sheets.task.headers.indexOf('step_no')])) return false;
      var docId = String(run[sheets.run.headers.indexOf('doc_id')] || '');
      var doc = docMap[docId];
      return doc && !wmsIsDeleted_(doc[sheets.docs.headers.indexOf('is_deleted')]) &&
        matrixRowAllowed_(docAuth.user, 'DOC', doc, null, docAuth.permission);
    }).map(function(match) {
      var task = match.row;
      var runId = String(task[sheets.task.headers.indexOf('workflow_run_id')] || '');
      var run = runMap[runId];
      var docId = String(run[sheets.run.headers.indexOf('doc_id')] || '');
      var doc = docMap[docId];
      var item = {
        task_id:task[sheets.task.headers.indexOf('task_id')],
        workflow_run_id:run[sheets.run.headers.indexOf('workflow_run_id')],
        step_no:task[sheets.task.headers.indexOf('step_no')],
        step_name:task[sheets.task.headers.indexOf('step_name_snapshot')],
        approver_emp_id:task[sheets.task.headers.indexOf('approver_emp_id')],
        approver_username:task[sheets.task.headers.indexOf('approver_username')],
        can_self_approve:wmsBoolean_(task[sheets.task.headers.indexOf('can_self_approve')]),
        can_decide:taskAuth.canEdit && docAuth.canEdit &&
          matrixScopeAllowed_(docAuth.user, 'DOC', doc, null, docAuth.permission, 'update'),
        requester_username:doc[sheets.docs.headers.indexOf('requested_by_username')],
        doc_id:doc[sheets.docs.headers.indexOf('doc_id')],
        doc_no:doc[sheets.docs.headers.indexOf('doc_no')],
        doc_type:doc[sheets.docs.headers.indexOf('doc_type')],
        doc_date:wmsCellValue_(doc[sheets.docs.headers.indexOf('doc_date')]),
        warehouse_id:doc[sheets.docs.headers.indexOf('warehouse_id')],
        partner_name_snapshot:doc[sheets.docs.headers.indexOf('partner_name_snapshot')],
        external_ref_no:doc[sheets.docs.headers.indexOf('external_ref_no')],
        reason:doc[sheets.docs.headers.indexOf('reason')],
        started_at:wmsCellValue_(run[sheets.run.headers.indexOf('started_at')])
      };
      return item;
    });
    return { success:true, list:list, canDecide:taskAuth.canEdit };
  } catch (err) {
    Logger.log('wmsWorkflowMyTasks: ' + err);
    return { success:false, message:'Không tải được công việc chờ duyệt: ' + err.message, list:[] };
  }
}

function wmsWorkflowDecide(token, taskId, decision, comment) {
  var taskAuth = wmsWorkflowAuth_(token, 'WMS_WF_TASK', 'update', WMS_WORKFLOW_INBOX_NAV_);
  if (!taskAuth.ok) return taskAuth.result;
  var normalizedDecision = String(decision || '').toUpperCase();
  if (['APPROVE', 'RETURN'].indexOf(normalizedDecision) < 0) return fail_('Quyết định duyệt không hợp lệ.');
  var note = wmsText_(comment, 1000);
  if (normalizedDecision === 'RETURN' && !note) return fail_('Bắt buộc nhập lý do trả phiếu.');
  try {
    var sheets = wmsWorkflowGetSheets_(typeof docSpreadsheet !== 'undefined' ? docSpreadsheet : null);
    var docPermission = wmsCanRead_(token, 'DOC', WMS_WORKFLOW_INBOX_NAV_);
    if (!docPermission.ok) return docPermission.result;
    return withLock_(function() {
      var taskMatches = wmsWorkflowReadByColumnFast_(sheets.task, 'task_id', taskId);
      if (taskMatches.length !== 1) return fail_('Không tìm thấy task cần xử lý.');
      var task = taskMatches[0].row;
      if (!matrixScopeAllowed_(taskAuth.user, 'WMS_WF_TASK', task, null, taskAuth.permission, 'update')) {
        return fail_('Task nằm ngoài phạm vi dữ liệu được phân quyền.', { code:'FORBIDDEN' });
      }
      if (String(task[sheets.task.headers.indexOf('approver_username')] || '').toLowerCase() !==
          String(taskAuth.user.username || '').toLowerCase()) {
        return fail_('Task này được giao cho người duyệt khác.', { code:'FORBIDDEN' });
      }
      if (String(task[sheets.task.headers.indexOf('task_status')] || '') !== WMS_WORKFLOW_TASK_STATUS_.PENDING) {
        return fail_('Task đã được xử lý hoặc không còn hiệu lực.');
      }
      var runId = String(task[sheets.task.headers.indexOf('workflow_run_id')] || '');
      var runMatches = wmsWorkflowReadByColumnFast_(sheets.run, 'workflow_run_id', runId);
      if (runMatches.length !== 1) return fail_('Không tìm thấy lần chạy workflow.');
      var run = runMatches[0].row;
      if (String(run[sheets.run.headers.indexOf('run_status')] || '') !== WMS_WORKFLOW_RUN_STATUS_.PENDING ||
          Number(run[sheets.run.headers.indexOf('current_step_no')]) !== Number(task[sheets.task.headers.indexOf('step_no')])) {
        return fail_('Task không còn là bước duyệt hiện tại.');
      }
      var docId = String(run[sheets.run.headers.indexOf('doc_id')] || '');
      var docMatches = wmsDocReadByValue_(sheets.docs, 'doc_id', docId);
      if (docMatches.length !== 1) return fail_('Không tìm thấy phiếu của task.');
      var doc = docMatches[0].row;
      if (!matrixScopeAllowed_(docPermission.user, 'DOC', doc, null, docPermission.permission, 'update')) {
        return fail_('Phiếu nằm ngoài phạm vi xử lý được phân quyền.', { code:'FORBIDDEN' });
      }
      if (String(doc[sheets.docs.headers.indexOf('workflow_run_id')] || '') !== runId ||
          String(doc[sheets.docs.headers.indexOf('doc_status')] || '') !== WMS_DOC_STATUS_.IN_APPROVAL) {
        return fail_('Phiếu đã đổi trạng thái; task này không còn hiệu lực.');
      }
      if (wmsIsDeleted_(doc[sheets.docs.headers.indexOf('is_deleted')]) ||
          String(doc[sheets.docs.headers.indexOf('posting_status')] || '') !== WMS_POSTING_STATUS_.NOT_POSTED) {
        return fail_('Phiếu đã bị xóa hoặc ghi sổ; task này không còn hiệu lực.');
      }
      var requester = String(doc[sheets.docs.headers.indexOf('requested_by_username')] || '').toLowerCase();
      if (normalizedDecision === 'APPROVE' && requester === String(taskAuth.user.username || '').toLowerCase() &&
          !wmsBoolean_(task[sheets.task.headers.indexOf('can_self_approve')])) {
        return fail_('Workflow không cho phép người đề nghị tự duyệt bước này.', { code:'FORBIDDEN' });
      }
      var now = new Date();
      var oldTask = task.slice(), oldRun = run.slice(), oldDoc = doc.slice();
      var taskStatus = sheets.task.headers.indexOf('task_status');
      var decisionIndex = sheets.task.headers.indexOf('decision');
      var actedAt = sheets.task.headers.indexOf('acted_at');
      var runStatus = sheets.run.headers.indexOf('run_status');
      var currentStep = sheets.run.headers.indexOf('current_step_no');
      var completedAt = sheets.run.headers.indexOf('completed_at');
      var workflowAction = normalizedDecision === 'APPROVE'
        ? WMS_WORKFLOW_ACTIONS_.APPROVE : WMS_WORKFLOW_ACTIONS_.RETURN_TO_REQUESTER;
      task[taskStatus] = normalizedDecision === 'APPROVE'
        ? WMS_WORKFLOW_TASK_STATUS_.APPROVED : WMS_WORKFLOW_TASK_STATUS_.RETURNED;
      task[decisionIndex] = note;
      task[actedAt] = now;
      var nextTask = null;
      if (normalizedDecision === 'RETURN') {
        run[runStatus] = WMS_WORKFLOW_RUN_STATUS_.RETURNED;
        doc[sheets.docs.headers.indexOf('approval_status')] = WMS_APPROVAL_STATUS_.RETURNED;
        doc[sheets.docs.headers.indexOf('doc_status')] = WMS_DOC_STATUS_.RETURNED;
        run[sheets.run.headers.indexOf('return_reason')] = note;
      } else {
        var stepRows = wmsWorkflowReadRows_(sheets.step).filter(function(row) {
          return String(row[sheets.step.headers.indexOf('workflow_version_id')] || '') ===
            String(run[sheets.run.headers.indexOf('workflow_version_id')] || '');
        }).sort(function(a, b) {
          return Number(a[sheets.step.headers.indexOf('step_no')]) -
            Number(b[sheets.step.headers.indexOf('step_no')]);
        });
        var nextStep = stepRows.find(function(row) {
          return Number(row[sheets.step.headers.indexOf('step_no')]) >
            Number(task[sheets.task.headers.indexOf('step_no')]);
        });
        if (nextStep) {
          run[currentStep] = Number(nextStep[sheets.step.headers.indexOf('step_no')]);
          nextTask = wmsWorkflowTaskRow_(sheets.task.headers, runId, nextStep, sheets.step.headers);
         } else {
           run[runStatus] = WMS_WORKFLOW_RUN_STATUS_.COMPLETED;
           run[completedAt] = now;
           doc[sheets.docs.headers.indexOf('approval_status')] = WMS_APPROVAL_STATUS_.APPROVED;
           doc[sheets.docs.headers.indexOf('doc_status')] = WMS_DOC_STATUS_.APPROVED;
           var finalStep = stepRows.find(function(row) {
             return Number(row[sheets.step.headers.indexOf('step_no')]) ===
               Number(task[sheets.task.headers.indexOf('step_no')]);
           });
           if (finalStep && wmsBoolean_(finalStep[sheets.step.headers.indexOf('is_stock_posting_step')])) {
             wmsPost(token, docId);
           }
         }
      }
      var actor = Object.assign({}, taskAuth.user, {
        emp_id:task[sheets.task.headers.indexOf('approver_emp_id')],
        fullName:task[sheets.task.headers.indexOf('approver_name_snapshot')]
      });
      var log = wmsWorkflowLogRow_(sheets.log.headers, runId, taskId, docId,
        workflowAction, note, actor, 'PENDING',
        normalizedDecision === 'RETURN' ? 'RETURNED' :
          (nextTask ? 'PENDING' : 'APPROVED'), now);
      doc[sheets.docs.headers.indexOf('updated_by')] = taskAuth.user.username;
      doc[sheets.docs.headers.indexOf('updated_at')] = now;
      var nextTaskRow = sheets.task.sheet.getLastRow() + 1;
      var logRow = sheets.log.sheet.getLastRow() + 1;
      var newTaskWritten = false;
      try {
        sheets.task.sheet.getRange(taskMatches[0].index + 2, 1, 1, sheets.task.headers.length).setValues([task]);
        sheets.run.sheet.getRange(runMatches[0].index + 2, 1, 1, sheets.run.headers.length).setValues([run]);
        if (nextTask) {
          sheets.task.sheet.getRange(nextTaskRow, 1, 1, sheets.task.headers.length).setValues([nextTask]);
          newTaskWritten = true;
        }
        sheets.docs.sheet.getRange(docMatches[0].index + 2, 1, 1, sheets.docs.headers.length).setValues([doc]);
        sheets.log.sheet.getRange(logRow, 1, 1, sheets.log.headers.length).setValues([log]);
      } catch (writeError) {
        var rollbackErrors = [];
        try { sheets.task.sheet.getRange(taskMatches[0].index + 2, 1, 1, sheets.task.headers.length).setValues([oldTask]); }
        catch (error) { rollbackErrors.push('khôi phục task: ' + error.message); }
        try { sheets.run.sheet.getRange(runMatches[0].index + 2, 1, 1, sheets.run.headers.length).setValues([oldRun]); }
        catch (error) { rollbackErrors.push('khôi phục run: ' + error.message); }
        try { sheets.docs.sheet.getRange(docMatches[0].index + 2, 1, 1, sheets.docs.headers.length).setValues([oldDoc]); }
        catch (error) { rollbackErrors.push('khôi phục phiếu: ' + error.message); }
        if (newTaskWritten) {
          try {
            var skippedTask = nextTask.slice();
            skippedTask[sheets.task.headers.indexOf('task_status')] = WMS_WORKFLOW_TASK_STATUS_.SKIPPED;
            sheets.task.sheet.getRange(nextTaskRow, 1, 1, sheets.task.headers.length).setValues([skippedTask]);
          } catch (error) { rollbackErrors.push('đóng task tiếp theo: ' + error.message); }
        }
        throw new Error(writeError.message + (rollbackErrors.length ? ' Rollback lỗi: ' + rollbackErrors.join('; ') : ''));
      }
      wmsAudit_(taskAuth.user, 'WMS_WF_TASK', workflowAction, taskId, note || docId);
      return {
        success:true,
        message:normalizedDecision === 'RETURN' ? 'Đã trả phiếu về người đề nghị.' :
          (nextTask ? 'Đã duyệt bước này và chuyển tới người duyệt tiếp theo.' :
            'Workflow đã hoàn tất; phiếu đã được duyệt nhưng chưa ghi sổ tồn kho.')
      };
    });
  } catch (err) {
    Logger.log('wmsWorkflowDecide: ' + err);
    return { success:false, message:'Không xử lý được task duyệt: ' + err.message };
  }
}
