/**
 * WMS_Core_Server.gs - Tiện ích dùng chung cho các màn hình quản lý kho.
 */

var WMS_DOC_TYPES_ = Object.freeze({
  PURCHASE_RECEIPT:'PURCHASE_RECEIPT',
  SALES_ISSUE:'SALES_ISSUE',
  CUSTOMER_RETURN:'CUSTOMER_RETURN',
  PRODUCTION_ISSUE:'PRODUCTION_ISSUE',
  PRODUCTION_RECEIPT:'PRODUCTION_RECEIPT',
  TRANSFER:'TRANSFER',
  STOCK_ADJUSTMENT:'STOCK_ADJUSTMENT'
});

var WMS_APPROVAL_STATUS_ = Object.freeze({
  DRAFT:'DRAFT',
  PENDING:'PENDING',
  RETURNED:'RETURNED',
  APPROVED:'APPROVED',
  CANCELLED:'CANCELLED'
});

var WMS_POSTING_STATUS_ = Object.freeze({
  NOT_POSTED:'NOT_POSTED',
  POSTING:'POSTING',
  POSTED:'POSTED',
  POST_FAILED:'POST_FAILED',
  REVERSED:'REVERSED'
});

var WMS_DOC_STATUS_ = Object.freeze({
  DRAFT:'DRAFT',
  IN_APPROVAL:'IN_APPROVAL',
  RETURNED:'RETURNED',
  APPROVED:'APPROVED',
  COMPLETED:'COMPLETED',
  CANCELLED:'CANCELLED'
});

var WMS_STOCK_STATUS_ = Object.freeze({
  AVAILABLE:'AVAILABLE',
  QUARANTINE:'QUARANTINE'
});

var WMS_LOT_QUALITY_STATUS_ = Object.freeze({
  PENDING:'PENDING',
  PASSED:'PASSED',
  FAILED:'FAILED',
  CONDITIONAL:'CONDITIONAL'
});

var WMS_LOT_STATUS_ = Object.freeze({
  ACTIVE:'ACTIVE',
  DEPLETED:'DEPLETED',
  BLOCKED:'BLOCKED',
  CLOSED:'CLOSED'
});

var WMS_QC_DISPOSITION_ = Object.freeze({
  RELEASE:'RELEASE',
  QUARANTINE:'QUARANTINE',
  REJECT:'REJECT',
  REWORK:'REWORK'
});

var WMS_LEDGER_TYPES_ = Object.freeze({
  RECEIPT:'RECEIPT',
  ISSUE:'ISSUE',
  TRANSFER_IN:'TRANSFER_IN',
  TRANSFER_OUT:'TRANSFER_OUT',
  QC_QUARANTINE:'QC_QUARANTINE',
  QC_RELEASE:'QC_RELEASE',
  ADJUSTMENT_IN:'ADJUSTMENT_IN',
  ADJUSTMENT_OUT:'ADJUSTMENT_OUT',
  REVERSAL:'REVERSAL'
});

var WMS_WORKFLOW_STATUS_ = Object.freeze({
  DRAFT:'DRAFT',
  ACTIVE:'ACTIVE',
  INACTIVE:'INACTIVE',
  ARCHIVED:'ARCHIVED'
});

var WMS_WORKFLOW_RUN_STATUS_ = Object.freeze({
  PENDING:'PENDING',
  RETURNED:'RETURNED',
  COMPLETED:'COMPLETED',
  CANCELLED:'CANCELLED',
  SUPERSEDED:'SUPERSEDED'
});

var WMS_WORKFLOW_TASK_STATUS_ = Object.freeze({
  PENDING:'PENDING',
  APPROVED:'APPROVED',
  RETURNED:'RETURNED',
  SKIPPED:'SKIPPED'
});

var WMS_WORKFLOW_ACTIONS_ = Object.freeze({
  SUBMIT:'SUBMIT',
  APPROVE:'APPROVE',
  RETURN_TO_REQUESTER:'RETURN_TO_REQUESTER',
  CHANGE_WORKFLOW:'CHANGE_WORKFLOW',
  STOCK_POSTED:'STOCK_POSTED',
  STOCK_REVERSED:'STOCK_REVERSED',
  CANCEL:'CANCEL'
});

var WMS_POSTING_PHASES_ = Object.freeze({
  STARTED:'STARTED',
  LEDGER_WRITTEN:'LEDGER_WRITTEN',
  BALANCE_WRITTEN:'BALANCE_WRITTEN',
  COMPLETED:'COMPLETED',
  FAILED:'FAILED'
});

var WMS_PRODUCTION_STATUS_ = Object.freeze({
  DRAFT:'DRAFT',
  IN_PROGRESS:'IN_PROGRESS',
  COMPLETED:'COMPLETED',
  CANCELLED:'CANCELLED'
});

var WMS_ATTACHMENT_STATUS_ = Object.freeze({
  ACTIVE:'ACTIVE',
  ARCHIVED:'ARCHIVED'
});

var WMS_USER_MAP_STATUS_ = Object.freeze({
  ACTIVE:'ACTIVE',
  INACTIVE:'INACTIVE'
});

function wmsPostingKey_(docId, workflowRunId, workflowStepId, postingType) {
  var parts = [docId, workflowRunId, workflowStepId, postingType].map(function(value) {
    return String(value == null ? '' : value).trim();
  });
  if (parts.some(function(value) { return !value; })) {
    throw new Error('Khóa ghi sổ cần doc_id, workflow_run_id, workflow_step_id và posting_type.');
  }
  return ['WMS_POST_V1'].concat(parts.map(encodeURIComponent)).join('|');
}

function wmsLedgerTransactionKey_(postingKey, lineId, leg) {
  var parts = [postingKey, lineId, leg].map(function(value) {
    return String(value == null ? '' : value).trim();
  });
  if (parts.some(function(value) { return !value; })) {
    throw new Error('Khóa giao dịch ledger cần posting_key, line_id và leg.');
  }
  return parts.map(encodeURIComponent).join('|');
}

/* Retry posting by posting_key: check stock under lock, record ledger once,
   reconcile WMS_BAL, then mark WMS_DOC posted. Never mutate posted ledger rows. */
function wmsGetSheet_(configKey, requiredHeaders, existingSpreadsheet) {
  if (typeof CONFIG === 'undefined' || !CONFIG.DB_WMS_PART1_ID) {
    throw new Error('Chưa cấu hình DB_WMS_PART1_ID.');
  }
  var sheetName = CONFIG[configKey];
  if (!sheetName) throw new Error('Chưa cấu hình tên sheet ' + configKey + '.');

  var spreadsheet = existingSpreadsheet || SpreadsheetApp.openById(CONFIG.DB_WMS_PART1_ID);
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error('Không tìm thấy sheet ' + sheetName + ' trong Spreadsheet WMS.');

  var lastColumn = sheet.getLastColumn();
  if (!lastColumn) throw new Error('Sheet ' + sheetName + ' chưa có hàng tiêu đề.');
  var headers = getSheetHeaders_(sheet);
  var missing = requiredHeaders.filter(function(header) {
    return headers.indexOf(header) === -1;
  });
  if (missing.length) {
    throw new Error('Sheet ' + sheetName + ' thiếu cột bắt buộc: ' + missing.join(', ') + '.');
  }
  return { spreadsheet: spreadsheet, sheet: sheet, headers: headers };
}

function wmsCellValue_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, 'GMT+7', 'dd/MM/yyyy HH:mm:ss');
  if (value == null) return '';
  return typeof value === 'string' ? value.replace(/^'/, '') : value;
}

function wmsRowsAsObjects_(headers, rows) {
  return rows.map(function(row) {
    var item = {};
    headers.forEach(function(header, index) {
      if (header) item[header] = wmsCellValue_(row[index]);
    });
    return item;
  });
}

function wmsIsDeleted_(value) {
  return value === true || ['TRUE', 'Y', '1', 'YES'].indexOf(String(value || '').trim().toUpperCase()) !== -1;
}

function wmsCanRead_(token, sheetName, navigationKey) {
  var bypass = isPermissionBypassed_(sheetName);
  var auth;
  if (bypass) {
    var sess = getSessionUserFromToken_(token);
    if (!sess.success) return { ok: false, result: sess };
    auth = { ok: true, user: sess.user };
  } else {
    auth = requireSessionAuth_(token, 'WMS', navigationKey);
    if (!auth.ok) return { ok: false, result: auth.res };
  }
  var permission = bypass ? { rules: [{ canRead: true, canCreate: true, canEdit: true, canDelete: true, conditions: [] }] } : getEffectivePermission_(auth.user.username, auth.user.role, sheetName);
  if (!bypass && !hasPermission_(auth.user, 'WMS', 'view', sheetName, navigationKey, permission)) {
    return {
      ok: false,
      result: {
        success: false,
        code: 'FORBIDDEN',
        message: 'Bạn không có quyền thực hiện thao tác này.',
        list: []
      }
    };
  }
  return {
    ok: true,
    user: auth.user,
    permission: permission,
    canRead: bypass ? true : hasPermission_(auth.user, 'WMS', 'view', sheetName, navigationKey, permission),
    canCreate: bypass ? true : hasPermission_(auth.user, 'WMS', 'create', sheetName, navigationKey, permission),
    canEdit: bypass ? true : hasPermission_(auth.user, 'WMS', 'update', sheetName, navigationKey, permission),
    canDelete: bypass ? true : hasPermission_(auth.user, 'WMS', 'delete', sheetName, navigationKey, permission)
  };
}

function wmsSoftDelete_(auth, source, tableName, idField, idValue, expectedUpdatedAt, label, beforeDelete) {
  var headers = source.headers;
  var idIndex = headers.indexOf(idField);
  var deletedIndex = headers.indexOf('is_deleted');
  var updatedAtIndex = headers.indexOf('updated_at');
  if (idIndex < 0 || deletedIndex < 0 || updatedAtIndex < 0 || headers.indexOf('updated_by') < 0) {
    throw new Error('Sheet ' + tableName + ' cần có các cột ' + idField + ', is_deleted, updated_by và updated_at.');
  }

  return withLock_(function() {
    var sheet = source.sheet;
    var lastRow = sheet.getLastRow();
    var rows = lastRow > 1
      ? sheet.getRange(2, 1, lastRow - 1, headers.length).getValues()
      : [];
    var targetIndex = rows.findIndex(function(row) {
      return String(row[idIndex] || '').trim().toLowerCase() === String(idValue || '').trim().toLowerCase();
    });
    if (targetIndex < 0) return fail_('Không tìm thấy ' + label + ' cần xóa.');

    var oldRow = rows[targetIndex];
    if (wmsIsDeleted_(oldRow[deletedIndex])) return fail_(label + ' đã được đánh dấu xóa.');
    if (String(expectedUpdatedAt || '') !== String(wmsCellValue_(oldRow[updatedAtIndex]) || '')) {
      return fail_(label + ' đã được người khác cập nhật. Vui lòng tải lại dữ liệu trước khi xóa.', {
        code: 'STALE_RECORD'
      });
    }
    if (!matrixScopeAllowed_(auth.user, tableName, oldRow, null, auth.permission, 'delete')) {
      return fail_(label + ' nằm ngoài phạm vi dữ liệu được phân quyền.', { code: 'FORBIDDEN' });
    }
    if (beforeDelete) {
      var validation = beforeDelete(oldRow, rows);
      if (validation) return validation;
    }

    var deletedRow = oldRow.slice();
    var now = new Date();
    headers.forEach(function(header, index) {
      if (header === 'is_deleted') deletedRow[index] = true;
      if (header === 'updated_by') deletedRow[index] = auth.user.username;
      if (header === 'updated_at') deletedRow[index] = now;
    });
    sheet.getRange(targetIndex + 2, 1, 1, headers.length).setValues([deletedRow]);
    var deletedId = String(oldRow[idIndex] || '').trim();
    wmsAudit_(auth.user, tableName, 'SOFT_DELETE', deletedId, label);
    return {
      success: true,
      message: 'Đã đánh dấu xóa ' + label + '. Dòng dữ liệu vẫn được giữ trong sheet.',
      id: deletedId,
      updated_at: wmsCellValue_(deletedRow[updatedAtIndex])
    };
  });
}

function wmsFilterReadableRows_(auth, sheetName, rows) {
  return rows.filter(function(row) {
    return matrixRowAllowed_(auth.user, sheetName, row, null, auth.permission);
  });
}

function wmsText_(value, maxLength) {
  var text = String(value == null ? '' : value).trim();
  if (maxLength && text.length > maxLength) {
    throw new Error('Giá trị vượt quá ' + maxLength + ' ký tự.');
  }
  return safeText_(text);
}

function wmsBoolean_(value) {
  return value === true || String(value).trim().toUpperCase() === 'TRUE' ||
    String(value).trim().toUpperCase() === 'Y' || String(value).trim() === '1';
}

function wmsOptionalNumber_(value, label) {
  if (value === '' || value == null) return '';
  var number = Number(value);
  if (!isFinite(number)) throw new Error(label + ' phải là số hợp lệ.');
  return number;
}

function wmsRowForWrite_(headers, values) {
  return headers.map(function(header) {
    return Object.prototype.hasOwnProperty.call(values, header) ? values[header] : '';
  });
}

function wmsAudit_(user, sheetName, action, key, detail) {
  if (typeof auditLog_ === 'function') {
    auditLog_(user, sheetName, action, key, detail);
  }
}

function wmsAuditBatch_(user, sheetName, action, records) {
  if (!records || !records.length) return;
  if (typeof auditLogBatch_ === 'function') {
    auditLogBatch_(user, sheetName, action, records);
    return;
  }
  records.forEach(function(record) {
    wmsAudit_(user, sheetName, action, record.key, record.detail);
  });
}

function keepAppWarm() {
  return true;
}
