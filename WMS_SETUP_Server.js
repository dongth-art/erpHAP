/**
 * Khởi tạo các bảng giao dịch và truy xuất nguồn gốc WMS.
 * Chạy setupWmsDataSheets() một lần trong Apps Script để tạo các sheet còn thiếu.
 */

var WMS_DATA_SHEET_SCHEMAS_ = [
  {
    configKey:'SHEET_WMS_DOC',
    headers:[
      'doc_id', 'doc_no', 'doc_type', 'doc_date', 'warehouse_id',
      'destination_warehouse_id', 'partner_id', 'partner_code_snapshot',
      'partner_name_snapshot', 'external_ref_type', 'external_ref_id',
      'external_ref_no', 'requested_by_username', 'requested_by_emp_id',
      'requested_by_name_snapshot', 'workflow_run_id', 'approval_status',
      'posting_status', 'doc_status', 'submitted_at', 'posted_at', 'completed_at',
      'reason', 'notes', 'reversal_of_doc_id', 'created_by', 'created_at',
      'updated_by', 'updated_at', 'is_deleted', 'selected_workflow_version_id'
    ]
  },
  {
    configKey:'SHEET_WMS_LINE',
    headers:[
      'line_id', 'doc_id', 'line_no', 'sku_id', 'sku_code_snapshot',
      'sku_name_snapshot', 'lot_id', 'warehouse_id', 'location_id',
      'destination_warehouse_id', 'destination_location_id', 'quantity_entered',
      'uom_entered', 'conversion_to_base', 'quantity_base', 'base_uom_snapshot',
      'stock_status_from', 'stock_status_to', 'unit_cost', 'line_amount',
      'production_batch_id', 'source_doc_line_id', 'qc_status_snapshot',
      'expiry_exception_confirmed', 'expiry_exception_reason',
      'expiry_confirmed_by_username', 'expiry_confirmed_by_emp_id',
      'expiry_confirmed_by_name_snapshot', 'expiry_confirmed_at', 'line_notes',
      'created_by', 'created_at', 'updated_by', 'updated_at', 'is_deleted'
    ]
  },
  {
    configKey:'SHEET_WMS_POST',
    headers:[
      'posting_id', 'posting_key', 'doc_id', 'workflow_run_id',
      'workflow_step_id', 'posting_type', 'status', 'payload_hash',
      'ledger_row_count', 'attempt_count', 'started_at', 'ledger_written_at',
      'balance_written_at', 'completed_at', 'last_error', 'created_by',
      'updated_at'
    ]
  },
  {
    configKey:'SHEET_WMS_LOT',
    headers:[
      'lot_id', 'internal_lot_no', 'sku_id', 'supplier_id',
      'supplier_code_snapshot', 'supplier_name_snapshot', 'supplier_lot_no',
      'manufacture_date', 'expiry_date', 'first_receipt_doc_id',
      'first_receipt_line_id', 'quality_status', 'lot_status', 'created_by',
      'created_at', 'updated_by', 'updated_at', 'is_deleted'
    ]
  },
  {
    configKey:'SHEET_WMS_QC',
    headers:[
      'qc_event_id', 'lot_id', 'doc_id', 'line_id', 'inspection_type',
      'result', 'disposition', 'reason', 'evidence_attachment_id',
      'decided_by_username', 'decided_by_emp_id', 'decided_by_name_snapshot',
      'decided_at', 'created_at'
    ]
  },
  {
    configKey:'SHEET_WMS_LEDGER',
    headers:[
      'ledger_id', 'transaction_key', 'doc_id', 'line_id', 'transaction_type',
      'posted_at', 'sku_id', 'warehouse_id', 'location_id', 'lot_id',
      'stock_status', 'quantity_delta_base', 'base_uom_snapshot', 'unit_cost',
      'value_delta', 'production_batch_id', 'reversal_of_ledger_id',
      'performed_by_username', 'performed_by_emp_id',
      'performed_by_name_snapshot', 'created_at'
    ]
  },
  {
    configKey:'SHEET_WMS_BAL',
    headers:[
      'balance_id', 'sku_id', 'warehouse_id', 'location_id', 'lot_id',
      'stock_status', 'quantity_on_hand_base', 'inventory_value',
      'moving_average_cost', 'last_ledger_id', 'last_updated_at'
    ]
  },
  {
    configKey:'SHEET_WMS_WF',
    headers:[
      'workflow_version_id', 'workflow_id', 'wf_code', 'wf_name', 'doc_type',
      'warehouse_scope_id', 'version_no', 'status', 'is_active',
      'created_by_username', 'created_by_emp_id', 'created_at',
      'updated_by_username', 'updated_at', 'activated_by_username',
      'activated_at', 'change_reason'
    ]
  },
  {
    configKey:'SHEET_WMS_WF_STEP',
    headers:[
      'workflow_step_id', 'workflow_version_id', 'step_no', 'step_name',
      'approver_emp_id', 'approver_username', 'approver_name_snapshot',
      'can_self_approve', 'is_stock_posting_step', 'created_at'
    ]
  },
  {
    configKey:'SHEET_WMS_WF_RUN',
    headers:[
      'workflow_run_id', 'doc_id', 'workflow_version_id', 'run_no',
      'run_status', 'current_step_no', 'started_by_username',
      'started_by_emp_id', 'started_at', 'completed_at', 'previous_run_id',
      'return_reason', 'created_at'
    ]
  },
  {
    configKey:'SHEET_WMS_WF_TASK',
    headers:[
      'task_id', 'workflow_run_id', 'workflow_step_id', 'step_no',
      'step_name_snapshot', 'approver_emp_id', 'approver_username',
      'approver_name_snapshot', 'can_self_approve', 'is_stock_posting_step',
      'task_status', 'acted_at', 'decision'
    ]
  },
  {
    configKey:'SHEET_WMS_WF_LOG',
    headers:[
      'action_id', 'workflow_run_id', 'task_id', 'doc_id', 'action_type',
      'comment', 'actor_username', 'actor_emp_id', 'actor_name_snapshot',
      'action_at', 'from_status', 'to_status'
    ]
  },
  {
    configKey:'SHEET_WMS_PROD',
    headers:[
      'production_batch_id', 'production_batch_no', 'external_order_ref',
      'status', 'started_at', 'completed_at', 'created_by_username',
      'created_by_emp_id', 'created_at', 'updated_by_username', 'updated_at',
      'is_deleted'
    ]
  },
  {
    configKey:'SHEET_WMS_GENEALOGY',
    headers:[
      'genealogy_id', 'production_batch_id', 'input_doc_id', 'input_line_id',
      'input_lot_id', 'input_quantity_base', 'output_doc_id', 'output_line_id',
      'output_lot_id', 'output_quantity_base', 'base_uom_snapshot',
      'yield_quantity_base', 'loss_quantity_base', 'created_by_username',
      'created_at'
    ]
  },
  {
    configKey:'SHEET_WMS_ATTACH',
    headers:[
      'attachment_id', 'doc_id', 'line_id', 'lot_id', 'production_batch_id',
      'file_id', 'file_name', 'mime_type', 'file_size', 'checksum',
      'document_category', 'uploaded_by_username', 'uploaded_at', 'status'
    ]
  },
  {
    configKey:'SHEET_WMS_USER_MAP',
    headers:[
      'mapping_id', 'username', 'emp_id', 'is_primary', 'status',
      'created_by', 'created_at', 'updated_by', 'updated_at'
    ]
  }
];

function wmsDataTextHeader_(header) {
  return /(?:^|_)(?:id|code|username|emp_id|uom|type|status)$/.test(header) ||
    ['doc_no', 'external_ref_no', 'internal_lot_no', 'supplier_lot_no',
      'production_batch_no'].indexOf(header) !== -1 ||
    header === 'transaction_key' || header === 'posting_key' ||
    /_code_snapshot$/.test(header) || header === 'base_uom_snapshot' ||
    header === 'barcode';
}

function wmsApplyDataTextFormats_(sheet, headers) {
  var rowCount = Math.max(sheet.getMaxRows ? sheet.getMaxRows() : sheet.getLastRow(), 1);
  headers.forEach(function(header, index) {
    if (wmsDataTextHeader_(header)) {
      sheet.getRange(1, index + 1, rowCount, 1).setNumberFormat('@');
    }
  });
}

function setupWmsDataSheets() {
  if (typeof CONFIG === 'undefined' || !CONFIG.DB_WMS_PART1_ID ||
      String(CONFIG.DB_WMS_PART1_ID).indexOf('NHẬP_ID') !== -1) {
    throw new Error('DB_WMS_PART1_ID chưa được cấu hình thành Spreadsheet ID hợp lệ.');
  }

  var spreadsheet = SpreadsheetApp.openById(CONFIG.DB_WMS_PART1_ID);
  var created = [];
  var alreadyReady = [];
  var errors = [];

  WMS_DATA_SHEET_SCHEMAS_.forEach(function(schema) {
    var sheetName = CONFIG[schema.configKey];
    if (!sheetName) {
      errors.push({ sheet:schema.configKey, message:'Chưa cấu hình tên sheet.' });
      return;
    }

    var sheet = spreadsheet.getSheetByName(sheetName);
    var isNew = !sheet;
    if (!sheet) sheet = spreadsheet.insertSheet(sheetName);

    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    if (lastRow === 0 || lastColumn === 0) {
      sheet.getRange(1, 1, 1, schema.headers.length).setValues([schema.headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, schema.headers.length)
        .setFontWeight('bold')
        .setBackground('#dbeafe');
      wmsApplyDataTextFormats_(sheet, schema.headers);
      sheet.autoResizeColumns(1, schema.headers.length);
      created.push(sheetName);
      return;
    }

    var actualHeaders = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
      .map(function(header) { return String(header || '').trim(); });
    var canAppendWorkflowColumn = schema.configKey === 'SHEET_WMS_DOC' &&
      actualHeaders.length === schema.headers.length - 1 &&
      schema.headers.slice(0, -1).every(function(header, index) {
        return actualHeaders[index] === header;
      });
    if (canAppendWorkflowColumn) {
      sheet.getRange(1, lastColumn + 1).setValue('selected_workflow_version_id');
      actualHeaders.push('selected_workflow_version_id');
    }
    var isCompatible = actualHeaders.length === schema.headers.length &&
      schema.headers.every(function(header, index) { return actualHeaders[index] === header; });
    if (!isCompatible) {
      errors.push({
        sheet:sheetName,
        message:'Header hiện tại khác schema dự kiến; sheet không bị thay đổi.'
      });
      return;
    }

    sheet.setFrozenRows(1);
    wmsApplyDataTextFormats_(sheet, schema.headers);
    alreadyReady.push(sheetName);
    if (isNew) created.push(sheetName);
  });

  return {
    success:errors.length === 0,
    message:errors.length
      ? 'Đã tạo ' + created.length + ' sheet; có ' + errors.length + ' sheet cần kiểm tra schema.'
      : 'Đã xác minh ' + (created.length + alreadyReady.length) + ' sheet WMS.',
    created:created,
    alreadyReady:alreadyReady,
    errors:errors
  };
}

function validateWmsDataSheets() {
  if (typeof CONFIG === 'undefined' || !CONFIG.DB_WMS_PART1_ID ||
      String(CONFIG.DB_WMS_PART1_ID).indexOf('NHẬP_ID') !== -1) {
    throw new Error('DB_WMS_PART1_ID chưa được cấu hình thành Spreadsheet ID hợp lệ.');
  }

  var spreadsheet = SpreadsheetApp.openById(CONFIG.DB_WMS_PART1_ID);
  var report = [];
  WMS_DATA_SHEET_SCHEMAS_.forEach(function(schema) {
    var sheetName = CONFIG[schema.configKey];
    var sheet = sheetName ? spreadsheet.getSheetByName(sheetName) : null;
    if (!sheet) {
      report.push({ sheet:sheetName || schema.configKey, status:'MISSING', missing:schema.headers.slice() });
      return;
    }
    var lastColumn = sheet.getLastColumn();
    if (!lastColumn) {
      report.push({ sheet:sheetName, status:'EMPTY', missing:schema.headers.slice() });
      return;
    }
    var actualHeaders = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
      .map(function(header) { return String(header || '').trim(); });
    var missing = schema.headers.filter(function(header) {
      return actualHeaders.indexOf(header) === -1;
    });
    var unexpected = actualHeaders.filter(function(header) {
      return header && schema.headers.indexOf(header) === -1;
    });
    var misordered = schema.headers.some(function(header, index) {
      var actualIndex = actualHeaders.indexOf(header);
      return actualIndex >= 0 && actualIndex !== index;
    });
    var duplicates = actualHeaders.filter(function(header, index) {
      return header && actualHeaders.indexOf(header) !== index;
    });
    report.push({
      sheet:sheetName,
      status:missing.length || unexpected.length || misordered || duplicates.length ? 'MISMATCH' : 'READY',
      missing:missing,
      unexpected:unexpected,
      misordered:misordered,
      duplicates:duplicates
    });
  });

  var readyCount = report.filter(function(item) { return item.status === 'READY'; }).length;
  return {
    success:readyCount === report.length,
    readyCount:readyCount,
    totalCount:report.length,
    report:report
  };
}
