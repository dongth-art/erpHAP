/**
 * Ghi sổ kho (posting) cho phiếu nhập/xuất đã duyệt xong bước ghi sổ.
 * Hợp đồng idempotency: xem ARCHITECTURE.md §14.4.
 * Không bao giờ sửa hoặc xóa dòng WMS_LEDGER đã ghi.
 */

var WMS_POST_HEADERS_ = [
  'posting_id', 'posting_key', 'doc_id', 'workflow_run_id',
  'workflow_step_id', 'posting_type', 'status', 'payload_hash',
  'ledger_row_count', 'attempt_count', 'started_at', 'ledger_written_at',
  'balance_written_at', 'completed_at', 'last_error', 'created_by',
  'updated_at'
];

var WMS_LEDGER_HEADERS_ = [
  'ledger_id', 'transaction_key', 'doc_id', 'line_id', 'transaction_type',
  'posted_at', 'sku_id', 'warehouse_id', 'location_id', 'lot_id',
  'stock_status', 'quantity_delta_base', 'base_uom_snapshot', 'unit_cost',
  'value_delta', 'production_batch_id', 'reversal_of_ledger_id',
  'performed_by_username', 'performed_by_emp_id',
  'performed_by_name_snapshot', 'created_at'
];

var WMS_BAL_HEADERS_ = [
  'balance_id', 'sku_id', 'warehouse_id', 'location_id', 'lot_id',
  'stock_status', 'quantity_on_hand_base', 'inventory_value',
  'moving_average_cost', 'last_ledger_id', 'last_updated_at'
];

/** Nguồn dữ liệu cần cho ghi sổ. Tái dùng hàm đọc của module workflow/doc. */
function wmsPostSheets_() {
  var spreadsheet = SpreadsheetApp.openById(CONFIG.DB_WMS_PART1_ID);
  return {
    spreadsheet:spreadsheet,
    docs:wmsDocGetSheet_(spreadsheet),
    lines:wmsGetSheet_('SHEET_WMS_LINE', WMS_LINE_HEADERS_, spreadsheet),
    post:wmsGetSheet_('SHEET_WMS_POST', WMS_POST_HEADERS_, spreadsheet),
    ledger:wmsGetSheet_('SHEET_WMS_LEDGER', WMS_LEDGER_HEADERS_, spreadsheet),
    bal:wmsGetSheet_('SHEET_WMS_BAL', WMS_BAL_HEADERS_, spreadsheet),
    run:wmsGetSheet_('SHEET_WMS_WF_RUN', WMS_WORKFLOW_HEADERS_.RUN, spreadsheet),
    task:wmsGetSheet_('SHEET_WMS_WF_TASK', WMS_WORKFLOW_HEADERS_.TASK, spreadsheet),
    log:wmsGetSheet_('SHEET_WMS_WF_LOG', WMS_WORKFLOW_HEADERS_.LOG, spreadsheet)
  };
}

function wmsPostNewId_(prefix) {
  return prefix + '-' + Utilities.getUuid().replace(/-/g, '').substring(0, 16).toUpperCase();
}

function wmsPostNumber_(value) {
  var number = Number(value == null ? 0 : value);
  return isFinite(number) ? number : 0;
}

/** Làm tròn 6 chữ số thập phân để tránh sai số dấu phẩy động khi cộng dồn tồn. */
function wmsPostRound_(value) {
  return Math.round(wmsPostNumber_(value) * 1000000) / 1000000;
}

function wmsPostBalanceKey_(skuId, warehouseId, locationId, lotId, stockStatus) {
  return [skuId, warehouseId, locationId, lotId, stockStatus].map(function(part) {
    return String(part == null ? '' : part).trim();
  }).join('|');
}

/** canonicalJson + SHA-256 (hex) để phát hiện dữ liệu phiếu đổi giữa hai lần ghi sổ. */
function wmsPostPayloadHash_(doc, lines) {
  var canonical = [
    String(doc.doc_id || ''),
    String(doc.doc_type || ''),
    String(doc.workflow_run_id || ''),
    lines.slice().sort(function(left, right) {
      return wmsPostNumber_(left.line_no) - wmsPostNumber_(right.line_no);
    }).map(function(line) {
      return [
        line.line_id, line.sku_id, line.warehouse_id, line.location_id,
        line.lot_id, wmsPostNumber_(line.quantity_base),
        wmsPostNumber_(line.unit_cost)
      ].join('~');
    }).join('||')
  ].join('|');
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, canonical,
    Utilities.Charset.UTF_8);
  return bytes.map(function(byte) {
    var hex = (byte < 0 ? byte + 256 : byte).toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  }).join('');
}

function wmsPostPostingType_(docType) {
  var type = String(docType || '').trim().toUpperCase();
  return type === WMS_DOC_TYPES_.SALES_ISSUE ? WMS_DOC_TYPES_.SALES_ISSUE : WMS_DOC_TYPES_.PURCHASE_RECEIPT;
}

/** Dòng ghi sổ hợp lệ: chưa xóa, có SKU + kho + số lượng nền dương. */
function wmsPostReadLines_(source, docId) {
  var lastRow = source.sheet.getLastRow();
  if (lastRow < 2) return [];
  var rows = source.sheet.getRange(2, 1, lastRow - 1, source.headers.length).getValues();
  var items = wmsRowsAsObjects_(source.headers, rows).filter(function(line) {
    return String(line.doc_id || '') === String(docId) && !wmsIsDeleted_(line.is_deleted);
  });
  items.forEach(function(line) {
    line.quantity_base = wmsPostNumber_(line.quantity_base);
    line.unit_cost = wmsPostNumber_(line.unit_cost);
    line.line_no = line.line_no === '' ? 0 : line.line_no;
  });
  items.sort(function(left, right) {
    return wmsPostNumber_(left.line_no) - wmsPostNumber_(right.line_no);
  });
  return items;
}

function wmsPostLineError_(line) {
  if (!String(line.sku_id || '').trim()) return 'thiếu mã SKU';
  if (!String(line.warehouse_id || '').trim()) return 'thiếu kho';
  if (!(line.quantity_base > 0)) return 'số lượng quy đổi phải lớn hơn 0';
  return '';
}

/** Ghi một dòng WMS_POST (tạo mới hoặc cập nhật theo index đã có). */
function wmsPostWritePostRow_(source, index, values) {
  var row = wmsRowForWrite_(source.headers, values);
  if (index == null || index < 0) {
    source.sheet.appendRow(row);
    return source.sheet.getLastRow();
  }
  source.sheet.getRange(index + 2, 1, 1, source.headers.length).setValues([row]);
  return index + 2;
}

/**
 * Lõi ghi sổ. Không tự ghi WMS_DOC — trả về `docPatch` để caller áp dụng
 * (nhờ vậy wmsWorkflowDecide chỉ ghi WMS_DOC đúng một lần).
 */
function wmsPostImpl_(sheets, doc, options) {
  var settings = options || {};
  var docId = String(doc.doc_id || '');
  var postingType = wmsPostPostingType_(doc.doc_type);
  var runId = String(doc.workflow_run_id || '');

  if ([WMS_DOC_TYPES_.PURCHASE_RECEIPT, WMS_DOC_TYPES_.SALES_ISSUE].indexOf(postingType) < 0) {
    return { success:false, code:'DOC_TYPE_UNSUPPORTED',
      message:'Loại phiếu ' + doc.doc_type + ' chưa hỗ trợ ghi sổ.' };
  }

  var lines = wmsPostReadLines_(sheets.lines, docId);
  if (!lines.length) {
    return { success:false, code:'NO_LINES', message:'Phiếu chưa có dòng hàng để ghi sổ.' };
  }
  var invalid = [];
  lines.forEach(function(line) {
    var reason = wmsPostLineError_(line);
    if (reason) invalid.push('dòng ' + (line.line_no || '?') + ': ' + reason);
  });
  if (invalid.length) {
    return { success:false, code:'LINE_NOT_POSTABLE',
      message:'Không ghi sổ được vì ' + invalid.join('; ') + '.' };
  }

  var stepId = String(settings.workflowStepId || '');
  if (!stepId) {
    var postingStep = (settings.versionSteps || []).filter(function(step) {
      return wmsBoolean_(step.is_stock_posting_step);
    })[0];
    stepId = postingStep ? String(postingStep.workflow_step_id || '') : '';
  }
  if (!stepId) {
    return { success:false, code:'RUN_INVALID',
      message:'Workflow chưa cấu hình bước ghi sổ kho.' };
  }
  if (!runId) {
    return { success:false, code:'RUN_INVALID', message:'Phiếu chưa có lần chạy workflow.' };
  }

  var postingKey = wmsPostingKey_(docId, runId, stepId, postingType);
  var payloadHash = wmsPostPayloadHash_(doc, lines);
  var now = new Date();
  var actor = settings.actor || {};

  // --- WMS_POST: tìm hoặc tạo dòng theo posting_key (dùng TextFinder) ---
  var postHeaderIndex = {};
  sheets.post.headers.forEach(function(h, idx) { postHeaderIndex[h] = idx + 1; });
  var postKeyCol = postHeaderIndex['posting_key'] || 2;
  var postFinder = sheets.post.sheet.getRange(2, postKeyCol, Math.max(0, sheets.post.sheet.getLastRow() - 1), 1).createTextFinder(postingKey);
  postFinder.matchEntireCell(true);
  var postMatch = postFinder.findNext();
  var postIndex = -1;
  var existing = null;
  if (postMatch) {
    var postRowNum = postMatch.getRow();
    var postRange = sheets.post.sheet.getRange(postRowNum, 1, 1, sheets.post.headers.length);
    var postVals = postRange.getValues()[0];
    postIndex = postRowNum - 2;
    existing = {};
    for (var ph = 0; ph < sheets.post.headers.length; ph += 1) {
      existing[sheets.post.headers[ph]] = postVals[ph];
    }
  }

  if (existing && String(existing.status || '') === WMS_POSTING_PHASES_.COMPLETED) {
    return {
      success:true,
      alreadyPosted:true,
      postingId:existing.posting_id,
      postingKey:postingKey,
      ledgerRowCount:wmsPostNumber_(existing.ledger_row_count),
      postingStatus:WMS_POSTING_STATUS_.POSTED,
      message:'Phiếu đã được ghi sổ trước đó.',
      docPatch:{}
    };
  }

  var attemptCount = wmsPostNumber_(existing && existing.attempt_count) + 1;
  var postingId = existing && existing.posting_id ? existing.posting_id : wmsPostNewId_('WMP');
  var postValues = {
    posting_id:postingId,
    posting_key:postingKey,
    doc_id:docId,
    workflow_run_id:runId,
    workflow_step_id:stepId,
    posting_type:postingType,
    status:WMS_POSTING_PHASES_.STARTED,
    payload_hash:payloadHash,
    ledger_row_count:0,
    attempt_count:attemptCount,
    started_at:wmsPostNumber_(existing && existing.attempt_count) > 0 && existing.started_at
      ? existing.started_at : now,
    ledger_written_at:existing ? existing.ledger_written_at : '',
    balance_written_at:existing ? existing.balance_written_at : '',
    completed_at:'',
    last_error:'',
    created_by:existing ? existing.created_by : (actor.username || ''),
    updated_at:now
  };
  var postRowIndex = wmsPostWritePostRow_(sheets.post, postIndex, postValues);

  var failPost = function(code, message) {
    postValues.status = WMS_POSTING_PHASES_.FAILED;
    postValues.last_error = String(message || '').substring(0, 900);
    postValues.updated_at = new Date();
    wmsPostWritePostRow_(sheets.post, postRowIndex - 2, postValues);
    return {
      success:false,
      code:code || 'POST_FAILED',
      message:message,
      postingId:postingId,
      postingKey:postingKey,
      postingStatus:WMS_POSTING_STATUS_.POST_FAILED,
      docPatch:{
        posting_status:WMS_POSTING_STATUS_.POST_FAILED,
        updated_by:actor.username || '',
        updated_at:new Date()
      }
    };
  };

  try {
    // --- Ledger: bỏ qua transaction_key đã tồn tại để chạy lại không nhân đôi tồn (TextFinder theo cột transaction_key) ---
    var ledgerHeaderIndex = {};
    sheets.ledger.headers.forEach(function(h, idx) { ledgerHeaderIndex[h] = idx + 1; });
    var ledgerTxCol = ledgerHeaderIndex['transaction_key'] || 2;
    var ledgerLastRow = sheets.ledger.sheet.getLastRow();
    var knownKeys = {};
    if (ledgerLastRow >= 2) {
      var ledgerTxRange = sheets.ledger.sheet.getRange(2, ledgerTxCol, ledgerLastRow - 1, 1);
      var ledgerTxValues = ledgerTxRange.getValues();
      for (var l = 0; l < ledgerTxValues.length; l += 1) {
        var ktx = ledgerTxValues[l][0];
        if (ktx !== '' && ktx !== null) {
          knownKeys[String(ktx)] = true;
        }
      }
    }

    var isIssue = postingType === WMS_DOC_TYPES_.SALES_ISSUE;
    var newLedgerRows = [];
    var deltas = [];
    var deltaIndex = {};

    lines.forEach(function(line) {
      var sign = isIssue ? -1 : 1;
      var quantityDelta = wmsPostRound_(sign * line.quantity_base);
      var valueDelta = wmsPostRound_(sign * line.quantity_base * line.unit_cost);
      var stockStatus = String((isIssue ? line.stock_status_from : line.stock_status_to) ||
        WMS_STOCK_STATUS_.AVAILABLE).trim() || WMS_STOCK_STATUS_.AVAILABLE;
      var transactionKey = wmsLedgerTransactionKey_(postingKey, line.line_id, 'MAIN');
      var ledgerId = wmsPostNewId_('WML');

      if (!knownKeys[transactionKey]) {
        knownKeys[transactionKey] = true;
        newLedgerRows.push(wmsRowForWrite_(sheets.ledger.headers, {
          ledger_id:ledgerId,
          transaction_key:transactionKey,
          doc_id:docId,
          line_id:line.line_id,
          transaction_type:postingType,
          posted_at:now,
          sku_id:line.sku_id,
          warehouse_id:line.warehouse_id,
          location_id:line.location_id,
          lot_id:line.lot_id,
          stock_status:stockStatus,
          quantity_delta_base:quantityDelta,
          base_uom_snapshot:line.base_uom_snapshot || '',
          unit_cost:line.unit_cost,
          value_delta:valueDelta,
          production_batch_id:line.production_batch_id || '',
          reversal_of_ledger_id:'',
          performed_by_username:actor.username || '',
          performed_by_emp_id:actor.emp_id || '',
          performed_by_name_snapshot:actor.fullName || actor.username || '',
          created_at:now
        }));
      }

      var balanceKey = wmsPostBalanceKey_(line.sku_id, line.warehouse_id, line.location_id,
        line.lot_id, stockStatus);
      if (!deltaIndex[balanceKey]) {
        deltaIndex[balanceKey] = {
          key:balanceKey,
          sku_id:line.sku_id,
          warehouse_id:line.warehouse_id,
          location_id:line.location_id,
          lot_id:line.lot_id,
          stock_status:stockStatus,
          quantity_delta:0,
          value_delta:0,
          unit_cost:line.unit_cost,
          last_ledger_id:ledgerId
        };
        deltas.push(deltaIndex[balanceKey]);
      }
      deltaIndex[balanceKey].quantity_delta =
        wmsPostRound_(deltaIndex[balanceKey].quantity_delta + quantityDelta);
      deltaIndex[balanceKey].value_delta =
        wmsPostRound_(deltaIndex[balanceKey].value_delta + valueDelta);
      deltaIndex[balanceKey].unit_cost = line.unit_cost;
      deltaIndex[balanceKey].last_ledger_id = ledgerId;
    });

    if (newLedgerRows.length) {
      sheets.ledger.sheet.getRange(sheets.ledger.sheet.getLastRow() + 1, 1,
        newLedgerRows.length, sheets.ledger.headers.length).setValues(newLedgerRows);
    }
    postValues.status = WMS_POSTING_PHASES_.LEDGER_WRITTEN;
    postValues.ledger_written_at = new Date();
    postValues.ledger_row_count = wmsPostNumber_(existing && existing.ledger_row_count) +
      newLedgerRows.length;
    postValues.updated_at = new Date();
    wmsPostWritePostRow_(sheets.post, postRowIndex - 2, postValues);

    // --- WMS_BAL: cộng dồn theo khóa, kiểm tra tồn âm trước khi ghi (load bal với TextFinder/map theo key) ---
    var balHeaderIndex = {};
    sheets.bal.headers.forEach(function(h, idx) { balHeaderIndex[h] = idx + 1; });
    var balLastRow = sheets.bal.sheet.getLastRow();
    var balRowsObjs = [];
    if (balLastRow >= 2) {
      var balValsAll = sheets.bal.sheet.getRange(2, 1, balLastRow - 1, sheets.bal.headers.length).getValues();
      for (var br = 0; br < balValsAll.length; br += 1) {
        var obj = {};
        for (var bh = 0; bh < sheets.bal.headers.length; bh += 1) {
          obj[sheets.bal.headers[bh]] = balValsAll[br][bh];
        }
        balRowsObjs.push(obj);
      }
    }
    var balIndex = {};
    balRowsObjs.forEach(function(row, index) {
      balIndex[wmsPostBalanceKey_(row.sku_id, row.warehouse_id, row.location_id,
        row.lot_id, row.stock_status)] = { row:row, index:index };
    });

    var shortages = [];
    deltas.forEach(function(delta) {
      var current = balIndex[delta.key];
      var available = current ? wmsPostNumber_(current.row.quantity_on_hand_base) : 0;
      if (isIssue && wmsPostRound_(available + delta.quantity_delta) < 0) {
        shortages.push({
          sku_id:delta.sku_id,
          warehouse_id:delta.warehouse_id,
          location_id:delta.location_id,
          lot_id:delta.lot_id,
          stock_status:delta.stock_status,
          available:wmsPostRound_(available),
          requested:wmsPostRound_(-delta.quantity_delta)
        });
      }
    });
    if (shortages.length) {
      return failPost('NEGATIVE_STOCK', 'Không đủ tồn kho cho ' +
        shortages.map(function(item) {
          return 'SKU ' + item.sku_id + ' tại kho ' + item.warehouse_id +
            ' (còn ' + item.available + ', cần ' + item.requested + ')';
        }).join('; ') + '.');
    }

    var balUpdates = [];
    var balAppends = [];
    deltas.forEach(function(delta) {
      var current = balIndex[delta.key];
      var quantity = wmsPostRound_(wmsPostNumber_(current && current.row.quantity_on_hand_base) +
        delta.quantity_delta);
      var value = wmsPostRound_(wmsPostNumber_(current && current.row.inventory_value) +
        delta.value_delta);
      var movingAverage = current && current.row.moving_average_cost !== '' &&
        current.row.moving_average_cost != null
        ? current.row.moving_average_cost : delta.unit_cost;
      var values = {
        balance_id:current ? current.row.balance_id : wmsPostNewId_('WMB'),
        sku_id:delta.sku_id,
        warehouse_id:delta.warehouse_id,
        location_id:delta.location_id,
        lot_id:delta.lot_id,
        stock_status:delta.stock_status,
        quantity_on_hand_base:quantity,
        inventory_value:value,
        moving_average_cost:movingAverage,
        last_ledger_id:delta.last_ledger_id,
        last_updated_at:new Date()
      };
      var row = wmsRowForWrite_(sheets.bal.headers, values);
      if (current) balUpdates.push({ index:current.index, row:row });
      else balAppends.push(row);
    });
    if (balUpdates.length) {
      var balUpdateRanges = [];
      var balUpdateValues = [];
      balUpdates.forEach(function(item) {
        balUpdateRanges.push(sheets.bal.sheet.getRange(item.index + 2, 1, 1, sheets.bal.headers.length));
        balUpdateValues.push([item.row]);
      });
      for (var bu = 0; bu < balUpdateRanges.length; bu += 1) {
        balUpdateRanges[bu].setValues(balUpdateValues[bu]);
      }
    }
    if (balAppends.length) {
      sheets.bal.sheet.getRange(sheets.bal.sheet.getLastRow() + 1, 1,
        balAppends.length, sheets.bal.headers.length).setValues(balAppends);
    }
    postValues.status = WMS_POSTING_PHASES_.BALANCE_WRITTEN;
    postValues.balance_written_at = new Date();
    postValues.updated_at = new Date();
    wmsPostWritePostRow_(sheets.post, postRowIndex - 2, postValues);

    postValues.status = WMS_POSTING_PHASES_.COMPLETED;
    postValues.completed_at = new Date();
    postValues.updated_at = new Date();
    wmsPostWritePostRow_(sheets.post, postRowIndex - 2, postValues);

    return {
      success:true,
      alreadyPosted:false,
      postingId:postingId,
      postingKey:postingKey,
      ledgerRowCount:postValues.ledger_row_count,
      attemptCount:attemptCount,
      postingStatus:WMS_POSTING_STATUS_.POSTED,
      message:'Đã ghi sổ kho cho phiếu ' + (doc.doc_no || docId) + '.',
      docPatch:{
        posting_status:WMS_POSTING_STATUS_.POSTED,
        posted_at:postValues.completed_at,
        doc_status:WMS_DOC_STATUS_.COMPLETED,
        completed_at:postValues.completed_at,
        updated_by:actor.username || '',
        updated_at:postValues.completed_at
      }
    };
  } catch (err) {
    Logger.log('wmsPostImpl_: ' + err);
    return failPost('POST_FAILED', 'Ghi sổ thất bại: ' + err.message);
  }
}

/** Ghi WMS_WF_LOG cho lần ghi sổ. Lỗi ghi log không được làm hỏng giao dịch chính. */
function wmsPostWriteLog_(sheets, doc, actor, action, comment, fromStatus, toStatus) {
  try {
    var log = wmsWorkflowLogRow_(sheets.log.headers,
      String(doc.workflow_run_id || ''), '', String(doc.doc_id || ''),
      action, comment, actor, fromStatus, toStatus, new Date());
    sheets.log.sheet.appendRow(log);
  } catch (err) {
    Logger.log('wmsPostWriteLog_: ' + err);
  }
}

/** Áp `docPatch` lên mảng dòng WMS_DOC đang giữ trong bộ nhớ. */
function wmsPostApplyDocPatch_(docSource, docRow, patch) {
  if (!patch) return;
  Object.keys(patch).forEach(function(header) {
    var index = docSource.headers.indexOf(header);
    if (index >= 0) docRow[index] = patch[header];
  });
}

/**
 * Gọi từ bên trong lock của wmsWorkflowDecide.
 * Ghi WMS_POST/WMS_LEDGER/WMS_BAL và SỬA TRỰC TIẾP mảng `docRow`;
 * caller chịu trách nhiệm ghi WMS_DOC (để chỉ ghi một lần).
 */
function wmsWorkflowPostLocked_(sheets, docRow, versionSteps, actor) {
  var doc = wmsRowsAsObjects_(sheets.docs.headers, [docRow])[0];
  var result = wmsPostImpl_(sheets, doc, {
    actor:actor,
    versionSteps:versionSteps || [],
    workflowStepId:''
  });
  wmsPostApplyDocPatch_(sheets.docs, docRow, result.docPatch);
  wmsPostWriteLog_(sheets, doc, actor,
    result.success ? WMS_WORKFLOW_ACTIONS_.STOCK_POSTED : WMS_WORKFLOW_ACTIONS_.POST_RETRY,
    result.message || '', WMS_APPROVAL_STATUS_.APPROVED,
    result.success ? WMS_DOC_STATUS_.COMPLETED : WMS_DOC_STATUS_.APPROVED);
  return result;
}

/**
 * API công khai: ghi sổ (kiêm thử lại khi POST_FAILED).
 * Idempotent — gọi lại khi đã POSTED trả alreadyPosted = true.
 */
function wmsPost(token, docId) {
  try {
    var sheets = wmsPostSheets_();
    var matches = wmsDocReadByValue_(sheets.docs, 'doc_id', docId);
    if (matches.length !== 1) return fail_('Không tìm thấy phiếu cần ghi sổ.', { code:'NOT_FOUND' });
    var docType = String(matches[0].row[sheets.docs.headers.indexOf('doc_type')] || '');
    var docAuth = wmsCanRead_(token, 'DOC', wmsDocNavigationKey_(docType));
    if (!docAuth.ok) return docAuth.result;
    // Quyền trên WMS_POST là tùy chọn: chỉ dùng để nhận diện quản trị viên.
    var postAuth = wmsCanRead_(token, 'WMS_POST', wmsDocNavigationKey_(docType));
    var adminPost = postAuth.ok && (postAuth.canEdit || postAuth.canDelete);

    return withLock_(function() {
      var docMatches = wmsDocReadByValue_(sheets.docs, 'doc_id', docId);
      if (docMatches.length !== 1) return fail_('Không tìm thấy phiếu cần ghi sổ.', { code:'NOT_FOUND' });
      var docRow = docMatches[0].row;
      var doc = wmsRowsAsObjects_(sheets.docs.headers, [docRow])[0];

      if (wmsIsDeleted_(doc.is_deleted)) return fail_('Phiếu đã được đánh dấu xóa.', { code:'NOT_FOUND' });
      if (String(doc.posting_status || '') === WMS_POSTING_STATUS_.POSTED) {
        return {
          success:true,
          alreadyPosted:true,
          message:'Phiếu đã được ghi sổ trước đó.',
          posting_status:WMS_POSTING_STATUS_.POSTED
        };
      }
      if ([WMS_POSTING_STATUS_.NOT_POSTED, WMS_POSTING_STATUS_.POST_FAILED]
        .indexOf(String(doc.posting_status || '')) < 0) {
        return fail_('Phiếu không ở trạng thái có thể ghi sổ.', { code:'ALREADY_POSTED' });
      }
      if (!matrixScopeAllowed_(docAuth.user, 'DOC', docRow, null, docAuth.permission, 'update')) {
        return fail_('Phiếu nằm ngoài phạm vi xử lý được phân quyền.', { code:'FORBIDDEN' });
      }

      // Bước ghi sổ của phiên bản workflow đang chạy trên phiếu.
      var runId = String(doc.workflow_run_id || '');
      if (!runId) return fail_('Phiếu chưa được gửi duyệt nên không thể ghi sổ.', { code:'NOT_APPROVED' });
      var runMatches = wmsWorkflowReadRows_(sheets.run).map(function(row, index) {
        return { row:row, index:index };
      }).filter(function(item) {
        return String(item.row[sheets.run.headers.indexOf('workflow_run_id')] || '') === runId;
      });
      if (runMatches.length !== 1) return fail_('Không tìm thấy lần chạy workflow của phiếu.', { code:'RUN_INVALID' });
      var runRow = runMatches[0].row;
      var versionId = String(runRow[sheets.run.headers.indexOf('workflow_version_id')] || '');
      var versionSteps = wmsWorkflowReadRows_(sheets.step).filter(function(row) {
        return String(row[sheets.step.headers.indexOf('workflow_version_id')] || '') === versionId;
      }).map(function(row) {
        return wmsRowsAsObjects_(sheets.step.headers, [row])[0];
      });
      var postingStep = versionSteps.filter(function(step) {
        return wmsBoolean_(step.is_stock_posting_step);
      })[0];
      if (!postingStep) return fail_('Workflow của phiếu không có bước ghi sổ kho.', { code:'RUN_INVALID' });

      // Điều kiện #4: run COMPLETED, hoặc run PENDING với task ghi sổ đã APPROVED.
      var runStatus = String(runRow[sheets.run.headers.indexOf('run_status')] || '');
      var postingStepTask = wmsWorkflowReadRows_(sheets.task).map(function(row) {
        return wmsRowsAsObjects_(sheets.task.headers, [row])[0];
      }).filter(function(task) {
        return String(task.workflow_run_id || '') === runId &&
          String(task.workflow_step_id || '') === String(postingStep.workflow_step_id || '');
      })[0];
      var postingStepApproved = postingStepTask &&
        String(postingStepTask.task_status || '') === WMS_WORKFLOW_TASK_STATUS_.APPROVED;
      if (runStatus !== WMS_WORKFLOW_RUN_STATUS_.COMPLETED && !postingStepApproved) {
        return fail_('Bước ghi sổ chưa được duyệt nên chưa thể ghi sổ kho.', { code:'NOT_APPROVED' });
      }

      var actor = Object.assign({}, docAuth.user, {
        emp_id:docAuth.user.emp_id || '',
        fullName:docAuth.user.fullName || docAuth.user.username
      });
      var isPostingApprover = postingStepTask &&
        String(postingStepTask.approver_username || '').toLowerCase() ===
          String(docAuth.user.username || '').toLowerCase();
      if (!isPostingApprover && !adminPost) {
        return fail_('Chỉ người duyệt bước ghi sổ hoặc quản trị viên được ghi sổ kho.', { code:'FORBIDDEN' });
      }

      var result = wmsPostImpl_(sheets, doc, {
        actor:actor,
        versionSteps:versionSteps,
        workflowStepId:String(postingStep.workflow_step_id || '')
      });
      wmsPostApplyDocPatch_(sheets.docs, docRow, result.docPatch);
      if (result.docPatch && Object.keys(result.docPatch).length) {
        sheets.docs.sheet.getRange(docMatches[0].index + 2, 1, 1, sheets.docs.headers.length)
          .setValues([docRow]);
      }
      wmsPostWriteLog_(sheets, doc, actor,
        result.success ? WMS_WORKFLOW_ACTIONS_.STOCK_POSTED : WMS_WORKFLOW_ACTIONS_.POST_RETRY,
        result.message || '', WMS_APPROVAL_STATUS_.APPROVED,
        result.success ? WMS_DOC_STATUS_.COMPLETED : WMS_DOC_STATUS_.APPROVED);
      if (result.success) wmsAudit_(docAuth.user, 'WMS_POST', 'POST', docId, result.postingKey || '');
      return Object.assign({ list:[] }, result);
    });
  } catch (err) {
    Logger.log('wmsPost: ' + err);
    return { success:false, message:'Không ghi sổ được phiếu: ' + err.message, list:[] };
  }
}
