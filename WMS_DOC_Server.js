/**
 * Draft warehouse documents. These APIs never write stock ledger or balances.
 */

var WMS_DOC_HEADERS_ = [
  'doc_id', 'doc_no', 'doc_type', 'doc_date', 'warehouse_id',
  'destination_warehouse_id', 'partner_id', 'partner_code_snapshot',
  'partner_name_snapshot', 'external_ref_type', 'external_ref_id',
  'external_ref_no', 'requested_by_username', 'requested_by_emp_id',
  'requested_by_name_snapshot', 'workflow_run_id', 'approval_status',
  'posting_status', 'doc_status', 'submitted_at', 'posted_at', 'completed_at',
  'reason', 'notes', 'reversal_of_doc_id', 'created_by', 'created_at',
  'updated_by', 'updated_at', 'is_deleted', 'selected_workflow_version_id'
];

var WMS_LINE_HEADERS_ = [
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
];

var WMS_DOC_VIEW_IN_ = 'view-wms-in';
var WMS_DOC_VIEW_OUT_ = 'view-wms-out';
var WMS_DOC_TYPES_ALLOWED_ = [WMS_DOC_TYPES_.PURCHASE_RECEIPT, WMS_DOC_TYPES_.SALES_ISSUE];

function wmsDocNavigationKey_(docType) {
  return docType === WMS_DOC_TYPES_.PURCHASE_RECEIPT ? WMS_DOC_VIEW_IN_ : WMS_DOC_VIEW_OUT_;
}

function wmsDocTypeValid_(docType) {
  return WMS_DOC_TYPES_ALLOWED_.indexOf(String(docType || '').trim().toUpperCase()) !== -1;
}

function wmsDocGetSheet_(spreadsheet) {
  var targetSpreadsheet = spreadsheet || SpreadsheetApp.openById(CONFIG.DB_WMS_PART1_ID);
  var legacyHeaders = WMS_DOC_HEADERS_.filter(function(header) {
    return header !== 'selected_workflow_version_id';
  });
  var source = wmsGetSheet_('SHEET_WMS_DOC', legacyHeaders, targetSpreadsheet);
  if (source.headers.indexOf('selected_workflow_version_id') < 0) {
    withLock_(function() {
      var current = wmsGetSheet_('SHEET_WMS_DOC', legacyHeaders, targetSpreadsheet);
      if (current.headers.indexOf('selected_workflow_version_id') < 0) {
        current.sheet.getRange(1, current.headers.length + 1).setValue('selected_workflow_version_id');
      }
    });
    source = wmsGetSheet_('SHEET_WMS_DOC', WMS_DOC_HEADERS_, targetSpreadsheet);
  }
  return source;
}

function wmsDocScopeError_(message) {
  var error = new Error(message);
  error.code = 'FORBIDDEN';
  return error;
}

function wmsDocLoadWorkflowOptions_(token, docType, docAuth) {
  var result = wmsWorkflowListAvailable_(
    token, docType, '', wmsDocNavigationKey_(docType), docAuth);
  if (result && result.success) {
    return { success:true, workflows:result.workflows || [], workflowAccess:true };
  }
  if (result && result.code === 'FORBIDDEN') {
    return { success:true, workflows:[], workflowAccess:false };
  }
  return result || { success:false, message:'Không tải được danh sách workflow.' };
}

function wmsDocAuth_(token, docType, action) {
  docType = String(docType || '').trim().toUpperCase();
  if (!wmsDocTypeValid_(docType)) {
    return { ok:false, result:fail_('Loại phiếu nhập/xuất không hợp lệ.') };
  }
  var docAuth = wmsCanRead_(token, 'DOC', wmsDocNavigationKey_(docType));
  if (!docAuth.ok) return docAuth;
  var lineAuth = null;
  if (action === 'create' || action === 'update' || action === 'delete') {
    var lineAuthResult = wmsCanRead_(token, 'LINE', wmsDocNavigationKey_(docType));
    if (!lineAuthResult.ok) return lineAuthResult;
    lineAuth = lineAuthResult;
  } else if (action === 'read') {
    lineAuth = docAuth;
  }
  var allowed = action === 'create' ? docAuth.canCreate && lineAuth.canCreate
    : action === 'update' ? docAuth.canEdit && lineAuth.canEdit
    : action === 'delete' ? docAuth.canDelete && lineAuth.canDelete : docAuth.canRead;
  if (!allowed) {
    return {
      ok:false,
      result:fail_('Bạn không có quyền ' +
        (action === 'create' ? 'tạo' : action === 'update' ? 'sửa' : 'xóa') +
        ' phiếu nhập/xuất.', { code:'FORBIDDEN' })
    };
  }
  return { ok:true, docAuth:docAuth, lineAuth:lineAuth };
}

function wmsDocLoadRows_(source) {
  var lastRow = source.sheet.getLastRow();
  return lastRow > 1
    ? source.sheet.getRange(2, 1, lastRow - 1, source.headers.length).getValues()
    : [];
}

function wmsDocReadByValue_(source, keyHeader, expectedValue) {
  return wmsDocReadByValues_(source, [keyHeader], [expectedValue]);
}

function wmsDocReadByDocIdFast_(source, docId) {
  var docIdIndex = source.headers.indexOf('doc_id');
  if (docIdIndex < 0) throw new Error('Sheet thiếu cột doc_id.');
  var lastRow = source.sheet.getLastRow();
  if (lastRow < 2) return [];
  var textFinder = source.sheet.getRange(2, docIdIndex + 1, lastRow - 1, 1).createTextFinder(String(docId || ''))
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

function wmsDocReadByValues_(source, keyHeaders, expectedValues) {
  var keyIndexes = keyHeaders.map(function(header) {
    var index = source.headers.indexOf(header);
    if (index < 0) throw new Error('Sheet thiếu cột ' + header + '.');
    return index;
  });
  var expected = {};
  expectedValues.forEach(function(value) {
    var key = String(value == null ? '' : value).replace(/^'/, '').trim().toLowerCase();
    if (key) expected[key] = true;
  });
  var lastRow = source.sheet.getLastRow();
  if (lastRow < 2) return [];
  var matches = {};
  keyIndexes.forEach(function(keyIndex) {
    source.sheet.getRange(2, keyIndex + 1, lastRow - 1, 1).getValues().forEach(function(cell, index) {
      var value = String(cell[0] == null ? '' : cell[0]).replace(/^'/, '').trim().toLowerCase();
      if (expected[value]) matches[index] = true;
    });
  });
  var rowIndexes = Object.keys(matches).map(Number).sort(function(left, right) {
    return left - right;
  });
  var result = [];
  var start = 0;
  while (start < rowIndexes.length) {
    var end = start + 1;
    while (end < rowIndexes.length && rowIndexes[end] === rowIndexes[end - 1] + 1) end++;
    var first = rowIndexes[start];
    var count = end - start;
    var rows = source.sheet.getRange(first + 2, 1, count, source.headers.length).getValues();
    rows.forEach(function(row, offset) {
      result.push({ index:first + offset, row:row });
    });
    start = end;
  }
  return result;
}

function wmsDocLoadReferenceSubset_(auth, warehouseId, rawLines) {
  var warehouseSource = wmsGetSheet_('SHEET_WMS_WAREHOUSE', WMS_WAREHOUSE_HEADERS_);
  var skuSource = wmsGetSkuSheet_();
  var locationSource = wmsGetSheet_('SHEET_WMS_LOCATION', WMS_LOCATION_HEADERS_);
  var skuLookups = [];
  var locationLookups = [];
  (rawLines || []).forEach(function(line) {
    if (line && (line.sku_id || line.sku_code)) skuLookups.push(line.sku_id || line.sku_code);
    if (line && line.location_id) locationLookups.push(line.location_id);
  });
  var warehouseRows = wmsDocReadByValue_(warehouseSource, 'warehouse_id', warehouseId);
  var skuRows = wmsDocReadByValues_(skuSource, ['sku_id', 'short_name', 'barcode'], skuLookups);
  var locationRows = wmsDocReadByValues_(locationSource, ['location_id'], locationLookups);
  var warehousePermission = getEffectivePermission_(auth.user.username, auth.user.role, 'WAREHOUSE');
  var skuPermission = getEffectivePermission_(auth.user.username, auth.user.role, 'SKU');
  var locationPermission = getEffectivePermission_(auth.user.username, auth.user.role, 'LOCATION');
  var warehouseObjects = warehouseRows.filter(function(entry) {
    var row = entry.row;
    return !wmsIsDeleted_(row[warehouseSource.headers.indexOf('is_deleted')]) &&
      wmsBoolean_(row[warehouseSource.headers.indexOf('is_active')]) &&
      matrixRowAllowed_(auth.user, 'WAREHOUSE', row, null, warehousePermission);
  }).map(function(entry) { return wmsRowsAsObjects_(warehouseSource.headers, [entry.row])[0]; });
  var skuObjects = skuRows.filter(function(entry) {
    var row = entry.row;
    return !wmsIsDeleted_(row[skuSource.headers.indexOf('is_deleted')]) &&
      String(row[skuSource.headers.indexOf('status')] || '').toUpperCase() === 'ACTIVE' &&
      matrixRowAllowed_(auth.user, 'SKU', row, null, skuPermission);
  }).map(function(entry) { return wmsRowsAsObjects_(skuSource.headers, [entry.row])[0]; });
  var locationObjects = locationRows.filter(function(entry) {
    var row = entry.row;
    return !wmsIsDeleted_(row[locationSource.headers.indexOf('is_deleted')]) &&
      String(row[locationSource.headers.indexOf('status')] || '').toUpperCase() === 'ACTIVE' &&
      matrixRowAllowed_(auth.user, 'LOCATION', row, null, locationPermission);
  }).map(function(entry) { return wmsRowsAsObjects_(locationSource.headers, [entry.row])[0]; });
  return { warehouses:warehouseObjects, skus:skuObjects, locations:locationObjects };
}

function wmsDocNextNumberFromValues_(values, docType, dateValue) {
  var prefix = docType === WMS_DOC_TYPES_.PURCHASE_RECEIPT ? 'NK' : 'XK';
  var base = prefix + '-' + String(dateValue || '').replace(/-/g, '');
  var max = 0;
  var expression = new RegExp('^' + base + '-(\\d{4,})$');
  values.forEach(function(value) {
    var match = String(value || '').replace(/^'/, '').trim().match(expression);
    if (match) max = Math.max(max, Number(match[1]));
  });
  return base + '-' + String(max + 1).padStart(4, '0');
}

function wmsDocReadColumn_(source, header) {
  var index = source.headers.indexOf(header);
  if (index < 0) throw new Error('Sheet thiếu cột ' + header + '.');
  var lastRow = source.sheet.getLastRow();
  return lastRow < 2 ? [] : source.sheet.getRange(2, index + 1, lastRow - 1, 1)
    .getValues().map(function(row) { return row[0]; });
}

function wmsDocFindHeader_(headers, record, docId) {
  var idIndex = headers.indexOf('doc_id');
  for (var i = 0; i < record.length; i++) {
    if (String(record[i][idIndex] || '').trim() === String(docId)) return i;
  }
  return -1;
}

function wmsDocNextNumber_(rows, headers, docType, dateValue) {
  var prefix = docType === WMS_DOC_TYPES_.PURCHASE_RECEIPT ? 'NK' : 'XK';
  var date = String(dateValue || '').replace(/-/g, '');
  var base = prefix + '-' + date;
  var numberIndex = headers.indexOf('doc_no');
  var max = 0;
  var expression = new RegExp('^' + base + '-(\\d{4,})$');
  rows.forEach(function(row) {
    var match = String(row[numberIndex] || '').replace(/^'/, '').trim().match(expression);
    if (match) max = Math.max(max, Number(match[1]));
  });
  return base + '-' + String(max + 1).padStart(4, '0');
}

function wmsDocLoadReferences_(auth) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'WMS_REF_DATA_' + auth.user.username;
  var cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  var warehouseSource = wmsGetSheet_('SHEET_WMS_WAREHOUSE', WMS_WAREHOUSE_HEADERS_);
  var skuSource = wmsGetSkuSheet_();
  var locationsSource = wmsGetSheet_('SHEET_WMS_LOCATION', WMS_LOCATION_HEADERS_);

  var warehouseRows = wmsDocLoadRows_(warehouseSource);
  var skuRows = wmsDocLoadRows_(skuSource);
  var locationRows = wmsDocLoadRows_(locationsSource);

  var warehousePermission = getEffectivePermission_(auth.user.username, auth.user.role, 'WAREHOUSE');
  var skuPermission = getEffectivePermission_(auth.user.username, auth.user.role, 'SKU');
  var locationPermission = getEffectivePermission_(auth.user.username, auth.user.role, 'LOCATION');

  var warehouseDeleted = warehouseSource.headers.indexOf('is_deleted');
  var skuDeleted = skuSource.headers.indexOf('is_deleted');
  var skuStatus = skuSource.headers.indexOf('status');
  var locationDeleted = locationsSource.headers.indexOf('is_deleted');

  var warehouses = wmsRowsAsObjects_(warehouseSource.headers, warehouseRows.filter(function(row) {
    return !wmsIsDeleted_(row[warehouseDeleted]) &&
      wmsBoolean_(row[warehouseSource.headers.indexOf('is_active')]) &&
      matrixRowAllowed_(auth.user, 'WAREHOUSE', row, null, warehousePermission);
  })).map(function(item) {
    return { warehouse_id: item.warehouse_id, warehouse_code: item.warehouse_code, warehouse_name: item.warehouse_name };
  });

  var skus = wmsRowsAsObjects_(skuSource.headers, skuRows.filter(function(row) {
    return !wmsIsDeleted_(row[skuDeleted]) &&
      String(row[skuStatus] || '').toUpperCase() === 'ACTIVE' &&
      matrixRowAllowed_(auth.user, 'SKU', row, null, skuPermission);
  })).map(function(item) {
    return { sku_id: item.sku_id, short_name: item.short_name, sku_name: item.sku_name, barcode: item.barcode, uom_base: item.uom_base };
  });

  var locations = wmsRowsAsObjects_(locationsSource.headers, locationRows.filter(function(row) {
    return !wmsIsDeleted_(row[locationDeleted]) &&
      String(row[locationsSource.headers.indexOf('status')] || '').toUpperCase() === 'ACTIVE' &&
      matrixRowAllowed_(auth.user, 'LOCATION', row, null, locationPermission);
  })).map(function(item) {
    return { location_id: item.location_id, warehouse_id: item.warehouse_id, location_code: item.location_code, location_name: item.location_name };
  });

  var res = { warehouses: warehouses, skus: skus, locations: locations };
  try { cache.put(cacheKey, JSON.stringify(res), 600); } catch(e) {}
  return res;
}

function wmsListDocOptions(token, docType) {
  var auth = wmsDocAuth_(token, docType, 'read');
  if (!auth.ok) return auth.result;
  try {
    var references = wmsDocLoadReferences_(auth.docAuth);
    var workflowOptions = wmsDocLoadWorkflowOptions_(token, docType, auth.docAuth);
    if (!workflowOptions.success) return workflowOptions;
    return {
      success:true,
      warehouses:references.warehouses,
      skus:references.skus,
      locations:references.locations,
      workflows:workflowOptions.workflows,
      workflowAccess:workflowOptions.workflowAccess
    };
  } catch (err) {
    Logger.log('wmsListDocOptions: ' + err);
    return { success:false, message:'Không tải được danh mục cho phiếu: ' + err.message };
  }
}

function wmsListDocs(token, docType, searchKey, page, pageSize, includeOptions) {
  var auth = wmsDocAuth_(token, docType, 'read');
  if (!auth.ok) return auth.result;
  try {
    var docs = wmsDocGetSheet_();
    var docRows = wmsDocLoadRows_(docs);
    var docIdIndex = docs.headers.indexOf('doc_id');
    var docTypeIndex = docs.headers.indexOf('doc_type');
    var deletedIndex = docs.headers.indexOf('is_deleted');
    var visibleDocs = docRows.filter(function(row, index) {
      if (String(row[docTypeIndex] || '').toUpperCase() !== String(docType).toUpperCase() ||
          wmsIsDeleted_(row[deletedIndex]) ||
          !matrixRowAllowed_(auth.docAuth.user, 'DOC', row, null, auth.docAuth.permission)) return false;
      return true;
    });
    var visibleDocObjects = wmsRowsAsObjects_(docs.headers, visibleDocs.reverse());
    var query = String(searchKey || '').trim().toLowerCase();
    var result = visibleDocObjects.filter(function(item) {
      return !query || [
        item.doc_no, item.doc_date, item.warehouse_id, item.partner_name_snapshot,
        item.partner_code_snapshot, item.external_ref_no, item.reason, item.notes,
        item.requested_by_name_snapshot, item.doc_status
      ].join(' ').toLowerCase().indexOf(query) !== -1;
    });

    page = Math.max(1, parseInt(page, 10) || 1);
    pageSize = Math.min(100, Math.max(1, parseInt(pageSize, 10) || 30));
    var total = result.length;
    var pages = Math.max(1, Math.ceil(total / pageSize));
    page = Math.min(page, pages);
    var pageRows = result.slice((page - 1) * pageSize, page * pageSize);

    // Lấy thông tin dòng hàng (WMS_LINE) chỉ cho các phiếu thuộc trang hiện tại
    var pageDocIds = {};
    pageRows.forEach(function(item) { pageDocIds[String(item.doc_id || '')] = true; });
    var docIdList = Object.keys(pageDocIds);

    var summaries = {};
    var detailLines = {};

    // NCC của Lô được đồng bộ sang WMS_LOT khi lưu; đọc lại để hiển thị cột "NCC Lô".
    var lotSupplierByLot = {};
    try {
      var lotSheet = docs.spreadsheet.getSheetByName(CONFIG.SHEET_WMS_LOT);
      if (lotSheet && lotSheet.getLastRow() > 1 && lotSheet.getLastColumn() > 0) {
        var lotLastColumn = lotSheet.getLastColumn();
        var lotHeaders = lotSheet.getRange(1, 1, 1, lotLastColumn).getDisplayValues()[0]
          .map(function(header) { return String(header || '').trim(); });
        var lotIdColumn = lotHeaders.indexOf('lot_id');
        var lotSupplierColumn = lotHeaders.indexOf('supplier_name_snapshot');
        if (lotIdColumn >= 0 && lotSupplierColumn >= 0) {
          var lotValues = lotSheet.getRange(2, 1, lotSheet.getLastRow() - 1, lotLastColumn).getDisplayValues();
          lotValues.forEach(function(values) {
            var lotKey = String(values[lotIdColumn] || '').trim();
            if (lotKey) lotSupplierByLot[lotKey] = values[lotSupplierColumn];
          });
        }
      }
    } catch (lotError) {
      Logger.log('wmsListDocs lot lookup: ' + lotError);
    }

    if (auth.lineAuth && docIdList.length) {
      var lines = wmsGetSheet_('SHEET_WMS_LINE', WMS_LINE_HEADERS_);
      var lineMatches = wmsDocReadByValues_(lines, ['doc_id'], docIdList);
      var lineDeleted = lines.headers.indexOf('is_deleted');
      var lineQty = lines.headers.indexOf('quantity_base');

      lineMatches.forEach(function(entry) {
        var row = entry.row;
        var id = String(row[lines.headers.indexOf('doc_id')] || '');
        if (wmsIsDeleted_(row[lineDeleted]) ||
            !matrixRowAllowed_(auth.lineAuth.user, 'LINE', row, null, auth.lineAuth.permission)) return;

        if (!summaries[id]) summaries[id] = { count: 0, quantity: 0 };
        summaries[id].count++;
        summaries[id].quantity += Number(row[lineQty]) || 0;

        if (!detailLines[id]) detailLines[id] = [];
        var lineLotId = row[lines.headers.indexOf('lot_id')];
        var lineSupplierColumn = lines.headers.indexOf('supplier_name_snapshot');
        detailLines[id].push({
          line_id: row[lines.headers.indexOf('line_id')],
          sku_id: row[lines.headers.indexOf('sku_id')],
          sku_code_snapshot: row[lines.headers.indexOf('sku_code_snapshot')],
          sku_name_snapshot: row[lines.headers.indexOf('sku_name_snapshot')],
          quantity_entered: row[lines.headers.indexOf('quantity_entered')],
          uom_entered: row[lines.headers.indexOf('uom_entered')],
          unit_cost: row[lines.headers.indexOf('unit_cost')],
          lot_id: lineLotId,
          supplier_name_snapshot: (lineSupplierColumn >= 0 ? row[lineSupplierColumn] : '') ||
            lotSupplierByLot[String(lineLotId || '').trim()] || '',
          production_batch_id: row[lines.headers.indexOf('production_batch_id')],
          location_id: row[lines.headers.indexOf('location_id')],
          destination_warehouse_id: row[lines.headers.indexOf('destination_warehouse_id')],
          destination_location_id: row[lines.headers.indexOf('destination_location_id')],
          line_notes: row[lines.headers.indexOf('line_notes')]
        });
      });
    }

    pageRows.forEach(function(item) {
      var summary = summaries[item.doc_id] || { count: 0, quantity: 0 };
      item.line_count = summary.count;
      item.quantity_total = summary.quantity;
      item.lines = detailLines[item.doc_id] || [];
    });

    // Bổ sung tiến trình duyệt động
    try { wmsDocEnrichApprovalProgress_(pageRows, docs.spreadsheet); } catch (e) { Logger.log('Enrich err: ' + e); }

    // Gộp luôn danh sách task duyệt của user hiện tại (tránh gọi API nối tiếp ở Frontend)
    var taskMap = {};
    try {
      var myTasksRes = wmsWorkflowMyTasks(token);
      if (myTasksRes && myTasksRes.success && myTasksRes.list) {
        myTasksRes.list.forEach(function(t) { if (t.doc_id) taskMap[t.doc_id] = t.task_id; });
      }
    } catch (e) { Logger.log('myTaskMap err: ' + e); }

    var response = {
      success: true,
      list: pageRows,
      total: total,
      page: page,
      page_size: pageSize,
      page_count: pages,
      myTaskMap: taskMap,
      permissions: {
        create: auth.docAuth.canCreate,
        update: auth.docAuth.canEdit,
        delete: auth.docAuth.canDelete
      }
    };

    if (includeOptions === true) {
      var references = wmsDocLoadReferences_(auth.docAuth);
      response.warehouses = references.warehouses;
      response.skus = references.skus;
      response.locations = references.locations;
      var workflowOptions = wmsDocLoadWorkflowOptions_(token, docType, auth.docAuth);
      if (!workflowOptions.success) return workflowOptions;
      response.workflows = workflowOptions.workflows;
      response.workflowAccess = workflowOptions.workflowAccess;
    }
    return response;
  } catch (err) {
    Logger.log('wmsListDocs: ' + err);
    return { success: false, message: 'Không tải được danh sách phiếu: ' + err.message, list: [] };
  }
}

function wmsDocResolveLines_(auth, warehouseId, rawLines, references) {
  if (!Array.isArray(rawLines) || !rawLines.length) throw new Error('Phiếu cần có ít nhất một dòng hàng.');
  if (rawLines.length > 500) throw new Error('Mỗi phiếu tối đa 500 dòng hàng.');
  var skuByCode = {};
  var skuById = {};
  references.skus.forEach(function(sku) {
    [sku.sku_id, sku.short_name, sku.barcode].forEach(function(code) {
      var key = String(code || '').trim().toUpperCase();
      if (key) skuByCode[key] = sku;
    });
    skuById[String(sku.sku_id)] = sku;
  });
  var locationById = {};
  references.locations.forEach(function(location) {
    locationById[String(location.location_id)] = location;
  });

  return rawLines.map(function(raw, index) {
    raw = raw || {};
    var lookup = String(raw.sku_id || raw.sku_code || raw.sku_code_snapshot || '').trim();
    var sku = skuById[lookup] || skuByCode[lookup.toUpperCase()];
    if (!sku) throw new Error('Dòng ' + (index + 1) + ': không tìm thấy mã SKU "' + lookup + '".');
    
    var quantity = wmsOptionalNumber_(raw.quantity !== undefined ? raw.quantity : raw.quantity_entered, 'Số lượng dòng ' + (index + 1));
    if (quantity === '' || quantity <= 0) {
      throw new Error('Dòng ' + (index + 1) + ': số lượng phải lớn hơn 0.');
    }
    var cost = wmsOptionalNumber_(raw.unit_cost, 'Đơn giá dòng ' + (index + 1));
    if (cost === '') cost = 0;
    if (cost < 0) throw new Error('Dòng ' + (index + 1) + ': đơn giá không được âm.');
    
    var locationId = wmsText_(raw.location_id, 100);
    var location = locationId ? locationById[locationId] : null;
    if (locationId && (!location || String(location.warehouse_id) !== String(warehouseId))) {
      throw new Error('Dòng ' + (index + 1) + ': vị trí không thuộc kho đã chọn hoặc ngoài phạm vi.');
    }
    var baseUom = String(sku.uom_base || '').trim();
    if (!baseUom) throw new Error('Dòng ' + (index + 1) + ': SKU chưa khai báo đơn vị cơ sở.');
    
    return {
      line_id: String(raw.line_id || '').trim() ||
        'LINE-' + Utilities.getUuid().replace(/-/g, '').substring(0, 16).toUpperCase(),
      existing_line_id: String(raw.line_id || '').trim(),
      line_no: index + 1,
      sku_id: sku.sku_id,
      sku_code_snapshot: sku.short_name,
      sku_name_snapshot: sku.sku_name,
      lot_id: wmsText_(raw.lot_id, 100),
      supplier_name_snapshot: wmsText_(raw.supplier_name_snapshot, 200), // <-- BỔ SUNG DÒNG NÀY
      production_batch_id: wmsText_(raw.production_batch_id, 100),
      warehouse_id: warehouseId,
      location_id: locationId,
      destination_warehouse_id: wmsText_(raw.destination_warehouse_id, 100),
      destination_location_id: wmsText_(raw.destination_location_id, 100),
      quantity_entered: quantity,
      uom_entered: baseUom,
      conversion_to_base: 1,
      quantity_base: quantity,
      base_uom_snapshot: baseUom,
      unit_cost: cost,
      line_amount: quantity * cost,
      qc_status_snapshot: '',
      line_notes: wmsText_(raw.line_notes || raw.notes, 1000),
      created_by: auth.user.username,
      created_at: new Date(),
      updated_by: auth.user.username,
      updated_at: new Date(),
      is_deleted: false
    };
  });
}

function wmsSaveDoc(token, docType, record) {
  docType = String(docType || '').trim().toUpperCase();
  var isCreate = !record || !String(record.doc_id || '').trim();
  var auth = wmsDocAuth_(token, docType, isCreate ? 'create' : 'update');
  if (!auth.ok) return auth.result;
  if (!record || typeof record !== 'object') return fail_('Dữ liệu phiếu không hợp lệ.');

  try {
    var docs = wmsDocGetSheet_();
    var lines = wmsGetSheet_('SHEET_WMS_LINE', WMS_LINE_HEADERS_);
    var warehouseId = wmsText_(record.warehouse_id, 100);
    var references = wmsDocLoadReferenceSubset_(auth.docAuth, warehouseId, record.lines);
    var warehouse = references.warehouses.find(function(item) {
      return String(item.warehouse_id) === warehouseId;
    });
    if (!warehouse) return fail_('Kho không hợp lệ hoặc nằm ngoài phạm vi dữ liệu được phân quyền.');
    var docDate = String(record.doc_date || '').trim();
    var dateParts = docDate.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    var parsedDate = dateParts
      ? new Date(Number(dateParts[1]), Number(dateParts[2]) - 1, Number(dateParts[3])) : null;
    if (!parsedDate || isNaN(parsedDate.getTime()) ||
        parsedDate.getFullYear() !== Number(dateParts[1]) ||
        parsedDate.getMonth() !== Number(dateParts[2]) - 1 ||
        parsedDate.getDate() !== Number(dateParts[3])) {
      return fail_('Ngày phiếu không hợp lệ.');
    }
    var partnerName = wmsText_(record.partner_name, 200);
    if (!partnerName) return fail_('Tên nhà cung cấp/khách hàng là bắt buộc.');
    var normalizedLines = wmsDocResolveLines_(auth.docAuth, warehouseId,
      record.lines, references);
    var partnerCode = wmsText_(record.partner_code, 100);
    var externalRef = wmsText_(record.external_ref_no, 150);
    var reason = wmsText_(record.reason, 500);
    var notes = wmsText_(record.notes, 2000);
    var selectedWorkflowProvided = Object.prototype.hasOwnProperty.call(
      record, 'selected_workflow_version_id');
    var selectedWorkflowVersionId = wmsText_(record.selected_workflow_version_id, 100);
    if (selectedWorkflowProvided && selectedWorkflowVersionId) {
        var workflowOptions = wmsWorkflowListAvailable_(token, docType, warehouseId,
          wmsDocNavigationKey_(docType), auth.docAuth);
      if (!workflowOptions.success) return workflowOptions;
      if (!workflowOptions.workflows.some(function(workflow) {
        return String(workflow.workflow_version_id) === selectedWorkflowVersionId;
      })) {
        return fail_('Workflow đã chọn không còn hoạt động hoặc không áp dụng cho loại phiếu/kho này.');
      }
    }

    var savedDocId = null;
    var lockResult = withLock_(function() {
      var docId = isCreate ? 'DOC-' + Utilities.getUuid().replace(/-/g, '').substring(0, 16).toUpperCase()
        : String(record.doc_id).trim();
      savedDocId = docId;
      var docMatches = isCreate ? [] : wmsDocReadByValue_(docs, 'doc_id', docId);
      if (!isCreate && !docMatches.length) return fail_('Không tìm thấy phiếu cần cập nhật.');
      if (docMatches.length > 1) throw new Error('Mã phiếu bị trùng trong sheet WMS_DOC.');
      var docMatch = docMatches.length ? docMatches[0] : null;
      var rowIndex = docMatch ? docMatch.index : -1;
      var oldRow = docMatch ? docMatch.row : null;
      if (oldRow) {
        if (wmsIsDeleted_(oldRow[docs.headers.indexOf('is_deleted')])) {
          return fail_('Phiếu đã được đánh dấu xóa.');
        }
        var statusIndex = docs.headers.indexOf('doc_status');
        if ([WMS_DOC_STATUS_.DRAFT, WMS_DOC_STATUS_.RETURNED].indexOf(String(oldRow[statusIndex] || '')) < 0 ||
            String(oldRow[docs.headers.indexOf('doc_type')] || '') !== docType) {
          return fail_('Chỉ được sửa phiếu nháp hoặc phiếu đã trả về cùng loại nghiệp vụ.');
        }
        if (String(record.expected_updated_at || '') !==
            String(wmsCellValue_(oldRow[docs.headers.indexOf('updated_at')]) || '')) {
          return fail_('Phiếu đã được cập nhật. Vui lòng tải lại trước khi sửa.', { code:'STALE_RECORD' });
        }
        if (!matrixScopeAllowed_(auth.docAuth.user, 'DOC', oldRow, null, auth.docAuth.permission, 'update')) {
          return fail_('Phiếu nằm ngoài phạm vi dữ liệu được phân quyền.', { code:'FORBIDDEN' });
        }
      }
      var now = new Date();
      var docNo = oldRow ? wmsCellValue_(oldRow[docs.headers.indexOf('doc_no')])
        : wmsDocNextNumberFromValues_(wmsDocReadColumn_(docs, 'doc_no'), docType, docDate);
      var values = {
        doc_id:docId,
        doc_no:docNo,
        doc_type:docType,
        doc_date:docDate,
        warehouse_id:warehouseId,
        destination_warehouse_id:'',
        partner_id:wmsText_(record.partner_id, 100),
        partner_code_snapshot:partnerCode,
        partner_name_snapshot:partnerName,
        external_ref_type:docType === WMS_DOC_TYPES_.PURCHASE_RECEIPT ? 'PURCHASE_ORDER' : 'SALES_ORDER',
        external_ref_id:'',
        external_ref_no:externalRef,
        requested_by_username:oldRow
          ? oldRow[docs.headers.indexOf('requested_by_username')] : auth.docAuth.user.username,
        requested_by_emp_id:oldRow ? oldRow[docs.headers.indexOf('requested_by_emp_id')] : '',
        requested_by_name_snapshot:oldRow
          ? oldRow[docs.headers.indexOf('requested_by_name_snapshot')]
          : String(auth.docAuth.user.fullName || auth.docAuth.user.username),
        workflow_run_id:'',
        approval_status:WMS_APPROVAL_STATUS_.DRAFT,
        posting_status:WMS_POSTING_STATUS_.NOT_POSTED,
        doc_status:WMS_DOC_STATUS_.DRAFT,
        submitted_at:'',
        posted_at:'',
        completed_at:'',
        reason:reason,
        notes:notes,
        reversal_of_doc_id:'',
        created_by:oldRow ? oldRow[docs.headers.indexOf('created_by')] : auth.docAuth.user.username,
        created_at:oldRow ? oldRow[docs.headers.indexOf('created_at')] : now,
        updated_by:auth.docAuth.user.username,
        updated_at:now,
        is_deleted:false,
        selected_workflow_version_id:selectedWorkflowProvided ? selectedWorkflowVersionId
          : oldRow ? oldRow[docs.headers.indexOf('selected_workflow_version_id')] : ''
      };
      var newRow = wmsRowForWrite_(docs.headers, values);
      if (!matrixScopeAllowed_(auth.docAuth.user, 'DOC', newRow, null,
          auth.docAuth.permission, isCreate ? 'create' : 'update')) {
        return fail_('Phiếu nằm ngoài phạm vi dữ liệu được phân quyền.', { code:'FORBIDDEN' });
      }

      var oldLineRows = [];
      var lineDocId = lines.headers.indexOf('doc_id');
      var lineDeleted = lines.headers.indexOf('is_deleted');
      if (oldRow) {
        oldLineRows = wmsDocReadByDocIdFast_(lines, docId).filter(function(item) {
          return !wmsIsDeleted_(item.row[lineDeleted]);
        });
      }
      var oldLinesById = {};
      oldLineRows.forEach(function(item) {
        var lineId = String(item.row[lines.headers.indexOf('line_id')] || '');
        if (!lineId || oldLinesById[lineId]) {
          throw new Error('Mã dòng bị thiếu hoặc trùng trong chi tiết phiếu.');
        }
        oldLinesById[lineId] = item;
      });
      var requestedLineIds = {};
      var newLineRows = [];
      var lineChanges = [];
      normalizedLines.forEach(function(line, index) {
        line.line_no = index + 1;
        var lineId = String(line.existing_line_id || '');
        var prior = lineId ? oldLinesById[lineId] : null;
        if (lineId && !prior) throw new Error('Một dòng phiếu không còn tồn tại. Vui lòng tải lại trước khi lưu.');
        if (prior && requestedLineIds[lineId]) throw new Error('Một dòng chi tiết bị gửi lặp.');
        if (prior) requestedLineIds[lineId] = true;
        var lineValues = prior
          ? Object.assign({}, wmsRowsAsObjects_(lines.headers, [prior.row])[0], line, { doc_id:docId })
          : Object.assign({}, line, { doc_id:docId });
        var row = wmsRowForWrite_(lines.headers, lineValues);
        if (!matrixScopeAllowed_(auth.lineAuth.user, 'LINE', row, null,
            auth.lineAuth.permission, prior ? 'update' : 'create')) {
          throw wmsDocScopeError_('Một hoặc nhiều dòng hàng nằm ngoài phạm vi dữ liệu được phân quyền.');
        }
        if (!prior) {
          newLineRows.push(row);
          return;
        }
        var old = prior.row;
        var changed = lines.headers.some(function(header, column) {
          if (['line_id', 'doc_id', 'created_by', 'created_at', 'updated_by', 'updated_at'].indexOf(header) !== -1) {
            return false;
          }
          return String(old[column] == null ? '' : old[column]) !==
            String(row[column] == null ? '' : row[column]);
        });
        if (changed) {
          if (!matrixScopeAllowed_(auth.lineAuth.user, 'LINE', old, null,
              auth.lineAuth.permission, 'update')) {
            throw wmsDocScopeError_('Một hoặc nhiều dòng phiếu nằm ngoài phạm vi sửa được phân quyền.');
          }
          row[lines.headers.indexOf('created_by')] = old[lines.headers.indexOf('created_by')];
          row[lines.headers.indexOf('created_at')] = old[lines.headers.indexOf('created_at')];
          row[lines.headers.indexOf('updated_by')] = auth.docAuth.user.username;
          row[lines.headers.indexOf('updated_at')] = now;
          lineChanges.push({ index:prior.index, row:row });
        }
      });
      var lineDeletes = oldLineRows.filter(function(item) {
        var lineId = String(item.row[lines.headers.indexOf('line_id')] || '');
        return !requestedLineIds[lineId];
      }).map(function(item) {
        if (!matrixScopeAllowed_(auth.lineAuth.user, 'LINE', item.row, null,
            auth.lineAuth.permission, 'delete')) {
          throw wmsDocScopeError_('Một hoặc nhiều dòng phiếu nằm ngoài phạm vi xóa được phân quyền.');
        }
        var deletedRow = item.row.slice();
        deletedRow[lineDeleted] = true;
        deletedRow[lines.headers.indexOf('updated_by')] = auth.docAuth.user.username;
        deletedRow[lines.headers.indexOf('updated_at')] = now;
        return { index:item.index, row:deletedRow, originalRow:item.row };
      });
      var docRowNumber = rowIndex < 0 ? docs.sheet.getLastRow() + 1 : rowIndex + 2;
      var lineStartRow = lines.sheet.getLastRow() + 1;
      var lineOriginals = lineChanges.concat(lineDeletes).map(function(item) {
        var original = oldLineRows.find(function(old) { return old.index === item.index; });
        return { index:item.index, row:item.originalRow || original.row.slice() };
      });
      var newLinesWritten = false;
      var documentWritten = false;
      try {
        if (newLineRows.length) {
          newLinesWritten = true;
          lines.sheet.getRange(lineStartRow, 1, newLineRows.length, lines.headers.length).setValues(newLineRows);
        }
        docs.sheet.getRange(docRowNumber, 1, 1, docs.headers.length).setValues([newRow]);
        documentWritten = true;
        wmsDocWriteRowRuns_(lines.sheet, lines.headers.length, lineChanges.concat(lineDeletes));
      } catch (writeError) {
        var rollbackErrors = [];
        try {
          wmsDocWriteRowRuns_(lines.sheet, lines.headers.length, lineOriginals);
        } catch (rollbackError) {
          rollbackErrors.push('Không thể khôi phục các dòng chi tiết đã cập nhật: ' + rollbackError.message);
        }
        if (newLinesWritten) {
          try {
            var discardedLines = newLineRows.map(function(row) {
              var discarded = row.slice();
              discarded[lines.headers.indexOf('is_deleted')] = true;
              discarded[lines.headers.indexOf('updated_by')] = auth.docAuth.user.username;
              discarded[lines.headers.indexOf('updated_at')] = now;
              return discarded;
            });
            lines.sheet.getRange(lineStartRow, 1, discardedLines.length, lines.headers.length)
              .setValues(discardedLines);
          } catch (rollbackError) {
            rollbackErrors.push('Không thể đánh dấu xóa các dòng chi tiết mới: ' + rollbackError.message);
          }
        }
        if (oldRow || documentWritten || docRowNumber <= docs.sheet.getLastRow()) {
          try {
            if (oldRow) {
              docs.sheet.getRange(docRowNumber, 1, 1, docs.headers.length).setValues([oldRow]);
            } else {
              var discardedDoc = newRow.slice();
              discardedDoc[docs.headers.indexOf('is_deleted')] = true;
              discardedDoc[docs.headers.indexOf('updated_by')] = auth.docAuth.user.username;
              discardedDoc[docs.headers.indexOf('updated_at')] = now;
              docs.sheet.getRange(docRowNumber, 1, 1, docs.headers.length).setValues([discardedDoc]);
            }
          } catch (rollbackError) {
            rollbackErrors.push('Không thể khôi phục trạng thái phiếu: ' + rollbackError.message);
          }
        }
        if (rollbackErrors.length) {
          throw new Error(writeError.message + ' ' + rollbackErrors.join(' '));
        }
        throw writeError;
      }
      wmsAudit_(auth.docAuth.user, 'DOC', isCreate ? 'CREATE_DRAFT' : 'UPDATE_DRAFT',
        docId, docNo);

      return {
        success:true,
        message:isCreate ? 'Đã lưu phiếu nháp.' : 'Đã cập nhật phiếu nháp.',
        doc_id:docId,
        doc:wmsRowsAsObjects_(docs.headers, [newRow])[0]
      };
    });
    if (lockResult && lockResult.success && savedDocId) {
      wmsSyncLots_(savedDocId, partnerCode, partnerName, normalizedLines, auth.docAuth.user);
    }
    return lockResult;
  } catch (err) {
    Logger.log('wmsSaveDoc: ' + err);
    var result = { success:false, message:'Không lưu được phiếu nháp: ' + err.message };
    if (err.code) result.code = err.code;
    return result;
  }
}

function wmsDocWriteRowRuns_(sheet, width, indexedRows) {
  if (!indexedRows.length) return;
  indexedRows.sort(function(a, b) { return a.index - b.index; });
  var start = 0;
  while (start < indexedRows.length) {
    var end = start + 1;
    while (end < indexedRows.length && indexedRows[end].index === indexedRows[end - 1].index + 1) end++;
    var run = indexedRows.slice(start, end);
    sheet.getRange(run[0].index + 2, 1, run.length, width).setValues(
      run.map(function(item) { return item.row; })
    );
    start = end;
  }
}

function wmsGetDoc(token, docType, docId) {
  var auth = wmsDocAuth_(token, docType, 'read');
  if (!auth.ok) return auth.result;
  try {
    var docs = wmsDocGetSheet_();
    var lines = wmsGetSheet_('SHEET_WMS_LINE', WMS_LINE_HEADERS_);
    var docMatches = wmsDocReadByValue_(docs, 'doc_id', docId);
    if (!docMatches.length || wmsIsDeleted_(docMatches[0].row[docs.headers.indexOf('is_deleted')])) {
      return fail_('Không tìm thấy phiếu.');
    }
    if (docMatches.length > 1) throw new Error('Mã phiếu bị trùng trong sheet WMS_DOC.');
    var row = docMatches[0].row;
    if (String(row[docs.headers.indexOf('doc_type')]) !== String(docType) ||
        !matrixRowAllowed_(auth.docAuth.user, 'DOC', row, null, auth.docAuth.permission)) {
      return fail_('Phiếu không tồn tại hoặc nằm ngoài phạm vi dữ liệu được phân quyền.', { code:'FORBIDDEN' });
    }
    var lineDeleted = lines.headers.indexOf('is_deleted');
    var visibleLines = [];
    if (auth.lineAuth) {
      visibleLines = wmsDocReadByDocIdFast_(lines, row[docs.headers.indexOf('doc_id')])
        .filter(function(entry) {
          var line = entry.row;
          return !wmsIsDeleted_(line[lineDeleted]) &&
            matrixRowAllowed_(auth.lineAuth.user, 'LINE', line, null, auth.lineAuth.permission);
        }).map(function(entry) {
          return entry.row;
        });
    }
    return {
      success:true,
      doc:wmsRowsAsObjects_(docs.headers, [row])[0],
      lines:wmsRowsAsObjects_(lines.headers, visibleLines)
    };
  } catch (err) {
    Logger.log('wmsGetDoc: ' + err);
    return { success:false, message:'Không tải được chi tiết phiếu: ' + err.message };
  }
}

function wmsDeleteDocs(token, docType, records) {
  var auth = wmsDocAuth_(token, docType, 'delete');
  if (!auth.ok) return auth.result;
  if (!Array.isArray(records) || !records.length) return fail_('Hãy chọn ít nhất một phiếu nháp.');
  if (records.length > 200) return fail_('Mỗi lần chỉ được xóa tối đa 200 phiếu.');

  try {
    var docs = wmsDocGetSheet_();
    var lines = wmsGetSheet_('SHEET_WMS_LINE', WMS_LINE_HEADERS_);
    return withLock_(function() {
      var docRows = wmsDocLoadRows_(docs);
      var lineRows = wmsDocLoadRows_(lines);
      var idIndex = docs.headers.indexOf('doc_id');
      var deletedIndex = docs.headers.indexOf('is_deleted');
      var seen = {};
      var selected = {};
      var removed = [];
      var errors = [];
      var docChanges = [];
      var now = new Date();
      records.forEach(function(record) {
        var id = String(record && record.doc_id || '').trim();
        var key = id.toLowerCase();
        if (!key || seen[key]) {
          errors.push({ doc_id:id, message:key ? 'Phiếu bị lặp trong yêu cầu.' : 'Thiếu mã phiếu.' });
          return;
        }
        seen[key] = true;
        var index = wmsDocFindHeader_(docs.headers, docRows, id);
        if (index < 0 || wmsIsDeleted_(docRows[index][deletedIndex])) {
          errors.push({ doc_id:id, message:'Không tìm thấy phiếu nháp.' });
          return;
        }
        var row = docRows[index];
        if (String(row[docs.headers.indexOf('doc_type')]) !== String(docType) ||
            String(row[docs.headers.indexOf('doc_status')]) !== WMS_DOC_STATUS_.DRAFT ||
            String(row[docs.headers.indexOf('posting_status')]) !== WMS_POSTING_STATUS_.NOT_POSTED) {
          errors.push({ doc_id:id, message:'Chỉ được xóa mềm phiếu nháp chưa ghi sổ.' });
          return;
        }
        if (String(record.expected_updated_at || '') !==
            String(wmsCellValue_(row[docs.headers.indexOf('updated_at')]) || '')) {
          errors.push({ doc_id:id, message:'Phiếu đã được cập nhật, vui lòng tải lại.', code:'STALE_RECORD' });
          return;
        }
        if (!matrixScopeAllowed_(auth.docAuth.user, 'DOC', row, null, auth.docAuth.permission, 'delete')) {
          errors.push({ doc_id:id, message:'Phiếu nằm ngoài phạm vi được phân quyền.', code:'FORBIDDEN' });
          return;
        }
        var inaccessibleLine = lineRows.some(function(line) {
          return String(line[lineDocId] || '') === id && !wmsIsDeleted_(line[lineDeleted]) &&
            !matrixScopeAllowed_(auth.lineAuth.user, 'LINE', line, null, auth.lineAuth.permission, 'delete');
        });
        if (inaccessibleLine) {
          errors.push({ doc_id:id, message:'Một hoặc nhiều dòng phiếu nằm ngoài phạm vi được phân quyền.', code:'FORBIDDEN' });
          return;
        }
        row[deletedIndex] = true;
        row[docs.headers.indexOf('updated_by')] = auth.docAuth.user.username;
        row[docs.headers.indexOf('updated_at')] = now;
        docChanges.push({ index:index, row:row });
        selected[id] = true;
        removed.push({ doc_id:id });
      });
      wmsDocWriteRowRuns_(docs.sheet, docs.headers.length, docChanges);

      var lineChanges = [];
      var lineDocId = lines.headers.indexOf('doc_id');
      var lineDeleted = lines.headers.indexOf('is_deleted');
      lineRows.forEach(function(row, index) {
        if (!selected[String(row[lineDocId] || '')] || wmsIsDeleted_(row[lineDeleted])) return;
        row[lineDeleted] = true;
        row[lines.headers.indexOf('updated_by')] = auth.docAuth.user.username;
        row[lines.headers.indexOf('updated_at')] = now;
        lineChanges.push({ index:index, row:row });
      });
      wmsDocWriteRowRuns_(lines.sheet, lines.headers.length, lineChanges);
      wmsAuditBatch_(auth.docAuth.user, 'DOC', 'SOFT_DELETE_DRAFT', removed.map(function(item) {
        return { key:item.doc_id, detail:'Draft ' + docType };
      }));
      return {
        success:true,
        message:'Đã chuyển ' + removed.length + ' phiếu nháp vào danh sách ẩn.',
        deleted:removed,
        errors:errors
      };
    });
  } catch (err) {
    Logger.log('wmsDeleteDocs: ' + err);
    return { success:false, message:'Không xóa được phiếu nháp: ' + err.message };
  }
}

/**
 * Hàm Helper Lookup Tiến Trình Duyệt Động (Global Scope chuẩn)
 */
function wmsDocEnrichApprovalProgress_(pageRows, spreadsheet) {
  if (!pageRows || !pageRows.length) return;
  var runIds = {};
  pageRows.forEach(function(item) {
    if (item.workflow_run_id) runIds[String(item.workflow_run_id)] = true;
  });
  if (!Object.keys(runIds).length) return;

  try {
    var runSheet = wmsGetSheet_('SHEET_WMS_WF_RUN', WMS_WORKFLOW_HEADERS_.RUN, spreadsheet);
    var stepSheet = wmsGetSheet_('SHEET_WMS_WF_STEP', WMS_WORKFLOW_HEADERS_.STEP, spreadsheet);
    var taskSheet = wmsGetSheet_('SHEET_WMS_WF_TASK', WMS_WORKFLOW_HEADERS_.TASK, spreadsheet);

    var runRows = wmsDocLoadRows_(runSheet);
    var stepRows = wmsDocLoadRows_(stepSheet);
    var taskRows = wmsDocLoadRows_(taskSheet);

    var runHeaders = runSheet.headers;
    var stepHeaders = stepSheet.headers;
    var taskHeaders = taskSheet.headers;

    var runMap = {};
    runRows.forEach(function(row) {
      var rId = String(row[runHeaders.indexOf('workflow_run_id')] || '');
      if (runIds[rId]) {
        runMap[rId] = {
          runId: rId,
          versionId: String(row[runHeaders.indexOf('workflow_version_id')] || ''),
          currentStepNo: Number(row[runHeaders.indexOf('current_step_no')]) || 1,
          status: String(row[runHeaders.indexOf('run_status')] || '')
        };
      }
    });

    var totalStepsMap = {};
    stepRows.forEach(function(row) {
      var vId = String(row[stepHeaders.indexOf('workflow_version_id')] || '');
      totalStepsMap[vId] = (totalStepsMap[vId] || 0) + 1;
    });

    var nextApproverMap = {};
    taskRows.forEach(function(row) {
      var tRunId = String(row[taskHeaders.indexOf('workflow_run_id')] || '');
      var tStatus = String(row[taskHeaders.indexOf('task_status')] || '');
      var tStepNo = Number(row[taskHeaders.indexOf('step_no')]);
      var tApprover = String(row[taskHeaders.indexOf('approver_name_snapshot')] || row[taskHeaders.indexOf('approver_username')] || '');

      if (runIds[tRunId] && tStatus === 'PENDING' && runMap[tRunId] && runMap[tRunId].currentStepNo === tStepNo) {
        nextApproverMap[tRunId] = tApprover;
      }
    });

    pageRows.forEach(function(item) {
      if (item.workflow_run_id && runMap[item.workflow_run_id]) {
        var r = runMap[item.workflow_run_id];
        item.currentStep = r.currentStepNo;
        item.totalSteps = totalStepsMap[r.versionId] || 1;
        item.nextApprover = nextApproverMap[item.workflow_run_id] || '';
      } else {
        item.currentStep = 0;
        item.totalSteps = 0;
        item.nextApprover = '';
      }
    });
  } catch (e) {
    Logger.log('Enrich progress err: ' + e);
  }
}

/**
 * Tự động đồng bộ/tham chiếu thông tin Lô hàng vào sheet WMS_LOT
 */
function wmsSyncLots_(docId, partnerCode, partnerName, lines, user) {
  if (!lines || !lines.length) return;

  var lotHeaders = [
    'lot_id', 'internal_lot_no', 'sku_id', 'supplier_id',
    'supplier_code_snapshot', 'supplier_name_snapshot', 'supplier_lot_no',
    'manufacture_date', 'expiry_date', 'first_receipt_doc_id',
    'first_receipt_line_id', 'quality_status', 'lot_status', 'created_by',
    'created_at', 'updated_by', 'updated_at', 'is_deleted'
  ];

  try {
    // Mở kho dữ liệu WMS thông qua CONFIG ID dùng chung
    var spreadsheet = SpreadsheetApp.openById(CONFIG.DB_WMS_PART1_ID);
    var sheetName = CONFIG.SHEET_WMS_LOT || 'WMS_LOT';
    var sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      Logger.log('wmsSyncLots_ ERROR: Không tìm thấy tab sheet có tên ' + sheetName);
      return;
    }

    var lastRow = sheet.getLastRow();
    var headers = [];
    
    // Nếu sheet chưa có dữ liệu hoặc tiêu đề, ghi hàng tiêu đề mẫu
    if (lastRow < 1) {
      sheet.appendRow(lotHeaders);
      headers = lotHeaders;
      lastRow = 1;
    } else {
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) {
        return String(h || '').trim();
      });
    }

    var existingRows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, headers.length).getValues() : [];
    
    var lotIndex = headers.indexOf('lot_id');
    if (lotIndex < 0) {
      Logger.log('wmsSyncLots_ ERROR: Không tìm thấy cột lot_id trong sheet WMS_LOT!');
      return;
    }

    var existingLotMap = {};
    existingRows.forEach(function(row, idx) {
      var lId = String(row[lotIndex] || '').trim();
      if (lId) existingLotMap[lId] = idx + 2; // Lưu số dòng tương ứng trong Google Sheet
    });

    var now = new Date();
    var appends = [];

    lines.forEach(function(line) {
      var lotId = String(line.lot_id || '').trim();
      if (!lotId) return; // Bỏ qua nếu dòng phiếu không có Số Lô

      // Ưu tiên NCC từ từng dòng Excel, nếu trống lấy NCC ở Header phiếu
      var supName = line.supplier_name_snapshot || partnerName || '';
      var supCode = partnerCode || '';

      if (existingLotMap[lotId]) {
        // Lô đã tồn tại -> Cập nhật thông tin Nhà cung cấp
        var rowNum = existingLotMap[lotId];
        var rowVal = existingRows[rowNum - 2];
        
        var idxSupCode = headers.indexOf('supplier_code_snapshot');
        var idxSupName = headers.indexOf('supplier_name_snapshot');
        var idxUpBy = headers.indexOf('updated_by');
        var idxUpAt = headers.indexOf('updated_at');

        var changed = false;
        if (idxSupCode >= 0 && String(rowVal[idxSupCode] || '') !== String(supCode)) { rowVal[idxSupCode] = supCode; changed = true; }
        if (idxSupName >= 0 && String(rowVal[idxSupName] || '') !== String(supName)) { rowVal[idxSupName] = supName; changed = true; }
        if (changed) {
          if (idxUpBy >= 0) rowVal[idxUpBy] = user ? user.username : 'SYSTEM';
          if (idxUpAt >= 0) rowVal[idxUpAt] = now;
          sheet.getRange(rowNum, 1, 1, headers.length).setValues([rowVal]);
        }
      } else {
        // Lô chưa tồn tại -> Tạo dòng mới trong sheet WMS_LOT
        var newLotValues = {
          lot_id: lotId,
          internal_lot_no: lotId,
          sku_id: line.sku_id || '',
          supplier_id: '',
          supplier_code_snapshot: supCode,
          supplier_name_snapshot: supName,
          supplier_lot_no: '',
          manufacture_date: '',
          expiry_date: '',
          first_receipt_doc_id: docId || '',
          first_receipt_line_id: line.line_id || '',
          quality_status: 'PENDING',
          lot_status: 'ACTIVE',
          created_by: user ? user.username : 'SYSTEM',
          created_at: now,
          updated_by: user ? user.username : 'SYSTEM',
          updated_at: now,
          is_deleted: false
        };

        // Ánh xạ dữ liệu theo đúng thứ tự tiêu đề cột thực tế của Sheet
        var rowToWrite = headers.map(function(h) {
          return newLotValues[h] !== undefined ? newLotValues[h] : '';
        });
        appends.push(rowToWrite);
      }
    });

    if (appends.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, appends.length, headers.length).setValues(appends);
    }

  } catch (e) {
    Logger.log('Lỗi wmsSyncLots_: ' + e.message + '\nStack: ' + e.stack);
  }
}