/**
 * WMS_LOCATION_Server.gs - Danh mục vị trí theo kho.
 */

var WMS_LOCATION_HEADERS_ = [
  'location_id', 'warehouse_id', 'location_code', 'location_name', 'location_type',
  'parent_location_id', 'status', 'created_by', 'created_at', 'updated_by', 'updated_at',
  'is_deleted'
];
var WMS_LOCATION_VIEW_ = 'view-wms-warehouse';

var LOCATION_CACHE_TTL_SEC = 120;
var _locationCache_ = null;
var _locationCacheTime_ = 0;

function getLocationCache_() {
  var now = Date.now();
  if (_locationCache_ && now - _locationCacheTime_ < LOCATION_CACHE_TTL_SEC * 1000) {
    return _locationCache_;
  }
  var cache = CacheService.getScriptCache();
  var raw = cache.get('WMS_LOCATION_CACHE');
  if (raw) {
    try {
      _locationCache_ = JSON.parse(raw);
      _locationCacheTime_ = now;
      return _locationCache_;
    } catch (e) {
      cache.remove('WMS_LOCATION_CACHE');
    }
  }
  return null;
}

function setLocationCache_(list) {
  _locationCache_ = list;
  _locationCacheTime_ = Date.now();
  try {
    CacheService.getScriptCache().put('WMS_LOCATION_CACHE', JSON.stringify(list), LOCATION_CACHE_TTL_SEC);
  } catch (e) { /* cache ignore */ }
}

function invalidateLocationCache_() {
  _locationCache_ = null;
  _locationCacheTime_ = 0;
  try { CacheService.getScriptCache().remove('WMS_LOCATION_CACHE'); } catch (e) { /* ignore */ }
}

function wmsCompareLocationCodes_(left, right) {
  var leftParts = String(left || '').toUpperCase().match(/\d+|\D+/g) || [];
  var rightParts = String(right || '').toUpperCase().match(/\d+|\D+/g) || [];
  var count = Math.min(leftParts.length, rightParts.length);
  for (var i = 0; i < count; i++) {
    var leftNumber = /^\d+$/.test(leftParts[i]);
    var rightNumber = /^\d+$/.test(rightParts[i]);
    var comparison;
    if (leftNumber && rightNumber) {
      comparison = Number(leftParts[i]) - Number(rightParts[i]);
      if (!comparison) comparison = leftParts[i].length - rightParts[i].length;
    } else {
      comparison = leftParts[i] < rightParts[i] ? -1 : (leftParts[i] > rightParts[i] ? 1 : 0);
    }
    if (comparison) return comparison;
  }
  return leftParts.length - rightParts.length;
}

function wmsSortLocationsByTree_(locations) {
  var byId = {};
  var children = {};
  var ordered = [];
  var visited = {};

  locations.forEach(function(location) {
    byId[String(location.location_id)] = location;
  });
  locations.forEach(function(location) {
    var parentId = String(location.parent_location_id || '');
    if (parentId && byId[parentId] && parentId !== String(location.location_id)) {
      (children[parentId] || (children[parentId] = [])).push(location);
    } else {
      (children.__roots || (children.__roots = [])).push(location);
    }
  });

  function sortSiblings(items) {
    items.sort(function(left, right) {
      var byCode = wmsCompareLocationCodes_(left.location_code, right.location_code);
      if (byCode) return byCode;
      var leftName = String(left.location_name || '').toUpperCase();
      var rightName = String(right.location_name || '').toUpperCase();
      if (leftName !== rightName) return leftName < rightName ? -1 : 1;
      return String(left.location_id).localeCompare(String(right.location_id));
    });
  }

  function appendBranch(location) {
    var id = String(location.location_id);
    if (visited[id]) return;
    visited[id] = true;
    ordered.push(location);
    var nested = children[id] || [];
    sortSiblings(nested);
    nested.forEach(appendBranch);
  }

  sortSiblings(children.__roots || []);
  (children.__roots || []).forEach(appendBranch);
  locations.slice().sort(function(left, right) {
    return wmsCompareLocationCodes_(left.location_code, right.location_code);
  }).forEach(appendBranch);
  return ordered;
}

function wmsListLocations(token, warehouseId) {
  var auth = wmsCanRead_(token, 'LOCATION', WMS_LOCATION_VIEW_);
  if (!auth.ok) return auth.result;
  if (!warehouseId) return fail_('Thiếu mã kho cần xem vị trí.');

  try {
    var cached = getLocationCache_();
    var rows;
    if (cached) {
      rows = cached;
    } else {
      var source = wmsGetSheet_('SHEET_WMS_LOCATION', WMS_LOCATION_HEADERS_);
      var lastRow = source.sheet.getLastRow();
      rows = lastRow > 1
        ? source.sheet.getRange(2, 1, lastRow - 1, source.headers.length).getValues()
        : [];
      setLocationCache_(rows);
    }
    var headers = WMS_LOCATION_HEADERS_;
    var warehouseIndex = headers.indexOf('warehouse_id');
    var deletedIndex = headers.indexOf('is_deleted');
    var readable = rows.filter(function(row) {
      return String(row[warehouseIndex] || '').trim() === String(warehouseId).trim() &&
        !wmsIsDeleted_(row[deletedIndex]) &&
        matrixRowAllowed_(auth.user, 'LOCATION', row, null, auth.permission);
    });
    return {
      success: true,
      list: wmsSortLocationsByTree_(wmsRowsAsObjects_(headers, readable)),
      permissions: { create: auth.canCreate, update: auth.canEdit, delete: auth.canDelete }
    };
  } catch (err) {
    Logger.log('wmsListLocations: ' + err);
    return { success: false, message: 'Không tải được vị trí kho: ' + err.message, list: [] };
  }
}

function wmsSaveLocation(token, record) {
  var auth = wmsCanRead_(token, 'LOCATION', WMS_LOCATION_VIEW_);
  if (!auth.ok) return auth.result;
  if (!record || typeof record !== 'object') return fail_('Dữ liệu vị trí không hợp lệ.');

  try {
    var warehouseId = wmsText_(record.warehouse_id, 100);
    if (!warehouseId) return fail_('Thiếu mã kho.');
    var source = wmsGetSheet_('SHEET_WMS_LOCATION', WMS_LOCATION_HEADERS_);
    var warehouseSheet = source.spreadsheet.getSheetByName(CONFIG.SHEET_WMS_WAREHOUSE);
    if (!warehouseSheet) throw new Error('Không tìm thấy sheet WAREHOUSE.');
    var warehouseHeaders = warehouseSheet.getRange(1, 1, 1, warehouseSheet.getLastColumn()).getDisplayValues()[0];
    var warehouseIdIndex = warehouseHeaders.indexOf('warehouse_id');
    if (warehouseIdIndex < 0) throw new Error('Sheet WAREHOUSE thiếu cột warehouse_id.');
    var warehouseDeletedIndex = warehouseHeaders.indexOf('is_deleted');
    if (warehouseDeletedIndex < 0) throw new Error('Sheet WAREHOUSE thiếu cột is_deleted.');
    var warehouseRows = warehouseSheet.getLastRow() > 1
      ? warehouseSheet.getRange(2, 1, warehouseSheet.getLastRow() - 1, warehouseHeaders.length).getValues()
      : [];
    var warehousePermission = getEffectivePermission_(auth.user.username, auth.user.role, 'WAREHOUSE');
    var warehouseExists = warehouseRows.some(function(row) {
      return String(row[warehouseIdIndex] || '').trim() === String(warehouseId).trim() &&
        !wmsIsDeleted_(row[warehouseDeletedIndex]) &&
        matrixRowAllowed_(auth.user, 'WAREHOUSE', row, null, warehousePermission);
    });
    if (!warehouseExists) return fail_('Kho không tồn tại hoặc nằm ngoài phạm vi được phân quyền.');
    var sheet = source.sheet;
    var headers = source.headers;
    var mode = String(record.mode || '').toUpperCase();
    var isCreate = mode === 'CREATE';
    if (!isCreate && mode !== 'UPDATE') return fail_('Chế độ lưu vị trí không hợp lệ.');
    if (isCreate && !auth.canCreate) return fail_('Bạn không có quyền tạo vị trí.', { code: 'FORBIDDEN' });
    if (!isCreate && !auth.canEdit) return fail_('Bạn không có quyền sửa vị trí.', { code: 'FORBIDDEN' });

    var code = wmsText_(record.location_code, 100);
    var name = wmsText_(record.location_name, 200);
    var type = wmsText_(record.location_type, 80);
    var parentId = wmsText_(record.parent_location_id, 100);
    var status = wmsText_(record.status || 'ACTIVE', 40).toUpperCase();
    if (!warehouseId || !code || !name || !type) {
      return fail_('Kho, mã vị trí, tên vị trí và loại vị trí là bắt buộc.');
    }
    if (status !== 'ACTIVE' && status !== 'INACTIVE') return fail_('Trạng thái vị trí không hợp lệ.');

    return withLock_(function() {
      var lastRow = sheet.getLastRow();
      var rows = lastRow > 1
        ? sheet.getRange(2, 1, lastRow - 1, headers.length).getValues()
        : [];
      var index = {
        id: headers.indexOf('location_id'),
        warehouse: headers.indexOf('warehouse_id'),
        parent: headers.indexOf('parent_location_id'),
        updatedAt: headers.indexOf('updated_at'),
        createdBy: headers.indexOf('created_by'),
        createdAt: headers.indexOf('created_at'),
        deleted: headers.indexOf('is_deleted')
      };
      var requestedId = isCreate ? '' : String(record.location_id || '').trim();
      var targetRow = -1;
      var existingById = {};

      rows.forEach(function(row, rowIndex) {
        var existingId = String(row[index.id] || '').trim();
        var existingWarehouse = String(row[index.warehouse] || '').trim();
        existingById[existingId] = {
          warehouseId: existingWarehouse,
          parentId: String(row[index.parent] || '').trim(),
          isDeleted: wmsIsDeleted_(row[index.deleted])
        };
        if (!isCreate && existingId.toLowerCase() === requestedId.toLowerCase()) {
          targetRow = rowIndex + 2;
        }
      });

      if (!isCreate && targetRow < 0) return fail_('Không tìm thấy vị trí cần cập nhật.');

      var oldRow = targetRow > 0 ? rows[targetRow - 2] : null;
      if (oldRow && wmsIsDeleted_(oldRow[index.deleted])) return fail_('Vị trí đã được đánh dấu xóa.');
      if (oldRow && String(oldRow[index.warehouse] || '').trim() !== warehouseId) {
        return fail_('Không thể chuyển vị trí sang kho khác.');
      }
      if (oldRow && (!Object.prototype.hasOwnProperty.call(record, 'expected_updated_at') ||
          String(record.expected_updated_at || '') !== String(wmsCellValue_(oldRow[index.updatedAt]) || ''))) {
        return fail_('Vị trí đã được người khác cập nhật. Vui lòng tải lại danh sách trước khi sửa.', {
          code: 'STALE_RECORD'
        });
      }
      if (parentId) {
        var parent = existingById[parentId];
        if (!parent || parent.isDeleted || parent.warehouseId !== warehouseId) {
          return fail_('Vị trí cha phải thuộc cùng kho.');
        }
        var currentParent = parentId;
        var visited = {};
        while (currentParent) {
          if (currentParent === requestedId) return fail_('Không thể gán vị trí con làm vị trí cha.');
          if (visited[currentParent]) return fail_('Cấu trúc vị trí hiện có chứa vòng lặp.');
          visited[currentParent] = true;
          currentParent = existingById[currentParent] ? existingById[currentParent].parentId : '';
        }
      }

      var locationId = isCreate
        ? 'LOC-' + Utilities.getUuid().replace(/-/g, '').substring(0, 12).toUpperCase()
        : requestedId;
      var now = new Date();
      var values = {
        location_id: locationId,
        warehouse_id: warehouseId,
        location_code: code,
        location_name: name,
        location_type: type,
        parent_location_id: parentId,
        status: status,
        created_by: oldRow ? oldRow[index.createdBy] : auth.user.username,
        created_at: oldRow ? oldRow[index.createdAt] : now,
        updated_by: auth.user.username,
        updated_at: now,
        is_deleted: false
      };
      var newRow = wmsRowForWrite_(headers, values);
      if (!matrixScopeAllowed_(auth.user, 'LOCATION', newRow, null, auth.permission, isCreate ? 'create' : 'update') ||
          (oldRow && !matrixScopeAllowed_(auth.user, 'LOCATION', oldRow, null, auth.permission, 'update'))) {
        return fail_('Vị trí nằm ngoài phạm vi dữ liệu được phân quyền.', { code: 'FORBIDDEN' });
      }

      var rowNumber = isCreate ? sheet.getLastRow() + 1 : targetRow;
      sheet.getRange(rowNumber, 1, 1, headers.length).setValues([newRow]);
      wmsAudit_(auth.user, 'LOCATION', isCreate ? 'CREATE' : 'UPDATE', locationId, name);
      invalidateLocationCache_();
      return {
        success: true,
        message: isCreate ? 'Đã tạo vị trí.' : 'Đã cập nhật vị trí.',
        location_id: locationId,
        location: wmsRowsAsObjects_(headers, [newRow])[0]
      };
    });
  } catch (err) {
    Logger.log('wmsSaveLocation: ' + err);
    return { success: false, message: 'Không lưu được vị trí kho: ' + err.message };
  }
}

function wmsDeleteLocation(token, locationId, expectedUpdatedAt) {
  var auth = wmsCanRead_(token, 'LOCATION', WMS_LOCATION_VIEW_);
  if (!auth.ok) return auth.result;
  if (!auth.canDelete) return fail_('Bạn không có quyền xóa vị trí.', { code: 'FORBIDDEN' });

  try {
    var source = wmsGetSheet_('SHEET_WMS_LOCATION', WMS_LOCATION_HEADERS_);
    var result = wmsSoftDelete_(auth, source, 'LOCATION', 'location_id', locationId,
      expectedUpdatedAt, 'vị trí', function(oldRow, rows) {
        var parentId = String(oldRow[source.headers.indexOf('location_id')] || '').trim();
        var parentIndex = source.headers.indexOf('parent_location_id');
        var deletedIndex = source.headers.indexOf('is_deleted');
        var hasChildren = rows.some(function(row) {
          return String(row[parentIndex] || '').trim() === parentId &&
            !wmsIsDeleted_(row[deletedIndex]);
        });
        return hasChildren ? fail_('Hãy đánh dấu xóa các vị trí con trước khi xóa vị trí này.') : null;
      });
    if (result && result.success) {
      invalidateLocationCache_();
    }
    return result;
  } catch (err) {
    Logger.log('wmsDeleteLocation: ' + err);
    return { success: false, message: 'Không đánh dấu xóa được vị trí: ' + err.message };
  }
}
