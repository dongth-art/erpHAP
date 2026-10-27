/**
 * WMS_WAREHOUSE_Server.gs - Danh mục kho.
 * Sheet WAREHOUSE: warehouse_id, warehouse_code, warehouse_name, address,
 * is_active, warehouse_purpose, storage_mode, temperature_min_c,
 * temperature_max_c, created_by, created_at, updated_by, updated_at.
 */

var WMS_WAREHOUSE_HEADERS_ = [
  'warehouse_id', 'warehouse_code', 'warehouse_name', 'address',
  'is_active', 'warehouse_purpose', 'storage_mode', 'temperature_min_c',
  'temperature_max_c', 'created_by', 'created_at', 'updated_by', 'updated_at', 'is_deleted'
];
var WMS_WAREHOUSE_VIEW_ = 'view-wms-warehouse';

var WAREHOUSE_CACHE_TTL_SEC = 120;
var _warehouseCache_ = null;
var _warehouseCacheTime_ = 0;

function getWarehouseCache_() {
  var now = Date.now();
  if (_warehouseCache_ && now - _warehouseCacheTime_ < WAREHOUSE_CACHE_TTL_SEC * 1000) {
    return _warehouseCache_;
  }
  var cache = CacheService.getScriptCache();
  var raw = cache.get('WMS_WAREHOUSE_CACHE');
  if (raw) {
    try {
      _warehouseCache_ = JSON.parse(raw);
      _warehouseCacheTime_ = now;
      return _warehouseCache_;
    } catch (e) {
      cache.remove('WMS_WAREHOUSE_CACHE');
    }
  }
  return null;
}

function setWarehouseCache_(list) {
  _warehouseCache_ = list;
  _warehouseCacheTime_ = Date.now();
  try {
    CacheService.getScriptCache().put('WMS_WAREHOUSE_CACHE', JSON.stringify(list), WAREHOUSE_CACHE_TTL_SEC);
  } catch (e) { /* cache ignore */ }
}

function invalidateWarehouseCache_() {
  _warehouseCache_ = null;
  _warehouseCacheTime_ = 0;
  try { CacheService.getScriptCache().remove('WMS_WAREHOUSE_CACHE'); } catch (e) { /* ignore */ }
}

function wmsListWarehouses(token, searchKey) {
  var auth = wmsCanRead_(token, 'WAREHOUSE', WMS_WAREHOUSE_VIEW_);
  if (!auth.ok) return auth.result;

  try {
    var cached = getWarehouseCache_();
    var rows;
    if (cached) {
      rows = cached;
    } else {
      var source = wmsGetSheet_('SHEET_WMS_WAREHOUSE', WMS_WAREHOUSE_HEADERS_);
      var lastRow = source.sheet.getLastRow();
      rows = lastRow > 1
        ? source.sheet.getRange(2, 1, lastRow - 1, source.headers.length).getValues()
        : [];
      setWarehouseCache_(rows);
    }
    var deletedIndex = source ? source.headers.indexOf('is_deleted') : WMS_WAREHOUSE_HEADERS_.indexOf('is_deleted');
    var readable = wmsFilterReadableRows_(auth, 'WAREHOUSE', rows.filter(function(row) {
      return !wmsIsDeleted_(row[deletedIndex]);
    }));
    var keyword = String(searchKey || '').trim().toLowerCase();
    var headers = source ? source.headers : WMS_WAREHOUSE_HEADERS_;
    var list = wmsRowsAsObjects_(headers, readable).filter(function(item) {
      return !keyword || [
        item.warehouse_id, item.warehouse_code, item.warehouse_name, item.address,
        item.warehouse_purpose, item.storage_mode
      ].join(' ').toLowerCase().indexOf(keyword) !== -1;
    });
    return {
      success: true,
      list: list,
      permissions: { create: auth.canCreate, update: auth.canEdit, delete: auth.canDelete }
    };
  } catch (err) {
    Logger.log('wmsListWarehouses: ' + err);
    return { success: false, message: 'Không tải được danh sách kho: ' + err.message, list: [] };
  }
}

function wmsSaveWarehouse(token, record) {
  var auth = wmsCanRead_(token, 'WAREHOUSE', WMS_WAREHOUSE_VIEW_);
  if (!auth.ok) return auth.result;
  if (!record || typeof record !== 'object') return fail_('Dữ liệu kho không hợp lệ.');

  try {
    var source = wmsGetSheet_('SHEET_WMS_WAREHOUSE', WMS_WAREHOUSE_HEADERS_);
    var sheet = source.sheet;
    var headers = source.headers;
    var idIndex = headers.indexOf('warehouse_id');
    var codeIndex = headers.indexOf('warehouse_code');
    var updatedAtIndex = headers.indexOf('updated_at');
    var deletedIndex = headers.indexOf('is_deleted');
    var mode = String(record.mode || '').toUpperCase();
    var isCreate = mode === 'CREATE';
    if (!isCreate && mode !== 'UPDATE') return fail_('Chế độ lưu kho không hợp lệ.');
    if (isCreate && !auth.canCreate) return fail_('Bạn không có quyền tạo kho.', { code: 'FORBIDDEN' });
    if (!isCreate && !auth.canEdit) return fail_('Bạn không có quyền sửa kho.', { code: 'FORBIDDEN' });

    var code = wmsText_(record.warehouse_code, 80);
    var name = wmsText_(record.warehouse_name, 200);
    var address = wmsText_(record.address, 1000);
    if (!code || !name) return fail_('Mã kho và tên kho là bắt buộc.');
    var purpose = wmsText_(record.warehouse_purpose, 80);
    var storageMode = wmsText_(record.storage_mode, 80);
    var temperatureMin = wmsOptionalNumber_(record.temperature_min_c, 'Nhiệt độ tối thiểu');
    var temperatureMax = wmsOptionalNumber_(record.temperature_max_c, 'Nhiệt độ tối đa');
    if (temperatureMin !== '' && temperatureMax !== '' && temperatureMin > temperatureMax) {
      return fail_('Nhiệt độ tối thiểu không thể lớn hơn nhiệt độ tối đa.');
    }

    return withLock_(function() {
      var lastRow = sheet.getLastRow();
      var rows = lastRow > 1
        ? sheet.getRange(2, 1, lastRow - 1, headers.length).getValues()
        : [];
      var requestedId = isCreate ? '' : String(record.warehouse_id || '').trim();
      var targetRow = -1;
      for (var i = 0; i < rows.length; i++) {
        var existingId = String(rows[i][idIndex] || '').replace(/^'/, '').trim();
        var existingCode = String(rows[i][codeIndex] || '').replace(/^'/, '').trim();
        if (isCreate && existingCode.toLowerCase() === code.replace(/^'/, '').toLowerCase()) {
          return fail_('Mã kho đã tồn tại: ' + record.warehouse_code);
        }
        if (!isCreate && existingId.toLowerCase() === requestedId.toLowerCase()) {
          targetRow = i + 2;
        } else if (!isCreate && existingCode.toLowerCase() === code.replace(/^'/, '').toLowerCase()) {
          return fail_('Mã kho đã được sử dụng bởi kho khác.');
        }
      }
      if (!isCreate && targetRow < 0) return fail_('Không tìm thấy kho cần cập nhật.');

      var oldRow = targetRow > 0 ? rows[targetRow - 2] : null;
      if (oldRow && wmsIsDeleted_(oldRow[deletedIndex])) return fail_('Kho đã được đánh dấu xóa.');
      if (oldRow && (!Object.prototype.hasOwnProperty.call(record, 'expected_updated_at') ||
          String(record.expected_updated_at || '') !== String(wmsCellValue_(oldRow[updatedAtIndex]) || ''))) {
        return fail_('Kho đã được người khác cập nhật. Vui lòng tải lại dữ liệu trước khi sửa.', {
          code: 'STALE_RECORD'
        });
      }
      var warehouseId = isCreate
        ? 'WH-' + Utilities.getUuid().replace(/-/g, '').substring(0, 12).toUpperCase()
        : requestedId;
      var values = {
        warehouse_id: wmsText_(warehouseId, 80),
        warehouse_code: code,
        warehouse_name: name,
        address: address,
        is_active: record.is_active === undefined ? true : wmsBoolean_(record.is_active),
        warehouse_purpose: purpose,
        storage_mode: storageMode,
        temperature_min_c: temperatureMin,
        temperature_max_c: temperatureMax,
        is_deleted: false
      };
      var now = new Date();
      var newRow = wmsRowForWrite_(headers, Object.assign({}, values, {
        created_by: oldRow ? oldRow[headers.indexOf('created_by')] : auth.user.username,
        created_at: oldRow ? oldRow[headers.indexOf('created_at')] : now,
        updated_by: auth.user.username,
        updated_at: now
      }));

      if (!matrixScopeAllowed_(auth.user, 'WAREHOUSE', newRow, null, auth.permission, isCreate ? 'create' : 'update') ||
          (oldRow && !matrixScopeAllowed_(auth.user, 'WAREHOUSE', oldRow, null, auth.permission, 'update'))) {
        return fail_('Kho nằm ngoài phạm vi dữ liệu được phân quyền.', { code: 'FORBIDDEN' });
      }
      var rowNumber = isCreate ? sheet.getLastRow() + 1 : targetRow;
      sheet.getRange(rowNumber, 1, 1, headers.length).setValues([newRow]);
      wmsAudit_(auth.user, 'WAREHOUSE', isCreate ? 'CREATE' : 'UPDATE', warehouseId, name);
      invalidateWarehouseCache_();
      return {
        success: true,
        message: isCreate ? 'Đã tạo kho.' : 'Đã cập nhật kho.',
        warehouse_id: warehouseId,
        warehouse: wmsRowsAsObjects_(headers, [newRow])[0]
      };
    });
  } catch (err) {
    Logger.log('wmsSaveWarehouse: ' + err);
    return { success: false, message: 'Không lưu được thông tin kho: ' + err.message };
  }
}

function wmsDeleteWarehouse(token, warehouseId, expectedUpdatedAt) {
  var auth = wmsCanRead_(token, 'WAREHOUSE', WMS_WAREHOUSE_VIEW_);
  if (!auth.ok) return auth.result;
  if (!auth.canDelete) return fail_('Bạn không có quyền xóa kho.', { code: 'FORBIDDEN' });

  try {
    var source = wmsGetSheet_('SHEET_WMS_WAREHOUSE', WMS_WAREHOUSE_HEADERS_);
    var result = wmsSoftDelete_(auth, source, 'WAREHOUSE', 'warehouse_id', warehouseId,
      expectedUpdatedAt, 'kho', function() {
        var locationSheet = source.spreadsheet.getSheetByName(CONFIG.SHEET_WMS_LOCATION);
        if (!locationSheet) throw new Error('Không tìm thấy sheet LOCATION.');
        var locationHeaders = locationSheet.getRange(1, 1, 1, locationSheet.getLastColumn())
          .getDisplayValues()[0].map(function(header) { return String(header || '').trim(); });
        var locationWarehouseIndex = locationHeaders.indexOf('warehouse_id');
        var locationDeletedIndex = locationHeaders.indexOf('is_deleted');
        if (locationWarehouseIndex < 0 || locationDeletedIndex < 0) {
          throw new Error('Sheet LOCATION cần có cột warehouse_id và is_deleted.');
        }
        var locationRows = locationSheet.getLastRow() > 1
          ? locationSheet.getRange(2, 1, locationSheet.getLastRow() - 1, locationHeaders.length).getValues()
          : [];
        var hasLocations = locationRows.some(function(row) {
          return String(row[locationWarehouseIndex] || '').trim() ===
            String(warehouseId || '').trim() &&
            !wmsIsDeleted_(row[locationDeletedIndex]);
        });
        return hasLocations ? fail_('Hãy đánh dấu xóa các vị trí trong kho trước khi xóa kho.') : null;
      });
    if (result && result.success) {
      invalidateWarehouseCache_();
    }
    return result;
  } catch (err) {
    Logger.log('wmsDeleteWarehouse: ' + err);
    return { success: false, message: 'Không đánh dấu xóa được kho: ' + err.message };
  }
}
