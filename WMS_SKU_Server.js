/**
 * WMS_SKU_Server.gs - Danh mục mã hàng (SKU).
 * Cấu trúc cột được lấy theo hàng tiêu đề trong sheet SKU.
 */

var WMS_SKU_HEADERS_ = [
  'sku_id', 'sku_name', 'sku_description', 'short_name', 'barcode', 'category_code', 'sub_category',
  'uom_base', 'uom_pack', 'pack_ratio', 'is_batch_managed', 'is_expiry_managed',
  'shelf_life_days', 'qc_inspection_req', 'min_stock', 'max_stock',
  'storage_condition', 'status', 'created_by', 'created_at', 'updated_by', 'updated_at',
  'is_deleted'
];
var WMS_SKU_VIEW_ = 'view-wms-sku';
var WMS_SKU_CATEGORIES_ = [
  '001 - Thành phẩm',
  '002 - Thành phẩm theo đơn hàng',
  '003 - Bán thành phẩm/Dở dang',
  '004 - Nguyên liệu thủy sản',
  '005 - Phụ gia/Gia vị',
  '006 - Vật tư bao trì trực tiếp',
  '007 - Vật tư bao bì gián tiếp',
  '008 - Vật tư tiêu hao sản xuất',
  '009 - Phụ tùng kỹ thuật',
  '010 - Vật tư tiêu hao kỹ thuật',
  '011 - Công cụ dụng cụ',
  '012 - Văn phòng phẩm/Thiết bị VP',
  '013 - Bảo hộ lao động',
  '014 - Nguyên liệu sản xuất',
  '015 - Khác',
  '016 - Máy móc thiết bị'
];

var SKU_CACHE_TTL_SEC = 120;
var _skuCache_ = null;
var _skuCacheTime_ = 0;

function getSkuCache_() {
  var now = Date.now();
  if (_skuCache_ && now - _skuCacheTime_ < SKU_CACHE_TTL_SEC * 1000) {
    return _skuCache_;
  }
  var cache = CacheService.getScriptCache();
  var raw = cache.get('WMS_SKU_CACHE');
  if (raw) {
    try {
      _skuCache_ = JSON.parse(raw);
      _skuCacheTime_ = now;
      return _skuCache_;
    } catch (e) {
      cache.remove('WMS_SKU_CACHE');
    }
  }
  return null;
}

function setSkuCache_(list) {
  _skuCache_ = list;
  _skuCacheTime_ = Date.now();
  try {
    CacheService.getScriptCache().put('WMS_SKU_CACHE', JSON.stringify(list), SKU_CACHE_TTL_SEC);
  } catch (e) { /* cache ignore */ }
}

function invalidateSkuCache_() {
  _skuCache_ = null;
  _skuCacheTime_ = 0;
  try { CacheService.getScriptCache().remove('WMS_SKU_CACHE'); } catch (e) { /* ignore */ }
}

function wmsGetSkuSheet_() {
  var requiredHeaders = WMS_SKU_HEADERS_.filter(function(header) {
    return header !== 'sku_description';
  });
  var source = wmsGetSheet_('SHEET_WMS_SKU', requiredHeaders);
  if (source.headers.indexOf('sku_description') < 0) {
    var descriptionColumn = source.headers.length + 1;
    source.sheet.getRange(1, descriptionColumn).setValue('sku_description');
    source.headers.push('sku_description');
  }
  return source;
}

function wmsSkuCategoryPrefix_(categoryCode) {
  var categoryIndex = WMS_SKU_CATEGORIES_.indexOf(String(categoryCode || '').trim());
  return categoryIndex < 0 ? '' : WMS_SKU_CATEGORIES_[categoryIndex].substring(0, 3);
}

function wmsSkuComparableShortName_(value, categoryCode) {
  var shortName = String(value == null ? '' : value).replace(/^'/, '').trim().toUpperCase();
  var prefix = wmsSkuCategoryPrefix_(categoryCode);
  if (!prefix) return shortName;
  var match = shortName.match(new RegExp('^0*' + String(Number(prefix)) + '\\.(\\d{1,4})$'));
  return match ? prefix + '.' + String(Number(match[1])).padStart(4, '0') : shortName;
}

function wmsNextSkuShortName_(categoryCode, rows, headers) {
  var prefix = wmsSkuCategoryPrefix_(categoryCode);
  if (!prefix) return '';
  var shortNameIndex = headers.indexOf('short_name');
  var highestSequence = 0;
  var numericPrefix = String(Number(prefix));
  var sequencePattern = new RegExp('^0*' + numericPrefix + '\\.(\\d{1,4})$');
  rows.forEach(function(row) {
    var rawValue = row[shortNameIndex];
    var normalizedValue = String(rawValue == null ? '' : rawValue).replace(/^'/, '').trim();
    var match = normalizedValue.match(sequencePattern);
    if (!match && typeof rawValue === 'number' && isFinite(rawValue)) {
      var numberParts = normalizedValue.split('.');
      if (numberParts.length === 2 && Number(numberParts[0]) === Number(prefix)) {
        match = ['', numberParts[1]];
      }
    }
    if (match) {
      var sequence = Number(String(match[1]).padEnd(4, '0'));
      highestSequence = Math.max(highestSequence, sequence);
    }
  });
  if (highestSequence >= 9999) return '';
  return prefix + '.' + String(highestSequence + 1).padStart(4, '0');
}

function wmsSuggestSkuCode(token, categoryCode) {
  var auth = wmsCanRead_(token, 'SKU', WMS_SKU_VIEW_);
  if (!auth.ok) return auth.result;
  if (!auth.canCreate && !auth.canEdit) {
    return fail_('Bạn không có quyền tạo hoặc sửa mã hàng.', { code: 'FORBIDDEN' });
  }
  if (!wmsSkuCategoryPrefix_(categoryCode)) return fail_('Nhóm hàng không hợp lệ.');

  try {
    var source = wmsGetSkuSheet_();
    var lastRow = source.sheet.getLastRow();
    var rows = lastRow > 1
      ? source.sheet.getRange(2, 1, lastRow - 1, source.headers.length).getValues()
      : [];
    var shortName = wmsNextSkuShortName_(categoryCode, rows, source.headers);
    if (!shortName) return fail_('Nhóm hàng đã sử dụng hết dải mã 4 số.');
    return { success: true, short_name: shortName };
  } catch (err) {
    Logger.log('wmsSuggestSkuCode: ' + err);
    return { success: false, message: 'Không tạo được mã hàng gợi ý: ' + err.message };
  }
}

function wmsImportSkus(token, records) {
  var auth = wmsCanRead_(token, 'SKU', WMS_SKU_VIEW_);
  if (!auth.ok) return auth.result;
  if (!auth.canCreate) return fail_('Bạn không có quyền tạo SKU.', { code: 'FORBIDDEN' });
  if (!Array.isArray(records) || !records.length) return fail_('Không có dòng dữ liệu hợp lệ để nhập.');
  if (records.length > 200) return fail_('Mỗi lần chỉ nhập tối đa 200 dòng SKU.');

  try {
    var source = wmsGetSkuSheet_();
    var sheet = source.sheet;
    var headers = source.headers;
    return withLock_(function() {
      var lastRow = sheet.getLastRow();
      var rows = lastRow > 1
        ? sheet.getRange(2, 1, lastRow - 1, headers.length).getValues()
        : [];
      var pendingRows = [];
      var imported = [];
      var errors = [];
      var idIndex = headers.indexOf('sku_id');
      var barcodeIndex = headers.indexOf('barcode');
      var shortNameIndex = headers.indexOf('short_name');
      var categoryIndex = headers.indexOf('category_code');
      var nameIndex = headers.indexOf('sku_name');
      var deletedIndex = headers.indexOf('is_deleted');
      var existingIdentifiers = {};
      var seenShortNames = {};
      var seenSkuNames = {};
      var highestSequenceByPrefix = {};
      rows.forEach(function(row) {
        var skuId = String(row[idIndex] || '').replace(/^'/, '').trim().toUpperCase();
        var barcode = String(row[barcodeIndex] || '').replace(/^'/, '').trim().toUpperCase();
        if (skuId) existingIdentifiers[skuId] = true;
        if (barcode) existingIdentifiers[barcode] = true;
        var existingCode = wmsSkuComparableShortName_(row[shortNameIndex], row[categoryIndex]);
        if (existingCode) seenShortNames[existingCode] = true;
        if (!wmsIsDeleted_(row[deletedIndex])) {
          var existingName = String(row[nameIndex] || '').replace(/^'/, '').trim().toLocaleLowerCase();
          if (existingName) seenSkuNames[existingName] = true;
        }
        var prefix = wmsSkuCategoryPrefix_(row[categoryIndex]);
        var codeMatch = existingCode.match(/^(\d{3})\.(\d{4})$/);
        if (prefix && codeMatch && codeMatch[1] === prefix) {
          highestSequenceByPrefix[prefix] = Math.max(
            highestSequenceByPrefix[prefix] || 0, Number(codeMatch[2])
          );
        }
      });

      records.forEach(function(record, index) {
        try {
          if (!record || typeof record !== 'object' || Array.isArray(record)) {
            throw new Error('Dữ liệu dòng không hợp lệ.');
          }
          var skuName = wmsText_(record.sku_name, 300);
          var shortName = wmsText_(record.short_name, 150);
          var categoryCode = wmsText_(record.category_code, 100);
          var baseUom = wmsText_(record.uom_base, 40);
          var skuDescription = wmsText_(record.sku_description, 2000);
          if (!skuName || !categoryCode || !baseUom) {
            throw new Error('Tên hàng, nhóm hàng và đơn vị cơ sở là bắt buộc.');
          }
          if (!wmsSkuCategoryPrefix_(categoryCode)) throw new Error('Nhóm hàng không hợp lệ.');

          var packRatio = record.pack_ratio === '' || record.pack_ratio == null
            ? 1 : Number(record.pack_ratio);
          var shelfLife = record.shelf_life_days === '' || record.shelf_life_days == null
            ? '' : Number(record.shelf_life_days);
          var minStock = record.min_stock === '' || record.min_stock == null
            ? '' : Number(record.min_stock);
          var maxStock = record.max_stock === '' || record.max_stock == null
            ? '' : Number(record.max_stock);
          if (!isFinite(packRatio) || packRatio <= 0) throw new Error('Quy cách đóng gói phải là số lớn hơn 0.');
          if (shelfLife !== '' && (!isFinite(shelfLife) || shelfLife < 0)) throw new Error('Hạn sử dụng phải là số ngày không âm.');
          if (minStock !== '' && (!isFinite(minStock) || minStock < 0)) throw new Error('Tồn tối thiểu phải là số không âm.');
          if (maxStock !== '' && (!isFinite(maxStock) || maxStock < 0)) throw new Error('Tồn tối đa phải là số không âm.');
          if (minStock !== '' && maxStock !== '' && minStock > maxStock) {
            throw new Error('Tồn tối thiểu không thể lớn hơn tồn tối đa.');
          }
          var status = String(record.status || 'ACTIVE').trim().toUpperCase();
          if (status !== 'ACTIVE' && status !== 'INACTIVE') throw new Error('Trạng thái SKU không hợp lệ.');

          if (!shortName) {
            var prefix = wmsSkuCategoryPrefix_(categoryCode);
            var nextSequence = (highestSequenceByPrefix[prefix] || 0) + 1;
            if (nextSequence > 9999) throw new Error('Nhóm hàng đã sử dụng hết dải mã 4 số.');
            shortName = prefix + '.' + String(nextSequence).padStart(4, '0');
          }
          var normalizedCode = wmsSkuComparableShortName_(shortName, categoryCode);
          if (seenShortNames[normalizedCode]) throw new Error('Mã hàng đã tồn tại. Vui lòng nhập mã khác.');
          var normalizedName = skuName.replace(/^'/, '').toLocaleLowerCase();
          if (seenSkuNames[normalizedName]) throw new Error('Tên hàng đã tồn tại. Vui lòng nhập tên khác.');

          var skuId = '';
          for (var attempt = 0; attempt < 5; attempt++) {
            skuId = 'SKU-' + Utilities.getUuid().replace(/-/g, '').substring(0, 12).toUpperCase();
            if (!existingIdentifiers[skuId]) break;
            skuId = '';
          }
          if (!skuId) throw new Error('Không thể tạo mã SKU duy nhất. Vui lòng thử lại.');

          var values = {
            sku_id: skuId,
            sku_name: skuName,
            sku_description: skuDescription,
            short_name: shortName,
            barcode: skuId,
            category_code: categoryCode,
            sub_category: wmsText_(record.sub_category, 150),
            uom_base: baseUom,
            uom_pack: wmsText_(record.uom_pack, 40),
            pack_ratio: packRatio,
            is_batch_managed: wmsBoolean_(record.is_batch_managed),
            is_expiry_managed: wmsBoolean_(record.is_expiry_managed),
            shelf_life_days: shelfLife,
            qc_inspection_req: wmsBoolean_(record.qc_inspection_req),
            min_stock: minStock,
            max_stock: maxStock,
            storage_condition: wmsText_(record.storage_condition, 300),
            status: status,
            created_by: auth.user.username,
            created_at: new Date(),
            updated_by: auth.user.username,
            updated_at: new Date(),
            is_deleted: false
          };
          var newRow = wmsRowForWrite_(headers, values);
          if (!matrixScopeAllowed_(auth.user, 'SKU', newRow, null, auth.permission, 'create')) {
            throw new Error('SKU nằm ngoài phạm vi dữ liệu được phân quyền.');
          }
          pendingRows.push(newRow);
          seenShortNames[normalizedCode] = true;
          seenSkuNames[normalizedName] = true;
          var generatedMatch = normalizedCode.match(/^(\d{3})\.(\d{4})$/);
          if (generatedMatch) {
            highestSequenceByPrefix[generatedMatch[1]] = Math.max(
              highestSequenceByPrefix[generatedMatch[1]] || 0, Number(generatedMatch[2])
            );
          }
          existingIdentifiers[skuId] = true;
          existingIdentifiers[String(values.barcode).toUpperCase()] = true;
          imported.push({ row: index + 1, sku_id: skuId, sku: wmsRowsAsObjects_(headers, [newRow])[0] });
        } catch (err) {
          errors.push({ row: index + 1, message: err.message || String(err) });
        }
      });

      if (pendingRows.length) {
        var firstRow = lastRow + 1;
        var shortNameColumn = shortNameIndex + 1;
        sheet.getRange(firstRow, shortNameColumn, pendingRows.length, 1).setNumberFormat('@');
        sheet.getRange(firstRow, 1, pendingRows.length, headers.length).setValues(pendingRows);
        wmsAuditBatch_(auth.user, 'SKU', 'CREATE', imported.map(function(item) {
          return { key: item.sku_id, detail: item.sku.sku_name };
        }));
      }
      return {
        success: true,
        message: 'Đã nhập ' + imported.length + '/' + records.length + ' SKU.',
        imported: imported,
        errors: errors
      };
    });
  } catch (err) {
    Logger.log('wmsImportSkus: ' + err);
    return { success: false, message: 'Không nhập được dữ liệu SKU: ' + err.message };
  }
}

function wmsListSkus(token, searchKey) {
  var auth = wmsCanRead_(token, 'SKU', WMS_SKU_VIEW_);
  if (!auth.ok) return auth.result;

  try {
    var cached = getSkuCache_();
    var rows;
    if (cached) {
      rows = cached;
    } else {
      var source = wmsGetSkuSheet_();
      var lastRow = source.sheet.getLastRow();
      rows = lastRow > 1
        ? source.sheet.getRange(2, 1, lastRow - 1, source.headers.length).getValues()
        : [];
      setSkuCache_(rows);
    }
    var headers = WMS_SKU_HEADERS_;
    var deletedIndex = headers.indexOf('is_deleted');
    var readable = wmsFilterReadableRows_(auth, 'SKU', rows.filter(function(row) {
      return !wmsIsDeleted_(row[deletedIndex]);
    }));
    var keyword = String(searchKey || '').trim().toLowerCase();
    var list = wmsRowsAsObjects_(headers, readable).filter(function(item) {
      return !keyword || [
        item.sku_id, item.sku_name, item.sku_description, item.short_name, item.barcode,
        item.category_code, item.sub_category, item.uom_base
      ].join(' ').toLowerCase().indexOf(keyword) !== -1;
    });
    return {
      success: true,
      list: list,
      permissions: { create: auth.canCreate, update: auth.canEdit, delete: auth.canDelete }
    };
  } catch (err) {
    Logger.log('wmsListSkus: ' + err);
    return { success: false, message: 'Không tải được danh sách SKU: ' + err.message, list: [] };
  }
}

function wmsSaveSku(token, record) {
  var auth = wmsCanRead_(token, 'SKU', WMS_SKU_VIEW_);
  if (!auth.ok) return auth.result;
  if (!record || typeof record !== 'object') return fail_('Dữ liệu SKU không hợp lệ.');

  try {
    var source = wmsGetSkuSheet_();
    var sheet = source.sheet;
    var headers = source.headers;
    var idIndex = headers.indexOf('sku_id');
    var barcodeIndex = headers.indexOf('barcode');
    var updatedAtIndex = headers.indexOf('updated_at');
    var deletedIndex = headers.indexOf('is_deleted');
    var mode = String(record.mode || '').toUpperCase();
    var isCreate = mode === 'CREATE';
    if (!isCreate && mode !== 'UPDATE') return fail_('Chế độ lưu SKU không hợp lệ.');
    if (isCreate && !auth.canCreate) return fail_('Bạn không có quyền tạo SKU.', { code: 'FORBIDDEN' });
    if (!isCreate && !auth.canEdit) return fail_('Bạn không có quyền sửa SKU.', { code: 'FORBIDDEN' });

    var requestedId = isCreate ? '' : wmsText_(record.sku_id, 100);
    var skuName = wmsText_(record.sku_name, 300);
    var shortName = wmsText_(record.short_name, 150);
    var categoryCode = wmsText_(record.category_code, 100);
    var skuDescription = wmsText_(record.sku_description, 2000);
    var baseUom = wmsText_(record.uom_base, 40);
    if (!isCreate && !requestedId) return fail_('Thiếu mã SKU cần cập nhật.');
    if (!skuName || (!shortName && !isCreate) || !categoryCode || !baseUom) {
      return fail_('Tên hàng, mã hàng, nhóm hàng và đơn vị cơ sở là bắt buộc.');
    }
    if (isCreate && !wmsSkuCategoryPrefix_(categoryCode)) return fail_('Nhóm hàng không hợp lệ.');

    var packRatio = record.pack_ratio === '' || record.pack_ratio == null ? 1 : Number(record.pack_ratio);
    var shelfLife = record.shelf_life_days === '' || record.shelf_life_days == null
      ? '' : Number(record.shelf_life_days);
    var minStock = record.min_stock === '' || record.min_stock == null ? '' : Number(record.min_stock);
    var maxStock = record.max_stock === '' || record.max_stock == null ? '' : Number(record.max_stock);
    if (!isFinite(packRatio) || packRatio <= 0) return fail_('Quy cách đóng gói phải là số lớn hơn 0.');
    if (shelfLife !== '' && (!isFinite(shelfLife) || shelfLife < 0)) return fail_('Hạn sử dụng phải là số ngày không âm.');
    if (minStock !== '' && (!isFinite(minStock) || minStock < 0)) return fail_('Tồn tối thiểu phải là số không âm.');
    if (maxStock !== '' && (!isFinite(maxStock) || maxStock < 0)) return fail_('Tồn tối đa phải là số không âm.');
    if (minStock !== '' && maxStock !== '' && minStock > maxStock) {
      return fail_('Tồn tối thiểu không thể lớn hơn tồn tối đa.');
    }
    var status = String(record.status || 'ACTIVE').trim().toUpperCase();
    if (status !== 'ACTIVE' && status !== 'INACTIVE') return fail_('Trạng thái SKU không hợp lệ.');

    return withLock_(function() {
      var lastRow = sheet.getLastRow();
      var rows = lastRow > 1
        ? sheet.getRange(2, 1, lastRow - 1, headers.length).getValues()
        : [];
      var targetRow = -1;
      for (var i = 0; i < rows.length; i++) {
        var existingId = String(rows[i][idIndex] || '').replace(/^'/, '').trim();
        if (!isCreate && existingId.toLowerCase() === requestedId.toLowerCase()) {
          targetRow = i + 2;
        }
      }
      if (!isCreate && targetRow < 0) return fail_('Không tìm thấy SKU cần cập nhật.');

      var oldRow = targetRow > 0 ? rows[targetRow - 2] : null;
      if (oldRow && wmsIsDeleted_(oldRow[deletedIndex])) return fail_('SKU đã được đánh dấu xóa.');
      if (isCreate && !shortName) {
        shortName = wmsNextSkuShortName_(categoryCode, rows, headers);
        if (!shortName) return fail_('Nhóm hàng đã sử dụng hết dải mã 4 số.');
      }
      if (!shortName) return fail_('Mã hàng không được để trống.');
      var normalizedShortName = wmsSkuComparableShortName_(shortName, categoryCode);
      var shortNameIndex = headers.indexOf('short_name');
      var categoryIndex = headers.indexOf('category_code');
      var duplicateShortName = rows.some(function(row, index) {
        return index !== targetRow - 2 &&
          wmsSkuComparableShortName_(row[shortNameIndex], row[categoryIndex]) === normalizedShortName;
      });
      if (duplicateShortName) return fail_('Mã hàng đã tồn tại. Vui lòng nhập mã khác.');
      var normalizedSkuName = skuName.replace(/^'/, '').toLocaleLowerCase();
      var duplicateSkuName = rows.some(function(row, index) {
        return index !== targetRow - 2 && !wmsIsDeleted_(row[deletedIndex]) &&
          String(row[headers.indexOf('sku_name')] || '').replace(/^'/, '').trim().toLocaleLowerCase() === normalizedSkuName;
      });
      if (duplicateSkuName) return fail_('Tên hàng đã tồn tại. Vui lòng nhập tên khác.');
      if (oldRow && (!Object.prototype.hasOwnProperty.call(record, 'expected_updated_at') ||
          String(record.expected_updated_at || '') !== String(wmsCellValue_(oldRow[updatedAtIndex]) || ''))) {
        return fail_('SKU đã được người khác cập nhật. Vui lòng tải lại dữ liệu trước khi sửa.', {
          code: 'STALE_RECORD'
        });
      }
      var skuId = requestedId;
      if (isCreate) {
        var existingIds = {};
        rows.forEach(function(row) {
          existingIds[String(row[idIndex] || '').replace(/^'/, '').trim().toUpperCase()] = true;
          if (barcodeIndex >= 0 && !wmsIsDeleted_(row[deletedIndex])) {
            existingIds[String(row[barcodeIndex] || '').replace(/^'/, '').trim().toUpperCase()] = true;
          }
        });
        for (var attempt = 0; attempt < 5; attempt++) {
          skuId = 'SKU-' + Utilities.getUuid().replace(/-/g, '').substring(0, 12).toUpperCase();
          if (!existingIds[skuId]) break;
          skuId = '';
        }
        if (!skuId) throw new Error('Không thể tạo mã SKU duy nhất. Vui lòng thử lại.');
      }
      var values = {
        sku_id: skuId,
        sku_name: skuName,
        sku_description: skuDescription,
        short_name: shortName,
        barcode: oldRow && barcodeIndex >= 0 && oldRow[barcodeIndex]
          ? oldRow[barcodeIndex] : skuId,
        category_code: categoryCode,
        sub_category: wmsText_(record.sub_category, 150),
        uom_base: baseUom,
        uom_pack: wmsText_(record.uom_pack, 40),
        pack_ratio: packRatio,
        is_batch_managed: wmsBoolean_(record.is_batch_managed),
        is_expiry_managed: wmsBoolean_(record.is_expiry_managed),
        shelf_life_days: shelfLife,
        qc_inspection_req: wmsBoolean_(record.qc_inspection_req),
        min_stock: minStock,
        max_stock: maxStock,
        storage_condition: wmsText_(record.storage_condition, 300),
        status: status,
        created_by: oldRow ? oldRow[headers.indexOf('created_by')] : auth.user.username,
        created_at: oldRow ? oldRow[headers.indexOf('created_at')] : new Date(),
        updated_by: auth.user.username,
        updated_at: new Date(),
        is_deleted: false
      };
      var newRow = wmsRowForWrite_(headers, values);
      if (!matrixScopeAllowed_(auth.user, 'SKU', newRow, null, auth.permission, isCreate ? 'create' : 'update') ||
          (oldRow && !matrixScopeAllowed_(auth.user, 'SKU', oldRow, null, auth.permission, 'update'))) {
        return fail_('SKU nằm ngoài phạm vi dữ liệu được phân quyền.', { code: 'FORBIDDEN' });
      }

      var rowNumber = isCreate ? sheet.getLastRow() + 1 : targetRow;
      var shortNameColumn = headers.indexOf('short_name');
      if (shortNameColumn >= 0) sheet.getRange(rowNumber, shortNameColumn + 1).setNumberFormat('@');
      sheet.getRange(rowNumber, 1, 1, headers.length).setValues([newRow]);
      wmsAudit_(auth.user, 'SKU', isCreate ? 'CREATE' : 'UPDATE', skuId, skuName);
      invalidateSkuCache_();
      return {
        success: true,
        message: isCreate ? 'Đã tạo SKU.' : 'Đã cập nhật SKU.',
        sku_id: skuId,
        sku: wmsRowsAsObjects_(headers, [newRow])[0]
      };
    });
  } catch (err) {
    Logger.log('wmsSaveSku: ' + err);
    return { success: false, message: 'Không lưu được thông tin SKU: ' + err.message };
  }
}

function wmsDeleteSku(token, skuId, expectedUpdatedAt) {
  var auth = wmsCanRead_(token, 'SKU', WMS_SKU_VIEW_);
  if (!auth.ok) return auth.result;
  if (!auth.canDelete) return fail_('Bạn không có quyền xóa SKU.', { code: 'FORBIDDEN' });

  try {
    var source = wmsGetSkuSheet_();
    var result = wmsSoftDelete_(auth, source, 'SKU', 'sku_id', skuId, expectedUpdatedAt, 'SKU');
    if (result && result.success) {
      invalidateSkuCache_();
    }
    return result;
  } catch (err) {
    Logger.log('wmsDeleteSku: ' + err);
    return { success: false, message: 'Không đánh dấu xóa được SKU: ' + err.message };
  }
}

function wmsDeleteSkus(token, records) {
  var auth = wmsCanRead_(token, 'SKU', WMS_SKU_VIEW_);
  if (!auth.ok) return auth.result;
  if (!auth.canDelete) return fail_('Bạn không có quyền xóa SKU.', { code: 'FORBIDDEN' });
  if (!Array.isArray(records) || !records.length) return fail_('Hãy chọn ít nhất một SKU cần xóa.');
  if (records.length > 200) return fail_('Mỗi lần chỉ được xóa tối đa 200 SKU.');

  try {
    var source = wmsGetSkuSheet_();
    var headers = source.headers;
    var idIndex = headers.indexOf('sku_id');
    var deletedIndex = headers.indexOf('is_deleted');
    var updatedAtIndex = headers.indexOf('updated_at');
    var updatedByIndex = headers.indexOf('updated_by');
    if (idIndex < 0 || deletedIndex < 0 || updatedAtIndex < 0 || updatedByIndex < 0) {
      throw new Error('Sheet SKU cần có các cột sku_id, is_deleted, updated_by và updated_at.');
    }

    return withLock_(function() {
      var sheet = source.sheet;
      var lastRow = sheet.getLastRow();
      var rows = lastRow > 1
        ? sheet.getRange(2, 1, lastRow - 1, headers.length).getValues()
        : [];
      var rowIndexesById = {};
      rows.forEach(function(row, index) {
        var key = String(row[idIndex] || '').trim().toLowerCase();
        if (key) rowIndexesById[key] = index;
      });

      var seen = {};
      var updates = [];
      var deleted = [];
      var errors = [];
      var now = new Date();
      records.forEach(function(record) {
        record = record || {};
        var skuId = String(record.sku_id || '').trim();
        var key = skuId.toLowerCase();
        if (!key) {
          errors.push({ sku_id:'', message:'Thiếu mã SKU.' });
          return;
        }
        if (seen[key]) {
          errors.push({ sku_id:skuId, message:'Mã SKU bị lặp trong yêu cầu.' });
          return;
        }
        seen[key] = true;
        var rowIndex = rowIndexesById[key];
        if (rowIndex == null) {
          errors.push({ sku_id:skuId, message:'Không tìm thấy SKU cần xóa.' });
          return;
        }
        var oldRow = rows[rowIndex];
        if (wmsIsDeleted_(oldRow[deletedIndex])) {
          errors.push({ sku_id:skuId, message:'SKU đã được đánh dấu xóa.' });
          return;
        }
        if (String(record.expected_updated_at || '') !== String(wmsCellValue_(oldRow[updatedAtIndex]) || '')) {
          errors.push({ sku_id:skuId, message:'SKU đã được cập nhật. Vui lòng tải lại dữ liệu trước khi xóa.', code:'STALE_RECORD' });
          return;
        }
        if (!matrixScopeAllowed_(auth.user, 'SKU', oldRow, null, auth.permission, 'delete')) {
          errors.push({ sku_id:skuId, message:'SKU nằm ngoài phạm vi dữ liệu được phân quyền.', code:'FORBIDDEN' });
          return;
        }
        var deletedRow = oldRow.slice();
        deletedRow[deletedIndex] = true;
        deletedRow[updatedByIndex] = auth.user.username;
        deletedRow[updatedAtIndex] = now;
        rows[rowIndex] = deletedRow;
        updates.push({ index:rowIndex, row:deletedRow });
        deleted.push({ sku_id:skuId });
      });

      updates.sort(function(a, b) { return a.index - b.index; });
      var runStart = 0;
      while (runStart < updates.length) {
        var runEnd = runStart + 1;
        while (runEnd < updates.length && updates[runEnd].index === updates[runEnd - 1].index + 1) runEnd++;
        var run = updates.slice(runStart, runEnd);
        sheet.getRange(run[0].index + 2, 1, run.length, headers.length).setValues(
          run.map(function(update) { return update.row; })
        );
        runStart = runEnd;
      }
      wmsAuditBatch_(auth.user, 'SKU', 'SOFT_DELETE', deleted.map(function(item) {
        return { key:item.sku_id, detail:'SKU' };
      }));
      invalidateSkuCache_();
      return {
        success:true,
        message:'Đã đánh dấu xóa ' + deleted.length + ' SKU. Dòng dữ liệu vẫn được giữ trong sheet.',
        deleted:deleted,
        errors:errors
      };
    });
  } catch (err) {
    Logger.log('wmsDeleteSkus: ' + err);
    return { success:false, message:'Không xóa được các SKU đã chọn: ' + err.message };
  }
}
