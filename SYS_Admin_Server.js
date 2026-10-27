/**
 * SYS_Admin_Server.gs - HAP HOLDING ERP SYSTEM
 * ĐỢT 2b: Nhật ký thao tác (SYS_AUDIT) và Thùng rác (khôi phục / xóa vĩnh viễn).
 * Chỉ SuperAdmin dùng được phần Thùng rác và xem Nhật ký.
 */

/* ============================================================
 * NHẬT KÝ THAO TÁC
 * ============================================================ */

function getAuditSheet_() {
  var c = (typeof CONFIG !== 'undefined') ? CONFIG : {};
  if (!c.DB_SYS_PART1_ID) return null;
  var ss = SpreadsheetApp.openById(c.DB_SYS_PART1_ID);
  var name = c.SHEET_SYS_AUDIT || 'SYS_AUDIT';
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(['time', 'username', 'module', 'action', 'record_key', 'detail']);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Trả về map tên cột -> chỉ số cột (0-based) của sheet SYS_AUDIT. Cache 1 lần/request. */
function getAuditHeaderMap_() {
  return requestMemo_('SYS_AUDIT_HEADER_MAP', function () {
    var sh = getAuditSheet_();
    if (!sh) return {};
    var lastColumn = Math.max(sh.getLastColumn(), 1);
    var headers = sh.getRange(1, 1, 1, lastColumn).getValues()[0];
    var map = {};
    for (var i = 0; i < headers.length; i++) {
      var name = String(headers[i] || '').trim();
      if (name && map[name] === undefined) map[name] = i;
    }
    return map;
  });
}

/** Ghi một dòng nhật ký. Không bao giờ làm hỏng thao tác chính nếu ghi lỗi. */
function auditLog_(user, moduleCode, action, key, detail) {
  try {
    var sh = getAuditSheet_();
    if (!sh) return;
    var u = (user && user.username) ? user.username : (typeof user === 'string' ? user : '');
    sh.appendRow([
      new Date(), u, moduleCode, action,
      safeText_(String(key == null ? '' : key).substring(0, 300)),
      safeText_(String(detail == null ? '' : detail).substring(0, 900))
    ]);
  } catch (e) {
    Logger.log('auditLog_: ' + e);
  }
}

function auditLogBatch_(user, moduleCode, action, records) {
  if (!records || !records.length) return;
  try {
    var sh = getAuditSheet_();
    if (!sh) return;
    var username = (user && user.username) ? user.username : (typeof user === 'string' ? user : '');
    var timestamp = new Date();
    var values = records.map(function(record) {
      return [
        timestamp,
        username,
        moduleCode,
        action,
        safeText_(String(record.key == null ? '' : record.key).substring(0, 300)),
        safeText_(String(record.detail == null ? '' : record.detail).substring(0, 900))
      ];
    });
    sh.getRange(sh.getLastRow() + 1, 1, values.length, values[0].length).setValues(values);
  } catch (e) {
    Logger.log('auditLogBatch_: ' + e);
  }
}

function fetchAuditLog(token, limit, keyword) {
  var auth = requireAuth_(token, 'SYS', 'audit_view', 'SYS_AUDIT', 'view-sys-admin-nhat-ky-thao-tac');
  if (!auth.ok) return auth.res;
  if (!isSuperAdminRole_(auth.user.role)) return { success: false, code: 'FORBIDDEN', message: 'Chỉ SuperAdmin được xem nhật ký hệ thống.' };

  try {
    limit = Math.min(2000, Math.max(1, parseInt(limit, 10) || 300));
    var kw = String(keyword || '').trim().toLowerCase();
    var sh = getAuditSheet_();
    if (!sh || sh.getLastRow() <= 1) return { success: true, list: [] };

    var hm = getAuditHeaderMap_();
    var colTime = hm['time'];
    var colUser = hm['username'];
    var colModule = hm['module'];
    var colAction = hm['action'];
    var colKey = hm['record_key'];
    var colDetail = hm['detail'];

    var last = sh.getLastRow();
    var window_ = Math.min(last - 1, 5000);           // chỉ quét 5000 dòng mới nhất
    var data = sh.getRange(last - window_ + 1, 1, window_, 6).getValues();
    var list = [];
    for (var i = data.length - 1; i >= 0 && list.length < limit; i--) {
      var r = data[i];
      var item = {
        time: r[colTime] instanceof Date ? Utilities.formatDate(r[colTime], 'GMT+7', 'dd/MM/yyyy HH:mm:ss') : String(r[colTime]),
        username: String(r[colUser]), module: String(r[colModule]), action: String(r[colAction]),
        record_key: String(r[colKey]), detail: String(r[colDetail])
      };
      if (kw) {
        var hay = (item.time + ' ' + item.username + ' ' + item.module + ' ' + item.action + ' ' + item.record_key + ' ' + item.detail).toLowerCase();
        if (hay.indexOf(kw) === -1) continue;
      }
      list.push(item);
    }
    return { success: true, list: list };
  } catch (err) {
    Logger.log('fetchAuditLog: ' + err);
    return { success: false, message: 'Không tải được nhật ký.', list: [] };
  }
}

/* ============================================================
 * THÙNG RÁC
 * Module: 'ID' (Hồ sơ CCCD) | 'EMP' (Nhân viên) | 'ATT' (Chấm công & Nghỉ phép)
 * ============================================================ */

function trashLabel_(code) {
  return { ID: 'Hồ sơ CCCD', EMP: 'Hồ sơ Nhân viên', ATT: 'Chấm công & Nghỉ phép' }[code] || code;
}

/** Bỏ hậu tố "#DEL..." mà xóa mềm gắn vào unique_key của chấm công */
function attOriginalKey_(k) {
  return String(k == null ? '' : k).replace(/#DEL\d+$/i, '');
}

/**
 * Quét một module, trả về { sheet, flagCol, items[], activeKeys{}, activeAlt{}, uniqueCol }.
 * items[i] = { row, key, title, sub, by, at, alt, origKey }
 * Đã đồng bộ: dùng header map để tra cứu cột thay vì chỉ số cột cứng.
 */
function trashScan_(code) {
  var res = { sheet: null, flagCol: 0, items: [], activeKeys: {}, activeAlt: {}, uniqueCol: -1 };
  var i;

  if (code === 'ID' || code === 'EMP') {
    var isEmp = (code === 'EMP');
    var sheet = isEmp ? getEMPSheet_(false) : getIDSheet_(getHRMIDConfig_(), false);
    if (!sheet || sheet.getLastRow() <= 1) {
      res.sheet = sheet;
      res.flagCol = isEmp ? empFlagCol_() : idFlagCol_();
      return res;
    }

    // Lấy header map để tra cứu cột theo tên
    var hm = getHeaderMap_(sheet);
    var flagCol = isEmp ? empFlagCol_() : idFlagCol_();
    res.sheet = sheet;
    res.flagCol = flagCol;

    var n = sheet.getLastRow() - 1;
    var width = flagCol + 2;
    var lastCol = sheet.getLastColumn();
    var data = sheet.getRange(2, 1, n, Math.min(width, Math.max(lastCol, 1))).getValues();

    // Xác định các cột cần thiết theo header
    var cccdIdx = hm['cccd_no'];
    var empIdIdx = hm['emp_id'];
    var fullNameIdx = hm['full_name'];
    var jobTitleIdx = hm['job_title'];
    var unitCodeIdx = hm['unit_code'];

    for (i = 0; i < data.length; i++) {
      var r = data[i];
      var del = (lastCol >= flagCol) && String(r[flagCol - 1]).trim().toUpperCase() === 'Y';
      var key, title, sub, alt;

      if (isEmp) {
        key = empIdIdx !== undefined ? cleanId_(r[empIdIdx]) : '';
        alt = cccdIdx !== undefined ? cleanId_(r[cccdIdx]) : '';
        title = fullNameIdx !== undefined ? String(r[fullNameIdx] || '') : '';
        var jt = jobTitleIdx !== undefined ? String(r[jobTitleIdx] || '') : '';
        var uc = unitCodeIdx !== undefined ? String(r[unitCodeIdx] || '') : '';
        sub = jt + (uc ? ' • ' + uc : '');
      } else {
        key = cccdIdx !== undefined ? cleanId_(r[cccdIdx]) : '';
        alt = '';
        title = fullNameIdx !== undefined ? String(r[fullNameIdx] || '') : '';
        sub = 'CCCD ' + key;
      }

      if (!del) {
        if (key) res.activeKeys[key] = true;
        if (isEmp && alt) res.activeAlt[alt] = true;
        continue;
      }
      res.items.push({
        row: i + 2, key: key,
        title: title,
        sub: sub,
        by: String(r[flagCol] || ''), at: fmtDateTime_(r[flagCol + 1]),
        alt: alt
      });
    }
    return res;
  }

  if (code === 'ATT') {
    var cfg = getHRMATTConfig_();
    var ss = SpreadsheetApp.openById(cfg.DB_HRM_PART1_ID);
    var sh = ss.getSheetByName(cfg.SHEET_HRM_ATT);
    res.sheet = sh;
    if (!sh || sh.getLastRow() <= 1) return res;

    var hm = getHeaderMap_(sh);
    var delIdx = hm['is_deleted'];
    var uIdx = hm['unique_key'];
    res.uniqueCol = uIdx === undefined ? -1 : uIdx + 1;
    var nn = sh.getLastRow() - 1;
    var col = function (idx) { return idx === undefined ? null : sh.getRange(2, idx + 1, nn, 1).getValues(); };
    var cDel = col(delIdx), cUk = col(uIdx), cId = col(hm['att_id']), cNm = col(hm['full_name']),
        cDt = col(hm['work_date']), cTp = col(hm['att_type']), cBy = col(hm['deleted_by']), cAt = col(hm['deleted_at']);

    for (i = 0; i < nn; i++) {
      var isDel = cDel && String(cDel[i][0]).trim().toUpperCase() === 'Y';
      var uk = cUk ? attOriginalKey_(cleanCompareKey_(cUk[i][0])) : '';
      if (!isDel) { if (uk) res.activeKeys[uk] = true; continue; }
      res.items.push({
        row: i + 2, key: getRawText_(cId ? cId[i][0] : ''),
        title: getRawText_(cNm ? cNm[i][0] : ''),
        sub: formatDateTimeValue_(cDt ? cDt[i][0] : '', 'date') + ' • ' + getRawText_(cTp ? cTp[i][0] : ''),
        by: getRawText_(cBy ? cBy[i][0] : ''), at: formatDateTimeValue_(cAt ? cAt[i][0] : '', 'full'),
        origKey: attOriginalKey_(getRawText_(cUk ? cUk[i][0] : ''))
      });
    }
    res.hm = hm;
    return res;
  }

  throw new Error('Module không hợp lệ');
}

function fetchTrashList(token, moduleCode) {
  var auth = requireAuth_(token, 'SYS', 'restore', 'SYS_TRASH', 'view-sys-admin-thung-rac-khoi-phuc');
  if (!auth.ok) return auth.res;
  if (!isSuperAdminRole_(auth.user.role)) return { success: false, code: 'FORBIDDEN', message: 'Chỉ SuperAdmin được truy cập Thùng rác.' };
  try {
    var scan = trashScan_(moduleCode);
    var list = scan.items.slice().reverse().map(function (it) {
      return { key: it.key, title: it.title, sub: it.sub, deleted_by: it.by, deleted_at: it.at };
    });
    return { success: true, list: list };
  } catch (err) {
    Logger.log('fetchTrashList: ' + err);
    return { success: false, message: 'Không tải được Thùng rác.', list: [] };
  }
}

function restoreFromTrash(token, moduleCode, keys) {
  var auth = requireAuth_(token, 'SYS', 'restore', 'SYS_TRASH', 'view-sys-admin-thung-rac-khoi-phuc');
  if (!auth.ok) return auth.res;
  if (!isSuperAdminRole_(auth.user.role)) return { success: false, code: 'FORBIDDEN', message: 'Chỉ SuperAdmin được khôi phục dữ liệu.' };
  try {
    if (!Array.isArray(keys) || keys.length === 0) return fail_('Chưa chọn mục nào để khôi phục.');

    return withLock_(function () {
      var scan = trashScan_(moduleCode);
      var want = {};
      keys.forEach(function (k) { var c = (moduleCode === 'ATT') ? cleanCompareKey_(k) : cleanId_(k); if (c) want[c] = true; });

      var rows = [], restoredKeys = [], skipped = [];
      scan.items.forEach(function (it) {
        var k = (moduleCode === 'ATT') ? cleanCompareKey_(it.key) : it.key;
        if (!want[k]) return;
        var conflict = false;
        if (moduleCode === 'ID') conflict = !!scan.activeKeys[it.key];
        else if (moduleCode === 'EMP') conflict = !!scan.activeKeys[it.key] || (it.alt && !!scan.activeAlt[it.alt]);
        else conflict = !!scan.activeKeys[cleanCompareKey_(it.origKey)];

        if (conflict) { skipped.push(it.key); return; }
        rows.push(it.row); restoredKeys.push(it.key);
        if (moduleCode === 'ATT' && scan.uniqueCol > 0) {
          scan.sheet.getRange(it.row, scan.uniqueCol).setValue(it.origKey ? "'" + it.origKey : '');
        }
        // chặn khôi phục 2 dòng cùng khóa trong 1 lần
        if (moduleCode === 'ID' || moduleCode === 'EMP') scan.activeKeys[it.key] = true;
        if (moduleCode === 'EMP' && it.alt) scan.activeAlt[it.alt] = true;
        if (moduleCode === 'ATT') scan.activeKeys[cleanCompareKey_(it.origKey)] = true;
      });

      if (rows.length === 0) {
        return fail_(skipped.length ? 'Không khôi phục được: đã có bản ghi đang hoạt động trùng khóa (' + sampleList_(skipped) + ').' : 'Không tìm thấy mục cần khôi phục.');
      }
      var flagCol = (moduleCode === 'ATT') ? scan.hm['is_deleted'] + 1 : scan.flagCol;
      restoreRows_(scan.sheet, rows, flagCol);
      auditLog_(auth.user, 'TRASH_' + moduleCode, 'RESTORE', sampleList_(restoredKeys, 20), 'Khôi phục ' + rows.length + ' mục (' + trashLabel_(moduleCode) + ')');

      return {
        success: true,
        message: 'Đã khôi phục ' + rows.length + ' mục.' + (skipped.length ? ' Bỏ qua ' + skipped.length + ' mục vì đã có bản ghi hoạt động trùng khóa (' + sampleList_(skipped) + ').' : '')
      };
    });
  } catch (err) {
    Logger.log('restoreFromTrash: ' + err);
    return fail_('Không thể khôi phục.');
  }
}

function purgeFromTrash(token, moduleCode, keys) {
  var auth = requireAuth_(token, 'SYS', 'restore', 'SYS_TRASH', 'view-sys-admin-thung-rac-khoi-phuc');
  if (!auth.ok) return auth.res;
  if (!isSuperAdminRole_(auth.user.role)) return { success: false, code: 'FORBIDDEN', message: 'Chỉ SuperAdmin được xóa vĩnh viễn dữ liệu.' };
  try {
    if (!Array.isArray(keys) || keys.length === 0) return fail_('Chưa chọn mục nào để xóa vĩnh viễn.');

    return withLock_(function () {
      var scan = trashScan_(moduleCode);
      var want = {};
      keys.forEach(function (k) { var c = (moduleCode === 'ATT') ? cleanCompareKey_(k) : cleanId_(k); if (c) want[c] = true; });

      var rows = [], purgedKeys = [];
      scan.items.forEach(function (it) {
        var k = (moduleCode === 'ATT') ? cleanCompareKey_(it.key) : it.key;
        if (want[k]) { rows.push(it.row); purgedKeys.push(it.key); }   // chỉ các dòng ĐANG ở Thùng rác
      });
      if (rows.length === 0) return fail_('Không tìm thấy mục nào trong Thùng rác khớp với lựa chọn.');

      var n = deleteRowsByIndexes_(scan.sheet, rows);
      auditLog_(auth.user, 'TRASH_' + moduleCode, 'PURGE', sampleList_(purgedKeys, 20), 'Xóa vĩnh viễn ' + n + ' mục (' + trashLabel_(moduleCode) + ')');
      return { success: true, message: 'Đã xóa vĩnh viễn ' + n + ' mục.' };
    });
  } catch (err) {
    Logger.log('purgeFromTrash: ' + err);
    return fail_('Không thể xóa vĩnh viễn.');
  }
}