/**
 * HRM_ID_Server.gs - HAP HOLDING ERP SYSTEM (Backend Hồ sơ CCCD/CMND)
 * ĐỢT 1: Đã gộp HRM_Server.gs (bản cũ) vào đây -> XÓA file HRM_Server.gs.
 * Mọi hàm công khai nhận "token" ở tham số đầu tiên và kiểm tra quyền qua requireAuth_ (SYS_Server.gs).
 */

/* ============================================================
 * TIỆN ÍCH NỘI BỘ
 * ============================================================ */

function getHRMIDConfig_() {
  var c = (typeof CONFIG !== 'undefined') ? CONFIG : {};
  return {
    dbId: c.DB_HRM_PART1_ID || '',
    sheetId: c.SHEET_HRM_ID || 'HRM_ID',
    sheetEmp: c.SHEET_HRM_EMP || 'HRM_EMP',
    driveFolderId: c.DRIVE_FOLDER_ID || ''
  };
}

/** dd/MM/yyyy hoặc d/M/yyyy -> yyyy-MM-dd (lấy từ HRM_Server cũ để không mất tính năng dán Excel) */
function parseVNFormattedDate_(dateStr) {
  if (!dateStr || typeof dateStr !== 'string') return dateStr || '';
  var str = dateStr.trim();
  if (str.indexOf('/') !== -1) {
    var p = str.split('/');
    if (p.length === 3) {
      return p[2] + '-' + ('0' + p[1]).slice(-2) + '-' + ('0' + p[0]).slice(-2);
    }
  }
  return str;
}

function getIDSheet_(cfg, createIfMissing) {
  if (!cfg.dbId) throw new Error('Chưa cấu hình DB_HRM_PART1_ID');
  var ss = SpreadsheetApp.openById(cfg.dbId);
  var sheet = ss.getSheetByName(cfg.sheetId);
  if (!sheet && createIfMissing) {
    sheet = ss.insertSheet(cfg.sheetId);
    sheet.appendRow(Object.keys(MAP_HRM_ID)); // 15 cột đúng thứ tự trong CORE_Config
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** Trả về map tên cột (lowercase, không dấu) -> chỉ số cột (0-based) của sheet HRM_ID. */
function getHRMIDHeaderMap_() {
  return requestMemo_('HRM_ID_HEADER_MAP', function () {
    try {
      var sheet = getIDSheet_(getHRMIDConfig_(), false);
      if (!sheet) return {};
      var lastColumn = Math.max(sheet.getLastColumn(), 1);
      var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
      var map = {};
      for (var i = 0; i < headers.length; i++) {
        var name = String(headers[i] || '').trim();
        if (name && map[name] === undefined) map[name] = i;
      }
      return map;
    } catch (e) {
      Logger.log('getHRMIDHeaderMap_: ' + e);
      return {};
    }
  });
}

/** Danh sách tên cột chuẩn của HRM_ID (15 cột dữ liệu + 3 cột xóa mềm). */
function HRMIDColumns_() {
  return Object.keys(MAP_HRM_ID);
}

/** Trả về chỉ số cột theo tên (lowercase). Dùng fallback index nếu thiếu. */
function hrmiColIdx_(name) {
  var hm = getHRMIDHeaderMap_();
  return hm[name] !== undefined ? hm[name] : HRMIDColumns_().indexOf(name);
}

function mapIdRow_(row) {
  return {
    raw_qr: row[hrmiColIdx_('raw_qr')] || '',
    cccd_no: cleanId_(row[hrmiColIdx_('cccd_no')]),
    cmnd_no: cleanId_(row[hrmiColIdx_('cmnd_no')]),
    full_name: row[hrmiColIdx_('full_name')] || '',
    gender: row[hrmiColIdx_('gender')] || '',
    dob: fmtDate_(row[hrmiColIdx_('dob')]),
    perm_addr: row[hrmiColIdx_('perm_addr')] || '',
    merged_addr: row[hrmiColIdx_('merged_addr')] || '',
    curr_addr: row[hrmiColIdx_('curr_addr')] || '',
    issue_date: fmtDate_(row[hrmiColIdx_('issue_date')]),
    issue_place: row[hrmiColIdx_('issue_place')] || '',
    img_front: row[hrmiColIdx_('img_front')] || '',
    img_back: row[hrmiColIdx_('img_back')] || '',
    created_by: row[hrmiColIdx_('created_by')] || '',
    created_at: row[hrmiColIdx_('created_at')] ? Utilities.formatDate(new Date(row[hrmiColIdx_('created_at')]), 'GMT+7', 'dd/MM/yyyy HH:mm') : ''
  };
}

function mapIdListRow_(row) {
  var item = mapIdRow_(row);
  item.has_img_front = !!item.img_front;
  item.has_img_back = !!item.img_back;
  delete item.img_front;
  delete item.img_back;
  return item;
}

/* ============================================================
 * ĐỌC DỮ LIỆU
 * ============================================================ */

function idFlagCol_() { return 16; }   // cột P: is_deleted | Q: deleted_by | R: deleted_at

/** Tìm dòng theo số CCCD. Ưu tiên dòng đang hoạt động; nếu chỉ có dòng trong Thùng rác thì báo deleted=true. */
function findIdRow_(sheet, cccd) {
  var last = sheet.getLastRow();
  if (last <= 1) return { row: -1, deleted: false };
  var cccdIdx = hrmiColIdx_('cccd_no');
  var col = sheet.getRange(2, cccdIdx + 1, last - 1, 1).getValues();
  var flags = readDeletedFlags_(sheet, idFlagCol_(), 2, last - 1);
  var delRow = -1;
  for (var i = 0; i < col.length; i++) {
    if (cleanId_(col[i][0]) === cccd) {
      if (!flags[i]) return { row: i + 2, deleted: false };
      if (delRow < 0) delRow = i + 2;
    }
  }
  return delRow > 0 ? { row: delRow, deleted: true } : { row: -1, deleted: false };
}

function fetchHRMIDListPaginated(token, page, pageSize) {
  var auth = requireAuth_(token, 'HRM', 'view', 'HRM_ID', 'view-hrm-id');
  if (!auth.ok) return Object.assign({ totalRows: 0 }, auth.res);
  var dataPermission = getEffectivePermission_(auth.user.username, auth.user.role, 'HRM_ID');
  var permissions = {
    create: hasPermission_(auth.user, 'HRM', 'create', 'HRM_ID', undefined, dataPermission),
    update: hasPermission_(auth.user, 'HRM', 'update', 'HRM_ID', undefined, dataPermission),
    delete: hasPermission_(auth.user, 'HRM', 'delete', 'HRM_ID', undefined, dataPermission),
    export: hasPermission_(auth.user, 'HRM', 'export', 'HRM_ID', undefined, dataPermission),
    import: hasPermission_(auth.user, 'HRM', 'import', 'HRM_ID', undefined, dataPermission),
    restore: hasPermission_(auth.user, 'HRM', 'restore', 'HRM_ID', undefined, dataPermission),
    addEmployee: hasPermission_(auth.user, 'HRM', 'ADD_TO_EMPLOYEE', 'HRM_ID', undefined, dataPermission)
  };

  try {
    // 1. KIỂM TRA MA TRẬN PHÂN QUYỀN (SYS_PERM)
    var perm = dataPermission;

    // Nếu bị chặn quyền READ (can_read = FALSE)
    if (!perm.canRead) {
      return { 
        success: false, 
        message: 'Tài khoản [' + auth.user.username + '] không có quyền xem dữ liệu bảng HRM_ID!', 
        list: [], 
        totalRows: 0 
      };
    }

    page = Math.max(1, parseInt(page, 10) || 1);
    pageSize = Math.min(200, Math.max(1, parseInt(pageSize, 10) || 20));

    var sheet = getIDSheet_(getHRMIDConfig_(), false);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, list: [], totalRows: 0, permissions: permissions };

    var n = sheet.getLastRow() - 1;
    var flags = readDeletedFlags_(sheet, idFlagCol_(), 2, n);
    var cccdIdx = hrmiColIdx_('cccd_no');
    var createdByIdx = hrmiColIdx_('created_by');
    var lastColumn = Math.max(sheet.getLastColumn(), 1);
    // Đọc toàn bộ dữ liệu từ cột 1 đến lastColumn
    var rawData = sheet.getRange(2, 1, n, lastColumn).getValues();

    var active = [];
    for (var i = 0; i < n; i++) {
      if (!flags[i] && cleanId_(rawData[i][cccdIdx]) !== '' &&
          matrixRowAllowed_(auth.user, 'HRM_ID', rawData[i], { owner: createdByIdx }, dataPermission)) {
        active.push(i + 2);
      }
    }

    var totalRows = active.length;
    var pageRows = active.slice((page - 1) * pageSize, page * pageSize);
    if (pageRows.length === 0) return { success: true, list: [], totalRows: totalRows, permissions: permissions };

    var list = [];
    pageRows.forEach(function (r) {
      var row = rawData[r - 2];
      if (cleanId_(row[cccdIdx]) !== '') list.push(mapIdListRow_(row));
    });

    return { 
      success: true, 
      list: list, 
      totalRows: totalRows,
      perm: perm,
      permissions: permissions
    };
  } catch (err) {
    Logger.log('fetchHRMIDListPaginated: ' + err);
    return { success: false, message: 'Lỗi tải dữ liệu', list: [], totalRows: 0 };
  }
}

function fetchAllHRMIDRecords(token) {
  var auth = requireAuth_(token, 'HRM', 'export', 'HRM_ID', 'view-hrm-id');
  if (!auth.ok) return auth.res;

  try {
    var dataPermission = getEffectivePermission_(auth.user.username, auth.user.role, 'HRM_ID');
    var sheet = getIDSheet_(getHRMIDConfig_(), false);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, list: [] };

    var n = sheet.getLastRow() - 1;
    var lastColumn = Math.max(sheet.getLastColumn(), 1);
    var data = sheet.getRange(2, 1, n, lastColumn).getValues();
    var flags = readDeletedFlags_(sheet, idFlagCol_(), 2, n);
    var cccdIdx = hrmiColIdx_('cccd_no');
    var createdByIdx = hrmiColIdx_('created_by');
    var list = [];
    for (var i = 0; i < data.length; i++) {
      if (!flags[i] && cleanId_(data[i][cccdIdx]) !== '' &&
          matrixRowAllowed_(auth.user, 'HRM_ID', data[i], { owner: createdByIdx }, dataPermission)) list.push(mapIdRow_(data[i]));
    }
    auditLog_(auth.user, 'HRM_ID', 'EXPORT', '', 'Xuất ' + list.length + ' hồ sơ CCCD');
    return { success: true, list: list };
  } catch (err) {
    Logger.log('fetchAllHRMIDRecords: ' + err);
    return { success: false, message: 'Lỗi tải dữ liệu xuất Excel', list: [] };
  }
}

function parseCCCDQRCode(token, qrText) {
  var auth = requireAuth_(token, 'HRM', 'view', 'HRM_ID', 'view-hrm-id');
  if (!auth.ok) return auth.res;

  try {
    if (!qrText || typeof qrText !== 'string') return fail_('Mã QR trống');
    var parts = qrText.split('|');
    if (parts.length < 6) return fail_('Định dạng mã QR không chuẩn');

    var toIso = function (s) {
      s = s ? s.trim() : '';
      return s.length === 8 ? s.substring(4, 8) + '-' + s.substring(2, 4) + '-' + s.substring(0, 2) : '';
    };

    return {
      success: true,
      data: {
        raw_qr: qrText,
        cccd_no: parts[0] ? parts[0].trim() : '',
        cmnd_no: parts[1] ? parts[1].trim() : '',
        full_name: parts[2] ? parts[2].trim() : '',
        dob: toIso(parts[3]),
        gender: parts[4] ? parts[4].trim() : '',
        perm_addr: parts[5] ? parts[5].trim() : '',
        issue_date: toIso(parts[6])
      }
    };
  } catch (e) {
    return fail_('Không thể giải mã QR');
  }
}

/* ============================================================
 * GHI DỮ LIỆU
 * ============================================================ */

function uploadCCCDImages(token, frontBase64, backBase64, cccdNo) {
  var auth = requireAuth_(token, 'HRM', 'create', 'HRM_ID', 'view-hrm-id');
  if (!auth.ok) return auth.res;

  try {
    var cfg = getHRMIDConfig_();
    var folder;
    if (cfg.driveFolderId) {
      folder = DriveApp.getFolderById(cfg.driveFolderId);
    } else {
      var it = DriveApp.getFoldersByName('HAP_ERP_CCCD_Storage');
      folder = it.hasNext() ? it.next() : DriveApp.createFolder('HAP_ERP_CCCD_Storage');
    }
    folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);

    var safeNo = cleanId_(cccdNo).replace(/[^0-9A-Za-z]/g, '').substring(0, 20) || 'NOID';
    var stamp = Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMddHHmmss');

    var saveOne = function (dataUrl, label) {
      if (!dataUrl) return '';
      var m = String(dataUrl).match(/^data:(image\/(?:jpeg|png|webp));base64,(.+)$/);
      if (!m) return '';
      if (m[2].length > 7000000) throw new Error('Ảnh vượt quá giới hạn dung lượng.');
      var ext = m[1] === 'image/png' ? 'png' : (m[1] === 'image/webp' ? 'webp' : 'jpg');
      var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], safeNo + '_' + label + '_' + stamp + '.' + ext);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
      return 'https://lh3.googleusercontent.com/d/' + file.getId();
    };

    return { success: true, frontUrl: saveOne(frontBase64, 'FRONT'), backUrl: saveOne(backBase64, 'BACK') };
  } catch (err) {
    Logger.log('uploadCCCDImages: ' + err);
    return fail_('Lỗi tải ảnh lên lưu trữ');
  }
}

function getHRMIDImage(token, cccdNo, side) {
  var auth = requireAuth_(token, 'HRM', 'view', 'HRM_ID', 'view-hrm-id');
  if (!auth.ok) return auth.res;
  if (side !== 'front' && side !== 'back') return fail_('Loại ảnh không hợp lệ.');

  try {
    var sheet = getIDSheet_(getHRMIDConfig_(), false);
    if (!sheet) return fail_('Không tìm thấy hồ sơ CCCD.');
    var found = findIdRow_(sheet, cleanId_(cccdNo));
    if (found.row < 2 || found.deleted) return fail_('Không tìm thấy hồ sơ CCCD đang hoạt động.');
    var lastColumn = Math.max(sheet.getLastColumn(), 1);
    var row = sheet.getRange(found.row, 1, 1, lastColumn).getValues()[0];
    var createdByIdx = hrmiColIdx_('created_by');
    if (!matrixRowAllowed_(auth.user, 'HRM_ID', row, { owner: createdByIdx })) {
      return { success: false, code: 'FORBIDDEN', message: 'Hồ sơ nằm ngoài phạm vi dữ liệu được phân quyền.' };
    }

    var url = String(row[side === 'front' ? hrmiColIdx_('img_front') : hrmiColIdx_('img_back')] || '');
    var match = url.match(/(?:googleusercontent\.com\/d\/|drive\.google\.com\/(?:file\/d\/|open\?id=))([A-Za-z0-9_-]+)/);
    if (!match) return fail_('Không tìm thấy tệp ảnh CCCD hợp lệ.');
    var file = DriveApp.getFileById(match[1]);
    var blob = file.getBlob();
    var mimeType = blob.getContentType();
    if (!/^image\/(?:jpeg|png|webp)$/i.test(mimeType)) return fail_('Định dạng ảnh CCCD không được hỗ trợ.');
    var bytes = blob.getBytes();
    if (bytes.length > 5000000) return fail_('Ảnh vượt quá giới hạn hiển thị.');
    return { success: true, dataUrl: 'data:' + mimeType + ';base64,' + Utilities.base64Encode(bytes) };
  } catch (err) {
    Logger.log('getHRMIDImage: ' + err);
    return fail_('Không thể tải ảnh CCCD.');
  }
}

/** Chủ dự án chạy một lần để thu hồi quyền chia sẻ liên kết của các ảnh CCCD đã tải lên trước đây. */
function secureExistingCCCDImageSharing() {
  if (!isOwnerRun_()) return { success: false, message: 'Không được phép.' };
  var sheet = getIDSheet_(getHRMIDConfig_(), false);
  if (!sheet || sheet.getLastRow() <= 1) return { success: true, secured: 0, errors: [] };

  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastColumn).getValues();
  var imgFrontIdx = hrmiColIdx_('img_front');
  var imgBackIdx = hrmiColIdx_('img_back');
  var seen = {}, secured = 0, errors = [];
  values.forEach(function (row) {
    [row[imgFrontIdx], row[imgBackIdx]].forEach(function (url) {
      var match = String(url || '').match(/(?:googleusercontent\.com\/d\/|drive\.google\.com\/(?:file\/d\/|open\?id=))([A-Za-z0-9_-]+)/);
      if (!match || seen[match[1]]) return;
      seen[match[1]] = true;
      try {
        DriveApp.getFileById(match[1]).setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
        secured++;
      } catch (err) {
        errors.push(match[1] + ': ' + err.toString());
      }
    });
  });
  return { success: errors.length === 0, secured: secured, errors: errors };
}

function saveHRMIDRecord(token, record) {
  var auth = requireSessionAuth_(token, 'HRM', 'view-hrm-id');
  if (!auth.ok) return auth.res;

  try {
    if (!record || !cleanId_(record.cccd_no)) return fail_('Thiếu số CCCD.');

    return withLock_(function () {
      var cfg = getHRMIDConfig_();
      var sheet = getIDSheet_(cfg, true);
      var cccd = cleanId_(record.cccd_no);
      var now = new Date();
      var userName = auth.user.username;
      var found = findIdRow_(sheet, cccd);

      var rowValues = [
        record.raw_qr ? safeText_(record.raw_qr) : '',
        "'" + cccd,
        record.cmnd_no ? "'" + cleanId_(record.cmnd_no) : '',
        safeText_(record.full_name),
        safeText_(record.gender),
        record.dob || '',
        safeText_(record.perm_addr),
        safeText_(record.merged_addr),
        safeText_(record.curr_addr),
        record.issue_date || '',
        safeText_(record.issue_place) || 'Cục Cảnh sát QLHC về TTXH',
        record.img_front || '',
        record.img_back || '',
        userName,
        now
      ];

      if (found.row > 0) {
        if (!hasPermission_(auth.user, 'HRM', 'update', 'HRM_ID')) {
          return fail_('Bạn không có quyền sửa hồ sơ.', { code: 'FORBIDDEN' });
        }
        var lastColumn = Math.max(sheet.getLastColumn(), 1);
        var old = sheet.getRange(found.row, 1, 1, lastColumn).getValues()[0];
        var createdByIdx = hrmiColIdx_('created_by');
        if (!matrixScopeAllowed_(auth.user, 'HRM_ID', old, { owner: createdByIdx }, null, 'update')) {
          return fail_('Hồ sơ nằm ngoài phạm vi dữ liệu được phân quyền.', { code: 'FORBIDDEN' });
        }
        if (found.deleted && !hasPermission_(auth.user, 'HRM', 'restore', 'HRM_ID')) {
          return fail_('Hồ sơ CCCD này đang nằm trong Thùng rác. Chỉ SuperAdmin được khôi phục.', { code: 'FORBIDDEN' });
        }
        var rawQrIdx = hrmiColIdx_('raw_qr');
        var imgFrontIdx = hrmiColIdx_('img_front');
        var imgBackIdx = hrmiColIdx_('img_back');
        var createdIdx = hrmiColIdx_('created_at');
        if (!rowValues[rawQrIdx]) rowValues[rawQrIdx] = old[rawQrIdx];
        if (!rowValues[imgFrontIdx]) rowValues[imgFrontIdx] = old[imgFrontIdx];
        if (!rowValues[imgBackIdx]) rowValues[imgBackIdx] = old[imgBackIdx];
        rowValues[createdByIdx] = old[createdByIdx] || userName;   // giữ người tạo ban đầu
        rowValues[createdIdx] = old[createdIdx] || now;           // giữ thời gian tạo ban đầu
        if (!matrixScopeAllowed_(auth.user, 'HRM_ID', rowValues, { owner: createdByIdx }, null, 'update')) {
          return fail_('Thông tin mới nằm ngoài phạm vi được phân quyền.', { code: 'FORBIDDEN' });
        }
        var changed = diffFields_(old, rowValues, ['raw_qr', 'cccd_no', 'cmnd_no', 'full_name', 'gender', 'dob', 'perm_addr', 'merged_addr', 'curr_addr', 'issue_date', 'issue_place', 'img_front', 'img_back']);

        sheet.getRange(found.row, 1, 1, lastColumn).setValues([rowValues]);
        if (found.deleted) restoreRows_(sheet, [found.row], idFlagCol_());
        auditLog_(auth.user, 'HRM_ID', found.deleted ? 'RESTORE_UPDATE' : 'UPDATE', cccd, changed ? 'Đổi: ' + changed : 'Không đổi nội dung');
        return { success: true, message: 'Cập nhật hồ sơ thành công!' };
      }

      if (!hasPermission_(auth.user, 'HRM', 'create', 'HRM_ID') ||
          !matrixScopeAllowed_(auth.user, 'HRM_ID', rowValues, { owner: createdByIdx }, null, 'create')) {
        return fail_('Thông tin mới nằm ngoài phạm vi được phân quyền.', { code: 'FORBIDDEN' });
      }
      var lastColumn = Math.max(sheet.getLastColumn(), 1);
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, lastColumn).setValues([rowValues]);
      auditLog_(auth.user, 'HRM_ID', 'CREATE', cccd, rowValues[hrmiColIdx_('full_name')]);
      return { success: true, message: 'Thêm mới hồ sơ thành công!' };
    });
  } catch (err) {
    Logger.log('saveHRMIDRecord: ' + err);
    return fail_('Không thể lưu hồ sơ.');
  }
}

function saveBulkHRMIDRecords(token, recordsList) {
  var auth = requireAuth_(token, 'HRM', 'import', 'HRM_ID', 'view-hrm-id');
  if (!auth.ok) return auth.res;

  try {
    if (!Array.isArray(recordsList) || recordsList.length === 0) return fail_('Không có dữ liệu hợp lệ để nhập.');
    if (recordsList.length > 2000) return fail_('Mỗi lần chỉ nhập tối đa 2000 dòng.');

    return withLock_(function () {
      var sheet = getIDSheet_(getHRMIDConfig_(), true);
      var now = new Date();
      var userName = auth.user.username;

      // Các CCCD đã có: 'A' = đang hoạt động, 'D' = đang trong Thùng rác
      var existing = {};
      var last = sheet.getLastRow();
      if (last > 1) {
        var col = sheet.getRange(2, 2, last - 1, 1).getValues();
        var flags = readDeletedFlags_(sheet, idFlagCol_(), 2, last - 1);
        for (var i = 0; i < col.length; i++) {
          var k = cleanId_(col[i][0]);
          if (!k) continue;
          if (!flags[i]) existing[k] = 'A'; else if (!existing[k]) existing[k] = 'D';
        }
      }

      var rows = [], dupe = [], trash = [], repeat = [], outOfScope = [], invalid = 0, seen = {};
      recordsList.forEach(function (r) {
        var c = r ? cleanId_(r.cccd_no) : '';
        if (!c) { invalid++; return; }
        if (seen[c]) { repeat.push(c); return; }
        seen[c] = true;
        if (existing[c] === 'A') { dupe.push(c); return; }
        if (existing[c] === 'D') { trash.push(c); return; }

        var newRow = [
          r.raw_qr ? safeText_(r.raw_qr) : '',
          "'" + c,
          r.cmnd_no ? "'" + cleanId_(r.cmnd_no) : '',
          safeText_(r.full_name),
          safeText_(r.gender) || 'Nam',
          parseVNFormattedDate_(r.dob) || '',
          safeText_(r.perm_addr),
          safeText_(r.merged_addr),
          safeText_(r.curr_addr),
          parseVNFormattedDate_(r.issue_date) || '',
          safeText_(r.issue_place) || 'Cục Cảnh sát QLHC về TTXH',
          r.img_front || '',
          r.img_back || '',
          userName,
          now
        ];
        if (!matrixScopeAllowed_(auth.user, 'HRM_ID', newRow, { owner: createdByIdx }, null, 'create')) {
          outOfScope.push(c);
          return;
        }
        rows.push(newRow);
      });

      var parts = [];
      if (dupe.length) parts.push(dupe.length + ' trùng CCCD đã có: ' + sampleList_(dupe));
      if (trash.length) parts.push(trash.length + ' đang nằm trong Thùng rác: ' + sampleList_(trash));
      if (repeat.length) parts.push(repeat.length + ' lặp lại trong dữ liệu dán');
      if (outOfScope.length) parts.push(outOfScope.length + ' ngoài phạm vi phân quyền: ' + sampleList_(outOfScope));
      if (invalid) parts.push(invalid + ' thiếu số CCCD');
      var skippedMsg = parts.length ? ' Bỏ qua ' + (dupe.length + trash.length + repeat.length + outOfScope.length + invalid) + ' dòng (' + parts.join('; ') + ').' : '';

      if (rows.length === 0) return fail_('Không có hồ sơ mới để nhập.' + skippedMsg);

      var lastColumn = Math.max(sheet.getLastColumn(), 1);
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, lastColumn).setValues(rows);
      auditLog_(auth.user, 'HRM_ID', 'IMPORT', sampleList_(rows.map(function (x) { return cleanId_(x[1]); }), 5), 'Nhập ' + rows.length + ' hồ sơ, bỏ qua ' + (recordsList.length - rows.length));
      return { success: true, message: 'Đã nhập thành công ' + rows.length + ' hồ sơ.' + skippedMsg };
    });
  } catch (err) {
    Logger.log('saveBulkHRMIDRecords: ' + err);
    return fail_('Không thể nhập dữ liệu từ Excel.');
  }
}

/* ============================================================
 * XÓA (an toàn hơn: chỉ xóa đúng các dòng cần xóa, không xóa-trắng-rồi-ghi-lại)
 * ============================================================ */

function deleteBulkHRMIDRecords(token, cccdList) {
  var auth = requireAuth_(token, 'HRM', 'delete', 'HRM_ID', 'view-hrm-id');
  if (!auth.ok) return auth.res;

  try {
    if (!Array.isArray(cccdList) || cccdList.length === 0) return fail_('Danh sách hồ sơ cần xóa bị trống.');

    return withLock_(function () {
      var sheet = getIDSheet_(getHRMIDConfig_(), false);
      if (!sheet || sheet.getLastRow() <= 1) return fail_('Bảng dữ liệu hiện tại đang trống.');

      var targets = {};
      cccdList.forEach(function (c) { var k = cleanId_(c); if (k) targets[k] = true; });

      var n = sheet.getLastRow() - 1;
      var lastColumn = Math.max(sheet.getLastColumn(), 1);
      var data = sheet.getRange(2, 1, n, lastColumn).getValues();
      var flags = readDeletedFlags_(sheet, idFlagCol_(), 2, n);
      var dataPermission = getEffectivePermission_(auth.user.username, auth.user.role, 'HRM_ID');
      var createdByIdx = hrmiColIdx_('created_by');
      var cccdIdxDel = hrmiColIdx_('cccd_no');
      var rowsToDelete = [], keys = [];
      for (var i = 0; i < n; i++) {
        var k = cleanId_(data[i][cccdIdxDel]);
        if (!flags[i] && targets[k] && matrixScopeAllowed_(auth.user, 'HRM_ID', data[i], { owner: createdByIdx }, dataPermission, 'delete')) {
          rowsToDelete.push(i + 2); keys.push(k);
        }
      }
      if (rowsToDelete.length === 0) return fail_('Không tìm thấy hồ sơ phù hợp để xóa.');

      var cnt = softDeleteRows_(sheet, rowsToDelete, idFlagCol_(), auth.user.username);
      auditLog_(auth.user, 'HRM_ID', 'DELETE', sampleList_(keys, 20), 'Chuyển ' + cnt + ' hồ sơ vào Thùng rác');
      return { success: true, message: 'Đã chuyển ' + cnt + ' hồ sơ vào Thùng rác (SuperAdmin có thể khôi phục).', deletedCount: cnt };
    });
  } catch (err) {
    Logger.log('deleteBulkHRMIDRecords: ' + err);
    return fail_('Không thể xóa dữ liệu.');
  }
}

function deleteHRMIDRecord(token, cccdNo) {
  return deleteBulkHRMIDRecords(token, [cccdNo]);
}

/* ============================================================
 * CHUYỂN HỒ SƠ CCCD SANG DANH SÁCH NHÂN SỰ (HRM_EMP)
 * ============================================================ */

function checkDuplicatesInHRMEMP_(cccdList) {
  try {
    if (!cccdList || !Array.isArray(cccdList) || cccdList.length === 0) return { success: true, duplicates: [] };

    var sheetEmp = getEMPSheet_(false);
    if (!sheetEmp || sheetEmp.getLastRow() <= 1) return { success: true, duplicates: [] };

    var n = sheetEmp.getLastRow() - 1;
    var lastColumn = Math.max(sheetEmp.getLastColumn(), 1);
    var data = sheetEmp.getRange(2, 1, n, lastColumn).getValues();
    var cccdIdx = hreColIdx_('cccd_no');
    var fullNameIdx = hreColIdx_('full_name');
    var flags = readDeletedFlags_(sheetEmp, empFlagCol_(), 2, n);
    var existing = {};
    for (var i = 0; i < data.length; i++) {
      var key = cleanId_(data[i][cccdIdx]);
      if (key && !flags[i]) existing[key] = data[i][fullNameIdx] || '';
    }

    var duplicates = [];
    for (var k = 0; k < cccdList.length; k++) {
      var t = cleanId_(cccdList[k]);
      if (existing[t] !== undefined) duplicates.push({ cccd_no: t, full_name: existing[t] });
    }
    return { success: true, duplicates: duplicates };
  } catch (err) {
    Logger.log('checkDuplicatesInHRMEMP_: ' + err);
    return { success: false, message: 'Không thể kiểm tra trùng lặp dữ liệu.' };
  }
}

function copyHRMIDToHRMEMP(token, cccdList, ignoreDuplicates) {
  var auth = requireAuth_(token, 'HRM', 'view', 'HRM_ID', 'view-hrm-id');
  if (!auth.ok) return auth.res;
  if (!hasPermission_(auth.user, 'HRM', 'ADD_TO_EMPLOYEE', 'HRM_ID')) {
    return { success: false, code: 'FORBIDDEN', message: 'Bạn chưa được cấp quyền tạo nhân sự từ hồ sơ CCCD.' };
  }

  try {
    if (!Array.isArray(cccdList) || cccdList.length === 0) return fail_('Danh sách hồ sơ chọn bị trống.');

    return withLock_(function () {
      var cfg = getHRMIDConfig_();
      var ss = SpreadsheetApp.openById(cfg.dbId);

      var sheetID = ss.getSheetByName(cfg.sheetId);
      if (!sheetID) return fail_('Không tìm thấy bảng dữ liệu Hồ sơ CCCD.');
      var lastRowID = sheetID.getLastRow();
      if (lastRowID <= 1) return fail_('Bảng Hồ sơ CCCD hiện tại chưa có dữ liệu.');

      var targetMap = {};
      cccdList.forEach(function (x) { var k = cleanId_(x); if (k) targetMap[k] = true; });

      var lastColumn = Math.max(sheetID.getLastColumn(), 1);
      var allIDData = sheetID.getRange(2, 1, lastRowID - 1, lastColumn).getValues();
      var idFlags = readDeletedFlags_(sheetID, idFlagCol_(), 2, lastRowID - 1);
      var dataPermission = getEffectivePermission_(auth.user.username, auth.user.role, 'HRM_ID');
      var createdByIdxCopy = hrmiColIdx_('created_by');
      var selected = [];
      for (var i = 0; i < allIDData.length; i++) {
        var row = allIDData[i];
        var cccd = cleanId_(row[hrmiColIdx_('cccd_no')]);
        if (!idFlags[i] && targetMap[cccd] &&
            matrixRowAllowed_(auth.user, 'HRM_ID', row, { owner: createdByIdxCopy }, dataPermission)) {
          selected.push({ cccd_no: cccd, full_name: row[hrmiColIdx_('full_name')] || '', gender: row[hrmiColIdx_('gender')] || '', dob: fmtDate_(row[hrmiColIdx_('dob')]), curr_addr: row[hrmiColIdx_('curr_addr')] || '' });
        }
      }
      if (selected.length === 0) return fail_('Không tìm thấy thông tin hồ sơ tương ứng.');

      if (!ignoreDuplicates) {
        var dup = checkDuplicatesInHRMEMP_(cccdList);
        if (dup.success && dup.duplicates.length > 0) {
          return { success: false, hasDuplicates: true, duplicates: dup.duplicates, message: 'Phát hiện số CCCD đã tồn tại trong Danh sách nhân sự.' };
        }
      }

      var sheetEmp = getEMPSheet_(true);

      // Sinh mã nhân viên tăng dần: 2026 + 4 chữ số (20260001...). Cột B = emp_id.
      var prefix = (typeof CONFIG !== 'undefined' && CONFIG.EMP_ID_PREFIX) ? String(CONFIG.EMP_ID_PREFIX) : '2026';
      var maxSeq = 0;
      var empLastRow = sheetEmp.getLastRow();
      if (empLastRow > 1) {
        var ids = sheetEmp.getRange(2, 2, empLastRow - 1, 1).getValues();
        for (var e = 0; e < ids.length; e++) {
          var idVal = cleanId_(ids[e][0]);
          if (idVal.indexOf(prefix) === 0 && idVal.length === 8) {
            var seq = parseInt(idVal.substring(4), 10);
            if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
          }
        }
      }

      var now = new Date();
      var rowsToAdd = selected.map(function (item) {
        maxSeq++;
        return empObjToRow_({
          cccd_no: item.cccd_no,
          emp_id: prefix + ('0000' + maxSeq).slice(-4),
          full_name: item.full_name,
          dob: item.dob,
          gender: item.gender,
          curr_addr: item.curr_addr
        }, auth.user.username, now, now);
      }).filter(function (row) {
        var empMap = getHRMEMPHeaderMap_();
        return matrixScopeAllowed_(auth.user, 'HRM_EMP', row, { owner: empMap['created_by'], dept: empMap['dept'], status: empMap['status'] }, null, 'create');
      });
      if (!rowsToAdd.length) return fail_('Không có hồ sơ nào nằm trong phạm vi được phân quyền để tạo nhân viên.');

      var empLastCol = Math.max(sheetEmp.getLastColumn(), 1);
      sheetEmp.getRange(sheetEmp.getLastRow() + 1, 1, rowsToAdd.length, empLastCol).setValues(rowsToAdd);
      auditLog_(auth.user, 'HRM_ID', 'COPY_TO_EMP', sampleList_(selected.map(function (x) { return x.cccd_no; }), 10), 'Tạo ' + rowsToAdd.length + ' hồ sơ nhân viên từ CCCD');
      return { success: true, message: 'Đã thêm thành công ' + rowsToAdd.length + ' nhân sự vào Danh sách nhân sự!', addedCount: rowsToAdd.length };
    });
  } catch (err) {
    Logger.log('copyHRMIDToHRMEMP: ' + err);
    return fail_('Lỗi khi cập nhật Danh sách nhân sự.');
  }
}