/**
 * DMS_Server.js - Phân hệ Quản lý Tài liệu (DMS)
 */
function dmsEnsureDocsColumns() {
  var desired = ['doc_id','doc_code','entity_code','activity_code','domain_code','doc_type_code','seq_no','doc_title','category','dept_code','current_version','file_url','file_id','status','effective_date','owner_emp_id','approver_emp_id','is_deleted','deleted_by','deleted_at','created_by','created_at','updated_by','updated_at'];
  var ss = SpreadsheetApp.openById(CONFIG.DB_DMS_PART1_ID);
  var sh = ss.getSheetByName(CONFIG.SHEET_DMS_DOCS);
  if (!sh) { sh = ss.insertSheet(CONFIG.SHEET_DMS_DOCS); sh.appendRow(desired); return 'Created sheet'; }
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var missing = desired.filter(function(h) { return headers.indexOf(h) === -1; });
  if (missing.length) {
    var startCol = headers.length + 1;
    sh.getRange(1, startCol, 1, missing.length).setValues([missing]);
    return 'Added: ' + missing.join(', ');
  }
  return 'All columns present';
}
function dmsUploadFile(token, name, mimeType, base64, folderId) {
  try {
    var bytes = Utilities.base64Decode(base64);
    var blob = Utilities.newBlob(bytes, mimeType || 'application/octet-stream', name);
    var folder = DriveApp.getFolderById(folderId);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, url: file.getUrl(), fileId: file.getId(), message: 'Đã upload file' };
  } catch (e) {
    return { success: false, message: 'Upload lỗi: ' + e.toString() };
  }
}
function dmsListEmployees(token) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.DB_HRM_PART1_ID);
    var sh = ss.getSheetByName(CONFIG.SHEET_HRM_EMP);
    if (!sh) return { success: true, list: [] };
    var rows = sh.getDataRange().getValues();
    var headers = rows[0];
    var ei = headers.indexOf('emp_id');
    var fi = headers.indexOf('full_name');
    if (ei < 0) ei = 0;
    if (fi < 0) fi = 1;
    var list = rows.slice(1).map(function(r) { return { emp_id: r[ei], full_name: r[fi] }; }).filter(function(x) { return x.emp_id; });
    return { success: true, list: list };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
function dmsAuth_(token, sheetName, action) {
  var actionMap = { read: 'view', edit: 'update', create: 'create', delete: 'delete' };
  var sysAction = actionMap[action] || action || 'view';
  var auth = requireSessionAuth_(token, 'DMS', null);
  if (!auth.ok) return { ok: false, result: auth.res };
  var perm = getEffectivePermission_(auth.user.username, auth.user.role, sheetName || 'DMS_DOCS');
  if (!perm.canRead && !perm.canCreate && !perm.canEdit && !perm.canDelete) {
    perm = { canRead: true, canCreate: true, canEdit: true, canDelete: true };
  }
  return { ok: true, user: auth.user, permission: {
    canRead: !!perm.canRead, canCreate: !!perm.canCreate, canEdit: !!perm.canEdit, canDelete: !!perm.canDelete
  } };
}

function dmsGetSheet_(sheetConfigKey, headers) {
  var ss = SpreadsheetApp.openById(CONFIG.DB_DMS_PART1_ID);
  var name = CONFIG[sheetConfigKey];
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); if (headers) sh.appendRow(headers); }
  return { sheet: sh, headers: sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] };
}

function dmsCellValue_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, 'GMT+7', 'dd/MM/yyyy HH:mm:ss');
  if (value == null) return '';
  return typeof value === 'string' ? value.replace(/^'/, '') : value;
}

function dmsRowsAsObjects_(headers, rows) {
  return rows.map(function(row) {
    var item = {};
    headers.forEach(function(header, index) {
      if (header) item[header] = dmsCellValue_(row[index]);
    });
    return item;
  });
}

function dmsGetEmployees_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('dms_employees');
  if (cached) { try { return JSON.parse(cached); } catch(e){} }
  var ss = SpreadsheetApp.openById(CONFIG.DB_HRM_PART1_ID);
  var sh = ss.getSheetByName(CONFIG.SHEET_HRM_EMP);
  if (!sh) return [];
  var rows = sh.getDataRange().getValues();
  var headers = rows[0];
  var ei = headers.indexOf('emp_id'); if (ei < 0) ei = 0;
  var fi = headers.indexOf('full_name'); if (fi < 0) fi = 1;
  var list = rows.slice(1).map(function(r){ return { emp_id: r[ei], full_name: r[fi] }; }).filter(function(x){ return x.emp_id; });
  cache.put('dms_employees', JSON.stringify(list), 600);
  return list;
}

function dmsEmpUsernameMap_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('dms_empusrmap');
  if (cached) { try { return JSON.parse(cached); } catch(e){} }
  var ss = SpreadsheetApp.openById(CONFIG.DB_HRM_PART1_ID);
  var sh = ss.getSheetByName(CONFIG.SHEET_HRM_EMP);
  var rows = sh.getDataRange().getValues();
  var h = rows[0];
  var ei = h.indexOf('emp_id'); if (ei<0) ei=0;
  var ci = h.indexOf('cccd_no'); if (ci<0) ci=-1;
  var cccdByEmp = {};
  for (var i=1;i<rows.length;i++){ if(rows[i][ei]) cccdByEmp[String(rows[i][ei])]= (ci>=0?rows[i][ci]:''); }
  var ss2 = SpreadsheetApp.openById(CONFIG.DB_SYS_PART1_ID);
  var sh2 = ss2.getSheetByName(CONFIG.SHEET_SYS_USR);
  var rows2 = sh2.getDataRange().getValues();
  var h2 = rows2[0];
  var ci2 = h2.indexOf('cccd_no'); if (ci2<0) ci2=-1;
  var ui = h2.indexOf('username'); if (ui<0) ui=0;
  var userByCccd = {};
  for (var j=1;j<rows2.length;j++){ if (ci2>=0 && rows2[j][ci2]) userByCccd[String(rows2[j][ci2])]=rows2[j][ui]; }
  var map = {};
  Object.keys(cccdByEmp).forEach(function(emp){ map[emp]= userByCccd[String(cccdByEmp[emp])] || ''; });
  cache.put('dms_empusrmap', JSON.stringify(map), 600);
  return map;
}

function dmsListDocs(token, page, pageSize) {
  try {
  var auth = requireSessionAuth_(token, 'DMS', null);
  var username = (auth && auth.ok) ? auth.user.username : 'admin';
  var isSuper = (auth && auth.ok) ? (auth.user.role === 'SuperAdmin' || String(username).toLowerCase() === 'admin') : true;
  var info = dmsGetSheet_('SHEET_DMS_DOCS', ['doc_id','doc_code','entity_code','activity_code','domain_code','doc_type_code','seq_no','doc_title','status','effective_date','owner_emp_id','reviewer_emp_id','approver_emp_id','file_url','file_id','is_deleted','deleted_by','deleted_at','created_by','created_at','updated_by','updated_at','approval_log']);
  var rows = info.sheet.getDataRange().getValues();
  var headers = rows[0];
  var data = rows.slice(1).filter(function(r) {
    var delVal = r[headers.indexOf('is_deleted')];
    return delVal !== true && delVal !== 'true' && delVal !== 1 && String(delVal).toLowerCase() !== 'true';
  });
  var list = dmsRowsAsObjects_(headers, data);
  var umap = dmsEmpUsernameMap_();
  list.forEach(function(d) {
    var st = String(d.status || '').toUpperCase();
    var ownerUser = umap[String(d.owner_emp_id)] || '';
    var reviewerUser = umap[String(d.reviewer_emp_id)] || '';
    var approverUser = umap[String(d.approver_emp_id)] || '';
    var can = isSuper;
    if (st === 'DRAFT') can = isSuper || (username === ownerUser);
    else if (st === 'PENDING_REVIEW') can = isSuper || (username === reviewerUser);
    else if (st === 'PENDING_APPROVAL') can = isSuper || (username === approverUser);
    else can = false;
    d.canApprove = can;
  });
  page = Math.max(1, parseInt(page, 10) || 1);
  pageSize = Math.min(100, Math.max(1, parseInt(pageSize, 10) || 30));
  var total = list.length;
  var pages = Math.max(1, Math.ceil(total / pageSize));
  page = Math.min(page, pages);
  var pageRows = list.slice((page - 1) * pageSize, page * pageSize);
  var perm = (auth && auth.ok && auth.permission) ? auth.permission : { canRead:true, canCreate:true, canEdit:true, canDelete:true };
  return { success: true, list: pageRows, total: total, page: page, page_size: pageSize, page_count: pages, employees: dmsGetEmployees_(),   permissions: { canRead: !!perm.canRead, canCreate: !!perm.canCreate, canEdit: !!perm.canEdit, canDelete: !!perm.canDelete, isSuper: isSuper } };
  } catch (e) {
    return { success: false, message: 'DMS load error: ' + (e && e.message ? e.message : e) };
  }
}

function dmsSaveDoc(token, record) {
  return withLock_(function() {
  try {
    var auth = requireSessionAuth_(token, 'DMS', null);
    var username = (auth && auth.ok) ? auth.user.username : 'admin';
    var isSuper = (auth && auth.ok) ? (auth.user.role === 'SuperAdmin' || String(username).toLowerCase() === 'admin') : true;
    if (!auth.ok) return auth.res;
  var info = dmsGetSheet_('SHEET_DMS_DOCS', ['doc_id','doc_code','entity_code','activity_code','domain_code','doc_type_code','seq_no','doc_title','status','effective_date','owner_emp_id','reviewer_emp_id','approver_emp_id','file_url','file_id','is_deleted','deleted_by','deleted_at','created_by','created_at','updated_by','updated_at']);
  var sh = info.sheet, headers = info.headers;
  var now = new Date();
  var user = auth.user.username;
  if (!record.doc_id) {
    record.doc_id = 'DMS-' + Date.now();
    // generate doc_code
    if (!record.seq_no) {
      var maxSeq = 0;
      var allRows = sh.getDataRange().getValues();
      var ecIdx = headers.indexOf('entity_code');
      var acIdx = headers.indexOf('activity_code');
      var dcIdx = headers.indexOf('domain_code');
      var tcIdx = headers.indexOf('doc_type_code');
      var sqIdx = headers.indexOf('seq_no');
      for (var r = 1; r < allRows.length; r++) {
        if (String(allRows[r][ecIdx]) === String(record.entity_code) &&
            String(allRows[r][acIdx]) === String(record.activity_code) &&
            String(allRows[r][dcIdx]) === String(record.domain_code) &&
            String(allRows[r][tcIdx]) === String(record.doc_type_code)) {
          var sq = parseInt(allRows[r][sqIdx], 10);
          if (!isNaN(sq) && sq > maxSeq) maxSeq = sq;
        }
      }
      record.seq_no = maxSeq + 1;
    }
    var seqStr = (parseInt(record.seq_no,10) < 10 ? '0' + record.seq_no : String(record.seq_no));
    record.doc_code = [record.entity_code, record.activity_code, record.domain_code, record.doc_type_code, seqStr].join('.');
    record.created_by = user;
    record.created_at = now;
    record.is_deleted = false;
    record.current_version = record.current_version || 'v1.0';
    record.status = record.status || 'DRAFT';
    var row = headers.map(function(h) { return record[h] !== undefined ? record[h] : ''; });
    sh.appendRow(row);
    return { success: true, message: 'Đã tạo tài liệu' };
  }
  // update
  var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  var curRowValues = sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(record.doc_id)) {
      var curStatus = String(curRowValues[i][headers.indexOf('status')] || '').toUpperCase();
      if (curStatus !== 'DRAFT' && !isSuper) return { success: false, message: 'Chỉ được sửa khi trạng thái Bản thảo' };
      record.updated_by = user;
      record.updated_at = now;
      var row = headers.map(function(h) { return record[h] !== undefined ? record[h] : curRowValues[i][headers.indexOf(h)]; });
      sh.getRange(i + 2, 1, 1, headers.length).setValues([row]);
      return { success: true, message: 'Đã cập nhật tài liệu' };
    }
  }
  return { success: false, message: 'Không tìm thấy tài liệu' };
  } catch (err) {
    Logger.log('dmsSaveDoc error: ' + err);
    return { success: false, message: 'Lỗi: ' + (err && err.message ? err.message : err) };
  }
  });
}

function dmsListVersions(token, docId) {
  var auth = dmsAuth_(token, 'DMS_VERSIONS', 'read');
  if (!auth.ok) return auth.result;
  var info = dmsGetSheet_('SHEET_DMS_VERSIONS', ['version_id','doc_id','version_tag','change_log','file_url','file_id','created_by','created_at']);
  var rows = info.sheet.getDataRange().getValues();
  var headers = rows[0];
  var data = rows.slice(1).filter(function(r) { return String(r[headers.indexOf('doc_id')]) === String(docId); });
  return { success: true, list: dmsRowsAsObjects_(headers, data) };
}

function dmsAddVersion(token, record) {
  var auth = dmsAuth_(token, 'DMS_VERSIONS', 'edit');
  if (!auth.ok) return auth.result;
  var info = dmsGetSheet_('SHEET_DMS_VERSIONS', ['version_id','doc_id','version_tag','change_log','file_url','file_id','created_by','created_at']);
  var sh = info.sheet, headers = info.headers;
  record.version_id = 'VER-' + Date.now();
  record.created_by = record.created_by || 'admin';
  record.created_at = new Date();
  var row = headers.map(function(h) { return record[h] !== undefined ? record[h] : ''; });
  sh.appendRow(row);
  return { success: true, message: 'Đã thêm phiên bản', version: record };
}

function dmsListComments(token, docId) {
  var auth = dmsAuth_(token, 'DMS_COMMENTS', 'read');
  if (!auth.ok) return auth.result;
  var info = dmsGetSheet_('SHEET_DMS_COMMENTS', ['comment_id','doc_id','version_tag','comment_type','content','created_by','created_at','resolution_status']);
  var rows = info.sheet.getDataRange().getValues();
  var headers = rows[0];
  var data = rows.slice(1).filter(function(r) { return String(r[headers.indexOf('doc_id')]) === String(docId); });
  return { success: true, list: dmsRowsAsObjects_(headers, data) };
}

function dmsAddComment(token, record) {
  var auth = dmsAuth_(token, 'DMS_COMMENTS', 'edit');
  if (!auth.ok) return auth.result;
  var info = dmsGetSheet_('SHEET_DMS_COMMENTS', ['comment_id','doc_id','version_tag','comment_type','content','created_by','created_at','resolution_status']);
  var sh = info.sheet, headers = info.headers;
  record.comment_id = 'CMT-' + Date.now();
  record.created_at = new Date();
  record.resolution_status = record.resolution_status || 'OPEN';
  var row = headers.map(function(h) { return record[h] !== undefined ? record[h] : ''; });
  sh.appendRow(row);
  return { success: true, message: 'Đã thêm bình luận', comment: record };
}

function dmsEmpUsername_(empId) {
  if (!empId) return '';
  var map = dmsEmpUsernameMap_();
  return map[String(empId)] || '';
}

function dmsApproveDoc(token, docId) {
  return withLock_(function() {
  var auth = requireSessionAuth_(token, 'DMS', null);
  if (!auth.ok) return auth.res;
  var username = auth.user.username;
  var info = dmsGetSheet_('SHEET_DMS_DOCS', ['doc_id','status','owner_emp_id','reviewer_emp_id','approver_emp_id','approval_log']);
  var sh = info.sheet, headers = info.headers;
  var rows = sh.getDataRange().getValues();
  var idIdx = headers.indexOf('doc_id');
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][idIdx]) !== String(docId)) continue;
    var status = String(rows[i][headers.indexOf('status')] || '').toUpperCase();
    var reviewerEmp = rows[i][headers.indexOf('reviewer_emp_id')];
    var approverEmp = rows[i][headers.indexOf('approver_emp_id')];
    var reviewerUser = dmsEmpUsername_(reviewerEmp);
    var approverUser = dmsEmpUsername_(approverEmp);
    var ownerEmp = rows[i][headers.indexOf('owner_emp_id')];
    var ownerUser = dmsEmpUsername_(ownerEmp);
    var nextStatus = '';
    if (status === 'DRAFT') {
      if (!(auth.user.role === 'SuperAdmin' || String(auth.user.username).toLowerCase() === 'admin') && username !== ownerUser) return { success: false, message: 'Chỉ người soạn được duyệt bước này' };
      nextStatus = 'PENDING_REVIEW';
    } else if (status === 'PENDING_REVIEW') {
      if (!(auth.user.role === 'SuperAdmin' || String(auth.user.username).toLowerCase() === 'admin') && username !== reviewerUser) return { success: false, message: 'Chỉ người xem xét được duyệt bước này' };
      nextStatus = 'PENDING_APPROVAL';
    } else if (status === 'PENDING_APPROVAL') {
      if (!(auth.user.role === 'SuperAdmin' || String(auth.user.username).toLowerCase() === 'admin') && username !== approverUser) return { success: false, message: 'Chỉ người duyệt được duyệt bước này' };
      nextStatus = 'APPROVED';
    } else {
      return { success: false, message: 'Tài liệu không ở trạng thái chờ duyệt' };
    }
    var logIdx = headers.indexOf('approval_log');
    var oldLog = logIdx >= 0 ? (rows[i][logIdx] || '') : '';
    var newLog = oldLog + (oldLog ? '\n' : '') + new Date().toISOString() + ' | ' + username + ' | ' + nextStatus;
    var stIdx = headers.indexOf('status');
    sh.getRange(i + 1, stIdx + 1).setValue(nextStatus);
    if (logIdx >= 0) sh.getRange(i + 1, logIdx + 1).setValue(newLog);
    return { success: true, message: 'Đã duyệt: ' + nextStatus, status: nextStatus };
  }
  return { success: false, message: 'Không tìm thấy tài liệu' };
  });
}

function dmsRejectDoc(token, docId, reason) {
  return withLock_(function() {
  var auth = requireSessionAuth_(token, 'DMS', null);
  if (!auth.ok) return auth.res;
  var username = auth.user.username;
  var info = dmsGetSheet_('SHEET_DMS_DOCS', ['doc_id','status','owner_emp_id','reviewer_emp_id','approver_emp_id','approval_log']);
  var sh = info.sheet, headers = info.headers;
  var rows = sh.getDataRange().getValues();
  var idIdx = headers.indexOf('doc_id');
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][idIdx]) !== String(docId)) continue;
    var status = String(rows[i][headers.indexOf('status')] || '').toUpperCase();
    var ownerEmp = rows[i][headers.indexOf('owner_emp_id')];
    var reviewerEmp = rows[i][headers.indexOf('reviewer_emp_id')];
    var approverEmp = rows[i][headers.indexOf('approver_emp_id')];
    var ownerUser = dmsEmpUsername_(ownerEmp);
    var reviewerUser = dmsEmpUsername_(reviewerEmp);
    var approverUser = dmsEmpUsername_(approverEmp);
    var allowed = (auth.user.role === 'SuperAdmin' || String(auth.user.username).toLowerCase() === 'admin');
    if (status === 'DRAFT') allowed = allowed || (username === ownerUser);
    else if (status === 'PENDING_REVIEW') allowed = allowed || (username === reviewerUser);
    else if (status === 'PENDING_APPROVAL') allowed = allowed || (username === approverUser);
    if (!allowed) return { success: false, message: 'Bạn không có quyền từ chối bước này' };
    var logIdx = headers.indexOf('approval_log');
    var oldLog = logIdx >= 0 ? (rows[i][logIdx] || '') : '';
    var newLog = oldLog + (oldLog ? '\n' : '') + new Date().toISOString() + ' | ' + username + ' | REJECTED' + (reason ? ' - ' + reason : '');
    var stIdx = headers.indexOf('status');
    sh.getRange(i + 1, stIdx + 1).setValue('DRAFT');
    if (logIdx >= 0) sh.getRange(i + 1, logIdx + 1).setValue(newLog);
    return { success: true, message: 'Đã từ chối, trả về Bản thảo' };
  }
  return { success: false, message: 'Không tìm thấy tài liệu' };
  });
}

function dmsDeleteDoc(token, docId) {
  return withLock_(function() {
  var auth = dmsAuth_(token, 'DMS_DOCS', 'delete');
  if (!auth.ok) return auth.result;
  var info = dmsGetSheet_('SHEET_DMS_DOCS', ['doc_id']);
  var sh = info.sheet, headers = info.headers;
  var idIdx = headers.indexOf('doc_id');
  var delIdx = headers.indexOf('is_deleted');
  var byIdx = headers.indexOf('deleted_by');
  var atIdx = headers.indexOf('deleted_at');
  var rows = sh.getDataRange().getValues();
   for (var i = 1; i < rows.length; i++) {
     if (String(rows[i][idIdx]) === String(docId)) {
       var curSt = String(rows[i][headers.indexOf('status')] || '').toUpperCase();
       var isSuperDel = (String(auth.user.username).toLowerCase() === 'admin');
       if (curSt !== 'DRAFT' && !isSuperDel) return { success: false, message: 'Chỉ được xóa khi trạng thái Bản thảo' };
      if (delIdx >= 0) sh.getRange(i + 1, delIdx + 1).setValue(true);
      if (byIdx >= 0) sh.getRange(i + 1, byIdx + 1).setValue(auth.user.username);
      if (atIdx >= 0) sh.getRange(i + 1, atIdx + 1).setValue(new Date());
      return { success: true, message: 'Đã đánh dấu xóa tài liệu' };
    }
  }
  return { success: false, message: 'Không tìm thấy tài liệu' };
  });
}

function dmsDeleteVersion(token, versionId) {
  var auth = dmsAuth_(token, 'DMS_VERSIONS', 'delete');
  if (!auth.ok) return auth.result;
  var info = dmsGetSheet_('SHEET_DMS_VERSIONS', ['version_id']);
  var sh = info.sheet, headers = info.headers;
  var idIdx = headers.indexOf('version_id');
  var rows = sh.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][idIdx]) === String(versionId)) {
      sh.deleteRow(i + 1);
      return { success: true, message: 'Đã xóa phiên bản' };
    }
  }
  return { success: false, message: 'Không tìm thấy phiên bản' };
}

function dmsListReadStatus(token, docId) {
  var auth = dmsAuth_(token, 'DMS_READ_ACK', 'read');
  if (!auth.ok) return auth.result;
  var info = dmsGetSheet_('SHEET_DMS_READ_ACK', ['ack_id','doc_id','version_tag','emp_id','username','acknowledged_at']);
  var rows = info.sheet.getDataRange().getValues();
  var headers = rows[0];
  var data = rows.slice(1).filter(function(r) { return String(r[headers.indexOf('doc_id')]) === String(docId); });
  return { success: true, list: dmsRowsAsObjects_(headers, data) };
}
function dmsAcknowledgeRead(token, docId, versionTag, username) {
  var auth = dmsAuth_(token, 'DMS_READ_ACK', 'edit');
  if (!auth.ok) return auth.result;
  var info = dmsGetSheet_('SHEET_DMS_READ_ACK', ['ack_id','doc_id','version_tag','emp_id','username','acknowledged_at']);
  var sh = info.sheet, headers = info.headers;
  var record = { ack_id: 'ACK-' + Date.now(), doc_id: docId, version_tag: versionTag, username: username, acknowledged_at: new Date() };
  var row = headers.map(function(h) { return record[h] !== undefined ? record[h] : ''; });
  sh.appendRow(row);
  return { success: true, message: 'Đã ghi nhận' };
}



