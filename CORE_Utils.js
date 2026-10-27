/**
 * CORE_Utils.gs - HAP HOLDING ERP SYSTEM
 * Tiện ích dùng chung cho mọi backend (khóa ghi, làm sạch dữ liệu, xóa mềm, xóa dòng an toàn).
 * Tất cả hàm có dấu "_" cuối tên nên trình duyệt không gọi trực tiếp được.
 */

/* ============================================================
 * BỘ NHỚ TẠM THEO REQUEST (giảm số lần đọc sheet trong cùng 1 request)
 * Mỗi lần gọi từ trình duyệt = 1 execution của Apps Script = 1 request.
 * Bộ nhớ tự xoá khi quá REQUEST_MEMO_TTL_MS_ để không dùng dữ liệu quá hạn.
 * ============================================================ */

var REQUEST_MEMO_TTL_MS_ = 60000;
var _requestMemo_ = {};
var _requestMemoAt_ = 0;

/** Lấy (hoặc tính 1 lần rồi ghi nhớ) giá trị theo khóa trong phạm vi 1 request. */
function requestMemo_(key, producer) {
  var now = Date.now();
  if (now - _requestMemoAt_ > REQUEST_MEMO_TTL_MS_) {
    _requestMemo_ = {};
    _requestMemoAt_ = 0;
  }
  if (Object.prototype.hasOwnProperty.call(_requestMemo_, key)) return _requestMemo_[key];
  var value = producer();
  _requestMemo_[key] = value;
  if (!_requestMemoAt_) _requestMemoAt_ = now;
  return value;
}

/** Xoá bộ nhớ tạm. Gọi khi bắt đầu 1 request mới hoặc ngay sau khi ghi dữ liệu phân quyền. */
function clearRequestMemo_() {
  _requestMemo_ = {};
  _requestMemoAt_ = Date.now();
}

/** Chạy fn() trong khóa toàn hệ thống để 2 người không ghi chồng nhau. */
function withLock_(fn, metrics) {
  var waitStartedAt = Date.now();
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    if (metrics) metrics.wait_ms = Date.now() - waitStartedAt;
    return { success: false, message: 'Hệ thống đang xử lý thao tác khác. Vui lòng thử lại sau ít giây.' };
  }
  if (metrics) metrics.wait_ms = Date.now() - waitStartedAt;
  var acquiredAt = Date.now();
  try {
    return fn();
  } finally {
    if (metrics) metrics.hold_ms = Date.now() - acquiredAt;
    lock.releaseLock();
  }
}

/** Bỏ dấu nháy đơn (') và khoảng trắng ở mã định danh (CCCD, mã NV...) */
function cleanId_(v) {
  return v == null ? '' : v.toString().replace(/'/g, '').trim();
}

/** Chặn "chèn công thức": ô bắt đầu bằng = + - @ sẽ bị Sheets hiểu là công thức */
function safeText_(v) {
  if (v == null) return '';
  var s = v.toString();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function fmtDate_(v) {
  if (!v) return '';
  return v instanceof Date ? Utilities.formatDate(v, 'GMT+7', 'yyyy-MM-dd') : v.toString();
}

function fmtDateTime_(v) {
  if (!v) return '';
  var d = (v instanceof Date) ? v : new Date(v);
  return isNaN(d.getTime()) ? v.toString() : Utilities.formatDate(d, 'GMT+7', 'dd/MM/yyyy HH:mm');
}

function fail_(msg, extra) {
  var r = { success: false, message: msg };
  if (extra) for (var k in extra) r[k] = extra[k];
  return r;
}

/** Liệt kê tối đa n phần tử đầu của mảng để đưa vào thông báo / nhật ký */
function sampleList_(arr, n) {
  n = n || 3;
  var s = arr.slice(0, n).join(', ');
  return arr.length > n ? s + '...' : s;
}

/** So sánh 2 dòng, trả về tên các cột có thay đổi (để ghi nhật ký) */
function diffFields_(oldRow, newRow, keys) {
  var norm = function (v) {
    if (v instanceof Date) return fmtDate_(v);
    return v == null ? '' : String(v).replace(/^'/, '').trim();
  };
  var out = [];
  for (var i = 0; i < keys.length; i++) {
    if (norm(oldRow[i]) !== norm(newRow[i])) out.push(keys[i]);
  }
  return out.join(',');
}

/** Đọc header hàng 1 của sheet dưới dạng giá trị thực tế (getValues), không dùng getDisplayValues. */
function getSheetHeaders_(sheet) {
  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  return sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
    .map(function(header) { return String(header || '').trim(); });
}

/** Lấy chỉ số cột theo tên header (trả về -1 nếu không tìm thấy). */
function getColumnIndexByName_(sheet, headerName) {
  var headers = getSheetHeaders_(sheet);
  return headers.indexOf(String(headerName || '').trim());
}

/** Gom các số dòng đã sắp xếp thành các đoạn liên tiếp [[bắt đầu, số dòng], ...] */
function groupRuns_(sortedRows) {
  var runs = [];
  if (!sortedRows.length) return runs;
  var start = sortedRows[0], count = 1;
  for (var j = 1; j < sortedRows.length; j++) {
    if (sortedRows[j] === sortedRows[j - 1] + 1) { count++; }
    else { runs.push([start, count]); start = sortedRows[j]; count = 1; }
  }
  runs.push([start, count]);
  return runs;
}

/**
 * Xóa hẳn các dòng (số dòng thật trong sheet): gom đoạn liên tiếp rồi xóa từ dưới lên.
 * Chỉ dùng cho "Xóa vĩnh viễn" trong Thùng rác.
 */
function deleteRowsByIndexes_(sheet, rowNumbers) {
  if (!rowNumbers || rowNumbers.length === 0) return 0;
  var rows = rowNumbers.slice().sort(function (a, b) { return a - b; });
  var runs = groupRuns_(rows);
  for (var k = runs.length - 1; k >= 0; k--) {
    sheet.deleteRows(runs[k][0], runs[k][1]);
  }
  return rows.length;
}

/* ============================================================
 * XÓA MỀM (3 cột cuối bảng: is_deleted | deleted_by | deleted_at)
 * ============================================================ */

/** Đảm bảo sheet có đủ 3 cột xóa mềm bắt đầu từ cột flagCol1 (số cột, tính từ 1) */
function ensureSoftCols_(sheet, flagCol1) {
  var need = flagCol1 + 2;
  if (sheet.getMaxColumns() < need) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), need - sheet.getMaxColumns());
  }
  var hdr = sheet.getRange(1, flagCol1, 1, 3).getValues()[0];
  if (String(hdr[0]).trim() !== 'is_deleted') {
    sheet.getRange(1, flagCol1, 1, 3).setValues([['is_deleted', 'deleted_by', 'deleted_at']]);
  }
}

/** Đọc cờ xóa mềm của nRows dòng liên tiếp. Chưa có cột cờ -> coi như chưa xóa dòng nào. */
function readDeletedFlags_(sheet, flagCol1, firstRow, nRows) {
  var out = [], i;
  if (nRows <= 0) return out;
  if (sheet.getLastColumn() < flagCol1) {
    for (i = 0; i < nRows; i++) out.push(false);
    return out;
  }
  var v = sheet.getRange(firstRow, flagCol1, nRows, 1).getValues();
  for (i = 0; i < v.length; i++) out.push(String(v[i][0]).trim().toUpperCase() === 'Y');
  return out;
}

/** Đánh dấu xóa mềm các dòng. Trả về số dòng đã đánh dấu. */
function softDeleteRows_(sheet, rowNumbers, flagCol1, username) {
  if (!rowNumbers.length) return 0;
  ensureSoftCols_(sheet, flagCol1);
  var rows = rowNumbers.slice().sort(function (a, b) { return a - b; });
  var now = new Date();
  groupRuns_(rows).forEach(function (run) {
    var vals = [];
    for (var i = 0; i < run[1]; i++) vals.push(['Y', username, now]);
    sheet.getRange(run[0], flagCol1, run[1], 3).setValues(vals);
  });
  return rows.length;
}

/** Bỏ đánh dấu xóa (khôi phục) */
function restoreRows_(sheet, rowNumbers, flagCol1) {
  var rows = rowNumbers.slice().sort(function (a, b) { return a - b; });
  groupRuns_(rows).forEach(function (run) {
    var vals = [];
    for (var i = 0; i < run[1]; i++) vals.push(['', '', '']);
    sheet.getRange(run[0], flagCol1, run[1], 3).setValues(vals);
  });
  return rows.length;
}