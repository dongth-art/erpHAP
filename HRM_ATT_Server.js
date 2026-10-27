/**
 * ============================================================================
 * HRM_ATT_Server.gs - SERVER XỬ LÝ CHẤM CÔNG, NGHỈ PHÉP & BATCH IMPORT EXCEL
 * ============================================================================
 */

function getHRMATTConfig_() {
  var globalConfig = (typeof CONFIG !== 'undefined') ? CONFIG : {};
  return {
    DB_SYS_PART1_ID: globalConfig.DB_SYS_PART1_ID || "",
    SHEET_SYS_USR: globalConfig.SHEET_SYS_USR || "SYS_USR",
    DB_HRM_PART1_ID: globalConfig.DB_HRM_PART1_ID || "",
    SHEET_HRM_ATT: globalConfig.SHEET_HRM_ATT || "HRM_ATT",
    SHEET_HRM_EMP: globalConfig.SHEET_HRM_EMP || "HRM_EMP"
  };
}

function buildUniqueKey_(empId, workDate, workShift) {
  var cleanEmp = (empId || "").toString().replace(/[\u200B-\u200D\uFEFF]/g, "").replace(/'/g, "").trim().toLowerCase();
  var cleanShift = (workShift || "Hành chính").toString().replace(/[\u200B-\u200D\uFEFF]/g, "").replace(/'/g, "").trim().toLowerCase();
  var dNum = dateToYYYYMMDD_(workDate);
  return cleanEmp + "_" + dNum + "_" + cleanShift;
}

function cleanCompareKey_(val) {
  if (val === null || val === undefined) return "";
  return val.toString().replace(/[\u200B-\u200D\uFEFF]/g, "").replace(/'/g, "").trim().toLowerCase();
}

function getRawText_(val) {
  if (val === null || val === undefined) return "";
  return val.toString().replace(/'/g, "").trim();
}

function dateToYYYYMMDD_(val) {
  if (!val && val !== 0) return 0;
  var d = null;

  if (val instanceof Date) {
    d = val;
  } else {
    var str = cleanCompareKey_(val);
    if (!str) return 0;
    if (str.indexOf("-") !== -1) {
      var p = str.split("-");
      if (p.length >= 3) d = new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
    } else if (str.indexOf("/") !== -1) {
      var p = str.split("/");
      if (p.length >= 3) d = new Date(parseInt(p[2], 10), parseInt(p[1], 10) - 1, parseInt(p[0], 10));
    } else {
      d = new Date(str);
    }
  }

  if (d && !isNaN(d.getTime())) {
    var yyyy = Utilities.formatDate(d, "GMT+7", "yyyy");
    var mm = Utilities.formatDate(d, "GMT+7", "MM");
    var dd = Utilities.formatDate(d, "GMT+7", "dd");
    return parseInt(yyyy + mm + dd, 10);
  }
  return 0;
}

function getHeaderMap_(sheet) {
  return requestMemo_('HRM_ATT_HEADER_MAP_' + sheet.getSheetName(), function () {
    var lastCol = sheet.getLastColumn();
    if (lastCol === 0) return {};
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var map = {};
    for (var i = 0; i < headers.length; i++) {
      var hKey = cleanCompareKey_(headers[i]);
      if (hKey) map[hKey] = i;
    }
    return map;
  });
}

function formatDateTimeValue_(val, type) {
  if (!val && val !== 0) return "";
  if (val instanceof Date) {
    if (type === "date") return Utilities.formatDate(val, "GMT+7", "yyyy-MM-dd");
    if (type === "month") return Utilities.formatDate(val, "GMT+7", "MM/yyyy");
    if (type === "time") return Utilities.formatDate(val, "GMT+7", "HH:mm:ss");
    if (type === "full") return Utilities.formatDate(val, "GMT+7", "dd/MM/yyyy HH:mm:ss");
  }
  var str = getRawText_(val);
  if (str.indexOf("GMT") !== -1 || str.indexOf("Indochina Time") !== -1) {
    var d = new Date(str);
    if (!isNaN(d.getTime())) {
      if (type === "date") return Utilities.formatDate(d, "GMT+7", "yyyy-MM-dd");
      if (type === "month") return Utilities.formatDate(d, "GMT+7", "MM/yyyy");
      if (type === "time") return Utilities.formatDate(d, "GMT+7", "HH:mm:ss");
      if (type === "full") return Utilities.formatDate(d, "GMT+7", "dd/MM/yyyy HH:mm:ss");
    }
  }
  return str;
}

/**
 * Hồ sơ nhân viên của CHÍNH người đang đăng nhập (lấy CCCD liên kết từ phiên, không nhận username từ client).
 */
function getEmployeeProfileByUsername(token) {
  var auth = requireAuth_(token, 'HRM', 'self', 'HRM_EMP', 'view-hrm-att');
  if (!auth.ok) return auth.res;

  try {
    if (auth.user.cccd === undefined) {
      return fail_('Phiên đăng nhập cũ chưa có thông tin CCCD. Vui lòng đăng xuất và đăng nhập lại.');
    }
    var cccdNo = cleanId_(auth.user.cccd);
    if (!cccdNo) return fail_("Tài khoản '" + auth.user.username + "' chưa được liên kết số CCCD!");

    var emp = empFindByCccd_(cccdNo);
    if (emp) return { success: true, data: emp, cccd_no: cccdNo };

    return fail_('Không tìm thấy Hồ sơ Nhân viên tương ứng với CCCD (' + cccdNo + ')!');
  } catch (err) {
    Logger.log('getEmployeeProfileByUsername: ' + err);
    return fail_('Lỗi dò tìm hồ sơ nhân viên.');
  }
}

function fetchHRMATTListPaginated(token, page, pageSize) {
  var auth = requireAuth_(token, 'HRM', 'view', 'HRM_ATT', 'view-hrm-att');
  if (!auth.ok) return Object.assign({ totalRows: 0 }, auth.res);
  var dataPermission = getEffectivePermission_(auth.user.username, auth.user.role, 'HRM_ATT');
  try {
    page = Math.max(1, parseInt(page, 10) || 1);
    pageSize = Math.min(3000, Math.max(1, parseInt(pageSize, 10) || 20));

    var cfg = getHRMATTConfig_();
    if (!cfg.DB_HRM_PART1_ID) return { success: false, message: "Chưa cấu hình DB Part 1", list: [], totalRows: 0 };

    var ss = SpreadsheetApp.openById(cfg.DB_HRM_PART1_ID);
    var sheet = ss.getSheetByName(cfg.SHEET_HRM_ATT);
    if (!sheet) return { success: true, list: [], totalRows: 0 };

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol === 0) return { success: true, list: [], totalRows: 0 };

    var headerMap = getHeaderMap_(sheet);
    var n = lastRow - 1;
    var delIdx = headerMap["is_deleted"];
    var rawAll = sheet.getRange(2, 1, n, lastCol).getValues();
    var ownerIdx = headerMap['created_by'];
    var deptIdx = headerMap['dept_code'];
    var statusIdx = headerMap['approval_status'];
    var active = [], i;
    if (delIdx === undefined) {
      for (i = 0; i < n; i++) {
        if (matrixRowAllowed_(auth.user, 'HRM_ATT', rawAll[i], { owner: ownerIdx, dept: deptIdx, status: statusIdx }, dataPermission)) active.push(i + 2);
      }
    } else {
      var flags = sheet.getRange(2, delIdx + 1, n, 1).getValues();
      for (i = 0; i < n; i++) {
        if (String(flags[i][0]).trim().toUpperCase() !== "Y" &&
            matrixRowAllowed_(auth.user, 'HRM_ATT', rawAll[i], { owner: ownerIdx, dept: deptIdx, status: statusIdx }, dataPermission)) active.push(i + 2);
      }
    }

    var totalRows = active.length;
    var pageRows = active.slice((page - 1) * pageSize, page * pageSize);
    if (pageRows.length === 0) return { success: true, list: [], totalRows: totalRows };

    var minR = pageRows[0], maxR = pageRows[pageRows.length - 1];
    // rawAll đã đọc toàn bộ sheet ở trên (index 0 = hàng row 2). Truy cập trực tiếp để tránh đọc trùng lại range con.
    var displayData = sheet.getRange(minR, 1, maxR - minR + 1, lastCol).getDisplayValues();
    var list = [];

    var getVal = function(row, colName) {
      var idx = headerMap[cleanCompareKey_(colName)];
      if (idx === undefined || row[idx] === undefined) return "";
      return row[idx];
    };
    var getDispVal = function(dispRow, colName) {
      var idx = headerMap[cleanCompareKey_(colName)];
      if (idx === undefined || dispRow[idx] === undefined) return "";
      return dispRow[idx];
    };

    pageRows.forEach(function(r) {
      var row = rawAll[r - 2];
      var dispRow = displayData[r - minR];

      var attId = getRawText_(getVal(row, "att_id"));
      var empId = getRawText_(getVal(row, "emp_id"));

      if (attId !== "" || empId !== "") {
        var dispCheckIn = getRawText_(getDispVal(dispRow, "check_in"));
        var dispCheckOut = getRawText_(getDispVal(dispRow, "check_out"));
        var dispWorkDate = getRawText_(getDispVal(dispRow, "work_date"));
        var dispWorkMonth = getRawText_(getDispVal(dispRow, "work_month"));

        list.push({
          att_id: attId,
          cccd_no: getRawText_(getVal(row, "cccd_no")),
          emp_id: empId,
          full_name: getRawText_(getVal(row, "full_name")),
          dept_code: getRawText_(getVal(row, "dept_code")),
          unit_code: getRawText_(getVal(row, "unit_code")),
          work_date: dispWorkDate || formatDateTimeValue_(getVal(row, "work_date"), "date"),
          work_month: dispWorkMonth || formatDateTimeValue_(getVal(row, "work_month"), "month"),
          work_shift: getRawText_(getVal(row, "work_shift")),
          att_type: getRawText_(getVal(row, "att_type")),
          work_day_value: getVal(row, "work_day_value") !== "" ? getVal(row, "work_day_value") : 1,
          check_in: dispCheckIn || formatDateTimeValue_(getVal(row, "check_in"), "time"),
          check_out: dispCheckOut || formatDateTimeValue_(getVal(row, "check_out"), "time"),
          meal_qty: getVal(row, "meal_qty") !== "" ? getVal(row, "meal_qty") : 0,
          location_latlong: getRawText_(getVal(row, "location_latlong")),
          has_selfie_img: !!getVal(row, "selfie_img"),
          leave_reason: getRawText_(getVal(row, "leave_reason")),
          reject_reason: getRawText_(getVal(row, "reject_reason")),
          approval_status: getRawText_(getVal(row, "approval_status")) || "Đã duyệt",
          created_by: getRawText_(getVal(row, "created_by")),
          created_at: formatDateTimeValue_(getVal(row, "created_at"), "full"),
          approval_chain_raw: getRawText_(getVal(row, "approval_chain")),
          approval_step_raw: getVal(row, "approval_step"),
          approval_due_text: formatDateTimeValue_(getVal(row, "approval_due_at"), "full"),
          is_auto_approved: getRawText_(getVal(row, "is_auto_approved")).toUpperCase() === "Y"
        });
      }
    });

    // ----- Thông tin duyệt theo cấp, tính riêng cho người đang xem -----
    var isSuperV = isSuperAdminRole_(auth.user.role);
    var attFeaturePermissions = {
      attendanceSelf: hasPermission_(auth.user, 'HRM', 'ATTENDANCE_SELF', 'HRM_ATT', undefined, dataPermission),
      attendanceProxy: hasPermission_(auth.user, 'HRM', 'ATTENDANCE_PROXY', 'HRM_ATT', undefined, dataPermission),
      leaveCreate: hasPermission_(auth.user, 'HRM', 'LEAVE_CREATE', 'HRM_ATT', undefined, dataPermission),
      import: hasPermission_(auth.user, 'HRM', 'import', 'HRM_ATT', undefined, dataPermission),
      approve: hasPermission_(auth.user, 'HRM', 'approve', 'HRM_ATT', undefined, dataPermission),
      reject: hasPermission_(auth.user, 'HRM', 'reject', 'HRM_ATT', undefined, dataPermission)
    };
    var myEmpV = auth.user.cccd ? empFindByCccd_(auth.user.cccd) : null;
    var myKeyV = myEmpV ? cleanCompareKey_(myEmpV.emp_id) : '';
    var namesV = null, hierV = {};
    list.forEach(function (it) {
      var pending = it.approval_status === "Chờ duyệt";
      var chain = parseChain_(it.approval_chain_raw);
      var step = parseInt(it.approval_step_raw, 10) || 1;
      if (pending && !chain.length) {                      // đơn cũ (trước khi có duyệt theo cấp)
        var hk = cleanCompareKey_(it.emp_id);
        if (hierV[hk] === undefined) hierV[hk] = getApprovalHierarchyForEmployee_(it.emp_id).hierarchy;
        chain = hierV[hk]; step = 1;
      }
      var cur = (pending && chain.length) ? chain[Math.min(step, chain.length) - 1] : '';
      if (cur && !namesV) namesV = empNameMap_();
      it.approval_total = chain.length;
      it.approval_step = pending ? step : 0;
      it.current_approver_id = cur;
      it.current_approver_name = (cur && namesV) ? (namesV[cleanCompareKey_(cur)] || '') : '';
      var isCurrentApprover = pending &&
        (isSuperV || (!!myKeyV && !!cur && cleanCompareKey_(cur) === myKeyV && cleanCompareKey_(it.emp_id) !== myKeyV));
      it.can_approve = isCurrentApprover && attFeaturePermissions.approve;
      it.can_reject = isCurrentApprover && attFeaturePermissions.reject;
      it.approval_due_at = it.approval_due_text;
      delete it.approval_chain_raw; delete it.approval_step_raw; delete it.approval_due_text;
    });

    return { success: true, list: list, totalRows: totalRows, permissions: attFeaturePermissions };

  } catch (err) {
    Logger.log('fetchHRMATTListPaginated: ' + err);
    return { success: false, message: "Lỗi tải dữ liệu chấm công", list: [], totalRows: 0 };
  }
}

function getHRMATTImage(token, attId) {
  var auth = requireAuth_(token, 'HRM', 'view', 'HRM_ATT', 'view-hrm-att');
  if (!auth.ok) return auth.res;
  var requestedId = cleanCompareKey_(attId);
  if (!requestedId) return { success: false, message: 'Mã chấm công không hợp lệ.' };

  try {
    var cfg = getHRMATTConfig_();
    var sheet = SpreadsheetApp.openById(cfg.DB_HRM_PART1_ID).getSheetByName(cfg.SHEET_HRM_ATT);
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Không tìm thấy bản ghi chấm công.' };

    var headerMap = getHeaderMap_(sheet);
    var idIdx = headerMap['att_id'];
    var imageIdx = headerMap['selfie_img'];
    if (idIdx === undefined || imageIdx === undefined) {
      return { success: false, message: 'Bản ghi không có ảnh chấm công.' };
    }

    var rowCount = sheet.getLastRow() - 1;
    var ids = sheet.getRange(2, idIdx + 1, rowCount, 1).getValues();
    var rowNumber = -1;
    for (var i = 0; i < ids.length; i++) {
      if (cleanCompareKey_(ids[i][0]) === requestedId) {
        rowNumber = i + 2;
        break;
      }
    }
    if (rowNumber < 0) return { success: false, message: 'Không tìm thấy bản ghi chấm công.' };

    var row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    var deletedIdx = headerMap['is_deleted'];
    if (deletedIdx !== undefined && String(row[deletedIdx]).trim().toUpperCase() === 'Y') {
      return { success: false, message: 'Không tìm thấy bản ghi chấm công.' };
    }
    var dataPermission = getEffectivePermission_(auth.user.username, auth.user.role, 'HRM_ATT');
    if (!matrixScopeAllowed_(auth.user, 'HRM_ATT', row, {
      owner: headerMap['created_by'], dept: headerMap['dept_code'], status: headerMap['approval_status']
    }, dataPermission, 'read')) {
      return { success: false, code: 'FORBIDDEN', message: 'Bản ghi nằm ngoài phạm vi dữ liệu được phân quyền.' };
    }

    return { success: true, image: getRawText_(row[imageIdx]) };
  } catch (err) {
    Logger.log('getHRMATTImage: ' + err);
    return { success: false, message: 'Không tải được ảnh chấm công.' };
  }
}

function saveHRMATTRecord(token, record) {
  var auth = requireSessionAuth_(token, 'HRM', 'view-hrm-att');
  if (!auth.ok) return auth.res;
  if (!record || typeof record !== 'object') return fail_('Dữ liệu chấm công không hợp lệ.');

  // Danh tính nhân viên lấy từ Hồ sơ Nhân viên trên server, KHÔNG tin dữ liệu client gửi lên
  var who = resolveEmpForWrite_(auth.user, record.emp_id);
  if (!who.ok) return who.res;
  record.emp_id = who.emp.emp_id;
  record.cccd_no = who.emp.cccd_no;
  record.full_name = who.emp.full_name;
  record.dept_code = who.emp.dept_code;
  record.unit_code = who.emp.unit_code;
  record.created_by = auth.user.username;
  record.is_proxy = !who.isSelf;
  delete record.unique_key;   // luôn do server tự tạo
  delete record.att_id;
  delete record.created_at;
  record.leave_reason = safeText_(record.leave_reason);
  record.reject_reason = safeText_(record.reject_reason);
  record.location_latlong = safeText_(record.location_latlong);

  // Mọi loại KHÁC "Chấm công bình thường" (phép năm, công tác, xin đi trễ...) đều phải duyệt theo từng cấp.
  // Không tin trạng thái duyệt do giao diện gửi lên.
  var attTypeIn = getRawText_(record.att_type) || "Chấm công bình thường";
  var featureActions = [];
  if (!who.isSelf) featureActions.push('ATTENDANCE_PROXY');
  if (attTypeIn === 'Chấm công bình thường') {
    if (who.isSelf) featureActions.push('ATTENDANCE_SELF');
  } else {
    featureActions.push('LEAVE_CREATE');
  }
  if (featureActions.some(function(action) {
    return !hasPermission_(auth.user, 'HRM', action, 'HRM_ATT');
  })) {
    return { success: false, code: 'FORBIDDEN', message: 'Bạn chưa được cấp quyền cho thao tác này.' };
  }
  var plan = (attTypeIn !== "Chấm công bình thường") ? buildApprovalPlan_(who.emp.emp_id, new Date()) : null;

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    var cfg = getHRMATTConfig_();
    var ss = SpreadsheetApp.openById(cfg.DB_HRM_PART1_ID);
    var sheet = ss.getSheetByName(cfg.SHEET_HRM_ATT);

    var requiredHeaders = [
      "att_id", "cccd_no", "emp_id", "full_name", "dept_code", "unit_code",
      "work_date", "work_month", "work_shift", "att_type", "work_day_value",
      "check_in", "check_out", "meal_qty", "location_latlong", "selfie_img",
      "leave_reason", "reject_reason", "approval_status", "created_by", "created_at", "unique_key"
    ];

    if (!sheet) {
      sheet = ss.insertSheet(cfg.SHEET_HRM_ATT);
      sheet.appendRow(requiredHeaders);
      SpreadsheetApp.flush();
    }

    var headerMap = getHeaderMap_(sheet);
    var lastCol = sheet.getLastColumn();

    if (headerMap["unique_key"] === undefined) {
      sheet.getRange(1, lastCol + 1).setValue("unique_key");
      SpreadsheetApp.flush();
      headerMap = getHeaderMap_(sheet);
      lastCol = sheet.getLastColumn();
    }
    if (plan) {
      headerMap = ensureATTApprovalCols_(sheet);
      lastCol = sheet.getLastColumn();
    }

    var now = new Date();
    var currentTimeStr = Utilities.formatDate(now, "GMT+7", "HH:mm:ss");

    var attTypeStr = getRawText_(record.att_type) || "Chấm công bình thường";
    var workShiftStr = getRawText_(record.work_shift) || "Hành chính";
    var targetDate = record.work_date || now;

    var currentUniqueKey = record.unique_key || buildUniqueKey_(record.emp_id, targetDate, workShiftStr);

    var targetDateObj = new Date();
    if (record.work_date && record.work_date.indexOf("-") !== -1) {
      var p = record.work_date.split("-");
      targetDateObj = new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
    } else if (record.work_date && record.work_date.indexOf("/") !== -1) {
      var p = record.work_date.split("/");
      targetDateObj = new Date(parseInt(p[2], 10), parseInt(p[1], 10) - 1, parseInt(p[0], 10));
    }

    var workDateFormatted = Utilities.formatDate(targetDateObj, "GMT+7", "yyyy-MM-dd");
    var workMonthFormatted = record.work_month || Utilities.formatDate(targetDateObj, "GMT+7", "MM/yyyy");

    var workDayVal = 1;
    if (record.work_day_value !== undefined && record.work_day_value !== "") {
      workDayVal = parseFloat(record.work_day_value);
    } else {
      if (workShiftStr === "Sáng" || workShiftStr === "Chiều" || attTypeStr === "Xin đi trễ" || attTypeStr === "Xin về sớm") {
        workDayVal = 0.5;
      } else if (attTypeStr === "Nghỉ không lương") {
        workDayVal = 0;
      }
    }

    var toSheetCodeText = function(val) {
      var str = getRawText_(val);
      return str ? "'" + str : "";
    };

    var existingRowIndex = -1;
    var uniqueKeyColIdx = headerMap["unique_key"] + 1;

    if (sheet.getLastRow() > 1) {
      var finder = sheet.getRange(2, uniqueKeyColIdx, sheet.getLastRow() - 1, 1)
                        .createTextFinder(currentUniqueKey)
                        .matchCase(false)
                        .matchEntireCell(true);
      var foundCell = finder.findNext();
      if (foundCell) {
        existingRowIndex = foundCell.getRow();
      }
    }

    var creatorName = getRawText_(record.created_by) || "system";
    var logPrefix = record.is_proxy ? ("Chấm công thay bởi user [" + creatorName + "]") : attTypeStr;
    var newLogEntry = "[" + Utilities.formatDate(now, "GMT+7", "dd/MM/yy HH:mm:ss") + "] " + logPrefix + " (" + currentTimeStr + ")";
    if (record.location_latlong) newLogEntry += " - GPS: " + record.location_latlong;

    if (existingRowIndex > 1) {
      var existingRowData = sheet.getRange(existingRowIndex, 1, 1, lastCol).getValues()[0];
      var existingDispData = sheet.getRange(existingRowIndex, 1, 1, lastCol).getDisplayValues()[0];
      if (!hasPermission_(auth.user, 'HRM', 'update', 'HRM_ATT') ||
          !matrixScopeAllowed_(auth.user, 'HRM_ATT', existingRowData, {
        owner: headerMap['created_by'], dept: headerMap['dept_code'], status: headerMap['approval_status']
      }, null, 'update')) {
        return fail_('Bản ghi nằm ngoài phạm vi dữ liệu được phân quyền.', { code: 'FORBIDDEN' });
      }

      var checkInIdx = headerMap["check_in"];
      var checkOutIdx = headerMap["check_out"];
      var logIdx = headerMap["reject_reason"];
      var gpsIdx = headerMap["location_latlong"];
      var selfieIdx = headerMap["selfie_img"];
      var attIdIdx = headerMap["att_id"];
      var mealIdx = headerMap["meal_qty"];
      var workValIdx = headerMap["work_day_value"];

      var existingAttId = getRawText_(existingRowData[attIdIdx]);
      var existingCheckIn = getRawText_(existingDispData[checkInIdx]) || formatDateTimeValue_(existingRowData[checkInIdx], "time");
      var existingLog = getRawText_(existingRowData[logIdx]);
      var existingGps = getRawText_(existingRowData[gpsIdx]);

      var updatedCheckIn = record.check_in || existingCheckIn || currentTimeStr;
      var updatedCheckOut = record.check_out || currentTimeStr;

      var updatedLog = record.reject_reason ? record.reject_reason : (existingLog ? (existingLog + " | " + newLogEntry) : newLogEntry);
      updatedLog = updatedLog.replace(/(\r\n|\n|\r)/gm, " | ");

      var updatedGps = record.location_latlong ? (existingGps ? (existingGps + " | " + record.location_latlong) : record.location_latlong) : existingGps;
      var updatedSelfie = record.selfie_base64 || record.selfie_img || getRawText_(existingRowData[selfieIdx]);

      var updatedRowData = existingRowData.slice();
      if (checkInIdx !== undefined) updatedRowData[checkInIdx] = getRawText_(updatedCheckIn);
      if (checkOutIdx !== undefined) updatedRowData[checkOutIdx] = getRawText_(updatedCheckOut);
      if (mealIdx !== undefined && record.meal_qty !== undefined) updatedRowData[mealIdx] = parseInt(record.meal_qty, 10) || 0;
      if (workValIdx !== undefined) updatedRowData[workValIdx] = workDayVal;
      if (logIdx !== undefined) updatedRowData[logIdx] = updatedLog;
      if (gpsIdx !== undefined && record.location_latlong) updatedRowData[gpsIdx] = updatedGps;
      if (selfieIdx !== undefined && updatedSelfie) updatedRowData[selfieIdx] = updatedSelfie;
      if (!matrixScopeAllowed_(auth.user, 'HRM_ATT', updatedRowData, {
        owner: headerMap['created_by'], dept: headerMap['dept_code'], status: headerMap['approval_status']
      }, null, 'update')) {
        return fail_('Thông tin cập nhật nằm ngoài phạm vi được phân quyền.', { code: 'FORBIDDEN' });
      }

      if (checkInIdx !== undefined) sheet.getRange(existingRowIndex, checkInIdx + 1).setValue(getRawText_(updatedCheckIn));
      if (checkOutIdx !== undefined) sheet.getRange(existingRowIndex, checkOutIdx + 1).setValue(getRawText_(updatedCheckOut));
      if (mealIdx !== undefined && record.meal_qty !== undefined) sheet.getRange(existingRowIndex, mealIdx + 1).setValue(parseInt(record.meal_qty, 10) || 0);
      if (workValIdx !== undefined) sheet.getRange(existingRowIndex, workValIdx + 1).setValue(workDayVal);
      if (logIdx !== undefined) sheet.getRange(existingRowIndex, logIdx + 1).setValue(updatedLog);
      if (gpsIdx !== undefined && record.location_latlong) sheet.getRange(existingRowIndex, gpsIdx + 1).setValue(updatedGps);
      if (selfieIdx !== undefined && updatedSelfie) sheet.getRange(existingRowIndex, selfieIdx + 1).setValue(updatedSelfie);

      SpreadsheetApp.flush();

      if (record.is_proxy) auditLog_(auth.user, 'HRM_ATT', 'PROXY_UPDATE', record.emp_id, 'Chấm công thay cho ' + record.emp_id + ' ngày ' + workDateFormatted);
      return {
        success: true,
        message: "Đã cập nhật dữ liệu chấm công vào bản ghi hiện có!",
        att_id: existingAttId,
        work_day_value: workDayVal,
        created_at: Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy HH:mm:ss")
      };

    } else {
      var autoAttId = record.att_id || ("ATT" + Utilities.formatDate(now, "GMT+7", "yyyyMMddHHmmss") + "_" + Utilities.getUuid().replace(/-/g, "").substring(0, 5));

      var finalCheckIn = "";
      var finalCheckOut = "";

      if (attTypeStr === "Chấm công bình thường" || record.is_proxy) {
        finalCheckIn = record.check_in || currentTimeStr;
        finalCheckOut = record.check_out || "";
      } else {
        finalCheckIn = record.check_in || "";
        finalCheckOut = record.check_out || "";
      }

      var rowData = new Array(lastCol);
      for (var k = 0; k < lastCol; k++) rowData[k] = "";

      var setValueByHeader = function(hName, val) {
        var idx = headerMap[cleanCompareKey_(hName)];
        if (idx !== undefined) rowData[idx] = val;
      };

      var cleanNewLogEntry = (record.reject_reason || newLogEntry).replace(/(\r\n|\n|\r)/gm, " | ");

      setValueByHeader("att_id", toSheetCodeText(autoAttId));
      setValueByHeader("cccd_no", toSheetCodeText(record.cccd_no));
      setValueByHeader("emp_id", toSheetCodeText(record.emp_id));
      setValueByHeader("full_name", getRawText_(record.full_name));
      setValueByHeader("dept_code", toSheetCodeText(record.dept_code));
      setValueByHeader("unit_code", toSheetCodeText(record.unit_code));
      setValueByHeader("work_date", workDateFormatted);
      setValueByHeader("work_month", workMonthFormatted);
      setValueByHeader("work_shift", workShiftStr);
      setValueByHeader("att_type", attTypeStr);
      setValueByHeader("work_day_value", workDayVal);
      setValueByHeader("check_in", getRawText_(finalCheckIn));
      setValueByHeader("check_out", getRawText_(finalCheckOut));
      setValueByHeader("meal_qty", parseInt(record.meal_qty, 10) || 0);
      setValueByHeader("location_latlong", record.location_latlong || "");
      setValueByHeader("selfie_img", record.selfie_base64 || record.selfie_img || "");
      setValueByHeader("leave_reason", record.leave_reason || "");
      setValueByHeader("reject_reason", cleanNewLogEntry);
      setValueByHeader("approval_status", plan ? plan.status : "Đã duyệt");
        if (plan) {
          setValueByHeader("approval_chain", plan.chainText);
          setValueByHeader("approval_step", plan.step);
          setValueByHeader("approval_due_at", plan.dueAt);
        }
      setValueByHeader("created_by", toSheetCodeText(creatorName));
      setValueByHeader("created_at", record.created_at || Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy HH:mm:ss"));
      setValueByHeader("unique_key", toSheetCodeText(currentUniqueKey));

      if (!hasPermission_(auth.user, 'HRM', 'create', 'HRM_ATT') ||
          !matrixScopeAllowed_(auth.user, 'HRM_ATT', rowData, {
        owner: headerMap['created_by'], dept: headerMap['dept_code'], status: headerMap['approval_status']
      }, null, 'create')) {
        return fail_('Thông tin mới nằm ngoài phạm vi được phân quyền.', { code: 'FORBIDDEN' });
      }
      sheet.appendRow(rowData);
      SpreadsheetApp.flush();

      if (record.is_proxy) auditLog_(auth.user, 'HRM_ATT', 'PROXY_CREATE', record.emp_id, 'Chấm công thay cho ' + record.emp_id + ' ngày ' + workDateFormatted);
      if (plan && !record.is_proxy) auditLog_(auth.user, 'HRM_ATT', 'LEAVE_SUBMIT', record.emp_id, attTypeIn + ' ngày ' + workDateFormatted + ' - ' + plan.status);
      return { 
        success: true, 
        message: "Tạo bản ghi chấm công mới thành công!", 
        att_id: autoAttId,
        work_day_value: workDayVal,
        created_at: Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy HH:mm:ss")
      };
    }

  } catch (err) {
    return { success: false, message: "Lỗi lưu chấm công: " + err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * IMPORT BATCH TỐC ĐỘ CAO (Xử lý hàng loạt trong bộ nhớ RAM và ghi 1 lần duy nhất)
 */
function importHRMATTExcelPaste(token, rawTextData) {
  var auth = requireAuth_(token, 'HRM', 'view', 'HRM_ATT', 'view-hrm-att');
  if (!auth.ok) return auth.res;
  if (!hasPermission_(auth.user, 'HRM', 'import', 'HRM_ATT')) {
    return { success: false, code: 'FORBIDDEN', message: 'Bạn chưa được cấp quyền nhập hàng loạt.' };
  }
  if (typeof rawTextData !== 'string' || rawTextData.length > 3000000) {
    return fail_('Dữ liệu dán quá lớn hoặc không hợp lệ!');
  }
  var activeUsername = auth.user.username;
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    if (!rawTextData || !rawTextData.trim()) {
      return { success: false, message: "Dữ liệu dán trống!" };
    }

    var cfg = getHRMATTConfig_();
    var ss = SpreadsheetApp.openById(cfg.DB_HRM_PART1_ID);
    var sheet = ss.getSheetByName(cfg.SHEET_HRM_ATT);

    var requiredHeaders = [
      "att_id", "cccd_no", "emp_id", "full_name", "dept_code", "unit_code",
      "work_date", "work_month", "work_shift", "att_type", "work_day_value",
      "check_in", "check_out", "meal_qty", "location_latlong", "selfie_img",
      "leave_reason", "reject_reason", "approval_status", "created_by", "created_at", "unique_key"
    ];

    if (!sheet) {
      sheet = ss.insertSheet(cfg.SHEET_HRM_ATT);
      sheet.appendRow(requiredHeaders);
      SpreadsheetApp.flush();
    }

    var headerMap = getHeaderMap_(sheet);
    var lastCol = sheet.getLastColumn();

    if (headerMap["unique_key"] === undefined) {
      sheet.getRange(1, lastCol + 1).setValue("unique_key");
      SpreadsheetApp.flush();
      headerMap = getHeaderMap_(sheet);
      lastCol = sheet.getLastColumn();
    }

    // 1. ĐỌC TOÀN BỘ DỮ LIỆU BẢNG HIỆN CÓ VỀ RAM (Chỉ tốn 1 lần I/O)
    var lastRow = sheet.getLastRow();
    var existingRows = [];
    if (lastRow > 1) {
      existingRows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    }

    // Map lưu vị trí dòng dựa vào unique_key
    var uniqueKeyMap = {};
    var uniqueKeyColIdx = headerMap["unique_key"];

    for (var r = 0; r < existingRows.length; r++) {
      var key = cleanCompareKey_(existingRows[r][uniqueKeyColIdx]);
      if (key) {
        uniqueKeyMap[key] = r;
      }
    }

    // Parse dữ liệu từ text copy/paste
    // Không trim() đầu văn bản/đầu dòng: ô đầu tiên (att_id) có thể để trống và dấu Tab đầu dòng là một cột
    var lines = rawTextData.replace(/^[\r\n]+/, "").replace(/\s+$/, "").split(/\r?\n/);
    var successCount = 0;
    var errorCount = 0;
    var currentUser = activeUsername || "admin";
    var attSeq = 0;
    var now = new Date();
    var currentTimeStr = Utilities.formatDate(now, "GMT+7", "HH:mm:ss");

    var toSheetCodeText = function(val) {
      var str = getRawText_(val);
      return str ? "'" + str : "";
    };

    // 2. XỬ LÝ VÀ ĐỐI CHIẾU TRÊN BỘ NHỚ RAM
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].replace(/\s+$/, "");
      if (!line.trim()) continue;

      var cols = line.split("\t");

      // Bỏ qua dòng Header nếu dán kèm
      if (i === 0 && (cols[0].toLowerCase().indexOf("att_id") !== -1 || cols[0].toLowerCase().indexOf("mã đơn") !== -1 || cols[2].toLowerCase().indexOf("emp_id") !== -1 || cols[2].toLowerCase().indexOf("mã nv") !== -1)) {
        continue;
      }

      while (cols.length < 22) cols.push("");

      var att_id           = getRawText_(cols[0]);
      var cccd_no          = getRawText_(cols[1]);
      var emp_id           = getRawText_(cols[2]);
      var full_name        = safeText_(getRawText_(cols[3]));
      var dept_code        = getRawText_(cols[4]);
      var unit_code        = getRawText_(cols[5]);
      var work_date        = getRawText_(cols[6]);
      var work_month       = getRawText_(cols[7]);
      var work_shift       = getRawText_(cols[8]) || "Hành chính";
      var att_type         = getRawText_(cols[9]) || "Chấm công bình thường";
      var work_day_value   = getRawText_(cols[10]);
      var check_in         = getRawText_(cols[11]);
      var check_out        = getRawText_(cols[12]);
      var meal_qty         = getRawText_(cols[13]);
      var location_latlong = getRawText_(cols[14]);
      var selfie_img       = getRawText_(cols[15]);
      var leave_reason     = safeText_(getRawText_(cols[16]));
      var reject_reason    = safeText_(getRawText_(cols[17]));
      var approval_status  = getRawText_(cols[18]) || "Đã duyệt";
      var created_at       = getRawText_(cols[20]);
      var unique_key       = getRawText_(cols[21]);

      if (!emp_id || !work_date) {
        errorCount++;
        continue;
      }

      var currentUniqueKey = unique_key || buildUniqueKey_(emp_id, work_date, work_shift);
      var cleanKey = cleanCompareKey_(currentUniqueKey);

      var targetDateObj = new Date();
      if (work_date.indexOf("-") !== -1) {
        var p = work_date.split("-");
        targetDateObj = new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
      } else if (work_date.indexOf("/") !== -1) {
        var p = work_date.split("/");
        targetDateObj = new Date(parseInt(p[2], 10), parseInt(p[1], 10) - 1, parseInt(p[0], 10));
      }

      var workDateFormatted = Utilities.formatDate(targetDateObj, "GMT+7", "yyyy-MM-dd");
      var workMonthFormatted = work_month || Utilities.formatDate(targetDateObj, "GMT+7", "MM/yyyy");

      var workDayVal = 1;
      if (work_day_value !== undefined && work_day_value !== "") {
        workDayVal = parseFloat(work_day_value);
      } else {
        if (work_shift === "Sáng" || work_shift === "Chiều" || att_type === "Xin đi trễ" || att_type === "Xin về sớm") {
          workDayVal = 0.5;
        } else if (att_type === "Nghỉ không lương") {
          workDayVal = 0;
        }
      }

      var newLogEntry = "[" + Utilities.formatDate(now, "GMT+7", "dd/MM/yy HH:mm:ss") + "] Import Excel bởi user [" + currentUser + "] (" + currentTimeStr + ")";
      if (location_latlong) newLogEntry += " - GPS: " + location_latlong;

      var cleanNewLog = (reject_reason || newLogEntry).replace(/(\r\n|\n|\r)/gm, " | ");

      // KIỂM TRA BẢN GHI ĐÃ TỒN TẠI TRONG RAM CHƯA
      if (uniqueKeyMap[cleanKey] !== undefined) {
        // --- CẬP NHẬT TRỰC TIẾP TRÊN MẢNG RAM ---
        var rIdx = uniqueKeyMap[cleanKey];
        var rowData = existingRows[rIdx].slice();
        if (!matrixScopeAllowed_(auth.user, 'HRM_ATT', existingRows[rIdx], {
          owner: headerMap['created_by'], dept: headerMap['dept_code'], status: headerMap['approval_status']
        }, null, 'update')) {
          errorCount++;
          continue;
        }

        var checkInIdx = headerMap["check_in"];
        var checkOutIdx = headerMap["check_out"];
        var logIdx = headerMap["reject_reason"];
        var gpsIdx = headerMap["location_latlong"];
        var mealIdx = headerMap["meal_qty"];
        var workValIdx = headerMap["work_day_value"];

        var existingLog = getRawText_(rowData[logIdx]);
        var existingGps = getRawText_(rowData[gpsIdx]);

        if (checkInIdx !== undefined) rowData[checkInIdx] = getRawText_(check_in || rowData[checkInIdx] || currentTimeStr);
        if (checkOutIdx !== undefined) rowData[checkOutIdx] = getRawText_(check_out || currentTimeStr);
        if (mealIdx !== undefined) rowData[mealIdx] = meal_qty !== "" ? parseInt(meal_qty, 10) : (rowData[mealIdx] || 0);
        if (workValIdx !== undefined) rowData[workValIdx] = workDayVal;
        if (logIdx !== undefined) rowData[logIdx] = reject_reason ? reject_reason : (existingLog ? (existingLog + " | " + newLogEntry) : newLogEntry);
        if (gpsIdx !== undefined && location_latlong) rowData[gpsIdx] = existingGps ? (existingGps + " | " + location_latlong) : location_latlong;

        if (!matrixScopeAllowed_(auth.user, 'HRM_ATT', rowData, {
          owner: headerMap['created_by'], dept: headerMap['dept_code'], status: headerMap['approval_status']
        }, null, 'update')) {
          errorCount++;
          continue;
        }
        existingRows[rIdx] = rowData;
        successCount++;

      } else {
        // --- TẠO MỚI DÒNG VÀ PUSH VÀO MẢNG RAM ---
        var autoAttId = att_id || ("ATT" + Utilities.formatDate(now, "GMT+7", "yyyyMMddHHmmss") + "_" + (++attSeq) + "_" + Utilities.getUuid().replace(/-/g, "").substring(0, 4));

        var newRow = new Array(lastCol);
        for (var k = 0; k < lastCol; k++) newRow[k] = "";

        var setValueByHeader = function(hName, val) {
          var idx = headerMap[cleanCompareKey_(hName)];
          if (idx !== undefined) newRow[idx] = val;
        };

        setValueByHeader("att_id", toSheetCodeText(autoAttId));
        setValueByHeader("cccd_no", toSheetCodeText(cccd_no));
        setValueByHeader("emp_id", toSheetCodeText(emp_id));
        setValueByHeader("full_name", full_name);
        setValueByHeader("dept_code", toSheetCodeText(dept_code));
        setValueByHeader("unit_code", toSheetCodeText(unit_code));
        setValueByHeader("work_date", workDateFormatted);
        setValueByHeader("work_month", workMonthFormatted);
        setValueByHeader("work_shift", work_shift);
        setValueByHeader("att_type", att_type);
        setValueByHeader("work_day_value", workDayVal);
        setValueByHeader("check_in", getRawText_(check_in));
        setValueByHeader("check_out", getRawText_(check_out));
        setValueByHeader("meal_qty", meal_qty !== "" ? parseInt(meal_qty, 10) : 0);
        setValueByHeader("location_latlong", location_latlong);
        setValueByHeader("selfie_img", selfie_img);
        setValueByHeader("leave_reason", leave_reason);
        setValueByHeader("reject_reason", cleanNewLog);
        setValueByHeader("approval_status", approval_status);
        setValueByHeader("created_by", toSheetCodeText(currentUser));
        setValueByHeader("created_at", created_at || Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy HH:mm:ss"));
        setValueByHeader("unique_key", toSheetCodeText(currentUniqueKey));

        if (!matrixScopeAllowed_(auth.user, 'HRM_ATT', newRow, {
          owner: headerMap['created_by'], dept: headerMap['dept_code'], status: headerMap['approval_status']
        }, null, 'create')) {
          errorCount++;
          continue;
        }
        existingRows.push(newRow);
        uniqueKeyMap[cleanKey] = existingRows.length - 1;
        successCount++;
      }
    }

    // 3. GHI ĐÈ LẠI TOÀN BỘ DỮ LIỆU ĐÃ TỐI ƯU VÀO SHEET (CHỈ 1 LẦN GHI DUY NHẤT)
    if (existingRows.length > 0) {
      sheet.getRange(2, 1, existingRows.length, lastCol).setValues(existingRows);
      SpreadsheetApp.flush();
    }

    auditLog_(auth.user, 'HRM_ATT', 'IMPORT', '', 'Import ' + successCount + ' dòng chấm công, lỗi ' + errorCount);
    return {
      success: true,
      message: "Import thành công " + successCount + " dòng bởi user [" + currentUser + "]!" + (errorCount > 0 ? " (Lỗi/Thất bại: " + errorCount + " dòng)" : "")
    };

  } catch (err) {
    return { success: false, message: "Lỗi xử lý import Excel Batch: " + err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/** Thêm 3 cột xóa mềm vào cuối bảng chấm công nếu chưa có. Trả về headerMap mới. */
function ensureATTSoftCols_(sheet) {
  var hm = getHeaderMap_(sheet);
  ['is_deleted', 'deleted_by', 'deleted_at'].forEach(function (h) {
    if (hm[h] === undefined) {
      var nextCol = sheet.getLastColumn() + 1;
      if (sheet.getMaxColumns() < nextCol) sheet.insertColumnsAfter(sheet.getMaxColumns(), nextCol - sheet.getMaxColumns());
      sheet.getRange(1, nextCol).setValue(h);
      hm = getHeaderMap_(sheet);
    }
  });
  return hm;
}

/** Ghi các giá trị (theo từng dòng) vào 1 cột, gom theo đoạn liên tiếp */
function writeColByRows_(sheet, col1, rowToValue) {
  var rows = Object.keys(rowToValue).map(Number).sort(function (a, b) { return a - b; });
  groupRuns_(rows).forEach(function (run) {
    var vals = [];
    for (var i = 0; i < run[1]; i++) vals.push([rowToValue[run[0] + i]]);
    sheet.getRange(run[0], col1, run[1], 1).setValues(vals);
  });
}

function deleteBulkHRMATTRecords(token, attIdList) {
  var auth = requireAuth_(token, 'HRM', 'delete', 'HRM_ATT', 'view-hrm-att');
  if (!auth.ok) return auth.res;

  try {
    if (!Array.isArray(attIdList) || attIdList.length === 0) return fail_('Danh sách mã cần xóa trống.');

    return withLock_(function () {
      var cfg = getHRMATTConfig_();
      var ss = SpreadsheetApp.openById(cfg.DB_HRM_PART1_ID);
      var sheet = ss.getSheetByName(cfg.SHEET_HRM_ATT);
      if (!sheet || sheet.getLastRow() <= 1) return fail_('Bảng dữ liệu trống hoặc không tồn tại.');

      var hm = ensureATTSoftCols_(sheet);
      var attIdIdx = hm['att_id'], uIdx = hm['unique_key'];
      if (attIdIdx === undefined) return fail_('Cấu trúc cột dữ liệu không hợp lệ!');

      var targets = {};
      attIdList.forEach(function (id) { var k = cleanCompareKey_(id); if (k) targets[k] = true; });

      var n = sheet.getLastRow() - 1;
      var ids = sheet.getRange(2, attIdIdx + 1, n, 1).getValues();
      var flags = sheet.getRange(2, hm['is_deleted'] + 1, n, 1).getValues();
      var ukeys = uIdx === undefined ? null : sheet.getRange(2, uIdx + 1, n, 1).getValues();
      var ownerIdx = hm['created_by'], deptIdx = hm['dept_code'], statusIdx = hm['approval_status'];
      var owners = ownerIdx === undefined ? null : sheet.getRange(2, ownerIdx + 1, n, 1).getValues();
      var depts = deptIdx === undefined ? null : sheet.getRange(2, deptIdx + 1, n, 1).getValues();
      var statuses = statusIdx === undefined ? null : sheet.getRange(2, statusIdx + 1, n, 1).getValues();
      var dataPermission = getEffectivePermission_(auth.user.username, auth.user.role, 'HRM_ATT');

      var now = new Date(), stamp = now.getTime();
      var rows = [], keys = [], mapFlag = {}, mapBy = {}, mapAt = {}, mapKey = {};
      for (var i = 0; i < n; i++) {
        if (String(flags[i][0]).trim().toUpperCase() === 'Y') continue;
        if (!targets[cleanCompareKey_(ids[i][0])]) continue;
        var scopeRow = [owners ? owners[i][0] : '', depts ? depts[i][0] : '', statuses ? statuses[i][0] : ''];
        if (!matrixScopeAllowed_(auth.user, 'HRM_ATT', scopeRow, { owner: 0, dept: 1, status: 2 }, dataPermission, 'delete')) continue;
        var r = i + 2;
        rows.push(r); keys.push(getRawText_(ids[i][0]));
        mapFlag[r] = 'Y'; mapBy[r] = auth.user.username; mapAt[r] = now;
        if (ukeys && getRawText_(ukeys[i][0])) mapKey[r] = getRawText_(ukeys[i][0]) + '#DEL' + stamp; // giải phóng khóa để chấm công lại được
      }
      if (rows.length === 0) return fail_('Không tìm thấy bản ghi phù hợp để xóa.');

      writeColByRows_(sheet, hm['is_deleted'] + 1, mapFlag);
      writeColByRows_(sheet, hm['deleted_by'] + 1, mapBy);
      writeColByRows_(sheet, hm['deleted_at'] + 1, mapAt);
      if (ukeys && Object.keys(mapKey).length) {
        writeColByRows_(sheet, uIdx + 1, mapKey);
      }

      auditLog_(auth.user, 'HRM_ATT', 'DELETE', sampleList_(keys, 20), 'Chuyển ' + rows.length + ' bản ghi vào Thùng rác');
      return { success: true, message: 'Đã chuyển ' + rows.length + ' bản ghi vào Thùng rác (SuperAdmin có thể khôi phục).' };
    });
  } catch (err) {
    Logger.log('deleteBulkHRMATTRecords: ' + err);
    return fail_('Không thể xóa dữ liệu.');
  }
}

/**
 * ============================================================================
 * NÂNG CẤP BỔ SUNG: MỞ RỘNG TÍNH NĂNG NGHỈ PHÉP BATCH (MÃ ĐƠN GỐC & DUYỆT 1 LẦN)
 * ============================================================================
 */



/**
 * Hàm lưu đơn nghỉ phép BATCH cho nhiều ngày cùng Mã Đơn Gốc (leave_batch_id)
 */
function submitHRMLeaveBatchRequest(token, payload) {
  var auth = requireSessionAuth_(token, 'HRM', 'view-hrm-att');
  if (!auth.ok) return auth.res;
  if (!payload || !payload.emp_id || !Array.isArray(payload.items) || payload.items.length === 0) {
    return fail_('Dữ liệu đơn xin nghỉ phép không hợp lệ!');
  }
  if (payload.items.length > 100) return fail_('Mỗi đơn chỉ được tối đa 100 ngày.');

  var who = resolveEmpForWrite_(auth.user, payload.emp_id);
  if (!who.ok) return who.res;
  var leaveFeatureActions = ['LEAVE_CREATE'];
  if (!who.isSelf) leaveFeatureActions.push('ATTENDANCE_PROXY');
  if (leaveFeatureActions.some(function(action) {
    return !hasPermission_(auth.user, 'HRM', action, 'HRM_ATT');
  }) || !hasPermission_(auth.user, 'HRM', 'create', 'HRM_ATT')) {
    return { success: false, code: 'FORBIDDEN', message: 'Bạn chưa được cấp quyền tạo đơn phép trong phạm vi này.' };
  }
  payload.emp_id = who.emp.emp_id;
  payload.cccd_no = who.emp.cccd_no;
  payload.full_name = who.emp.full_name;
  payload.dept_code = who.emp.dept_code;
  payload.unit_code = who.emp.unit_code;
  payload.created_by = auth.user.username;
  payload.leave_reason = safeText_(payload.leave_reason);
  payload.items.forEach(function (it) { if (it) it.leave_reason = safeText_(it.leave_reason); });

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    if (!payload || !payload.emp_id || !payload.items || !Array.isArray(payload.items) || payload.items.length === 0) {
      return { success: false, message: "Dữ liệu đơn xin nghỉ phép không hợp lệ!" };
    }

    var cfg = getHRMATTConfig_();
    var ss = SpreadsheetApp.openById(cfg.DB_HRM_PART1_ID);
    var sheet = ss.getSheetByName(cfg.SHEET_HRM_ATT);

    var requiredHeaders = [
      "att_id", "cccd_no", "emp_id", "full_name", "dept_code", "unit_code",
      "work_date", "work_month", "work_shift", "att_type", "work_day_value",
      "check_in", "check_out", "meal_qty", "location_latlong", "selfie_img",
      "leave_reason", "reject_reason", "approval_status", "created_by", "created_at", "unique_key"
    ];

    if (!sheet) {
      sheet = ss.insertSheet(cfg.SHEET_HRM_ATT);
      sheet.appendRow(requiredHeaders);
      SpreadsheetApp.flush();
    }

    var headerMap = getHeaderMap_(sheet);
    var lastCol = sheet.getLastColumn();

    if (headerMap["unique_key"] === undefined) {
      sheet.getRange(1, lastCol + 1).setValue("unique_key");
      SpreadsheetApp.flush();
      headerMap = getHeaderMap_(sheet);
      lastCol = sheet.getLastColumn();
    }

    var now = new Date();
    var batchTimestamp = Utilities.formatDate(now, "GMT+7", "yyyyMMddHHmmss");
    var batchId = "LEAVE_" + batchTimestamp + "_" + Math.floor(Math.random() * 8999 + 1000);
    
    // Lập chuỗi duyệt theo sơ đồ tổ chức (manager_id) tại thời điểm nộp đơn
    var plan = buildApprovalPlan_(payload.emp_id, now);
    var initialStatus = plan.status;
    headerMap = ensureATTApprovalCols_(sheet);
    lastCol = sheet.getLastColumn();

    var toSheetCodeText = function(val) {
      var str = getRawText_(val);
      return str ? "'" + str : "";
    };

    var creatorName = getRawText_(payload.created_by) || "system";
    var logMsg = "[" + Utilities.formatDate(now, "GMT+7", "dd/MM/yy HH:mm:ss") + "] Gửi đơn phép nhóm [" + batchId + "] bởi user [" + creatorName + "]";
    if (plan.chain.length) logMsg += " - Chờ duyệt cấp 1/" + plan.chain.length + " bởi [" + plan.chain[0] + "]";

    var items = payload.items;
    var successCount = 0;

    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var workDateStr = item.work_date;
      var workShiftStr = item.work_shift || "Cả ngày";
      var attTypeStr = item.att_type || "Phép năm";
      
      var targetDateObj = new Date();
      if (workDateStr && workDateStr.indexOf("-") !== -1) {
        var p = workDateStr.split("-");
        targetDateObj = new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
      } else if (workDateStr && workDateStr.indexOf("/") !== -1) {
        var p = workDateStr.split("/");
        targetDateObj = new Date(parseInt(p[2], 10), parseInt(p[1], 10) - 1, parseInt(p[0], 10));
      }

      var workDateFormatted = Utilities.formatDate(targetDateObj, "GMT+7", "yyyy-MM-dd");
      var workMonthFormatted = Utilities.formatDate(targetDateObj, "GMT+7", "MM/yyyy");

      var workDayVal = 1;
      if (item.work_day_value !== undefined && item.work_day_value !== "") {
        workDayVal = parseFloat(item.work_day_value);
      } else {
        if (workShiftStr === "Sáng" || workShiftStr === "Chiều" || attTypeStr === "Xin đi trễ" || attTypeStr === "Xin về sớm") {
          workDayVal = 0.5;
        } else if (attTypeStr === "Nghỉ không lương") {
          workDayVal = 0;
        }
      }

      var currentUniqueKey = buildUniqueKey_(payload.emp_id, workDateFormatted, workShiftStr);
      var subAttId = batchId + "_" + (i + 1);

      var rowData = new Array(lastCol);
      for (var k = 0; k < lastCol; k++) rowData[k] = "";

      var setValueByHeader = function(hName, val) {
        var idx = headerMap[cleanCompareKey_(hName)];
        if (idx !== undefined) rowData[idx] = val;
      };

      setValueByHeader("att_id", toSheetCodeText(subAttId));
      setValueByHeader("cccd_no", toSheetCodeText(payload.cccd_no));
      setValueByHeader("emp_id", toSheetCodeText(payload.emp_id));
      setValueByHeader("full_name", getRawText_(payload.full_name));
      setValueByHeader("dept_code", toSheetCodeText(payload.dept_code));
      setValueByHeader("unit_code", toSheetCodeText(payload.unit_code));
      setValueByHeader("work_date", workDateFormatted);
      setValueByHeader("work_month", workMonthFormatted);
      setValueByHeader("work_shift", workShiftStr);
      setValueByHeader("att_type", attTypeStr);
      setValueByHeader("work_day_value", workDayVal);
      setValueByHeader("check_in", getRawText_(item.check_in || ""));
      setValueByHeader("check_out", getRawText_(item.check_out || ""));
      setValueByHeader("meal_qty", 0);
      setValueByHeader("location_latlong", "");
      setValueByHeader("selfie_img", "");
      setValueByHeader("leave_reason", item.leave_reason || payload.leave_reason || "");
      setValueByHeader("reject_reason", logMsg);
      setValueByHeader("approval_status", initialStatus);
      setValueByHeader("approval_chain", plan.chainText);
      setValueByHeader("approval_step", plan.step);
      setValueByHeader("approval_due_at", plan.dueAt);
      setValueByHeader("created_by", toSheetCodeText(creatorName));
      setValueByHeader("created_at", Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy HH:mm:ss"));
      setValueByHeader("unique_key", toSheetCodeText(currentUniqueKey));

      if (!matrixScopeAllowed_(auth.user, 'HRM_ATT', rowData, {
        owner: headerMap['created_by'], dept: headerMap['dept_code'], status: headerMap['approval_status']
      }, null, 'create')) {
        return fail_('Thông tin đơn nằm ngoài phạm vi được phân quyền.', { code: 'FORBIDDEN' });
      }
      sheet.appendRow(rowData);
      successCount++;
    }

    SpreadsheetApp.flush();

    auditLog_(auth.user, 'HRM_ATT', 'LEAVE_SUBMIT', batchId, 'Đơn phép ' + successCount + ' ngày của ' + payload.emp_id + ' - ' + initialStatus);
    return {
      success: true,
      message: "Gửi đơn phép thành công cho " + successCount + " ngày (Mã đơn: " + batchId + ")!",
      batch_id: batchId,
      approval_status: initialStatus,
      approval_total: plan.chain.length
    };

  } catch (err) {
    return { success: false, message: "Lỗi gửi đơn phép: " + err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * DUYỆT / TỪ CHỐI đơn phép THEO TỪNG CẤP.
 *  - Chuỗi duyệt = quản lý trực tiếp -> quản lý của quản lý -> ... (theo manager_id), chốt tại thời điểm nộp đơn.
 *  - Mỗi lần chỉ người ở đúng cấp hiện tại được duyệt. Duyệt xong đơn chuyển lên cấp kế tiếp.
 *  - Từ chối ở bất kỳ cấp nào = kết thúc đơn.
 *  - Cấp cuối duyệt xong = hoàn tất. Nếu cấp cuối không xử lý sau 24 giờ -> hệ thống tự động duyệt (runLeaveAutoApprove_).
 *  - SuperAdmin có thể duyệt vượt cấp (hoàn tất ngay) hoặc từ chối. Không ai (trừ SuperAdmin) được tự duyệt đơn của mình.
 */
function approveHRMLeaveBatchRecord(token, batchOrAttId, actionStatus, rejectNote) {
  if (actionStatus !== 'APPROVE' && actionStatus !== 'REJECT') return fail_('Hành động duyệt không hợp lệ!');
  var permissionAction = actionStatus === 'REJECT' ? 'reject' : 'approve';
  var auth = requireAuth_(token, 'HRM', permissionAction, 'HRM_ATT', 'view-hrm-att');
  if (!auth.ok) return auth.res;
  var user = auth.user;

  try {
    if (!batchOrAttId) return fail_('Thiếu mã đơn cần duyệt!');

    return withLock_(function () {
      var cfg = getHRMATTConfig_();
      var ss = SpreadsheetApp.openById(cfg.DB_HRM_PART1_ID);
      var sheet = ss.getSheetByName(cfg.SHEET_HRM_ATT);
      if (!sheet || sheet.getLastRow() <= 1) return fail_('Bảng dữ liệu trống!');

      var hm = ensureATTApprovalCols_(sheet);
      var attIdIdx = hm['att_id'], statusIdx = hm['approval_status'], logIdx = hm['reject_reason'], empIdIdx = hm['emp_id'];
      var chainIdx = hm['approval_chain'], stepIdx = hm['approval_step'], dueIdx = hm['approval_due_at'], delIdx = hm['is_deleted'];
      if (attIdIdx === undefined || statusIdx === undefined || empIdIdx === undefined) return fail_('Cấu trúc cột dữ liệu không hợp lệ!');

      // Các mã đơn cần xử lý
      var targets = [];
      if (Array.isArray(batchOrAttId)) {
        batchOrAttId.forEach(function (t) { var k = cleanCompareKey_(t); if (k) targets.push(k); });
      } else {
        var prefix = cleanCompareKey_(batchOrAttId);
        if (prefix.indexOf('_') !== -1 && prefix.indexOf('leave_') !== -1) prefix = prefix.split('_').slice(0, 3).join('_');
        if (prefix) targets.push(prefix);
      }
      if (targets.length === 0) return fail_('Thiếu mã đơn cần duyệt!');
      if (targets.length > 200) return fail_('Mỗi lần chỉ xử lý tối đa 200 đơn.');
      var matches = function (rowAttId) {
        for (var t = 0; t < targets.length; t++) {
          if (rowAttId === targets[t] || rowAttId.indexOf(targets[t] + '_') === 0) return true;
        }
        return false;
      };

      // Danh tính người duyệt (từ phiên đăng nhập)
      var isSuper = isSuperAdminRole_(user.role);
      var myEmp = user.cccd ? empFindByCccd_(user.cccd) : null;
      var myKey = myEmp ? cleanCompareKey_(myEmp.emp_id) : '';
      var names = empNameMap_();
      var nameOf = function (id) { var nm = names[cleanCompareKey_(id)]; return nm ? nm + ' (' + id + ')' : String(id); };
      var hierCache = {};
      var chainOf = function (empId) {
        var k = cleanCompareKey_(empId);
        if (hierCache[k] === undefined) hierCache[k] = getApprovalHierarchyForEmployee_(empId).hierarchy;
        return hierCache[k];
      };

      // Chỉ đọc các cột cần thiết (tránh đọc cả ảnh selfie)
      var n = sheet.getLastRow() - 1;
      var col = function (idx) { return idx === undefined ? null : sheet.getRange(2, idx + 1, n, 1).getValues(); };
      var cAtt = col(attIdIdx), cSt = col(statusIdx), cEmp = col(empIdIdx), cLog = col(logIdx),
          cChain = col(chainIdx), cStep = col(stepIdx), cDel = col(delIdx),
          cOwner = col(hm['created_by']), cDept = col(hm['dept_code']);

      var now = new Date();
      var nowStr = Utilities.formatDate(now, 'GMT+7', 'dd/MM/yy HH:mm:ss');
      var note = String(rejectNote || '').replace(/[\r\n]+/g, ' ').substring(0, 300);
      var dueNext = new Date(now.getTime() + APPROVAL_AUTO_HOURS_ * 3600 * 1000);
      var mStatus = {}, mStep = {}, mDue = {}, mChain = {}, mLog = {};
      var matched = 0, updated = 0, notPending = 0, denied = [], outcome = null;

      for (var i = 0; i < n; i++) {
        if (cDel && String(cDel[i][0]).trim().toUpperCase() === 'Y') continue;
        if (!matches(cleanCompareKey_(cAtt[i][0]))) continue;
        matched++;
        var scopeRow = [cOwner ? cOwner[i][0] : '', cDept ? cDept[i][0] : '', cSt[i][0]];
        if (!matrixScopeAllowed_(user, 'HRM_ATT', scopeRow, { owner: 0, dept: 1, status: 2 }, null, 'read')) {
          denied.push('scope');
          continue;
        }

        if ((getRawText_(cSt[i][0]) || 'Đã duyệt') !== 'Chờ duyệt') { notPending++; continue; }

        var rowEmpId = getRawText_(cEmp[i][0]);
        var rowEmpKey = cleanCompareKey_(rowEmpId);
        var chain = cChain ? parseChain_(getRawText_(cChain[i][0])) : [];
        var step = cStep ? (parseInt(cStep[i][0], 10) || 1) : 1;
        var legacy = false;
        if (!chain.length) { chain = chainOf(rowEmpId); step = 1; legacy = true; }   // đơn cũ chưa có chuỗi duyệt
        step = Math.min(Math.max(step, 1), Math.max(chain.length, 1));

        if (!isSuper && myKey && rowEmpKey === myKey) { denied.push('self'); continue; }

        var level = 0;
        if (isSuper) {
          level = -1;
        } else if (myKey && chain.length && cleanCompareKey_(chain[step - 1]) === myKey) {
          level = step;
        } else {
          var pos = chain.map(function (x) { return cleanCompareKey_(x); }).indexOf(myKey);
          denied.push(!myKey || pos === -1 ? 'notin' : (pos + 1 > step ? 'wait' : 'passed'));
          continue;
        }

        var r = i + 2, entry, total = chain.length;
        var who = '[' + user.username + (myEmp ? ' - ' + myEmp.emp_id : '') + ']';
        if (actionStatus === 'REJECT') {
          mStatus[r] = 'Từ chối'; mStep[r] = ''; mDue[r] = '';
          entry = '[' + nowStr + '] TỪ CHỐI' + (level === -1 ? ' (SuperAdmin)' : ' cấp ' + level + '/' + total) + ' bởi ' + who + (note ? ' - Lý do: ' + note : '');
          outcome = { type: 'reject' };
        } else if (level === -1 || step >= total) {
          mStatus[r] = 'Đã duyệt'; mStep[r] = ''; mDue[r] = '';
          entry = '[' + nowStr + '] DUYỆT ' + (level === -1 ? 'vượt cấp (SuperAdmin)' : 'cấp cuối ' + step + '/' + total + ' - hoàn tất') + ' bởi ' + who;
          outcome = { type: 'final' };
        } else {
          var next = step + 1;
          mStep[r] = next;
          mDue[r] = (next === total) ? dueNext : '';   // tới cấp cuối: bắt đầu đếm 24 giờ tự động duyệt
          entry = '[' + nowStr + '] DUYỆT cấp ' + step + '/' + total + ' bởi ' + who + ' -> chuyển cấp ' + next + '/' + total + ': ' + nameOf(chain[next - 1]) +
                  (next === total ? ' (tự động duyệt nếu quá ' + APPROVAL_AUTO_HOURS_ + ' giờ)' : '');
          outcome = { type: 'advance', next: next, total: total, approver: nameOf(chain[next - 1]) };
        }
        if (legacy) mChain[r] = chain.join('>');
        var oldLog = cLog ? getRawText_(cLog[i][0]) : '';
        mLog[r] = oldLog ? oldLog + ' | ' + entry : entry;
        updated++;
      }

      if (updated === 0) {
        if (matched === 0) return fail_('Không tìm thấy đơn phép phù hợp.');
        if (denied.length) {
          var msgs = {
            self: 'Bạn không thể tự duyệt đơn của chính mình.',
            wait: 'Chưa đến lượt bạn duyệt: đơn đang chờ cấp duyệt trước đó.',
            passed: 'Bạn đã duyệt cấp của mình rồi, đơn đang chờ cấp trên.',
            notin: 'Bạn không thuộc chuỗi duyệt của đơn này.',
            scope: 'Đơn nằm ngoài phạm vi dữ liệu được phân quyền.'
          };
          return fail_(msgs[denied[0]], { code: 'FORBIDDEN' });
        }
        return fail_('Đơn này đã được xử lý xong, không còn ở trạng thái chờ duyệt.');
      }

      if (logIdx !== undefined) writeColByRows_(sheet, logIdx + 1, mLog);
      if (Object.keys(mStatus).length) writeColByRows_(sheet, statusIdx + 1, mStatus);
      if (stepIdx !== undefined && Object.keys(mStep).length) writeColByRows_(sheet, stepIdx + 1, mStep);
      if (dueIdx !== undefined && Object.keys(mDue).length) writeColByRows_(sheet, dueIdx + 1, mDue);
      if (chainIdx !== undefined && Object.keys(mChain).length) writeColByRows_(sheet, chainIdx + 1, mChain);
      SpreadsheetApp.flush();

      var message;
      if (outcome.type === 'reject') message = 'Đã từ chối ' + updated + ' bản ghi.';
      else if (outcome.type === 'final') message = 'Đã duyệt hoàn tất ' + updated + ' bản ghi.';
      else message = 'Đã duyệt cấp của bạn (' + updated + ' bản ghi). Đơn chuyển lên cấp ' + outcome.next + '/' + outcome.total + ': ' + outcome.approver + '.';

      auditLog_(user, 'HRM_ATT', outcome.type === 'reject' ? 'REJECT' : (outcome.type === 'final' ? 'APPROVE' : 'APPROVE_LEVEL'),
                sampleList_(targets, 5), message + (note ? ' - ' + note : ''));
      return { success: true, message: message, updated_count: updated, final: outcome.type !== 'advance' };
    });
  } catch (err) {
    Logger.log('approveHRMLeaveBatchRecord: ' + err);
    return fail_('Lỗi duyệt đơn phép.');
  }
}

/**
 * Xác định nhân viên được ghi dữ liệu và quyền ghi:
 *  - Chính mình (CCCD liên kết của tài khoản trùng CCCD nhân viên) -> được phép.
 *  - Quyền ghi thay người khác được kiểm tra bằng hành động riêng sau khi xác định nhân viên.
 */
function resolveEmpForWrite_(user, empId) {
  var emp = empFindByEmpId_(empId);
  if (!emp) {
    return { ok: false, res: fail_('Không tìm thấy nhân viên có mã "' + cleanId_(empId) + '" trong Hồ sơ Nhân viên.') };
  }
  var isSelf = !!user.cccd && cleanId_(user.cccd) === cleanId_(emp.cccd_no);
  return { ok: true, emp: emp, isSelf: isSelf };
}

/**
 * CHẠY MỘT LẦN (chủ dự án, trong trình soạn thảo): đánh lại mã cho các bản ghi chấm công bị TRÙNG att_id
 * (lỗi cũ: 2 bản ghi trong cùng 1 giây cùng mã). Bản ghi đầu tiên giữ nguyên, các bản trùng được thêm hậu tố _D2, _D3...
 */
function fixDuplicateAttIds() {
  if (!isOwnerRun_()) return { success: false, message: 'Không được phép.' };
  var cfg = getHRMATTConfig_();
  var sheet = SpreadsheetApp.openById(cfg.DB_HRM_PART1_ID).getSheetByName(cfg.SHEET_HRM_ATT);
  if (!sheet || sheet.getLastRow() <= 1) { Logger.log('Bảng chấm công trống.'); return { success: true, changed: 0 }; }

  var idx = getHeaderMap_(sheet)['att_id'];
  if (idx === undefined) { Logger.log('Không có cột att_id.'); return { success: false }; }

  var n = sheet.getLastRow() - 1;
  var range = sheet.getRange(2, idx + 1, n, 1);
  var vals = range.getValues();
  var seen = {}, changed = 0;
  for (var i = 0; i < n; i++) {
    var raw = getRawText_(vals[i][0]);
    var key = cleanCompareKey_(raw);
    if (!key) continue;
    if (seen[key]) {
      seen[key]++;
      var newId = raw + '_D' + seen[key];
      while (seen[cleanCompareKey_(newId)]) { seen[key]++; newId = raw + '_D' + seen[key]; }
      vals[i][0] = "'" + newId;
      seen[cleanCompareKey_(newId)] = 1;
      changed++;
    } else {
      seen[key] = 1;
    }
  }
  if (changed > 0) range.setValues(vals);
  Logger.log('Đã đánh lại mã cho ' + changed + ' bản ghi bị trùng att_id.');
  return { success: true, changed: changed };
}

/* ============================================================
 * ĐỢT 2c - DUYỆT PHÉP THEO TỪNG CẤP
 * ============================================================ */

var APPROVAL_AUTO_HOURS_ = 24;   // Cấp cuối tự động duyệt sau 24 giờ kể từ khi cấp trước duyệt

function parseChain_(text) {
  return String(text || '').split('>').map(function (x) { return cleanId_(x); }).filter(function (x) { return x; });
}

/** Thêm 4 cột duyệt theo cấp vào cuối bảng chấm công nếu chưa có. Trả về headerMap mới. */
function ensureATTApprovalCols_(sheet) {
  var hm = getHeaderMap_(sheet);
  ['approval_chain', 'approval_step', 'approval_due_at', 'is_auto_approved'].forEach(function (h) {
    if (hm[h] === undefined) {
      var nextCol = sheet.getLastColumn() + 1;
      if (sheet.getMaxColumns() < nextCol) sheet.insertColumnsAfter(sheet.getMaxColumns(), nextCol - sheet.getMaxColumns());
      sheet.getRange(1, nextCol).setValue(h);
      hm = getHeaderMap_(sheet);
    }
  });
  return hm;
}

/**
 * Kế hoạch duyệt của một nhân viên tại thời điểm nộp đơn.
 * Không có quản lý nào -> tự động "Đã duyệt". Chuỗi chỉ có 1 cấp -> cấp đó cũng là cấp cuối nên đếm 24 giờ ngay từ lúc nộp.
 */
function buildApprovalPlan_(empId, now) {
  var chain = getApprovalHierarchyForEmployee_(empId).hierarchy;
  if (!chain.length) return { status: 'Đã duyệt', chain: [], chainText: '', step: '', dueAt: '' };
  return {
    status: 'Chờ duyệt',
    chain: chain,
    chainText: chain.join('>'),
    step: 1,
    dueAt: chain.length === 1 ? new Date(now.getTime() + APPROVAL_AUTO_HOURS_ * 3600 * 1000) : ''
  };
}

/**
 * Tự động duyệt các đơn ở CẤP CUỐI đã quá hạn (due_at <= bây giờ).
 * Đánh dấu is_auto_approved = Y và ghi rõ vào nhật ký của đơn.
 */
function runLeaveAutoApprove_() {
  return withLock_(function () {
    var cfg = getHRMATTConfig_();
    var sheet = SpreadsheetApp.openById(cfg.DB_HRM_PART1_ID).getSheetByName(cfg.SHEET_HRM_ATT);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, count: 0 };

    var hm = getHeaderMap_(sheet);
    var stIdx = hm['approval_status'], chIdx = hm['approval_chain'], spIdx = hm['approval_step'], dueIdx = hm['approval_due_at'];
    if (stIdx === undefined || chIdx === undefined || spIdx === undefined || dueIdx === undefined) return { success: true, count: 0 };

    var autoIdx = hm['is_auto_approved'], logIdx = hm['reject_reason'], idIdx = hm['att_id'], delIdx = hm['is_deleted'];
    var n = sheet.getLastRow() - 1;
    var col = function (idx) { return idx === undefined ? null : sheet.getRange(2, idx + 1, n, 1).getValues(); };
    var cSt = col(stIdx), cCh = col(chIdx), cSp = col(spIdx), cDue = col(dueIdx), cLog = col(logIdx), cId = col(idIdx), cDel = col(delIdx);

    var now = new Date();
    var nowStr = Utilities.formatDate(now, 'GMT+7', 'dd/MM/yy HH:mm:ss');
    var names = null;
    var mStatus = {}, mStep = {}, mDue = {}, mAuto = {}, mLog = {}, batches = {};

    for (var i = 0; i < n; i++) {
      if (cDel && String(cDel[i][0]).trim().toUpperCase() === 'Y') continue;
      if (getRawText_(cSt[i][0]) !== 'Chờ duyệt') continue;
      var dv = cDue[i][0];
      if (dv === '' || dv == null) continue;
      var due = (dv instanceof Date) ? dv : new Date(dv);
      if (isNaN(due.getTime()) || due.getTime() > now.getTime()) continue;

      var chain = parseChain_(getRawText_(cCh[i][0]));
      var step = parseInt(cSp[i][0], 10) || 0;
      if (!chain.length || step !== chain.length) continue;          // chỉ áp dụng cho cấp cuối

      if (!names) names = empNameMap_();
      var last = chain[chain.length - 1];
      var lastName = names[cleanCompareKey_(last)] ? names[cleanCompareKey_(last)] + ' (' + last + ')' : last;
      var r = i + 2;
      mStatus[r] = 'Đã duyệt'; mStep[r] = ''; mDue[r] = ''; mAuto[r] = 'Y';
      var entry = '[' + nowStr + '] TỰ ĐỘNG DUYỆT bởi hệ thống: cấp cuối ' + step + '/' + chain.length + ' [' + lastName + '] chưa xử lý sau ' +
                  APPROVAL_AUTO_HOURS_ + ' giờ kể từ cấp duyệt trước (hạn ' + Utilities.formatDate(due, 'GMT+7', 'dd/MM/yy HH:mm') + ')';
      var oldLog = cLog ? getRawText_(cLog[i][0]) : '';
      mLog[r] = oldLog ? oldLog + ' | ' + entry : entry;
      if (cId) batches[getRawText_(cId[i][0]).split('_').slice(0, 3).join('_')] = true;
    }

    var count = Object.keys(mStatus).length;
    if (count === 0) return { success: true, count: 0 };

    var autoCol = hm['is_auto_approved'];
    if (autoCol === undefined) autoCol = ensureATTApprovalCols_(sheet)['is_auto_approved'];
    writeColByRows_(sheet, stIdx + 1, mStatus);
    writeColByRows_(sheet, spIdx + 1, mStep);
    writeColByRows_(sheet, dueIdx + 1, mDue);
    writeColByRows_(sheet, autoCol + 1, mAuto);
    if (logIdx !== undefined) writeColByRows_(sheet, logIdx + 1, mLog);
    SpreadsheetApp.flush();

    auditLog_('SYSTEM', 'HRM_ATT', 'AUTO_APPROVE', sampleList_(Object.keys(batches), 5), 'Tự động duyệt ' + count + ' bản ghi phép (cấp cuối quá ' + APPROVAL_AUTO_HOURS_ + ' giờ)');
    return { success: true, count: count };
  });
}

/** Hàm chạy theo hẹn giờ (30 phút/lần). An toàn nếu bị gọi thêm: chỉ áp dụng đúng quy tắc 24 giờ. */
function autoApproveTick() {
  try {
    var c = CacheService.getScriptCache();
    if (c.get('LV_TICK')) return;          // đã chạy trong vòng 60 giây qua
    c.put('LV_TICK', '1', 60);
    runLeaveAutoApprove_();
  } catch (e) { Logger.log('autoApproveTick: ' + e); }
}

/** CHẠY MỘT LẦN (chủ dự án, trong trình soạn thảo): cài bộ hẹn giờ kiểm tra tự động duyệt mỗi 30 phút. */
function installAutoApproveTrigger() {
  if (!isOwnerRun_()) return { success: false, message: 'Không được phép.' };
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'autoApproveTick') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('autoApproveTick').timeBased().everyMinutes(30).create();
  Logger.log('Đã cài bộ hẹn giờ autoApproveTick: chạy mỗi 30 phút.');
  return { success: true };
}

function removeAutoApproveTrigger() {
  if (!isOwnerRun_()) return { success: false, message: 'Không được phép.' };
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'autoApproveTick') { ScriptApp.deleteTrigger(t); n++; }
  });
  Logger.log('Đã gỡ ' + n + ' bộ hẹn giờ.');
  return { success: true, removed: n };
}
