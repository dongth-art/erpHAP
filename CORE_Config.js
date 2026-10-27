/**
 * CORE_Config.gs - HAP HOLDING ERP SYSTEM
 * Quản lý Cấu hình ID Database, Mã phân hệ và Ánh xạ Tiêu đề Cột
 */

const CONFIG = {
  // 1. THÔNG TIN CÔNG TY & LOGO
  COMPANY_NAME: "HAP HOLDING",
  COMPANY_LOGO_URL: "https://lh3.googleusercontent.com/d/1f4pu_IXo9UA_TQIBYbcsCBBdLESQCCkW",

  // 2. DUNG LƯỢNG & GIỚI HẠN
  MAX_ROWS_PER_SHEET: 30000,

  // 3. ID FILE DATABASE SYSTEM (SYS)
  DB_SYS_PART1_ID: "1YUMxBFVD0v5KJpqefgt4Dgd3JOiffwgJdJBCKGb_UH8",
  DB_SYS_PART2_ID: "NHẬP_ID_FILE_HAP_SYS_PART2_VÀO_ĐÂY",

  // 4. ID FILE DATABASE HUMAN RESOURCES (HRM)
  DB_HRM_PART1_ID: "1LZ30absUD1iI1zw4DldVqWmaumxC3njg7kQkGdu5a8k",
  DB_HRM_PART2_ID: "NHẬP_ID_FILE_HAP_HRM_PART2_VÀO_ĐÂY",

  // 5. ID FILE DATABASE WAREHOUSE MANAGEMENT (WMS)
  DB_WMS_PART1_ID: "1vacZW8zNB_8EvkPYd0GSt-koxG4eKKGch87Vv5I1kAE",

  // Không có bảng nào được bỏ qua kiểm tra phân quyền ma trận.
  BYPASS_PERMISSION_SHEETS: [],
  DB_DMS_PART1_ID: "13aC5K5nCMLyP5tvrZjNG1pytNGjxoY14kwH2oRLJFUc",
  SHEET_SYS_USR: "SYS_USR",
  SHEET_SYS_AUDIT: "SYS_AUDIT",   // Nhật ký thao tác (tự tạo khi có thao tác đầu tiên)
  SHEET_SYS_MOD: "SYS_MOD",
  SHEET_SYS_PERM: "SYS_PERM",
  SHEET_HRM_ID: "HRM_ID",
  SHEET_HRM_EMP: "HRM_EMP",
  SHEET_HRM_ATT: "HRM_ATT",
  SHEET_WMS_WAREHOUSE: "WAREHOUSE",
  SHEET_WMS_LOCATION: "LOCATION",
  SHEET_WMS_SKU: "SKU",
  SHEET_WMS_DOC: "WMS_DOC",
  SHEET_WMS_LINE: "WMS_LINE",
  SHEET_WMS_POST: "WMS_POST",
  SHEET_WMS_LOT: "WMS_LOT",
  SHEET_WMS_QC: "WMS_QC",
  SHEET_WMS_LEDGER: "WMS_LEDGER",
  SHEET_WMS_BAL: "WMS_BAL",
  SHEET_WMS_WF: "WMS_WF",
  SHEET_WMS_WF_STEP: "WMS_WF_STEP",
  SHEET_WMS_WF_RUN: "WMS_WF_RUN",
  SHEET_WMS_WF_TASK: "WMS_WF_TASK",
  SHEET_WMS_WF_LOG: "WMS_WF_LOG",
  SHEET_WMS_PROD: "WMS_PROD",
  SHEET_WMS_GENEALOGY: "WMS_GENEALOGY",
  SHEET_WMS_ATTACH: "WMS_ATTACH",
  SHEET_WMS_USER_MAP: "WMS_USER_MAP",
  SHEET_DMS_DOCS: "DMS_DOCS",
  SHEET_DMS_VERSIONS: "DMS_VERSIONS",
  SHEET_DMS_COMMENTS: "DMS_COMMENTS",
  SHEET_DMS_READ_ACK: "DMS_READ_ACK",

  // 7. QUY ƯỚC MÃ NHÂN VIÊN: 2026 = năm thành lập công ty, cố định. Mã dạng 20260001, 20260002...
  EMP_ID_PREFIX: "2026",

  // 8. (Tùy chọn) ID thư mục Drive lưu ảnh CCCD. Để trống = tự dùng/tạo thư mục "HAP_ERP_CCCD_Storage"
  DRIVE_FOLDER_ID: ""
};

// Từ điển ánh xạ hiển thị Tiếng Việt cho Bảng HRM_ID
const MAP_HRM_ID = {
  raw_qr: "Mã QR CCCD Gốc",
  cccd_no: "Số CCCD",
  cmnd_no: "Số CMND Cũ",
  full_name: "Họ và Tên",
  gender: "Giới Tính",
  dob: "Ngày Sinh",
  perm_addr: "Nơi Thường Trú",
  merged_addr: "Nơi Thường Trú Sau Sáp Nhập",
  curr_addr: "Nơi Ở Hiện Tại",
  issue_date: "Ngày Cấp CCCD",
  issue_place: "Nơi Cấp",
  img_front: "Ảnh CCCD Mặt Trước",
  img_back: "Ảnh CCCD Mặt Sau",
  created_by: "Người Tạo",
  created_at: "Thời Gian Tạo"
};

// Từ điển ánh xạ hiển thị Tiếng Việt cho Bảng SYS_USR
const MAP_SYS_USR = {
  username: "Tên Đăng Nhập",
  password_hash: "Mật Khẩu Mã Hóa",
  cccd_no: "Số CCCD Liên Kết",
  full_name: "Họ và Tên",
  role: "Vai Trò",
  module_access: "Phân Quyền Phân Hệ",
  is_active: "Trạng Thái",
  last_login: "Đăng Nhập Lần Cuối",
  created_by: "Người Tạo",
  created_at: "Thời Gian Tạo",
  nav_access: "Quyền Truy Cập Chức Năng"
};