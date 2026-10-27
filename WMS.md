# Kiến trúc phân hệ WMS — HAP HOLDING ERP

Phân hệ Quản lý Kho (Warehouse Management System) được xây dựng trên nền tảng Google Apps Script (runtime V8) + Google Sheets làm cơ sở dữ liệu. Toàn bộ mã nguồn WMS nằm trong thư mục `C:\Users\hadon\Desktop\HAP\erpHAP\`.

---

## 1. Tổng quan phân hệ WMS

| Mục | Chi tiết |
|-----|----------|
| Nền tảng | Google Apps Script (V8) + Google Sheets + Google Drive |
| Múi giờ | Asia/Ho_Chi_Minh (GMT+7) |
| Ngôn ngữ hiển thị | Tiếng Việt |
| Cổng vào | `CORE_Router.js` → `handleApiRoute('WMS', action, payload)` |
| Khóa toàn cục | `LockService.getScriptLock()` qua `withLock_()` trong `CORE_Utils.js` |
| Cache | `CacheService` (TTL 120s) cho bảng tham chiếu ít thay đổi |

Phân hệ WMS hiện đã triển khai đầy đủ: danh mục kho, vị trí, SKU, phiếu nhập/xuất, luồng duyệt nhiều cấp và ghi sổ kho idempotent.

---

## 2. Danh mục tệp và trách nhiệm

### 2.1 Backend `.js` (8 tệp)

| Tệp | Trách nhiệm chính | Sheet liên quan |
|-----|-------------------|-----------------|
| `WMS_Core_Server.js` | Enum trạng thái, helpers dùng chung, khóa ghi sổ idempotent | Tất cả sheet WMS |
| `WMS_WAREHOUSE_Server.js` | Danh mục kho: CRUD kho | `WAREHOUSE` |
| `WMS_LOCATION_Server.js` | Danh mục vị trí theo kho, sắp xếp dạng cây | `LOCATION` |
| `WMS_SKU_Server.js` | Danh mục hàng hóa: sinh mã ngắn, import/export | `SKU` |
| `WMS_DOC_Server.js` | Phiếu nhập/xuất + dòng chi tiết, cấp số phiếu, CRUD nháp | `WMS_DOC`, `WMS_LINE` |
| `WMS_POST_Server.js` | Ghi sổ kho idempotent (WMS_POST, WMS_LEDGER, WMS_BAL) | `WMS_POST`, `WMS_LEDGER`, `WMS_BAL` |
| `WMS_WORKFLOW_Server.js` | Luồng duyệt có phiên bản, bước duyệt, run/task/log | `WMS_WF`, `WMS_WF_STEP`, `WMS_WF_RUN`, `WMS_WF_TASK`, `WMS_WF_LOG`, `WMS_USER_MAP` |
| `WMS_SETUP_Server.js` | Khởi tạo 16 sheet WMS, validate schema | 16 sheet WMS |

### 2.2 Frontend `.html` (8 tệp)

| Tệp | Vai trò | View container |
|-----|---------|----------------|
| `UI_WMS_WAREHOUSE.html` | Tab Kho + Sơ đồ vị trí (4 cấp) | `view-wms-warehouse` |
| `UI_WMS_SKU.html` | Danh mục SKU + import Excel + QR mã | `view-wms-sku` |
| `UI_WMS_DOCS_IN.html`, `UI_WMS_DOCS_OUT.html`, `UI_WMS_DOCS_MODALS.html`, `UI_WMS_DOCS_SCRIPT.html`, `UI_WMS_DOCS_STYLE.html` | Phiếu nhập (`view-wms-in`) & xuất (`view-wms-out`) | `view-wms-in`, `view-wms-out` |
| `UI_WMS_WORKFLOW.html` | Cấu hình luồng duyệt + hộp thư duyệt | `view-wms-wf`, `view-wms-wf-inbox` |

### 2.3 Core (được chia sẻ toàn hệ thống)

| Tệp | Vai trò |
|-----|---------|
| `CORE_Config.js` | `CONFIG` (ID spreadsheet, tên sheet, hằng số) |
| `CORE_Router.js` | `doGet`, `include`, `handleApiRoute`, routing module |
| `CORE_Utils.js` | `withLock_`, `safeText_`, `fail_`, `fmtDate_`, `requestMemo_` |

---

## 3. Mô hình dữ liệu WMS (16 sheet)

Tất cả sheet được khai báo trong `CORE_Config.js` và khởi tạo bởi `setupWmsDataSheets()` trong `WMS_SETUP_Server.js`. Hàng 1 là header, dữ liệu từ hàng 2. Mọi bảng đều có cột `is_deleted` (soft delete) và cột audit `created_by | created_at | updated_by | updated_at`.

| Sheet | Khóa chính | Mục đích |
|-------|-----------|----------|
| `WAREHOUSE` | `warehouse_id` | Kho: mã, tên, địa chỉ, mục đích, chế độ bảo quản, dải nhiệt độ |
| `LOCATION` | `location_id` | Vị trí trong kho, có `parent_location_id` (cây vị trí 4 cấp) |
| `SKU` | `sku_id` | Hàng hóa: 16 nhóm, `uom_base/uom_pack/pack_ratio`, `is_batch_managed`, `is_expiry_managed`, `qc_inspection_req`, min/max stock |
| `WMS_DOC` | `doc_id` | Header phiếu (31 cột) — trạng thái duyệt/ghi sổ/phiếu |
| `WMS_LINE` | `line_id` | Dòng phiếu (35 cột) — quy đổi đơn vị, lô, vị trí, giá vốn |
| `WMS_POST` | `posting_id` | Bản ghi ghi sổ, idempotent qua `posting_key` |
| `WMS_LEDGER` | `ledger_id` | Bút toán kho, idempotent qua `transaction_key` |
| `WMS_BAL` | `balance_id` | Tồn kho tổng hợp theo sku/kho/vị trí/lô/trạng thái |
| `WMS_LOT` | `lot_id` | Lô: hạn dùng, nhà cung cấp, trạng thái chất lượng |
| `WMS_QC` | `qc_event_id` | Sự kiện kiểm định chất lượng |
| `WMS_WF` | `workflow_version_id` | Phiên bản workflow (bất biến sau khi kích hoạt) |
| `WMS_WF_STEP` | `workflow_step_id` | Bước duyệt (đánh dấu `is_stock_posting_step` cho bước ghi sổ) |
| `WMS_WF_RUN` | `workflow_run_id` | Lần chạy cho một phiếu |
| `WMS_WF_TASK` | `task_id` | Task giao cho người duyệt |
| `WMS_WF_LOG` | `action_id` | Nhật ký quyết định duyệt |
| `WMS_USER_MAP` | `mapping_id` | Cầu nối `username` ↔ `emp_id` |

---

## 4. Enum trạng thái (`WMS_Core_Server.js`)

| Enum | Giá trị |
|------|---------|
| `WMS_DOC_TYPES_` | `PURCHASE_RECEIPT`, `SALES_ISSUE`, `CUSTOMER_RETURN`, `PRODUCTION_ISSUE`, `PRODUCTION_RECEIPT`, `TRANSFER`, `STOCK_ADJUSTMENT` (chỉ 2 loại đầu có UI) |
| `WMS_APPROVAL_STATUS_` | `DRAFT`, `PENDING`, `RETURNED`, `APPROVED`, `CANCELLED` |
| `WMS_POSTING_STATUS_` | `NOT_POSTED`, `POSTING`, `POSTED`, `POST_FAILED`, `REVERSED` |
| `WMS_DOC_STATUS_` | `DRAFT`, `IN_APPROVAL`, `RETURNED`, `APPROVED`, `COMPLETED`, `CANCELLED` |
| `WMS_STOCK_STATUS_` | `AVAILABLE`, `QUARANTINE` |
| `WMS_LOT_QUALITY_STATUS_` | `PENDING`, `PASSED`, `FAILED`, `CONDITIONAL` |
| `WMS_LOT_STATUS_` | `ACTIVE`, `DEPLETED`, `BLOCKED`, `CLOSED` |
| `WMS_QC_DISPOSITION_` | `RELEASE`, `QUARANTINE`, `REJECT`, `REWORK` |
| `WMS_LEDGER_TYPES_` | `RECEIPT`, `ISSUE`, `TRANSFER_IN`, `TRANSFER_OUT`, `QC_QUARANTINE`, `QC_RELEASE`, `ADJUSTMENT_IN`, `ADJUSTMENT_OUT`, `REVERSAL` |
| `WMS_WORKFLOW_STATUS_` | `DRAFT`, `ACTIVE`, `INACTIVE`, `ARCHIVED` |
| `WMS_WORKFLOW_RUN_STATUS_` | `PENDING`, `RETURNED`, `COMPLETED`, `CANCELLED`, `SUPERSEDED` |
| `WMS_WORKFLOW_TASK_STATUS_` | `PENDING`, `APPROVED`, `RETURNED`, `SKIPPED` |
| `WMS_WORKFLOW_ACTIONS_` | `SUBMIT`, `APPROVE`, `RETURN_TO_REQUESTER`, `CHANGE_WORKFLOW`, `STOCK_POSTED`, `STOCK_REVERSED`, `CANCEL` |
| `WMS_POSTING_PHASES_` | `STARTED`, `LEDGER_WRITTEN`, `BALANCE_WRITTEN`, `COMPLETED`, `FAILED` |
| `WMS_PRODUCTION_STATUS_` | `DRAFT`, `IN_PROGRESS`, `COMPLETED`, `CANCELLED` |
| `WMS_ATTACHMENT_STATUS_` | `ACTIVE`, `ARCHIVED` |
| `WMS_USER_MAP_STATUS_` | `ACTIVE`, `INACTIVE` |

---

## 5. Vòng đời phiếu nhập/xuất

```
[DRAFT] --wmsSaveDoc--> [DRAFT/RETURNED]
  |--wmsWorkflowSubmit--> [IN_APPROVAL]
       |--duyệt bước 1..n-1--> [IN_APPROVAL]
       |--duyệt bước cuối--> [APPROVED]
       |--RETURN (bắt buộc lý do)--> [RETURNED]
  [APPROVED] --wmsPost--> [COMPLETED] (ghi sổ tồn kho)
  [DRAFT/RETURNED] --xóa mềm--> [xóa]
```

Trạng thái lưu ở 3 cột độc lập trong `WMS_DOC`:
- `doc_status`: `DRAFT, IN_APPROVAL, RETURNED, APPROVED, COMPLETED, CANCELLED`
- `approval_status`: `DRAFT, PENDING, RETURNED, APPROVED, CANCELLED`
- `posting_status`: `NOT_POSTED, POSTING, POSTED, POST_FAILED, REVERSED`

---

## 6. Loại phiếu (`WMS_DOC_TYPES_`)

| Loại | Mã | Trạng thái | View |
|------|-----|-----------|------|
| Nhập kho | `PURCHASE_RECEIPT` | ✅ Có UI | `view-wms-in` |
| Xuất kho | `SALES_ISSUE` | ✅ Có UI | `view-wms-out` |
| Hoàn trả từ khách hàng | `CUSTOMER_RETURN` | 🚧 Chỉ khai báo, chưa có UI | — |
| Xuất sản xuất | `PRODUCTION_ISSUE` | 🚧 Chỉ khai báo | — |
| Nhập sản xuất | `PRODUCTION_RECEIPT` | 🚧 Chỉ khai báo | — |
| Chuyển kho | `TRANSFER` | 🚧 Chỉ khai báo | — |
| Điều chỉnh tồn | `STOCK_ADJUSTMENT` | 🚧 Chỉ khai báo | — |

---

## 7. Ghi sổ kho (Posting)

Thiết kế đã có đầy đủ hợp đồng và hàm thực thi `wmsPost` trong `WMS_POST_Server.js`:

- **Khóa idempotent**: `posting_key = WMS_POST_V1|<doc_id>|<run_id>|<step_id>|<posting_type>`
- **Transaction key**: `<posting_key>|<line_id>|<leg>` (tránh nhân đôi bút toán khi retry)
- **Pha ghi sổ**: `STARTED → LEDGER_WRITTEN → BALANCE_WRITTEN → COMPLETED` (hoặc `FAILED`)
- **Chống tồn âm**: kiểm tra `WMS_BAL` trước khi ghi cho phiếu xuất (`SALES_ISSUE`)
- **Không sửa/xóa dòng WMS_LEDGER đã ghi**: reversal được dùng để hoàn tác

Hàm chính: `wmsPost(token, docId)` — idempotent, gọi lại khi đã `POSTED` trả `alreadyPosted = true`.

---

## 8. Luồng duyệt (Workflow)

- **Versioned**: `WMS_WF` + `WMS_WF_STEP` bất biến sau khi `ACTIVE`; sửa = tạo phiên bản mới
- **Mô hình**: tuần tự, đúng một người duyệt mỗi bước (bổ sung duyệt song song kiểu "một-đủ" là mở rộng sau)
- **Bước ghi sổ**: bước có `is_stock_posting_step = true` cho phép `wmsPost`
- **Trạng thái run**: `PENDING, RETURNED, COMPLETED, CANCELLED, SUPERSEDED`
- **Trạng thái task**: `PENDING, APPROVED, RETURNED, SKIPPED`
- **Chống tự duyệt**: người duyệt không thể duyệt phiếu của mình trừ khi `can_self_approve = true`
- **Toàn bộ nằm trong `withLock_`**; rollback thủ công nếu lỗi giữa chừng

---

## 9. Phân quyền (RBAC + ABAC)

Hai lớp kiểm soát (theo `ARCHITECTURE.md` §6):

1. **Token hợp lệ** (regex 64 ký tự hex) + `CacheService` có `SES_<token>`
2. **SYS_USR còn `is_active`**
3. **Quyền MENU**: `nav_access` chứa `navigationKey`
4. **Quyền HÀNH ĐỘNG**: `canRead/Create/Edit/Delete` hoặc `custom_actions`
5. **Phạm vi dòng dữ liệu**: `matrixScopeAllowed_` (dựa trên `MATRIX_RULE_V2` trong `SYS_PERM`)

### Bí danh tên sheet WMS
WMS dùng tên sheet ngắn (`DOC`, `LINE`) trong khi `SYS_PERM` lưu tên thật (`WMS_DOC`, `WMS_LINE`). Helper `getPermissionSheetNames_()` trong `SYS_Server.js` cầu nối hai cách viết.

### Vai trò
| Vai trò | Quyền đặc biệt |
|---------|---------------|
| `SuperAdmin` | Toàn quyền; duy nhất vào `view-sys-admin` (Thùng rác, Nhật ký) |
| `ADMIN` | Quản lý tài khoản USER, ma trận phân quyền, quyền menu |
| `DIRECTOR`, `DEPARTMENT_MANAGER`, `EMPLOYEE` | Theo ma trận phân quyền |

---

## 10. Router WMS (`CORE_Router.js` → `routeWmsModule`)

| Action | Hàm backend |
|--------|-------------|
| `listWarehouses` | `wmsListWarehouses` |
| `saveWarehouse` | `wmsSaveWarehouse` |
| `deleteWarehouse` | `wmsDeleteWarehouse` |
| `listLocations` | `wmsListLocations` |
| `saveLocation` | `wmsSaveLocation` |
| `deleteLocation` | `wmsDeleteLocation` |
| `listSkus` | `wmsListSkus` |
| `suggestSkuCode` | `wmsSuggestSkuCode` |
| `saveSku` | `wmsSaveSku` |
| `importSkus` | `wmsImportSkus` |
| `deleteSku` | `wmsDeleteSku` |
| `deleteSkus` | `wmsDeleteSkus` |
| `listDocOptions` | `wmsListDocOptions` |
| `listDocs` | `wmsListDocs` |
| `getDoc` | `wmsGetDoc` |
| `saveDoc` | `wmsSaveDoc` |
| `deleteDocs` | `wmsDeleteDocs` |
| `getWorkflowSetup` | `wmsWorkflowGetSetup` |
| `saveWorkflow` | `wmsWorkflowSave` |
| `activateWorkflow` | `wmsWorkflowActivate` |
| `setWorkflowStatus` | `wmsWorkflowSetStatus` |
| `saveWorkflowMapping` | `wmsWorkflowSaveMapping` |
| `setWorkflowMappingStatus` | `wmsWorkflowSetMappingStatus` |
| `listMyWorkflowTasks` | `wmsWorkflowMyTasks` |
| `submitWorkflow` | `wmsWorkflowSubmit` |
| `decideWorkflowTask` | `wmsWorkflowDecide` |

---

## 11. Quy ước lập trình WMS

- Hàm có hậu tố `_` = nội bộ, không gọi từ trình duyệt được (ví dụ `wmsCanRead_`, `wmsSoftDelete_`)
- Hàm không có hậu tố = public, gọi được từ client qua `google.script.run` hoặc `handleApiRoute`
- Mọi thao tác ghi bọc trong `withLock_()` để tránh ghi chồng
- Kiểm tra optimistic concurrency qua `expected_updated_at` → `STALE_RECORD`
- Xóa mềm qua `wmsSoftDelete_()` (đánh dấu `is_deleted = true`, giữ dòng trong sheet)
- Cache 2 cấp (bộ nhớ + `CacheService`) TTL 120s, invalidate sau `save`/`update`/`delete`
- Ghi theo nguyên dòng (`setValues([row])`) để giảm số lần gọi API Sheets

---

## 12. Trạng thái hiện tại & khuyến nghị phát triển

| Thành phần | Trạng thái |
|------------|-----------|
| Danh mục kho/vị trí/SKU | ✅ Hoàn toàn |
| Phiếu nhập/xuất (nháp, duyệt, xóa) | ✅ Hoàn toàn |
| Luồng duyệt nhiều cấp | ✅ Hoàn toàn |
| Ghi sổ kho (`wmsPost`) | ✅ Có hàm thực thi |
| Cập nhật tồn kho (`WMS_BAL`) sau duyệt | ⚠️ Chưa được gọi tự động sau khi phiếu `APPROVED` |
| Sổ cái (`WMS_LEDGER`) | ✅ Có schema + hàm ghi |
| Báo cáo tồn kho (`view-wms-inv`) | 🚧 Chỉ là khung giao diện |
| Loại phiếu ngoài 2 loại chính | 🚧 Chỉ khai báo enum, chưa có UI/luồng |

**Khuyến nghị:** Khi tiếp tục phát triển, cần đảm bảo sau khi phiếu đạt `APPROVED`, gọi `wmsPost(token, docId)` để cập nhật `WMS_LEDGER` và `WMS_BAL`, đồng thời chuyển `posting_status` sang `POSTED` và `doc_status` sang `COMPLETED`.

---

*Tài liệu được cập nhật từ mã nguồn phân hệ WMS của dự án erpHAP (ngày 25/10/2026).*
