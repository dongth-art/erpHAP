/**
 * CORE_Router.gs - HAP HOLDING ERP SYSTEM
 * Điều hướng giao diện và các lệnh gọi API Backend
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('UI_Index')
    .evaluate()
    .setTitle('HAP HOLDING - Hệ Thống Quản Trị Doanh Nghiệp')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // <--- Thêm dòng này
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Endpoint tiếp nhận yêu cầu tập trung từ Client
 */
function handleApiRoute(module, action, payload) {
  try {
    clearRequestMemo_(); // Bắt đầu request mới: bỏ bộ nhớ tạm của request trước.
    payload = payload || {};
    switch (module) {
      case 'SYS':
        return routeSysModule(action, payload);
      case 'HRM':
        return routeHrmModule(action, payload);
      case 'WMS':
        return routeWmsModule(action, payload);
      case 'DMS':
        return routeDmsModule(action, payload);
      default:
        return { success: false, message: 'Module không hợp lệ: ' + module };
    }
  } catch (err) {
    return { success: false, message: 'Lỗi CORE Router: ' + err.toString() };
  }
}

function routeSysModule(action, payload) {
  switch (action) {
    case 'verifyLogin':
      return typeof verifyLogin === 'function' ? verifyLogin(payload.username, payload.password) : { success: false, message: 'Hàm verifyLogin chưa khai báo' };
    
    case 'getSysUsersList':
      return typeof getSysUsersList === 'function'
        ? getSysUsersList(payload.token)
        : { success: false, message: 'Hàm getSysUsersList chưa khai báo trong SYS_Server' };

    case 'saveSysUser':
      return typeof saveSysUser === 'function'
        ? saveSysUser(payload.token, payload)
        : { success: false, message: 'Hàm saveSysUser chưa khai báo' };

    case 'updateSysUserStatus':
      return typeof updateSysUserStatus === 'function'
        ? updateSysUserStatus(payload.token, payload.username, payload.status)
        : { success: false, message: 'Hàm updateSysUserStatus chưa khai báo' };

    case 'getSysUserLoginHistory':
      return typeof getSysUserLoginHistory === 'function'
        ? getSysUserLoginHistory(payload.token, payload.username)
        : { success: false, message: 'Hàm getSysUserLoginHistory chưa khai báo' };
    
    case 'getSystemSchema':
      if (typeof getSystemSchemaApi === 'function') return getSystemSchemaApi(payload.token);
      return { success: false, message: 'Hàm getSystemSchema chưa khai báo' };

    case 'savePermissionRule':
      if (typeof savePermissionRuleApi === 'function') return savePermissionRuleApi(payload ? payload.token : null, payload.rule);
      return { success: false, message: 'Hàm savePermissionRule chưa khai báo' };

    case 'savePermissionRulesBatch':
      if (typeof savePermissionRulesBatchApi === 'function') return savePermissionRulesBatchApi(payload ? payload.token : null, payload.rules);
      return { success: false, message: 'Hàm savePermissionRulesBatch chưa khai báo' };

    case 'deletePermissionRules':
      if (typeof deletePermissionRulesApi === 'function') {
        return deletePermissionRulesApi(
          payload ? payload.token : null,
          payload ? payload.targetValue : null,
          payload ? payload.deletions : null
        );
      }
      return { success: false, message: 'Hàm deletePermissionRules chưa khai báo' };

    case 'getPermissionRule':
      if (typeof getPermissionRuleApi === 'function') {
        return getPermissionRuleApi(payload ? payload.token : null, payload.targetType, payload.targetValue, payload.sheetName);
      }
      return { success: false, message: 'Hàm getPermissionRuleApi chưa khai báo' };

    case 'saveUserNavigationAccess':
      return typeof saveUserNavigationAccessApi === 'function'
        ? saveUserNavigationAccessApi(payload ? payload.token : null, payload.username, payload.viewIds)
        : { success: false, message: 'Hàm saveUserNavigationAccessApi chưa khai báo' };
    // ============================================================
    // BỔ SUNG ROUTE MỚI CHO MASTER-DETAIL USER & PHÂN QUYỀN TÍCH HỢP
    // ============================================================
    case 'fetchSysUsersPaginated':
      return typeof fetchSysUsersPaginated === 'function' 
        ? fetchSysUsersPaginated(payload.token, payload.page, payload.pageSize, payload.searchKey) 
        : { success: false, message: 'Hàm fetchSysUsersPaginated chưa khai báo' };

    case 'getDistinctColumnValues':
      return typeof getDistinctColumnValuesApi === 'function'
        ? getDistinctColumnValuesApi(payload.token, payload.sheetName, payload.columnIndex)
        : { success: false, message: 'getDistinctColumnValuesApi chưa khai báo' };
      return typeof getUserPermissionsDetailApi === 'function' 
        ? getUserPermissionsDetailApi(payload.token, payload.targetUsername) 
        : { success: false, message: 'Hàm getUserPermissionsDetailApi chưa khai báo' };

    case 'saveUserFullPermissions':
      return typeof saveUserFullPermissionsApi === 'function' 
        ? saveUserFullPermissionsApi(payload.token, payload) 
        : { success: false, message: 'Hàm saveUserFullPermissionsApi chưa khai báo' };
    
    default:
      return { success: false, message: 'Action SYS không hợp lệ: ' + action };
  }
}

function routeHrmModule(action, payload) {
  switch (action) {
    case 'fetchHRMEMPListPaginated':
      return typeof fetchHRMEMPListPaginated === 'function'
        ? fetchHRMEMPListPaginated(payload.token, payload.page, payload.pageSize || payload.limit)
        : { success: false, message: 'Hàm fetchHRMEMPListPaginated chưa sẵn sàng' };
    
    // BỔ SUNG ROUTE LẤY DANH SÁCH HRM_ID ĐÃ QUA BỘ LỌC PHÂN QUYỀN
    case 'fetchHrmIdList':
      if (typeof fetchHrmIdListApi === 'function') {
        return fetchHrmIdListApi(payload);
      } else if (typeof fetchHrmIdList === 'function') {
        return fetchHrmIdList(payload);
      } else {
        return { success: false, message: 'Hàm fetchHrmIdListApi chưa khai báo trong HRM_Server' };
      }

    default:
      return { success: false, message: 'Action HRM không hợp lệ: ' + action };
  }
}

function routeDmsModule(action, payload) {
  switch (action) {
    case 'listDocs': return typeof dmsListDocs === 'function' ? dmsListDocs(payload.token, payload.page, payload.pageSize) : { success: false, message: 'dmsListDocs chưa khai báo' };
    case 'saveDoc': return typeof dmsSaveDoc === 'function' ? dmsSaveDoc(payload.token, payload.record) : { success: false, message: 'dmsSaveDoc chưa khai báo' };
    case 'uploadFile': return typeof dmsUploadFile === 'function' ? dmsUploadFile(payload.token, payload.name, payload.mimeType, payload.base64, payload.folderId) : { success: false, message: 'dmsUploadFile chưa khai báo' };
    case 'listEmployees': return typeof dmsListEmployees === 'function' ? dmsListEmployees(payload.token) : { success: false, message: 'dmsListEmployees chưa khai báo' };
    case 'deleteDoc': return typeof dmsDeleteDoc === 'function' ? dmsDeleteDoc(payload.token, payload.docId) : { success: false, message: 'dmsDeleteDoc chưa khai báo' };
    case 'approveDoc': return typeof dmsApproveDoc === 'function' ? dmsApproveDoc(payload.token, payload.docId) : { success: false, message: 'dmsApproveDoc chưa khai báo' };
    case 'rejectDoc': return typeof dmsRejectDoc === 'function' ? dmsRejectDoc(payload.token, payload.docId, payload.reason) : { success: false, message: 'dmsRejectDoc chưa khai báo' };
    case 'listVersions': return typeof dmsListVersions === 'function' ? dmsListVersions(payload.token, payload.docId) : { success: false, message: 'dmsListVersions chưa khai báo' };
    case 'addVersion': return typeof dmsAddVersion === 'function' ? dmsAddVersion(payload.token, payload.record) : { success: false, message: 'dmsAddVersion chưa khai báo' };
    case 'deleteVersion': return typeof dmsDeleteVersion === 'function' ? dmsDeleteVersion(payload.token, payload.versionId) : { success: false, message: 'dmsDeleteVersion chưa khai báo' };
    case 'listComments': return typeof dmsListComments === 'function' ? dmsListComments(payload.token, payload.docId) : { success: false, message: 'dmsListComments chưa khai báo' };
    case 'addComment': return typeof dmsAddComment === 'function' ? dmsAddComment(payload.token, payload.record) : { success: false, message: 'dmsAddComment chưa khai báo' };
    case 'ackRead': return typeof dmsAcknowledgeRead === 'function' ? dmsAcknowledgeRead(payload.token, payload.docId, payload.versionTag, payload.username) : { success: false, message: 'dmsAcknowledgeRead chưa khai báo' };
    case 'listReadStatus': return typeof dmsListReadStatus === 'function' ? dmsListReadStatus(payload.token, payload.docId) : { success: false, message: 'dmsListReadStatus chưa khai báo' };
    default: return { success: false, message: 'Action DMS không hợp lệ: ' + action };
  }
}
function routeWmsModule(action, payload) {
  switch (action) {
    case 'listWarehouses':
      return typeof wmsListWarehouses === 'function'
        ? wmsListWarehouses(payload.token, payload.searchKey)
        : { success: false, message: 'Hàm wmsListWarehouses chưa khai báo' };

    case 'saveWarehouse':
      return typeof wmsSaveWarehouse === 'function'
        ? wmsSaveWarehouse(payload.token, payload.record)
        : { success: false, message: 'Hàm wmsSaveWarehouse chưa khai báo' };

    case 'deleteWarehouse':
      return typeof wmsDeleteWarehouse === 'function'
        ? wmsDeleteWarehouse(payload.token, payload.warehouseId, payload.expectedUpdatedAt)
        : { success: false, message: 'Hàm wmsDeleteWarehouse chưa khai báo' };

    case 'listLocations':
      return typeof wmsListLocations === 'function'
        ? wmsListLocations(payload.token, payload.warehouseId)
        : { success: false, message: 'Hàm wmsListLocations chưa khai báo' };

    case 'saveLocation':
      return typeof wmsSaveLocation === 'function'
        ? wmsSaveLocation(payload.token, payload.record)
        : { success: false, message: 'Hàm wmsSaveLocation chưa khai báo' };

    case 'deleteLocation':
      return typeof wmsDeleteLocation === 'function'
        ? wmsDeleteLocation(payload.token, payload.locationId, payload.expectedUpdatedAt)
        : { success: false, message: 'Hàm wmsDeleteLocation chưa khai báo' };

    case 'listSkus':
      return typeof wmsListSkus === 'function'
        ? wmsListSkus(payload.token, payload.searchKey)
        : { success: false, message: 'Hàm wmsListSkus chưa khai báo' };

    case 'suggestSkuCode':
      return typeof wmsSuggestSkuCode === 'function'
        ? wmsSuggestSkuCode(payload.token, payload.categoryCode)
        : { success: false, message: 'Hàm wmsSuggestSkuCode chưa khai báo' };

    case 'saveSku':
      return typeof wmsSaveSku === 'function'
        ? wmsSaveSku(payload.token, payload.record)
        : { success: false, message: 'Hàm wmsSaveSku chưa khai báo' };

    case 'importSkus':
      return typeof wmsImportSkus === 'function'
        ? wmsImportSkus(payload.token, payload.records)
        : { success: false, message: 'Hàm wmsImportSkus chưa khai báo' };

    case 'deleteSku':
      return typeof wmsDeleteSku === 'function'
        ? wmsDeleteSku(payload.token, payload.skuId, payload.expectedUpdatedAt)
        : { success: false, message: 'Hàm wmsDeleteSku chưa khai báo' };

    case 'deleteSkus':
      return typeof wmsDeleteSkus === 'function'
        ? wmsDeleteSkus(payload.token, payload.records)
        : { success: false, message: 'Hàm wmsDeleteSkus chưa khai báo' };

    case 'listDocOptions':
      return typeof wmsListDocOptions === 'function'
        ? wmsListDocOptions(payload.token, payload.docType)
        : { success: false, message: 'Hàm wmsListDocOptions chưa khai báo' };

    case 'listDocs':
      return typeof wmsListDocs === 'function'
        ? wmsListDocs(payload.token, payload.docType, payload.searchKey, payload.page, payload.pageSize, payload.includeOptions)
        : { success: false, message: 'Hàm wmsListDocs chưa khai báo' };

    case 'getDoc':
      return typeof wmsGetDoc === 'function'
        ? wmsGetDoc(payload.token, payload.docType, payload.docId)
        : { success: false, message: 'Hàm wmsGetDoc chưa khai báo' };

    case 'saveDoc':
      return typeof wmsSaveDoc === 'function'
        ? wmsSaveDoc(payload.token, payload.docType, payload.record)
        : { success: false, message: 'Hàm wmsSaveDoc chưa khai báo' };

    case 'deleteDocs':
      return typeof wmsDeleteDocs === 'function'
        ? wmsDeleteDocs(payload.token, payload.docType, payload.records)
        : { success: false, message: 'Hàm wmsDeleteDocs chưa khai báo' };

    case 'getWorkflowSetup':
      return typeof wmsWorkflowGetSetup === 'function'
        ? wmsWorkflowGetSetup(payload.token, payload.since)
        : { success:false, message:'Hàm wmsWorkflowGetSetup chưa khai báo' };

    case 'saveWorkflow':
      return typeof wmsWorkflowSave === 'function'
        ? wmsWorkflowSave(payload.token, payload.record)
        : { success:false, message:'Hàm wmsWorkflowSave chưa khai báo' };

    case 'activateWorkflow':
      return typeof wmsWorkflowActivate === 'function'
        ? wmsWorkflowActivate(payload.token, payload.workflowVersionId)
        : { success:false, message:'Hàm wmsWorkflowActivate chưa khai báo' };

    case 'setWorkflowStatus':
      return typeof wmsWorkflowSetStatus === 'function'
        ? wmsWorkflowSetStatus(payload.token, payload.workflowVersionId, payload.action)
        : { success:false, message:'Hàm wmsWorkflowSetStatus chưa khai báo' };

    case 'saveWorkflowMapping':
      return typeof wmsWorkflowSaveMapping === 'function'
        ? wmsWorkflowSaveMapping(payload.token, payload.record)
        : { success:false, message:'Hàm wmsWorkflowSaveMapping chưa khai báo' };

    case 'setWorkflowMappingStatus':
      return typeof wmsWorkflowSetMappingStatus === 'function'
        ? wmsWorkflowSetMappingStatus(payload.token, payload.mappingId, payload.status)
        : { success:false, message:'Hàm wmsWorkflowSetMappingStatus chưa khai báo' };

    case 'listMyWorkflowTasks':
      return typeof wmsWorkflowMyTasks === 'function'
        ? wmsWorkflowMyTasks(payload.token)
        : { success:false, message:'Hàm wmsWorkflowMyTasks chưa khai báo' };

    case 'submitWorkflow':
      return typeof wmsWorkflowSubmit === 'function'
        ? wmsWorkflowSubmit(payload.token, payload.docId, payload.workflowVersionId)
        : { success:false, message:'Hàm wmsWorkflowSubmit chưa khai báo' };

    case 'decideWorkflowTask':
      return typeof wmsWorkflowDecide === 'function'
        ? wmsWorkflowDecide(payload.token, payload.taskId, payload.decision, payload.comment)
        : { success:false, message:'Hàm wmsWorkflowDecide chưa khai báo' };

    default:
      return { success: false, message: 'Action WMS không hợp lệ: ' + action };
  }
}