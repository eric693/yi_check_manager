// WorklogHandlers.gs - 工作日誌 API Handler 函數

// ==================== 員工功能 ====================

/**
 * ✅ 處理提交工作日誌
 */
function handleSubmitWorklog(params) {
  try {
    Logger.log('═══════════════════════════════════════');
    Logger.log('📝 handleSubmitWorklog 開始');
    Logger.log('═══════════════════════════════════════');
    
    // 驗證 Session
    if (!params.token) {
      Logger.log('❌ 缺少 token');
      return { ok: false, msg: "缺少認證 token" };
    }
    
    const session = checkSession_(params.token);
    
    if (!session.ok || !session.user) {
      Logger.log('❌ Session 無效');
      return { ok: false, msg: "未授權或 session 已過期" };
    }
    
    Logger.log('✅ Session 驗證成功');
    Logger.log('   使用者: ' + session.user.name);
    
    // 取得參數
    const userId = session.user.userId;
    const userName = session.user.name;
    const department = session.user.dept || '未分配部門';
    const date = params.date;
    const hours = params.hours;
    const content = params.content;
    
    Logger.log('📥 收到的參數:');
    Logger.log('   日期: ' + date);
    Logger.log('   時數: ' + hours);
    Logger.log('   內容長度: ' + (content ? content.length : 0));
    
    // 呼叫核心函數
    const result = submitWorklog(userId, userName, department, date, hours, content);
    
    Logger.log('📤 處理結果: ' + result.success);
    Logger.log('═══════════════════════════════════════');
    
    return {
      ok: result.success,
      msg: result.message,
      worklogId: result.worklogId
    };
    
  } catch (error) {
    Logger.log('❌ handleSubmitWorklog 錯誤: ' + error);
    return { ok: false, msg: error.message };
  }
}

/**
 * ✅ 處理取得工作日誌列表
 */
function handleGetWorklogs(params) {
  try {
    // 驗證 Session
    if (!params.token) {
      return { ok: false, msg: "缺少認證 token" };
    }
    
    const session = checkSession_(params.token);
    
    if (!session.ok || !session.user) {
      return { ok: false, msg: "未授權或 session 已過期" };
    }
    
    const userId = session.user.userId;
    const limit = parseInt(params.limit) || 30;
    
    Logger.log('📋 查詢工作日誌: ' + session.user.name);
    
    const result = getWorklogs(userId, limit);
    
    return {
      ok: result.success,
      worklogs: result.worklogs,
      total: result.total,
      msg: result.message || '查詢成功'
    };
    
  } catch (error) {
    Logger.log('❌ handleGetWorklogs 錯誤: ' + error);
    return { ok: false, msg: error.message };
  }
}

/**
 * ✅ 處理取得工作日誌詳情
 */
function handleGetWorklogDetail(params) {
  try {
    // 驗證 Session
    if (!params.token) {
      return { ok: false, msg: "缺少認證 token" };
    }
    
    const session = checkSession_(params.token);
    
    if (!session.ok || !session.user) {
      return { ok: false, msg: "未授權或 session 已過期" };
    }
    
    if (!params.id) {
      return { ok: false, msg: "缺少工作日誌 ID" };
    }
    
    const result = getWorklogDetail(params.id);
    
    // 檢查權限（只能查看自己的工作日誌）
    if (result.success && result.worklog.userId !== session.user.userId && session.user.dept !== '管理員') {
      return { ok: false, msg: "沒有權限查看此工作日誌" };
    }
    
    return {
      ok: result.success,
      worklog: result.worklog,
      msg: result.message
    };
    
  } catch (error) {
    Logger.log('❌ handleGetWorklogDetail 錯誤: ' + error);
    return { ok: false, msg: error.message };
  }
}

// ==================== 管理員功能 ====================

/**
 * ✅ 處理取得待審核工作日誌
 */
function handleGetPendingWorklogs(params) {
  try {
    Logger.log('📋 handleGetPendingWorklogs 開始');
    
    // 驗證 Session
    if (!params.token) {
      return { ok: false, msg: "缺少認證 token" };
    }
    
    const session = checkSession_(params.token);
    
    if (!session.ok || !session.user) {
      return { ok: false, msg: "未授權或 session 已過期" };
    }
    
    // 驗證管理員權限
    if (session.user.dept !== '管理員') {
      Logger.log('❌ 非管理員嘗試存取');
      return { ok: false, msg: "需要管理員權限" };
    }
    
    Logger.log('✅ 管理員權限驗證通過: ' + session.user.name);
    
    const result = getPendingWorklogs();
    
    Logger.log('📤 找到 ' + (result.worklogs ? result.worklogs.length : 0) + ' 筆待審核工作日誌');
    
    return {
      ok: result.success,
      worklogs: result.worklogs,
      total: result.total,
      msg: result.message || '查詢成功'
    };
    
  } catch (error) {
    Logger.log('❌ handleGetPendingWorklogs 錯誤: ' + error);
    return { ok: false, msg: error.message };
  }
}

/**
 * ✅ 處理審核工作日誌（修正版）
 */
function handleReviewWorklog(params) {
  try {
    Logger.log('═══════════════════════════════════════');
    Logger.log('📝 handleReviewWorklog 開始');
    Logger.log('═══════════════════════════════════════');
    
    // 驗證 Session
    if (!params.token) {
      Logger.log('❌ 缺少 token');
      return { ok: false, msg: "缺少認證 token" };
    }
    
    const session = checkSession_(params.token);
    
    if (!session.ok || !session.user) {
      Logger.log('❌ Session 無效');
      return { ok: false, msg: "未授權或 session 已過期" };
    }
    
    // 驗證管理員權限
    if (session.user.dept !== '管理員') {
      Logger.log('❌ 非管理員嘗試審核');
      return { ok: false, msg: "需要管理員權限" };
    }
    
    Logger.log('✅ 管理員權限驗證通過: ' + session.user.name);
    
    // ⭐ 修改：改用新的參數名稱
    const worklogId = params.worklogId || params.id;  // 支援兩種參數名
    const action = params.reviewAction || params.action;  // 支援兩種參數名
    const comment = params.reviewComment || params.comment || '';  // 支援兩種參數名
    
    Logger.log('📥 收到的參數:');
    Logger.log('   工作日誌ID: ' + worklogId);
    Logger.log('   審核動作: ' + action);
    Logger.log('   審核意見: ' + comment);
    
    if (!worklogId) {
      Logger.log('❌ 缺少工作日誌 ID');
      return { ok: false, msg: "缺少工作日誌 ID" };
    }
    
    if (!action || (action !== 'approve' && action !== 'reject')) {
      Logger.log('❌ 無效的審核動作: ' + action);
      Logger.log('   params 內容: ' + JSON.stringify(params));
      return { ok: false, msg: "無效的審核動作: " + action };
    }
    
    // 呼叫核心函數
    const result = reviewWorklog(
      worklogId,
      action,
      session.user.userId,
      session.user.name,
      comment
    );
    
    Logger.log('📤 審核結果: ' + result.success);
    Logger.log('═══════════════════════════════════════');
    
    return {
      ok: result.success,
      msg: result.message
    };
    
  } catch (error) {
    Logger.log('❌ handleReviewWorklog 錯誤: ' + error);
    return { ok: false, msg: error.message };
  }
}
/**
 * ✅ 處理取得工作日誌報表
 */
function handleGetWorklogReport(params) {
  try {
    Logger.log('📊 handleGetWorklogReport 開始');
    
    // 驗證 Session
    if (!params.token) {
      return { ok: false, msg: "缺少認證 token" };
    }
    
    const session = checkSession_(params.token);
    
    if (!session.ok || !session.user) {
      return { ok: false, msg: "未授權或 session 已過期" };
    }
    
    // 驗證管理員權限
    if (session.user.dept !== '管理員') {
      return { ok: false, msg: "需要管理員權限" };
    }
    
    const employeeId = params.employeeId;
    const yearMonth = params.yearMonth;
    
    if (!employeeId || !yearMonth) {
      return { ok: false, msg: "缺少必要參數" };
    }
    
    Logger.log('   員工ID: ' + employeeId);
    Logger.log('   年月: ' + yearMonth);
    
    const result = getWorklogReport(employeeId, yearMonth);
    
    return {
      ok: result.success,
      worklogs: result.worklogs,
      summary: result.summary,
      msg: result.message || '查詢成功'
    };
    
  } catch (error) {
    Logger.log('❌ handleGetWorklogReport 錯誤: ' + error);
    return { ok: false, msg: error.message };
  }
}

// ==================== 測試函數 ====================

/**
 * 🧪 測試提交工作日誌 API
 */
function testHandleSubmitWorklog() {
  Logger.log('🧪 測試 handleSubmitWorklog');
  
  const testParams = {
    token: '你的有效token',  // ⚠️ 替換成有效的 token
    date: '2026-01-16',
    hours: '8.5',
    content: '測試工作日誌內容：完成系統開發、修復 bug、參與會議討論。'
  };
  
  const result = handleSubmitWorklog(testParams);
  Logger.log('結果: ' + JSON.stringify(result, null, 2));
}

/**
 * 🧪 測試查詢工作日誌 API
 */
function testHandleGetWorklogs() {
  Logger.log('🧪 測試 handleGetWorklogs');
  
  const testParams = {
    token: '71cff111-bdd2-4c44-ae34-ba86265d1c78',
    limit: 10
  };
  
  const result = handleGetWorklogs(testParams);
  Logger.log('結果: ' + JSON.stringify(result, null, 2));
}

/**
 * 🧪 測試審核工作日誌 API
 */
function testHandleReviewWorklog() {
  Logger.log('🧪 測試 handleReviewWorklog');
  
  const testParams = {
    token: '71cff111-bdd2-4c44-ae34-ba86265d1c78',  // ⚠️ 需要管理員權限
    id: 'WL_1234567890',  // ⚠️ 替換成實際的工作日誌 ID
    action: 'approve',
    comment: '工作內容詳實，核准通過'
  };
  
  const result = handleReviewWorklog(testParams);
  Logger.log('結果: ' + JSON.stringify(result, null, 2));
}

/**
 * ✅ 處理取得全部員工工作日誌報表
 */
function handleGetAllWorklogReport(params) {
  try {
    Logger.log('📊 handleGetAllWorklogReport 開始');
    
    // 驗證 Session
    if (!params.token) {
      return { ok: false, msg: "缺少認證 token" };
    }
    
    const session = checkSession_(params.token);
    
    if (!session.ok || !session.user) {
      return { ok: false, msg: "未授權或 session 已過期" };
    }
    
    // 驗證管理員權限
    if (session.user.dept !== '管理員') {
      return { ok: false, msg: "需要管理員權限" };
    }
    
    const yearMonth = params.yearMonth;
    
    if (!yearMonth) {
      return { ok: false, msg: "缺少年月參數" };
    }
    
    Logger.log('   年月: ' + yearMonth);
    
    const result = getAllWorklogReport(yearMonth);
    
    return {
      ok: result.success,
      worklogs: result.worklogs,
      total: result.total,
      msg: result.message || '查詢成功'
    };
    
  } catch (error) {
    Logger.log('❌ handleGetAllWorklogReport 錯誤: ' + error);
    return { ok: false, msg: error.message };
  }
}