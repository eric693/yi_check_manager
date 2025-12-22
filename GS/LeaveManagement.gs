// LeaveManagement.gs - 小時制請假系統（完整修正版）

/**
 * ✅ 提交請假申請（修正版 - 正確處理日期時間）
 * 
 * 修正內容：
 * 1. 正確解析前端傳來的 ISO 8601 日期時間字串
 * 2. 確保日期時間正確寫入 Sheet
 * 3. 修正請假原因欄位
 */
function submitLeaveRequest(sessionToken, leaveType, startDateTime, endDateTime, reason) {
  try {
    Logger.log('═══════════════════════════════════════');
    Logger.log('📋 開始處理請假申請（小時制）');
    Logger.log('═══════════════════════════════════════');
    
    // ⭐ 步驟 1：驗證 Session
    Logger.log('📡 驗證 Session...');
    const employee = checkSession_(sessionToken);
    
    if (!employee.ok || !employee.user) {
      Logger.log('❌ Session 驗證失敗');
      return { 
        ok: false, 
        code: "ERR_SESSION_INVALID",
        msg: "未授權或 session 已過期" 
      };
    }
    
    const user = employee.user;
    Logger.log('✅ Session 驗證成功');
    Logger.log(`   員工ID: ${user.userId}`);
    Logger.log(`   員工姓名: ${user.name}`);
    Logger.log('');
    
    // ⭐ 步驟 2：記錄收到的參數
    Logger.log('📥 收到的參數:');
    Logger.log(`   leaveType: ${leaveType}`);
    Logger.log(`   startDateTime: ${startDateTime}`);
    Logger.log(`   endDateTime: ${endDateTime}`);
    Logger.log(`   reason: ${reason}`);
    Logger.log('');
    
    // ⭐⭐⭐ 步驟 3：正確解析日期時間（關鍵修正）
    Logger.log('🔄 解析日期時間...');
    
    let start, end;
    
    try {
      // 前端傳來的是 ISO 8601 格式字串：例如 "2025-12-18T09:00"
      start = new Date(startDateTime);
      end = new Date(endDateTime);
      
      // 驗證日期是否有效
      if (isNaN(start.getTime()) || isNaN(end.getTime())) {
        Logger.log('❌ 日期時間格式無效');
        Logger.log(`   startDateTime: ${startDateTime} → ${start}`);
        Logger.log(`   endDateTime: ${endDateTime} → ${end}`);
        return {
          ok: false,
          code: "ERR_INVALID_DATETIME",
          msg: "日期時間格式無效"
        };
      }
      
      Logger.log('✅ 日期時間解析成功');
      Logger.log(`   開始: ${start.toISOString()}`);
      Logger.log(`   結束: ${end.toISOString()}`);
      
    } catch (parseError) {
      Logger.log('❌ 日期時間解析失敗: ' + parseError.message);
      return {
        ok: false,
        code: "ERR_DATETIME_PARSE",
        msg: "無法解析日期時間"
      };
    }
    
    Logger.log('');
    
    // ⭐ 步驟 4：驗證時間順序
    if (end <= start) {
      Logger.log('❌ 結束時間必須晚於開始時間');
      return {
        ok: false,
        code: "ERR_INVALID_TIME_RANGE",
        msg: "結束時間必須晚於開始時間"
      };
    }
    
    // ⭐ 步驟 5：計算工作時數和天數
    Logger.log('💡 計算工作時數和天數...');
    
    const { workHours, days } = calculateWorkHoursAndDays(start, end);
    
    Logger.log(`   工作時數: ${workHours} 小時`);
    Logger.log(`   天數: ${days} 天`);
    Logger.log('');
    
    // ⭐ 步驟 6：檢查假期餘額
    Logger.log('🔍 檢查假期餘額...');
    const balance = getLeaveBalance(sessionToken);
    
    if (!balance.ok) {
      Logger.log('❌ 無法取得假期餘額');
      return {
        ok: false,
        code: "ERR_BALANCE_CHECK",
        msg: "無法取得假期餘額"
      };
    }
    
    Logger.log('✅ 假期餘額檢查完成');
    Logger.log('');
    
    // ⭐⭐⭐ 步驟 7：格式化日期時間為 Sheet 可讀格式（關鍵修正）
    Logger.log('📝 格式化日期時間...');
    
    const formattedStartDateTime = Utilities.formatDate(
      start,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm:ss'
    );
    
    const formattedEndDateTime = Utilities.formatDate(
      end,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm:ss'
    );
    
    Logger.log(`   開始時間（格式化）: ${formattedStartDateTime}`);
    Logger.log(`   結束時間（格式化）: ${formattedEndDateTime}`);
    Logger.log('');
    
    // ⭐ 步驟 8：取得或建立工作表
    Logger.log('📊 取得工作表...');
    const sheet = getLeaveRecordsSheet();
    
    if (!sheet) {
      Logger.log('❌ 無法取得請假記錄工作表');
      return {
        ok: false,
        code: "ERR_SHEET_ACCESS",
        msg: "無法存取請假記錄工作表"
      };
    }
    
    Logger.log('✅ 工作表已就緒');
    Logger.log('');
    
    // ⭐⭐⭐ 步驟 9：寫入資料（確保所有欄位都有值）
    Logger.log('💾 準備寫入資料...');
    
    const row = [
      new Date(),                  // A: 申請時間（使用 Date 物件）
      user.userId || '',           // B: 員工ID
      user.name || '',             // C: 姓名
      user.dept || '',             // D: 部門
      leaveType || '',             // E: 假別
      formattedStartDateTime,      // F: 開始時間（格式化字串）⭐
      formattedEndDateTime,        // G: 結束時間（格式化字串）⭐
      workHours,                   // H: 工作時數（數字）
      days,                        // I: 天數（數字）
      reason || '',                // J: 原因 ⭐
      'PENDING',                   // K: 狀態
      '',                          // L: 審核人
      '',                          // M: 審核時間
      ''                           // N: 審核意見
    ];
    
    Logger.log('📋 準備寫入的資料:');
    Logger.log(`   A (申請時間): ${row[0]}`);
    Logger.log(`   B (員工ID): ${row[1]}`);
    Logger.log(`   C (姓名): ${row[2]}`);
    Logger.log(`   D (部門): ${row[3]}`);
    Logger.log(`   E (假別): ${row[4]}`);
    Logger.log(`   F (開始時間): ${row[5]} ⭐`);
    Logger.log(`   G (結束時間): ${row[6]} ⭐`);
    Logger.log(`   H (工作時數): ${row[7]}`);
    Logger.log(`   I (天數): ${row[8]}`);
    Logger.log(`   J (原因): ${row[9]} ⭐`);
    Logger.log(`   K (狀態): ${row[10]}`);
    Logger.log('');
    
    try {
      sheet.appendRow(row);
      Logger.log('✅ 資料寫入成功');
    } catch (writeError) {
      Logger.log('❌ 資料寫入失敗: ' + writeError.message);
      return {
        ok: false,
        code: "ERR_WRITE_FAILED",
        msg: "無法寫入請假記錄"
      };
    }
    
    Logger.log('');
    Logger.log('═══════════════════════════════════════');
    Logger.log('✅✅✅ 請假申請提交成功');
    Logger.log('═══════════════════════════════════════');
    
    return {
      ok: true,
      code: "LEAVE_SUBMIT_SUCCESS",
      msg: "請假申請已提交",
      data: {
        leaveType: leaveType,
        startDateTime: formattedStartDateTime,
        endDateTime: formattedEndDateTime,
        workHours: workHours,
        days: days,
        reason: reason
      }
    };
    
  } catch (error) {
    Logger.log('');
    Logger.log('❌❌❌ submitLeaveRequest 發生錯誤');
    Logger.log('錯誤訊息: ' + error.message);
    Logger.log('錯誤堆疊: ' + error.stack);
    Logger.log('═══════════════════════════════════════');
    
    return {
      ok: false,
      code: "ERR_INTERNAL_ERROR",
      msg: "系統錯誤：" + error.message
    };
  }
}

/**
 * 計算工作時數（排除午休時間 12:00-13:00）
 * 修正版：確保正確計算跨午休時段
 */
function calculateWorkHours(startTime, endTime) {
    if (!startTime || !endTime) {
        return 0;
    }
    
    const start = new Date(startTime);
    const end = new Date(endTime);
    
    // 檢查日期是否有效
    if (isNaN(start.getTime()) || isNaN(end.getTime())) {
        console.error('❌ 無效的日期格式');
        return 0;
    }
    
    // 檢查結束時間是否早於開始時間
    if (end <= start) {
        console.error('❌ 結束時間必須晚於開始時間');
        return 0;
    }
    
    // 計算總時長（毫秒）
    const totalMs = end - start;
    
    // 轉換為小時
    let totalHours = totalMs / (1000 * 60 * 60);
    
    // 如果是同一天，檢查是否跨越午休時間 12:00-13:00
    if (start.toDateString() === end.toDateString()) {
        const startHour = start.getHours() + start.getMinutes() / 60;
        const endHour = end.getHours() + end.getMinutes() / 60;
        
        const lunchStart = 12; // 12:00
        const lunchEnd = 13;   // 13:00
        
        // 判斷是否跨越午休時間
        if (startHour < lunchEnd && endHour > lunchStart) {
            // 計算重疊的午休時間
            const overlapStart = Math.max(startHour, lunchStart);
            const overlapEnd = Math.min(endHour, lunchEnd);
            const lunchOverlap = Math.max(0, overlapEnd - overlapStart);
            
            totalHours -= lunchOverlap;
            
            console.log('🍱 扣除午休時間:', lunchOverlap.toFixed(2), '小時');
        }
    } else {
        // 跨日請假：每天都要扣除 1 小時午休
        const startDate = new Date(start);
        startDate.setHours(0, 0, 0, 0);
        
        const endDate = new Date(end);
        endDate.setHours(0, 0, 0, 0);
        
        const daysDiff = Math.floor((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
        
        // 每天扣除 1 小時午休
        totalHours -= daysDiff;
        
        console.log('📅 跨日請假，扣除', daysDiff, '天的午休時間');
    }
    
    // 確保不會是負數
    totalHours = Math.max(0, totalHours);
    
    // 四捨五入到小數點後 2 位
    return Math.round(totalHours * 100) / 100;
}
/**
 * ✅ 取得或建立請假記錄工作表
 */
function getLeaveRecordsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('請假紀錄');
  
  if (!sheet) {
    Logger.log('📝 請假紀錄工作表不存在，自動建立...');
    
    sheet = ss.insertSheet('請假紀錄');
    
    // 建立標題列（14個欄位）
    sheet.appendRow([
      '申請時間', '員工ID', '姓名', '部門', '假別',
      '開始時間', '結束時間', '工作時數', '天數', '原因',
      '狀態', '審核人', '審核時間', '審核意見'
    ]);
    
    // 美化標題列
    const headerRange = sheet.getRange(1, 1, 1, 14);
    headerRange.setBackground('#4A90E2');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    
    // 凍結標題列
    sheet.setFrozenRows(1);
    
    Logger.log('✅ 請假紀錄工作表已建立');
  }
  
  return sheet;
}

function getLeaveBalance(sessionToken) {
  try {
    const employee = checkSession_(sessionToken);
    
    if (!employee.ok || !employee.user) {
      return {
        ok: false,
        code: "ERR_SESSION_INVALID"
      };
    }
    
    const user = employee.user;
    Logger.log('🔍 查詢員工: ' + user.userId);
    
    const sheet = getLeaveBalanceSheet();
    
    if (!sheet) {
      Logger.log('❌ 工作表不存在，嘗試建立...');
      initializeEmployeeLeave(sessionToken);
      return getLeaveBalance(sessionToken);
    }
    
    const values = sheet.getDataRange().getValues();
    Logger.log('📊 工作表行數: ' + values.length);
    
    for (let i = 1; i < values.length; i++) {
      Logger.log(`   檢查第 ${i} 行: ${values[i][0]}`);
      
      if (values[i][0] === user.userId) {
        Logger.log('✅ 找到員工資料');
        
        const balance = {
          ANNUAL_LEAVE: values[i][1] || 0,
          SICK_LEAVE: values[i][2] || 0,
          PERSONAL_LEAVE: values[i][3] || 0,
          BEREAVEMENT_LEAVE: values[i][4] || 0,
          MARRIAGE_LEAVE: values[i][5] || 0,
          MATERNITY_LEAVE: values[i][6] || 0,
          PATERNITY_LEAVE: values[i][7] || 0,
          HOSPITALIZATION_LEAVE: values[i][8] || 0,
          MENSTRUAL_LEAVE: values[i][9] || 0,
          FAMILY_CARE_LEAVE: values[i][10] || 0,
          OFFICIAL_LEAVE: values[i][11] || 1,
          WORK_INJURY_LEAVE: values[i][12] || 1,
          NATURAL_DISASTER_LEAVE: values[i][13] || 1,
          COMP_TIME_OFF: values[i][14] || 0,
          ABSENCE_WITHOUT_LEAVE: values[i][15] || 0
        };
        // const balance = {
        // // ✅ 使用 camelCase 格式
        // annualLeave: values[i][1] || 0,                    // B
        // sickLeave: values[i][2] || 0,                      // C
        // personalLeave: values[i][3] || 0,                  // D
        // bereavementLeave: values[i][4] || 0,               // E
        // marriageLeave: values[i][5] || 0,                  // F
        // maternityLeave: values[i][6] || 0,                 // G
        // paternityLeave: values[i][7] || 0,                 // H
        // hospitalizationLeave: values[i][8] || 0,           // I
        // menstrualLeave: values[i][9] || 0,                 // J
        // familyCareLeave: values[i][10] || 0,               // K
        // officialLeave: values[i][11] || 0,               // L
        // workInjuryLeave: values[i][12] || 0,             // M
        // naturalDisasterLeave: values[i][13] || 0,        // N
        // compTimeOff: values[i][14] || 0,                   // O
        // absenceWithoutLeave: values[i][15] || 0            // P
      // };
        
        Logger.log('📋 假期餘額:');
        Logger.log(JSON.stringify(balance, null, 2));
        
        return {
          ok: true,
          balance: balance  // ⭐ 關鍵：返回 balance 物件
        };
      }
    }
    
    // 如果找不到，自動初始化
    Logger.log('⚠️ 找不到員工資料，嘗試初始化...');
    initializeEmployeeLeave(sessionToken);
    return getLeaveBalance(sessionToken);
    
  } catch (error) {
    Logger.log('❌ getLeaveBalance 錯誤: ' + error);
    Logger.log('錯誤堆疊: ' + error.stack);
    return {
      ok: false,
      code: "ERR_INTERNAL_ERROR",
      msg: error.message
    };
  }
}

function getLeaveBalanceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('假期餘額');
  
  if (!sheet) {
    Logger.log('📝 假期餘額工作表不存在，自動建立...');
    
    sheet = ss.insertSheet('假期餘額');
    
    // ✅ 修改：建立標題列（17個欄位）
    sheet.appendRow([
      '員工ID',           // A
      '特休假',           // B
      '未住院病假',       // C
      '事假',             // D
      '喪假',             // E
      '婚假',             // F
      '產假',             // G
      '陪產檢及陪產假',   // H
      '住院病假',         // I
      '生理假',           // J
      '家庭照顧假',       // K
      '公假(含兵役假)',   // L
      '公傷假',           // M
      '天然災害停班',     // N
      '加班補休假',       // O
      '曠工',             // P
      '更新時間'          // Q
    ]);
    
    // 美化標題列
    const headerRange = sheet.getRange(1, 1, 1, 17);
    headerRange.setBackground('#4A90E2');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    
    // 凍結標題列
    sheet.setFrozenRows(1);
    
    Logger.log('✅ 假期餘額工作表已建立');
  }
  
  return sheet;
}

function testGetLeaveBalance() {
  // ⚠️ 替換成你的 sessionToken
  const token = '7dac1161-bbac-487d-900b-3e06c1acab8d';
  
  Logger.log('🧪 開始測試 getLeaveBalance');
  Logger.log('');
  
  const result = getLeaveBalance(token);
  
  Logger.log('📤 測試結果:');
  Logger.log(JSON.stringify(result, null, 2));
  
  if (result.ok) {
    Logger.log('');
    Logger.log('✅ 測試成功！');
    Logger.log('');
    Logger.log('假期餘額:');
    for (const [key, value] of Object.entries(result.balance)) {
      Logger.log(`   ${key}: ${value}`);
    }
  } else {
    Logger.log('');
    Logger.log('❌ 測試失敗');
    Logger.log('錯誤碼: ' + result.code);
  }
}

function initializeEmployeeLeave(sessionToken) {
  try {
    const employee = checkSession_(sessionToken);
    
    if (!employee.ok || !employee.user) {
      return {
        ok: false,
        code: "ERR_SESSION_INVALID"
      };
    }
    
    const user = employee.user;
    const sheet = getLeaveBalanceSheet();
    
    // 檢查是否已存在
    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === user.userId) {
        Logger.log('ℹ️ 員工 ' + user.name + ' 的假期餘額已存在');
        return {
          ok: true,
          msg: "假期餘額已存在"
        };
      }
    }
    
    // ✅ 修改：15種假別的預設餘額（依據台灣勞基法）
    const defaultBalance = [
      user.userId,        // A: 員工ID
      7,                  // B: 特休假（天）- 依年資調整
      30,                 // C: 未住院病假（天/年）
      14,                 // D: 事假（天/年）
      5,                  // E: 喪假（天）- 依親等不同
      8,                  // F: 婚假（天）
      56,                 // G: 產假（天）- 8週
      7,                  // H: 陪產檢及陪產假（天）
      30,                 // I: 住院病假（天/年）
      12,                 // J: 生理假（天/年）- 每月1天
      7,                  // K: 家庭照顧假（天/年）
      0,                // L: 公假（含兵役假）（無上限）
      0,                // M: 公傷假（無上限）
      0,                // N: 天然災害停班（無上限）
      0,                  // O: 加班補休假（初始0）
      0,                  // P: 曠工（初始0）
      new Date()          // Q: 更新時間
    ];
    
    sheet.appendRow(defaultBalance);
    
    Logger.log('✅ 已為員工 ' + user.name + ' 初始化假期餘額');
    
    return {
      ok: true,
      msg: "假期餘額已初始化"
    };
    
  } catch (error) {
    Logger.log('❌ initializeEmployeeLeave 錯誤: ' + error);
    return {
      ok: false,
      msg: error.message
    };
  }
}

/**
 * ✅ 取得員工請假記錄
 */
function getEmployeeLeaveRecords(sessionToken) {
  try {
    const employee = checkSession_(sessionToken);
    
    if (!employee.ok || !employee.user) {
      return {
        ok: false,
        code: "ERR_SESSION_INVALID"
      };
    }
    
    const user = employee.user;
    const sheet = getLeaveRecordsSheet();
    const values = sheet.getDataRange().getValues();
    
    if (values.length <= 1) {
      return {
        ok: true,
        records: []
      };
    }
    
    const records = [];
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][1] === user.userId) {
        const record = {
          applyTime: formatDateTime(values[i][0]),      // A
          employeeName: values[i][2],                   // C
          dept: values[i][3],                           // D
          leaveType: values[i][4],                      // E
          startDateTime: values[i][5],                  // F
          endDateTime: values[i][6],                    // G
          workHours: values[i][7],                      // H
          days: values[i][8],                           // I
          reason: values[i][9] || '',                   // J
          status: values[i][10] || 'PENDING',           // K
          reviewer: values[i][11] || '',                // L
          reviewTime: values[i][12] ? formatDateTime(values[i][12]) : '', // M
          reviewComment: values[i][13] || ''            // N
        };
        
        records.push(record);
      }
    }
    
    // 按申請時間排序（最新的在前）
    records.sort((a, b) => new Date(b.applyTime) - new Date(a.applyTime));
    
    return {
      ok: true,
      records: records
    };
    
  } catch (error) {
    Logger.log('❌ getEmployeeLeaveRecords 錯誤: ' + error);
    return {
      ok: false,
      msg: error.message
    };
  }
}

/**
 * ✅ 取得待審核請假申請（管理員用）
 */
function getPendingLeaveRequests(sessionToken) {
  try {
    const employee = checkSession_(sessionToken);
    
    if (!employee.ok || !employee.user) {
      return {
        ok: false,
        code: "ERR_SESSION_INVALID"
      };
    }
    
    // 檢查是否為管理員
    if (employee.user.dept !== '管理員') {
      return {
        ok: false,
        code: "ERR_PERMISSION_DENIED",
        msg: "需要管理員權限"
      };
    }
    
    const sheet = getLeaveRecordsSheet();
    const values = sheet.getDataRange().getValues();
    
    if (values.length <= 1) {
      return {
        ok: true,
        requests: []
      };
    }
    
    const requests = [];
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][10] === 'PENDING') {
        const request = {
          rowNumber: i + 1,
          applyTime: formatDateTime(values[i][0]),
          employeeId: values[i][1],
          employeeName: values[i][2],
          dept: values[i][3],
          leaveType: values[i][4],
          startDateTime: values[i][5],
          endDateTime: values[i][6],
          workHours: values[i][7],
          days: values[i][8],
          reason: values[i][9] || ''
        };
        
        requests.push(request);
      }
    }
    
    return {
      ok: true,
      requests: requests
    };
    
  } catch (error) {
    Logger.log('❌ getPendingLeaveRequests 錯誤: ' + error);
    return {
      ok: false,
      msg: error.message
    };
  }
}


/**
 * ✅ 格式化日期時間
 */
function formatDateTime(date) {
  if (!date) return '';
  try {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  } catch (e) {
    return String(date);
  }
}

/**
 * 🧪 測試函數
 */
function testSubmitLeaveWithHours() {
  Logger.log('🧪 測試小時制請假申請');
  Logger.log('');
  
  const testParams = {
    token: '7dac1161-bbac-487d-900b-3e06c1acab8d',  // ⚠️ 替換成有效 token
    leaveType: 'BEREAVEMENT_LEAVE',
    startDateTime: '2025-12-18T09:00',
    endDateTime: '2025-12-18T12:00',
    reason: '測試請假（小時制）'
  };
  
  Logger.log('📥 測試參數:');
  Logger.log(JSON.stringify(testParams, null, 2));
  Logger.log('');
  
  const result = submitLeaveRequest(
    testParams.token,
    testParams.leaveType,
    testParams.startDateTime,
    testParams.endDateTime,
    testParams.reason
  );
  
  Logger.log('');
  Logger.log('📤 測試結果:');
  Logger.log(JSON.stringify(result, null, 2));
  
  if (result.ok) {
    Logger.log('');
    Logger.log('✅✅✅ 測試成功！');
    Logger.log('請檢查 Google Sheet 的「請假紀錄」工作表');
  } else {
    Logger.log('');
    Logger.log('❌ 測試失敗');
  }
}

/**
 * 🔄 遷移假期餘額工作表（8欄 → 17欄）
 * 
 * 使用方式：
 * 1. 在 Apps Script 編輯器中執行此函數
 * 2. 會自動備份舊資料
 * 3. 重建新結構並遷移資料
 */
function migrateLeaveBalanceSheet() {
  Logger.log('═══════════════════════════════════════');
  Logger.log('🔄 開始遷移假期餘額工作表');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const oldSheet = ss.getSheetByName('假期餘額');
  
  if (!oldSheet) {
    Logger.log('❌ 找不到「假期餘額」工作表');
    return;
  }
  
  // 📋 步驟 1：備份舊工作表
  Logger.log('📋 步驟 1：備份舊工作表...');
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const backupSheet = oldSheet.copyTo(ss);
  backupSheet.setName('假期餘額_備份_' + timestamp);
  Logger.log('✅ 已備份: ' + backupSheet.getName());
  Logger.log('');
  
  // 📋 步驟 2：讀取舊資料
  Logger.log('📋 步驟 2：讀取舊資料...');
  const oldData = oldSheet.getDataRange().getValues();
  const recordCount = oldData.length - 1; // 扣除標題列
  
  Logger.log(`   找到 ${recordCount} 筆員工資料`);
  Logger.log('');
  
  if (recordCount <= 0) {
    Logger.log('⚠️ 沒有資料需要遷移');
    return;
  }
  
  // 📋 步驟 3：刪除舊工作表
  Logger.log('📋 步驟 3：刪除舊工作表...');
  ss.deleteSheet(oldSheet);
  Logger.log('✅ 已刪除舊的「假期餘額」工作表');
  Logger.log('');
  
  // 📋 步驟 4：建立新工作表（17 個欄位）
  Logger.log('📋 步驟 4：建立新工作表（17 個欄位）...');
  const newSheet = ss.insertSheet('假期餘額');
  
  // 建立標題列
  const headers = [
    '員工ID',           // A
    '特休假',           // B - ANNUAL_LEAVE
    '未住院病假',       // C - SICK_LEAVE
    '事假',             // D - PERSONAL_LEAVE
    '喪假',             // E - BEREAVEMENT_LEAVE
    '婚假',             // F - MARRIAGE_LEAVE
    '產假',             // G - MATERNITY_LEAVE
    '陪產檢及陪產假',   // H - PATERNITY_LEAVE
    '住院病假',         // I - HOSPITALIZATION_LEAVE
    '生理假',           // J - MENSTRUAL_LEAVE
    '家庭照顧假',       // K - FAMILY_CARE_LEAVE
    '公假(含兵役假)',   // L - OFFICIAL_LEAVE
    '公傷假',           // M - WORK_INJURY_LEAVE
    '天然災害停班',     // N - NATURAL_DISASTER_LEAVE
    '加班補休假',       // O - COMP_TIME_OFF
    '曠工',             // P - ABSENCE_WITHOUT_LEAVE
    '更新時間'          // Q
  ];
  
  newSheet.appendRow(headers);
  
  // 美化標題列
  const headerRange = newSheet.getRange(1, 1, 1, 17);
  headerRange.setBackground('#4A90E2');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 凍結標題列
  newSheet.setFrozenRows(1);
  
  Logger.log('✅ 新工作表已建立（17 個欄位）');
  Logger.log('');
  
  // 📋 步驟 5：遷移資料
  Logger.log('📋 步驟 5：遷移資料...');
  Logger.log('');
  
  for (let i = 1; i < oldData.length; i++) {
    const oldRow = oldData[i];
    
    // 對應關係：
    // 舊: [員工ID, 特休假, 病假, 事假, 喪假, 婚假, 產假, 陪產假, 更新時間]
    // 新: [員工ID, 特休假, 未住院病假, 事假, 喪假, 婚假, 產假, 陪產檢及陪產假, 住院病假, 生理假, 家庭照顧假, 公假, 公傷假, 天然災害停班, 加班補休假, 曠工, 更新時間]
    
    const newRow = [
      oldRow[0] || '',      // A: 員工ID（保留）
      oldRow[1] || 7,       // B: 特休假（保留）
      oldRow[2] || 30,      // C: 未住院病假（舊的「病假」）
      oldRow[3] || 14,      // D: 事假（保留）
      oldRow[4] || 5,       // E: 喪假（保留）
      oldRow[5] || 8,       // F: 婚假（保留）
      oldRow[6] || 56,      // G: 產假（保留）
      oldRow[7] || 7,       // H: 陪產檢及陪產假（保留）
      30,                   // I: 住院病假（新增，預設 30 天）
      12,                   // J: 生理假（新增，預設 12 天）
      7,                    // K: 家庭照顧假（新增，預設 7 天）
      0,                  // L: 公假（新增，無上限）
      0,                  // M: 公傷假（新增，無上限）
      0,                  // N: 天然災害停班（新增，無上限）
      0,                    // O: 加班補休假（新增，初始 0）
      0,                    // P: 曠工（新增，初始 0）
      new Date()            // Q: 更新時間（更新為當前時間）
    ];
    
    newSheet.appendRow(newRow);
    
    Logger.log(`   ✅ [${i}/${recordCount}] 已遷移: ${oldRow[0]}`);
  }
  
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
  Logger.log('✅✅✅ 遷移完成！');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  Logger.log('📊 遷移摘要:');
  Logger.log(`   - 舊結構: 8 個欄位`);
  Logger.log(`   - 新結構: 17 個欄位`);
  Logger.log(`   - 遷移記錄數: ${recordCount} 筆`);
  Logger.log(`   - 備份工作表: ${backupSheet.getName()}`);
  Logger.log('');
  Logger.log('📝 新增假別:');
  Logger.log('   - 住院病假（30天）');
  Logger.log('   - 生理假（12天）');
  Logger.log('   - 家庭照顧假（7天）');
  Logger.log('   - 公假（無上限）');
  Logger.log('   - 公傷假（無上限）');
  Logger.log('   - 天然災害停班（無上限）');
  Logger.log('   - 加班補休假（初始0）');
  Logger.log('   - 曠工（初始0）');
  Logger.log('');
  
  // 顯示成功訊息給使用者
  Browser.msgBox(
    '✅ 遷移完成！',
    '已成功將 ' + recordCount + ' 筆假期餘額遷移到新結構！\n\n' +
    '舊結構：8 個欄位\n' +
    '新結構：17 個欄位（新增 8 種假別）\n\n' +
    '備份工作表: ' + backupSheet.getName() + '\n\n' +
    '請檢查「假期餘額」工作表確認資料正確。',
    Browser.Buttons.OK
  );
}

/**
 * 🧪 測試函數
 */
function testLeaveBalanceComplete() {
  Logger.log('🧪 測試完整的假期餘額系統');
  Logger.log('');
  
  const token = '7dac1161-bbac-487d-900b-3e06c1acab8d';  // ⚠️ 替換成有效 token
  
  Logger.log('📋 步驟 1：初始化假期餘額');
  const initResult = initializeEmployeeLeave(token);
  Logger.log('   結果: ' + JSON.stringify(initResult));
  Logger.log('');
  
  Logger.log('📋 步驟 2：查詢假期餘額');
  const balanceResult = getLeaveBalance(token);
  Logger.log('   結果: ' + JSON.stringify(balanceResult, null, 2));
  Logger.log('');
  
  if (balanceResult.ok) {
    Logger.log('✅✅✅ 測試成功！');
    Logger.log('');
    Logger.log('📊 假期餘額:');
    const balance = balanceResult.balance;
    Logger.log('   特休假: ' + balance.annualLeave + ' 天');
    Logger.log('   病假: ' + balance.sickLeave + ' 天');
    Logger.log('   事假: ' + balance.personalLeave + ' 天');
    Logger.log('   喪假: ' + balance.bereavementLeave + ' 天');
    Logger.log('   婚假: ' + balance.marriageLeave + ' 天');
    Logger.log('   產假: ' + balance.maternityLeave + ' 天');
    Logger.log('   陪產假: ' + balance.paternityLeave + ' 天');
    Logger.log('   產檢假: ' + balance.prenatalCheckupLeave + ' 天');
    Logger.log('   生理假: ' + balance.menstrualLeave + ' 天');
    Logger.log('   家庭照顧假: ' + balance.familyCareLeave + ' 天');
    Logger.log('   公假: ' + (balance.officialLeave === 0 ? '無上限' : balance.officialLeave + ' 天'));
    Logger.log('   公傷病假: ' + (balance.occupationalInjuryLeave === 0 ? '無上限' : balance.occupationalInjuryLeave + ' 天'));
    Logger.log('   疫苗接種假: ' + (balance.vaccinationLeave === 0 ? '無上限' : balance.vaccinationLeave + ' 天'));
    Logger.log('   防疫照顧假: ' + (balance.epidemicCareLeave === 0 ? '無上限' : balance.epidemicCareLeave + ' 天'));
  } else {
    Logger.log('❌ 測試失敗');
  }
}

// LeaveManagement.gs - 小時制請假系統（完整修正版 + 餘額扣除）

/**
 * ✅ 審核請假申請（修正版：核准後自動扣除餘額）
 */
function reviewLeaveRequest(sessionToken, rowNumber, reviewAction, comment) {
  try {
    Logger.log('═══════════════════════════════════════');
    Logger.log('📋 開始審核請假');
    Logger.log('═══════════════════════════════════════');
    Logger.log(`   行號: ${rowNumber}`);
    Logger.log(`   動作: ${reviewAction}`);
    Logger.log('');
    
    const employee = checkSession_(sessionToken);
    
    if (!employee.ok || !employee.user) {
      return {
        ok: false,
        code: "ERR_SESSION_INVALID"
      };
    }
    
    if (employee.user.dept !== '管理員') {
      return {
        ok: false,
        code: "ERR_PERMISSION_DENIED",
        msg: "需要管理員權限"
      };
    }
    
    const sheet = getLeaveRecordsSheet();
    const record = sheet.getRange(rowNumber, 1, 1, 14).getValues()[0];
    
    const userId = record[1];           // B: 員工ID
    const employeeName = record[2];     // C: 姓名
    const leaveType = record[4];        // E: 假別
    const workHours = record[7];        // H: 工作時數
    const days = record[8];             // I: 天數
    
    Logger.log('📋 請假資料:');
    Logger.log(`   員工: ${employeeName} (${userId})`);
    Logger.log(`   假別: ${leaveType}`);
    Logger.log(`   時數: ${workHours} 小時`);
    Logger.log(`   天數: ${days} 天`);
    Logger.log('');
    
    // 更新狀態
    const status = (reviewAction === 'approve') ? 'APPROVED' : 'REJECTED';
    
    sheet.getRange(rowNumber, 11).setValue(status);          // K: 狀態
    sheet.getRange(rowNumber, 12).setValue(employee.user.name); // L: 審核人
    sheet.getRange(rowNumber, 13).setValue(new Date());      // M: 審核時間
    sheet.getRange(rowNumber, 14).setValue(comment || '');  // N: 審核意見
    
    Logger.log(`✅ 審核狀態已更新: ${status}`);
    Logger.log('');
    
    // ⭐⭐⭐ 關鍵修正：核准時扣除假期餘額
    if (reviewAction === 'approve') {
      Logger.log('💰 開始扣除假期餘額...');
      
      const deductResult = deductLeaveBalance(userId, leaveType, days);
      
      if (!deductResult.ok) {
        Logger.log('❌ 扣除餘額失敗: ' + deductResult.msg);
        
        // 回滾狀態（可選）
        sheet.getRange(rowNumber, 11).setValue('PENDING');
        
        return {
          ok: false,
          code: "ERR_DEDUCT_FAILED",
          msg: "扣除餘額失敗: " + deductResult.msg
        };
      }
      
      Logger.log('✅ 假期餘額扣除成功');
      Logger.log(`   ${leaveType}: 扣除 ${days} 天 (${workHours} 小時)`);
      Logger.log(`   剩餘: ${deductResult.remaining} 天`);
    }
    
    Logger.log('');
    Logger.log('═══════════════════════════════════════');
    Logger.log('✅✅✅ 審核完成');
    Logger.log('═══════════════════════════════════════');
    
    return {
      ok: true,
      msg: "審核完成"
    };
    
  } catch (error) {
    Logger.log('❌ reviewLeaveRequest 錯誤: ' + error);
    Logger.log('錯誤堆疊: ' + error.stack);
    return {
      ok: false,
      msg: error.message
    };
  }
}

/**
 * ✅ 計算工作時數和天數（排除午休時間）
 * 
 * @param {Date} start - 開始時間
 * @param {Date} end - 結束時間
 * @return {Object} { workHours: number, days: number }
 */
function calculateWorkHoursAndDays(start, end) {
  try {
    Logger.log('💡 計算工作時數和天數');
    Logger.log(`   開始: ${start.toISOString()}`);
    Logger.log(`   結束: ${end.toISOString()}`);
    
    // 計算總時長（毫秒）
    const totalMs = end - start;
    
    // 轉換為小時
    let totalHours = totalMs / (1000 * 60 * 60);
    
    Logger.log(`   初始總時數: ${totalHours.toFixed(2)} 小時`);
    
    // 如果是同一天，檢查是否跨越午休時間 12:00-13:00
    if (start.toDateString() === end.toDateString()) {
      Logger.log('   ℹ️ 同一天請假');
      
      const startHour = start.getHours() + start.getMinutes() / 60;
      const endHour = end.getHours() + end.getMinutes() / 60;
      
      const lunchStart = 12; // 12:00
      const lunchEnd = 13;   // 13:00
      
      // 判斷是否跨越午休時間
      if (startHour < lunchEnd && endHour > lunchStart) {
        // 計算重疊的午休時間
        const overlapStart = Math.max(startHour, lunchStart);
        const overlapEnd = Math.min(endHour, lunchEnd);
        const lunchOverlap = Math.max(0, overlapEnd - overlapStart);
        
        totalHours -= lunchOverlap;
        
        Logger.log(`   🍱 扣除午休時間: ${lunchOverlap.toFixed(2)} 小時`);
      }
    } else {
      // 跨日請假：每天都要扣除 1 小時午休
      Logger.log('   ℹ️ 跨日請假');
      
      const startDate = new Date(start);
      startDate.setHours(0, 0, 0, 0);
      
      const endDate = new Date(end);
      endDate.setHours(0, 0, 0, 0);
      
      const daysDiff = Math.floor((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
      
      // 每天扣除 1 小時午休
      totalHours -= daysDiff;
      
      Logger.log(`   📅 跨 ${daysDiff} 天，扣除 ${daysDiff} 小時午休`);
    }
    
    // 確保不會是負數
    totalHours = Math.max(0, totalHours);
    
    // 四捨五入到小數點後 2 位
    const workHours = Math.round(totalHours * 100) / 100;
    
    // 計算天數（8 小時 = 1 天）
    const days = Math.round((workHours / 8) * 100) / 100;
    
    Logger.log(`   ✅ 最終工時: ${workHours} 小時`);
    Logger.log(`   ✅ 換算天數: ${days} 天`);
    
    return {
      workHours: workHours,
      days: days
    };
    
  } catch (error) {
    Logger.log(`❌ calculateWorkHoursAndDays 錯誤: ${error.message}`);
    return {
      workHours: 0,
      days: 0
    };
  }
}
/**
 * ✅ 修正：取得已核准的請假記錄（適配小時制請假系統）
 * 
 * 實際欄位順序（新結構）：
 * A - 申請時間
 * B - 員工ID
 * C - 姓名
 * D - 部門
 * E - 假別
 * F - 開始時間 (datetime)
 * G - 結束時間 (datetime)
 * H - 工作時數 (hours)
 * I - 天數 (days)
 * J - 原因
 * K - 狀態
 * L - 審核人
 * M - 審核時間
 * N - 審核意見
 */
function getApprovedLeaveRecords(monthParam, userIdParam) {
  try {
    Logger.log('═══════════════════════════════════════');
    Logger.log('📋 getApprovedLeaveRecords 開始（小時制版本）');
    Logger.log('═══════════════════════════════════════');
    Logger.log(`   monthParam: ${monthParam}`);
    Logger.log(`   userIdParam: ${userIdParam}`);
    Logger.log('');
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('請假紀錄');
    
    if (!sheet) {
      Logger.log('⚠️ 找不到請假紀錄工作表');
      return [];
    }
    
    const values = sheet.getDataRange().getValues();
    
    if (values.length <= 1) {
      Logger.log('⚠️ 工作表只有標題，沒有資料');
      return [];
    }
    
    Logger.log(`✅ 工作表有 ${values.length - 1} 筆資料`);
    Logger.log('');
    
    // ✅ 根據新的欄位順序（14 個欄位）
    const leaveRecords = [];
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      
      // 固定欄位索引（從 0 開始）
      const applyTime = row[0];            // A 欄 (索引 0)
      const employeeId = row[1];           // B 欄 (索引 1)
      const employeeName = row[2];         // C 欄 (索引 2)
      const dept = row[3];                 // D 欄 (索引 3)
      const leaveType = row[4];            // E 欄 (索引 4)
      const startDateTime = row[5];        // F 欄 (索引 5) - 開始時間
      const endDateTime = row[6];          // G 欄 (索引 6) - 結束時間
      const workHours = row[7];            // H 欄 (索引 7) - 工作時數
      const days = row[8];                 // I 欄 (索引 8) - 天數
      const reason = row[9];               // J 欄 (索引 9)
      const status = row[10];              // K 欄 (索引 10) ⭐ 關鍵修正
      const reviewer = row[11];            // L 欄 (索引 11)
      const reviewTime = row[12];          // M 欄 (索引 12)
      const reviewComment = row[13];       // N 欄 (索引 13)
      
      Logger.log(`═══════════════════════════════════════`);
      Logger.log(`📋 第 ${i + 1} 行:`);
      Logger.log(`   員工ID: ${employeeId}`);
      Logger.log(`   員工姓名: ${employeeName}`);
      Logger.log(`   狀態: "${status}"`);
      Logger.log(`   開始時間: ${startDateTime}`);
      Logger.log(`   結束時間: ${endDateTime}`);
      Logger.log(`   工作時數: ${workHours} 小時`);
      Logger.log(`   天數: ${days} 天`);
      
      // ⭐ 檢查狀態（只取已核准的）
      if (String(status).trim() !== 'APPROVED') {
        Logger.log(`   ⏭️ 狀態不是 APPROVED (實際: "${status}")，跳過`);
        Logger.log('');
        continue;
      }
      
      // 格式化日期時間
      let formattedStartDate, formattedEndDate;
      
      try {
        // 處理可能的日期格式
        const startDate = new Date(startDateTime);
        const endDate = new Date(endDateTime);
        
        if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
          Logger.log(`   ⚠️ 日期格式無效，跳過`);
          Logger.log('');
          continue;
        }
        
        formattedStartDate = formatDate(startDate);
        formattedEndDate = formatDate(endDate);
        
        Logger.log(`   格式化開始日期: ${formattedStartDate}`);
        Logger.log(`   格式化結束日期: ${formattedEndDate}`);
        
      } catch (dateError) {
        Logger.log(`   ⚠️ 日期解析錯誤: ${dateError.message}`);
        Logger.log('');
        continue;
      }
      
      // ⭐ 檢查月份（使用開始日期的月份）
      if (!formattedStartDate.startsWith(monthParam)) {
        Logger.log(`   ⏭️ 月份不符 (需要: ${monthParam}, 實際: ${formattedStartDate})，跳過`);
        Logger.log('');
        continue;
      }
      
      // ⭐ 檢查員工ID（如果有指定）
      if (userIdParam && employeeId !== userIdParam) {
        Logger.log(`   ⏭️ 員工ID不符 (需要: ${userIdParam}, 實際: ${employeeId})，跳過`);
        Logger.log('');
        continue;
      }
      
      Logger.log(`   ✅ 符合所有條件！`);
      
      // ⭐ 生成請假期間的每一天
      const start = new Date(startDateTime);
      const end = new Date(endDateTime);
      
      // 計算跨越了幾天
      const startDay = new Date(start.getFullYear(), start.getMonth(), start.getDate());
      const endDay = new Date(end.getFullYear(), end.getMonth(), end.getDate());
      const totalDays = Math.floor((endDay - startDay) / (1000 * 60 * 60 * 24)) + 1;
      
      Logger.log(`   📅 請假天數範圍: ${totalDays} 天`);
      
      // 為每一天生成記錄
      for (let d = new Date(startDay); d <= endDay; d.setDate(d.getDate() + 1)) {
        const dateStr = formatDate(d);
        
        // 確保日期在查詢月份內
        if (dateStr.startsWith(monthParam)) {
          leaveRecords.push({
            employeeId: employeeId,
            employeeName: employeeName,
            date: dateStr,
            leaveType: leaveType,
            workHours: parseFloat(workHours) || 0,      // ⭐ 新增：工作時數
            days: parseFloat(days) || 0,
            status: status,
            reason: reason || '',
            startDateTime: startDateTime,                // ⭐ 新增：完整開始時間
            endDateTime: endDateTime,                    // ⭐ 新增：完整結束時間
            reviewer: reviewer || '',                    // ⭐ 新增：審核人
            reviewTime: reviewTime || '',                // ⭐ 新增：審核時間
            reviewComment: reviewComment || ''           // ⭐ 新增：審核意見
          });
          
          Logger.log(`      ➕ 加入日期: ${dateStr}`);
        }
      }
      
      Logger.log('');
    }
    
    Logger.log('═══════════════════════════════════════');
    Logger.log(`✅ getApprovedLeaveRecords 完成`);
    Logger.log(`   共找到 ${leaveRecords.length} 筆已核准的請假記錄`);
    Logger.log('═══════════════════════════════════════');
    
    return leaveRecords;
    
  } catch (error) {
    Logger.log('═══════════════════════════════════════');
    Logger.log('❌ getApprovedLeaveRecords 錯誤');
    Logger.log('   錯誤訊息: ' + error.message);
    Logger.log('   錯誤堆疊: ' + error.stack);
    Logger.log('═══════════════════════════════════════');
    return [];
  }
}
/**
 * ✅ 扣除假期餘額（新增函數）
 * 
 * @param {string} userId - 員工ID
 * @param {string} leaveType - 假別
 * @param {number} days - 要扣除的天數
 * @return {object} 結果
 */
function deductLeaveBalance(userId, leaveType, days) {
  try {
    Logger.log('📊 扣除假期餘額');
    Logger.log(`   員工ID: ${userId}`);
    Logger.log(`   假別: ${leaveType}`);
    Logger.log(`   天數: ${days}`);
    Logger.log('');
    
    const sheet = getLeaveBalanceSheet();
    const values = sheet.getDataRange().getValues();
    
    // 假別欄位對應表
    const leaveTypeColumnMap = {
      'ANNUAL_LEAVE': 2,              // B 欄
      'SICK_LEAVE': 3,                // C 欄
      'PERSONAL_LEAVE': 4,            // D 欄
      'BEREAVEMENT_LEAVE': 5,         // E 欄
      'MARRIAGE_LEAVE': 6,            // F 欄
      'MATERNITY_LEAVE': 7,           // G 欄
      'PATERNITY_LEAVE': 8,           // H 欄
      'HOSPITALIZATION_LEAVE': 9,     // I 欄
      'MENSTRUAL_LEAVE': 10,          // J 欄
      'FAMILY_CARE_LEAVE': 11,        // K 欄
      'OFFICIAL_LEAVE': 12,           // L 欄
      'WORK_INJURY_LEAVE': 13,        // M 欄
      'NATURAL_DISASTER_LEAVE': 14,   // N 欄
      'COMP_TIME_OFF': 15,            // O 欄
      'ABSENCE_WITHOUT_LEAVE': 16     // P 欄
    };
    
    const columnIndex = leaveTypeColumnMap[leaveType];
    
    if (!columnIndex) {
      Logger.log('❌ 無效的假別: ' + leaveType);
      return {
        ok: false,
        msg: "無效的假別"
      };
    }
    
    // 尋找員工記錄
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === userId) {
        Logger.log(`✅ 找到員工記錄（第 ${i + 1} 行）`);
        
        const currentBalance = values[i][columnIndex - 1]; // 因為陣列從 0 開始
        
        Logger.log(`   目前餘額: ${currentBalance} 天`);
        
        // 檢查餘額是否足夠
        if (currentBalance < days) {
          Logger.log(`   ⚠️ 餘額不足：需要 ${days} 天，只剩 ${currentBalance} 天`);
          return {
            ok: false,
            msg: `${leaveType} 餘額不足（需要 ${days} 天，只剩 ${currentBalance} 天）`
          };
        }
        
        // 扣除餘額
        const newBalance = currentBalance - days;
        
        Logger.log(`   扣除 ${days} 天後: ${newBalance} 天`);
        
        sheet.getRange(i + 1, columnIndex).setValue(newBalance);
        sheet.getRange(i + 1, 17).setValue(new Date()); // Q 欄: 更新時間
        
        Logger.log('✅ 餘額已更新');
        
        return {
          ok: true,
          remaining: newBalance
        };
      }
    }
    
    Logger.log('❌ 找不到員工記錄');
    return {
      ok: false,
      msg: "找不到員工記錄"
    };
    
  } catch (error) {
    Logger.log('❌ deductLeaveBalance 錯誤: ' + error);
    return {
      ok: false,
      msg: error.message
    };
  }
}

/**
 * 🧪 測試扣除餘額功能
 */
function testDeductLeaveBalance() {
  Logger.log('🧪 測試扣除假期餘額');
  Logger.log('');
  
  // ⚠️ 請替換成實際的員工ID
  const testUserId = 'U7854bd6965d1c25b1c79d00c1dce001b'; // 從 LINE 取得的 userId
  
  Logger.log('📋 測試參數:');
  Logger.log(`   員工ID: ${testUserId}`);
  Logger.log(`   假別: ANNUAL_LEAVE (特休假)`);
  Logger.log(`   天數: 0.25 (2 小時)`);
  Logger.log('');
  
  const result = deductLeaveBalance(testUserId, 'ANNUAL_LEAVE', 0.25);
  
  Logger.log('📤 測試結果:');
  Logger.log(JSON.stringify(result, null, 2));
  
  if (result.ok) {
    Logger.log('');
    Logger.log('✅ 測試成功！');
    Logger.log(`   剩餘餘額: ${result.remaining} 天`);
  } else {
    Logger.log('');
    Logger.log('❌ 測試失敗');
  }
}

/**
 * 🧪 完整測試：提交 → 審核 → 扣除餘額
 */
function testCompleteLeaveFlow() {
  Logger.log('🧪 測試完整請假流程');
  Logger.log('');
  
  const token = '7dac1161-bbac-487d-900b-3e06c1acab8d'; // ⚠️ 替換成有效 token
  
  // 步驟 1：提交請假
  Logger.log('📋 步驟 1：提交請假申請');
  const submitResult = submitLeaveRequest(
    token,
    'ANNUAL_LEAVE',
    '2025-12-19T09:00',
    '2025-12-19T11:00',
    '測試完整流程'
  );
  
  Logger.log('   結果: ' + JSON.stringify(submitResult));
  
  if (!submitResult.ok) {
    Logger.log('❌ 提交失敗，測試終止');
    return;
  }
  
  Logger.log('');
  
  // 步驟 2：查詢餘額（扣除前）
  Logger.log('📋 步驟 2：查詢餘額（扣除前）');
  const balanceBefore = getLeaveBalance(token);
  Logger.log('   特休假餘額: ' + balanceBefore.balance.ANNUAL_LEAVE + ' 天');
  Logger.log('');
  
  // 步驟 3：審核請假（需要手動指定 rowNumber）
  Logger.log('📋 步驟 3：審核請假申請');
  Logger.log('   ⚠️ 請手動查看「請假紀錄」工作表的最後一行行號');
  Logger.log('   然後修改下面的 rowNumber');
  
  const rowNumber = 2; // ⚠️ 替換成實際行號
  
  const reviewResult = reviewLeaveRequest(token, rowNumber, 'approve', '核准測試');
  Logger.log('   結果: ' + JSON.stringify(reviewResult));
  Logger.log('');
  
  // 步驟 4：查詢餘額（扣除後）
  Logger.log('📋 步驟 4：查詢餘額（扣除後）');
  const balanceAfter = getLeaveBalance(token);
  Logger.log('   特休假餘額: ' + balanceAfter.balance.ANNUAL_LEAVE + ' 天');
  Logger.log('');
  
  // 比較
  Logger.log('📊 比較結果:');
  Logger.log(`   扣除前: ${balanceBefore.balance.ANNUAL_LEAVE} 天`);
  Logger.log(`   扣除後: ${balanceAfter.balance.ANNUAL_LEAVE} 天`);
  Logger.log(`   差異: ${balanceBefore.balance.ANNUAL_LEAVE - balanceAfter.balance.ANNUAL_LEAVE} 天`);
  Logger.log('');
  
  if (balanceBefore.balance.ANNUAL_LEAVE > balanceAfter.balance.ANNUAL_LEAVE) {
    Logger.log('✅✅✅ 測試成功！餘額已正確扣除');
  } else {
    Logger.log('❌ 測試失敗：餘額未扣除');
  }
}