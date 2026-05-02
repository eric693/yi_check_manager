/**
 * 排班管理模組
 * 負責處理員工排班的所有邏輯
 */

// ==================== ⭐ 格式化函數 (新增) ====================

function formatDateOnly(dateValue) {
  if (!dateValue) return "";
  
  let date;
  if (typeof dateValue === 'string') {
    // 已經是 YYYY-MM-DD
    if (/^\d{4}-\d{2}-\d{2}$/.test(dateValue)) {
      return dateValue;
    }
    // ⭐ 新增：支援 YYYY/M/D 或 YYYY/MM/DD 格式
    if (/^\d{4}\/\d{1,2}\/\d{1,2}$/.test(dateValue)) {
      const parts = dateValue.split('/');
      return `${parts[0]}-${String(parts[1]).padStart(2, '0')}-${String(parts[2]).padStart(2, '0')}`;
    }
    date = new Date(dateValue);
  } else if (dateValue instanceof Date) {
    date = dateValue;
  } else {
    return String(dateValue);
  }
  
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatTimeOnly(timeValue) {
    // ⭐ 修正：改用 == null
    if (timeValue == null || timeValue === '') return '00:00';
    
    if (typeof timeValue === 'string' && /^\d{2}:\d{2}$/.test(timeValue)) {
        return timeValue;
    }
    
    // ⭐ 新增：處理 "0:00"
    if (typeof timeValue === 'string' && /^\d{1}:\d{2}$/.test(timeValue)) {
        return '0' + timeValue;
    }
    
    if (typeof timeValue === 'string' && /^\d{2}:\d{2}:\d{2}$/.test(timeValue)) {
        return timeValue.substring(0, 5);
    }
    
    if (timeValue instanceof Date) {
        const hours = String(timeValue.getHours()).padStart(2, '0');
        const minutes = String(timeValue.getMinutes()).padStart(2, '0');
        return `${hours}:${minutes}`;
    }
    
    if (typeof timeValue === 'string') {
        try {
            const date = new Date(timeValue);
            const hours = String(date.getHours()).padStart(2, '0');
            const minutes = String(date.getMinutes()).padStart(2, '0');
            return `${hours}:${minutes}`;
        } catch (e) {
            return '00:00';
        }
    }
    
    return String(timeValue);
}

/**
 * ⭐ 格式化完整日期時間
 */
function formatDateTime(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

// ==================== 原有功能 ====================

/**
 * 取得排班工作表
 */
function getShiftSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('排班表');
  
  if (!sheet) {
    sheet = ss.insertSheet('排班表');
    const headers = [
      '排班ID',
      '員工ID', 
      '員工姓名',
      '日期',
      '班別',
      '上班時間',
      '下班時間',
      '地點',
      '備註',
      '建立時間',
      '建立者',
      '最後修改時間',
      '最後修改者',
      '狀態'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

/**
 * ✅ 新增排班（修正版）
 */
function addShift(shiftData) {
  try {
    const sheet = getShiftSheet();
    const userId = Session.getActiveUser().getEmail();
    
    // 驗證必填欄位
    if (!shiftData.employeeId || !shiftData.date || !shiftData.shiftType) {
      return {
        success: false,
        message: '請填寫所有必填欄位'
      };
    }
    
    // ⭐⭐⭐ 修正：傳入班別參數
    const isDuplicate = checkDuplicateShift(
      shiftData.employeeId, 
      shiftData.date, 
      shiftData.shiftType
    );
    
    if (isDuplicate) {
      return {
        success: false,
        message: '該員工在此日期已有此班別的排班'
      };
    }
    
    const shiftId = 'SHIFT-' + Utilities.getUuid();
    const timestamp = formatDateTime(new Date());
    
    const rowData = [
      shiftId,
      shiftData.employeeId,
      shiftData.employeeName || '',
      formatDateOnly(shiftData.date),
      shiftData.shiftType,
      formatTimeOnly(shiftData.startTime),
      formatTimeOnly(shiftData.endTime),
      shiftData.location || '',
      shiftData.note || '',
      timestamp,
      userId,
      timestamp,
      userId,
      '正常'
    ];
    
    sheet.appendRow(rowData);
    
    // 發送LINE通知
    try {
      sendShiftNotification(shiftData.employeeId, shiftData);
    } catch (e) {
      Logger.log('發送排班通知失敗: ' + e);
    }
    
    return {
      success: true,
      message: '排班新增成功',
      shiftId: shiftId
    };
    
  } catch (error) {
    Logger.log('新增排班錯誤: ' + error);
    return {
      success: false,
      message: '新增排班失敗: ' + error.message
    };
  }
}

/**
 * ✅ 統一版：檢查重複排班
 * 重複定義：同一員工 + 同一日期 + 同一班別
 * 
 * @param {string} employeeId - 員工ID
 * @param {string} date - 日期 (YYYY-MM-DD)
 * @param {string} shiftType - 班別
 * @returns {boolean} true=重複, false=不重複
 */
function checkDuplicateShift(employeeId, date, shiftType) {
  try {
    const sheet = getShiftSheet();
    const data = sheet.getDataRange().getValues();
    
    const targetDate = formatDateOnly(date);
    
    Logger.log(`🔍 檢查重複: ${employeeId} - ${targetDate} - ${shiftType}`);
    
    for (let i = 1; i < data.length; i++) {
      // 跳過已刪除的記錄
      if (data[i][13] === '已刪除') continue;
      
      const shiftDate = formatDateOnly(data[i][3]);
      
      // ⭐⭐⭐ 比較：員工ID + 日期 + 班別
      if (data[i][1] === employeeId && 
          shiftDate === targetDate && 
          data[i][4] === shiftType) {
        Logger.log(`⚠️ 發現重複: Row ${i + 1}`);
        return true;
      }
    }
    
    Logger.log(`✅ 無重複`);
    return false;
    
  } catch (error) {
    Logger.log('❌ checkDuplicateShift 錯誤: ' + error);
    return false; // 錯誤時允許新增
  }
}


/**
 * ✅ 批量新增排班（精細重複檢查版 - 已統一邏輯）
 */
function batchAddShifts(shiftsArray) {
  try {
    const sheet = getShiftSheet();
    const userId = Session.getActiveUser().getEmail();
    const timestamp = formatDateTime(new Date());
    const results = {
      success: 0,
      failed: 0,
      errors: []
    };
    
    Logger.log('═══════════════════════════════════════');
    Logger.log('📦 開始批量新增（精細重複檢查）');
    Logger.log('   總筆數: ' + shiftsArray.length);
    Logger.log('   重複定義: 員工ID + 日期 + 班別');
    Logger.log('═══════════════════════════════════════');
    
    // 檢查鍵：員工ID_日期_班別
    const processedInBatch = new Set();
    
    // 預先載入已存在的排班
    const existingShifts = new Set();
    const existingData = sheet.getDataRange().getValues();
    
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][13] !== '已刪除') {
        const existingKey = `${existingData[i][1]}_${formatDateOnly(existingData[i][3])}_${existingData[i][4]}`;
        existingShifts.add(existingKey);
      }
    }
    
    Logger.log('📊 工作表中已有 ' + existingShifts.size + ' 個排班');
    Logger.log('');
    
    // 處理每筆資料
    shiftsArray.forEach((shiftData, index) => {
      try {
        Logger.log(`📝 處理第 ${index + 1}/${shiftsArray.length} 筆`);
        Logger.log(`   員工: ${shiftData.employeeName}`);
        Logger.log(`   日期: ${shiftData.date}`);
        Logger.log(`   班別: ${shiftData.shiftType}`);
        
        const formattedDate = formatDateOnly(shiftData.date);
        const key = `${shiftData.employeeId}_${formattedDate}_${shiftData.shiftType}`;
        
        Logger.log(`   檢查鍵: ${key}`);
        
        // 檢查 1：本批次中是否已處理過
        if (processedInBatch.has(key)) {
          Logger.log(`   ❌ 批次內重複`);
          results.failed++;
          results.errors.push(
            `第 ${index + 1} 筆 (${shiftData.employeeName} ${formattedDate} ${shiftData.shiftType}): 批次中已有相同的排班`
          );
          return;
        }
        
        // 檢查 2：工作表中是否已存在
        if (existingShifts.has(key)) {
          Logger.log(`   ❌ 工作表中已存在`);
          results.failed++;
          results.errors.push(
            `第 ${index + 1} 筆 (${shiftData.employeeName} ${formattedDate} ${shiftData.shiftType}): 該員工在此日期已有此班別`
          );
          return;
        }
        
        // 標記為已處理
        processedInBatch.add(key);
        
        // 新增到工作表
        const shiftId = 'SHIFT-' + Utilities.getUuid();
        
        const rowData = [
          shiftId,
          shiftData.employeeId,
          shiftData.employeeName || '',
          formattedDate,
          shiftData.shiftType,
          formatTimeOnly(shiftData.startTime),
          formatTimeOnly(shiftData.endTime),
          shiftData.location || '',
          shiftData.note || '',
          timestamp,
          userId,
          timestamp,
          userId,
          '正常'
        ];
        
        sheet.appendRow(rowData);
        results.success++;
        
        Logger.log(`   ✅ 新增成功`);
        
      } catch (e) {
        Logger.log(`   ❌ 例外錯誤: ${e.message}`);
        results.failed++;
        results.errors.push(`第 ${index + 1} 筆: ${e.message}`);
      }
    });
    
    Logger.log('');
    Logger.log('═══════════════════════════════════════');
    Logger.log('📊 批量新增完成');
    Logger.log(`   ✅ 成功: ${results.success} 筆`);
    Logger.log(`   ❌ 失敗: ${results.failed} 筆`);
    Logger.log('═══════════════════════════════════════');
    
    return {
      success: true,
      message: `批量新增完成: 成功 ${results.success} 筆, 失敗 ${results.failed} 筆`,
      results: results
    };
    
  } catch (error) {
    Logger.log('❌ batchAddShifts 整體錯誤: ' + error);
    return {
      success: false,
      message: '批量新增失敗: ' + error.message
    };
  }
}
/**
 * 查詢排班
 */
function getShifts(filters) {
  try {
    const sheet = getShiftSheet();
    const data = sheet.getDataRange().getValues();
    const shifts = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (row[13] === '已刪除') continue;
      
      const shiftDate = formatDateOnly(row[3]);
      
      if (filters) {
        if (filters.employeeId && row[1] !== filters.employeeId) continue;
        if (filters.startDate && shiftDate < formatDateOnly(filters.startDate)) continue;
        if (filters.endDate && shiftDate > formatDateOnly(filters.endDate)) continue;
        if (filters.shiftType && row[4] !== filters.shiftType) continue;
        if (filters.location && row[7] !== filters.location) continue;
      }
      
      shifts.push({
        shiftId: row[0],
        employeeId: row[1],
        employeeName: row[2],
        date: formatDateOnly(row[3]),
        shiftType: row[4],
        startTime: formatTimeOnly(row[5]),
        endTime: formatTimeOnly(row[6]),
        location: row[7],
        note: row[8],
        createdAt: row[9],
        createdBy: row[10],
        updatedAt: row[11],
        updatedBy: row[12],
        status: row[13]
      });
    }
    
    shifts.sort((a, b) => new Date(b.date) - new Date(a.date));
    
    return {
      success: true,
      data: shifts,
      count: shifts.length
    };
    
  } catch (error) {
    Logger.log('查詢排班錯誤: ' + error);
    return {
      success: false,
      message: '查詢排班失敗: ' + error.message,
      data: []
    };
  }
}


/**
 * 取得單一排班詳情 (⭐ 已修正 - 格式化回傳資料)
 */
function getShiftById(shiftId) {
  try {
    const sheet = getShiftSheet();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === shiftId) {
        // ✅ 格式化回傳資料
        return {
          success: true,
          data: {
            shiftId: data[i][0],
            employeeId: data[i][1],
            employeeName: data[i][2],
            date: formatDateOnly(data[i][3]),
            shiftType: data[i][4],
            startTime: formatTimeOnly(data[i][5]),
            endTime: formatTimeOnly(data[i][6]),
            location: data[i][7],
            note: data[i][8],
            createdAt: data[i][9],
            createdBy: data[i][10],
            updatedAt: data[i][11],
            updatedBy: data[i][12],
            status: data[i][13]
          }
        };
      }
    }
    
    return {
      success: false,
      message: '找不到該排班記錄'
    };
    
  } catch (error) {
    Logger.log('查詢排班詳情錯誤: ' + error);
    return {
      success: false,
      message: '查詢失敗: ' + error.message
    };
  }
}

/**
 * 更新排班
 */
function updateShift(shiftId, updateData) {
  try {
    const sheet = getShiftSheet();
    const data = sheet.getDataRange().getValues();
    const userId = Session.getActiveUser().getEmail();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === shiftId) {
        if (updateData.date) sheet.getRange(i + 1, 4).setValue(formatDateOnly(updateData.date));
        if (updateData.shiftType) sheet.getRange(i + 1, 5).setValue(updateData.shiftType);
        if (updateData.startTime) sheet.getRange(i + 1, 6).setValue(formatTimeOnly(updateData.startTime));
        if (updateData.endTime) sheet.getRange(i + 1, 7).setValue(formatTimeOnly(updateData.endTime));
        if (updateData.location) sheet.getRange(i + 1, 8).setValue(updateData.location);
        if (updateData.note !== undefined) sheet.getRange(i + 1, 9).setValue(updateData.note);
        
        sheet.getRange(i + 1, 12).setValue(formatDateTime(new Date()));
        sheet.getRange(i + 1, 13).setValue(userId);
        
        return {
          success: true,
          message: '排班更新成功'
        };
      }
    }
    
    return {
      success: false,
      message: '找不到該排班記錄'
    };
    
  } catch (error) {
    Logger.log('更新排班錯誤: ' + error);
    return {
      success: false,
      message: '更新失敗: ' + error.message
    };
  }
}

/**
 * 刪除排班
 */
function deleteShift(shiftId) {
  try {
    const sheet = getShiftSheet();
    const data = sheet.getDataRange().getValues();
    const userId = Session.getActiveUser().getEmail();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === shiftId) {
        sheet.getRange(i + 1, 14).setValue('已刪除');
        sheet.getRange(i + 1, 12).setValue(formatDateTime(new Date()));
        sheet.getRange(i + 1, 13).setValue(userId);
        
        return {
          success: true,
          message: '排班刪除成功'
        };
      }
    }
    
    return {
      success: false,
      message: '找不到該排班記錄'
    };
    
  } catch (error) {
    Logger.log('刪除排班錯誤: ' + error);
    return {
      success: false,
      message: '刪除失敗: ' + error.message
    };
  }
}


/**
 * 取得員工的排班資訊（用於打卡驗證） (⭐ 已修正 - 格式化回傳資料)
 */
function getEmployeeShiftForDate(employeeId, date) {
  try {
    const sheet = getShiftSheet();
    const data = sheet.getDataRange().getValues();
    
    const targetDate = formatDateOnly(date);
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === employeeId && data[i][13] !== '已刪除') {
        const shiftDate = formatDateOnly(data[i][3]);
        
        if (shiftDate === targetDate) {
          // ✅ 格式化回傳資料
          return {
            success: true,
            hasShift: true,
            data: {
              shiftId: data[i][0],
              shiftType: data[i][4],
              startTime: formatTimeOnly(data[i][5]),
              endTime: formatTimeOnly(data[i][6]),
              location: data[i][7]
            }
          };
        }
      }
    }
    
    return {
      success: true,
      hasShift: false,
      message: '今日無排班'
    };
    
  } catch (error) {
    Logger.log('查詢員工排班錯誤: ' + error);
    return {
      success: false,
      message: '查詢失敗: ' + error.message
    };
  }
}

/**
 * 取得本週排班統計
 */
function getWeeklyShiftStats() {
  try {
    const sheet = getShiftSheet();
    const data = sheet.getDataRange().getValues();
    const today = new Date();
    const startOfWeek = new Date(today);
    startOfWeek.setDate(today.getDate() - today.getDay());
    const endOfWeek = new Date(startOfWeek);
    endOfWeek.setDate(startOfWeek.getDate() + 6);
    
    const startDateStr = formatDateOnly(startOfWeek);
    const endDateStr = formatDateOnly(endOfWeek);
    
    const stats = {
      totalShifts: 0,
      byShiftType: {},
      byEmployee: {}
    };
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][13] === '已刪除') continue;
      
      const shiftDate = formatDateOnly(data[i][3]);
      if (shiftDate >= startDateStr && shiftDate <= endDateStr) {
        stats.totalShifts++;
        
        const shiftType = data[i][4];
        stats.byShiftType[shiftType] = (stats.byShiftType[shiftType] || 0) + 1;
        
        const employeeName = data[i][2];
        stats.byEmployee[employeeName] = (stats.byEmployee[employeeName] || 0) + 1;
      }
    }
    
    return {
      success: true,
      data: stats
    };
    
  } catch (error) {
    Logger.log('取得排班統計錯誤: ' + error);
    return {
      success: false,
      message: '取得統計失敗: ' + error.message
    };
  }
}

/**
 * 匯出排班資料
 */
function exportShifts(filters) {
  try {
    const result = getShifts(filters);
    if (!result.success) {
      return result;
    }
    
    return {
      success: true,
      data: result.data,
      filename: `排班表_${formatDateOnly(new Date()).replace(/-/g, '')}.csv`
    };
    
  } catch (error) {
    Logger.log('匯出排班錯誤: ' + error);
    return {
      success: false,
      message: '匯出失敗: ' + error.message
    };
  }
}

/**
 * 發送排班通知（透過LINE）
 */
function sendShiftNotification(employeeId, shiftData) {
  try {
    // 取得員工的LINE User ID
    const userInfo = getUserInfoByEmployeeId(employeeId);
    if (!userInfo || !userInfo.lineUserId) {
      Logger.log('找不到員工的LINE ID');
      return;
    }
    
    const message = `您好！您有新的排班通知：\n\n` +
                   `日期: ${shiftData.date}\n` +
                   `班別: ${shiftData.shiftType}\n` +
                   `上班時間: ${shiftData.startTime}\n` +
                   `下班時間: ${shiftData.endTime}\n` +
                   `地點: ${shiftData.location}\n` +
                   `${shiftData.note ? '備註: ' + shiftData.note : ''}`;
    
    sendLineMessage(userInfo.lineUserId, message);
    
  } catch (error) {
    Logger.log('發送排班通知錯誤: ' + error);
  }
}

/**
 * 從員工ID取得使用者資訊
 */
function getUserInfoByEmployeeId(employeeId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName('使用者資料');
    if (!userSheet) return null;
    
    const data = userSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === employeeId) {
        return {
          lineUserId: data[i][1],
          name: data[i][2],
          email: data[i][3]
        };
      }
    }
    
    return null;
  } catch (error) {
    Logger.log('取得使用者資訊錯誤: ' + error);
    return null;
  }
}

// ==================== 測試函數 ====================

/**
 * 測試時間格式化
 */
function testTimeFormatting() {
  const testCases = [
    "08:00",
    "08:00:00",
    new Date("2025-10-24T00:00:00"),
    "1899-12-30T01:00:00.000Z"
  ];
  
  Logger.log("=== 時間格式化測試 ===");
  testCases.forEach(test => {
    Logger.log(`輸入: ${test} → 輸出: ${formatTimeOnly(test)}`);
  });
}

/**
 * 測試排班系統
 */
function testShiftSystem() {
  Logger.log('===== 測試排班系統 =====');
  
  const testShift = {
    employeeId: 'TEST001',
    employeeName: '測試員工',
    date: '2025-10-25',
    shiftType: '早班',
    startTime: '08:00',
    endTime: '16:00',
    location: '測試地點',
    note: '測試備註'
  };
  
  const addResult = addShift(testShift);
  Logger.log('新增結果: ' + JSON.stringify(addResult));
  
  const queryResult = getShifts({ employeeId: 'TEST001' });
  Logger.log('查詢結果: ' + JSON.stringify(queryResult));
}


function testSingleShift() {
  const testData = {
    employeeId: 'Ue76b65367821240ac26387d2972a5adf',
    employeeName: '測試員工',
    date: '2026-02-20',
    shiftType: '廚房A班',
    startTime: '11:00',
    endTime: '20:00',
    location: '總公司',
    note: '測試'
  };
  
  Logger.log('測試單筆新增');
  const result = addShift(testData);
  Logger.log('結果: ' + JSON.stringify(result));
}

function checkExistingShifts() {
  const sheet = getShiftSheet();
  const data = sheet.getDataRange().getValues();
  
  Logger.log('現有排班數量: ' + (data.length - 1));
  
  // 檢查日期格式
  for (let i = 1; i <= Math.min(5, data.length - 1); i++) {
    Logger.log(`Row ${i + 1}:`);
    Logger.log(`  日期原始: ${data[i][3]}`);
    Logger.log(`  日期類型: ${typeof data[i][3]}`);
    Logger.log(`  格式化後: ${formatDateOnly(data[i][3])}`);
  }
}

// ==================== 班別設定 CRUD ====================

const BUILT_IN_SHIFT_SEED = [
  { name: '廚房A班', startTime: '11:00', endTime: '20:00', category: '廚房' },
  { name: '廚房B班', startTime: '11:30', endTime: '20:30', category: '廚房' },
  { name: '廚房C班', startTime: '12:00', endTime: '21:00', category: '廚房' },
  { name: '廚房D班', startTime: '13:00', endTime: '22:00', category: '廚房' },
  { name: '廚房E班', startTime: '14:00', endTime: '23:00', category: '廚房' },
  { name: '廚房F班', startTime: '15:00', endTime: '00:00', category: '廚房' },
  { name: '廚房G班', startTime: '11:30', endTime: '15:00', category: '廚房' },
  { name: '廚房H班', startTime: '18:00', endTime: '23:00', category: '廚房' },
  { name: '廚房I班', startTime: '18:00', endTime: '00:00', category: '廚房' },
  { name: '外場A1班', startTime: '11:00', endTime: '20:00', category: '外場' },
  { name: '外場A2班', startTime: '11:30', endTime: '16:30', category: '外場' },
  { name: '外場A3班', startTime: '11:30', endTime: '17:00', category: '外場' },
  { name: '外場A4班', startTime: '11:30', endTime: '20:30', category: '外場' },
  { name: '外場B1班', startTime: '16:00', endTime: '01:00', category: '外場' },
  { name: '外場B2班', startTime: '17:00', endTime: '01:00', category: '外場' },
  { name: '外場B3班', startTime: '18:00', endTime: '01:00', category: '外場' },
  { name: '外場B4班', startTime: '19:00', endTime: '01:00', category: '外場' },
  { name: '年假',     startTime: '00:00', endTime: '00:00', category: '假別' },
  { name: '過年假',   startTime: '00:00', endTime: '00:00', category: '假別' },
  { name: '國定假日', startTime: '00:00', endTime: '00:00', category: '假別' },
  { name: '排休',     startTime: '00:00', endTime: '00:00', category: '假別' }
];

function getShiftTypeSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_SHIFT_TYPES);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SHIFT_TYPES);
    const headers = ['班別ID', '班別名稱', '上班時間', '下班時間', '類別', '建立時間', '狀態'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#34a853')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);

    // 強制 C、D 欄（上班/下班時間）儲存為純文字，防止 Sheets 自動轉換時間格式
    sheet.getRange('C:D').setNumberFormat('@STRING@');

    const timestamp = formatDateTime(new Date());
    const rows = BUILT_IN_SHIFT_SEED.map(t => [
      'ST-' + Utilities.getUuid(),
      t.name, t.startTime, t.endTime, t.category,
      timestamp, '啟用'
    ]);
    sheet.getRange(2, 1, rows.length, 7).setValues(rows);
    Logger.log('✅ 班別設定工作表建立並初始化完成');
  }

  return sheet;
}

function getShiftTypesData() {
  try {
    const sheet = getShiftTypeSheet_();
    const data = sheet.getDataRange().getValues();
    const types = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i][6] === '停用') continue;
      types.push({
        id:        data[i][0],
        name:      data[i][1],
        startTime: formatTimeOnly(data[i][2]),
        endTime:   formatTimeOnly(data[i][3]),
        category:  data[i][4],
        createdAt: String(data[i][5]),
        status:    data[i][6]
      });
    }

    return { success: true, data: types };
  } catch (error) {
    Logger.log('❌ getShiftTypesData 錯誤: ' + error);
    return { success: false, message: error.message, data: [] };
  }
}

function addShiftTypeRecord(data) {
  try {
    const sheet = getShiftTypeSheet_();
    const existing = sheet.getDataRange().getValues();

    for (let i = 1; i < existing.length; i++) {
      if (existing[i][1] === data.name && existing[i][6] !== '停用') {
        return { success: false, message: `班別「${data.name}」已存在` };
      }
    }

    const id = 'ST-' + Utilities.getUuid();
    const lastRow = sheet.getLastRow() + 1;
    sheet.getRange(lastRow, 1, 1, 7).setValues([[
      id,
      data.name,
      data.startTime || '00:00',
      data.endTime   || '00:00',
      data.category  || '自訂',
      formatDateTime(new Date()),
      '啟用'
    ]]);
    // 確保時間欄位保持純文字格式
    sheet.getRange(lastRow, 3, 1, 2).setNumberFormat('@STRING@');

    return { success: true, message: '班別新增成功', id: id };
  } catch (error) {
    Logger.log('❌ addShiftTypeRecord 錯誤: ' + error);
    return { success: false, message: error.message };
  }
}

function updateShiftTypeRecord(id, data) {
  try {
    const sheet = getShiftTypeSheet_();
    const rows = sheet.getDataRange().getValues();

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) {
        if (data.name      !== undefined) sheet.getRange(i + 1, 2).setValue(data.name);
        if (data.startTime !== undefined) {
          sheet.getRange(i + 1, 3).setNumberFormat('@STRING@').setValue(data.startTime);
        }
        if (data.endTime   !== undefined) {
          sheet.getRange(i + 1, 4).setNumberFormat('@STRING@').setValue(data.endTime);
        }
        if (data.category  !== undefined) sheet.getRange(i + 1, 5).setValue(data.category);
        return { success: true, message: '班別更新成功' };
      }
    }

    return { success: false, message: '找不到該班別' };
  } catch (error) {
    Logger.log('❌ updateShiftTypeRecord 錯誤: ' + error);
    return { success: false, message: error.message };
  }
}

function deleteShiftTypeRecord(id) {
  try {
    const sheet = getShiftTypeSheet_();
    const rows = sheet.getDataRange().getValues();

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) {
        sheet.getRange(i + 1, 7).setValue('停用');
        return { success: true, message: '班別刪除成功' };
      }
    }

    return { success: false, message: '找不到該班別' };
  } catch (error) {
    Logger.log('❌ deleteShiftTypeRecord 錯誤: ' + error);
    return { success: false, message: error.message };
  }
}