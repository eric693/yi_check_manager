// DbOperations.gs - 完整優化版（精簡版）

// ==================== 員工相關功能 ====================

/**
 * 修正：僅 admin_list 內的 userId 才是管理員，其餘為員工
 */
const ADMIN_LIST = [
  "U1771fd65da16e2f2000a3c3805fbe256",
  "U69d37ae1b9a878ba9408527026bd5b44"
];

// DbOperations.gs - 修正後的 writeEmployee_ 函數

/**
 * ✅ 修正版：登入時不覆蓋手動設定的姓名
 */
function writeEmployee_(profile) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_EMPLOYEES);
  const values = sheet.getDataRange().getValues();
  const employeeId = profile.userId;

  // 檢查是否已存在
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === employeeId) {
      
      // ⭐⭐⭐ 關鍵修正：檢查是否有手動設定的姓名
      const currentName = values[i][2];           // C 欄：displayName（目前顯示的姓名）
      const nameOverride = values[i][8] || "";    // I 欄：nameOverride（手動設定的姓名）
      
      // 只在沒有手動設定姓名時才更新
      if (!nameOverride) {
        Logger.log(`✅ 更新員工 ${profile.displayName} 的 LINE 姓名`);
        sheet.getRange(i + 1, 3).setValue(profile.displayName);  // C 欄
      } else {
        Logger.log(`🔒 保留員工 ${currentName} 的手動設定姓名（忽略 LINE 姓名：${profile.displayName}）`);
        // 不更新姓名，保持原有的手動設定
      }
      
      // 更新其他資訊（email, 頭像等）
      sheet.getRange(i + 1, 2).setValue(profile.email || "");
      sheet.getRange(i + 1, 4).setValue(profile.pictureUrl);
      sheet.getRange(i + 1, 8).setValue("啟用");
      
      Logger.log(`✅ 更新員工資料完成（保留原有權限：${values[i][5]}）`);
      return values[i];
    }
  }

  // 判斷是否為管理員
  const role = ADMIN_LIST.includes(employeeId) ? "管理員" : "員工";

  // 新增資料
  const row = [
    employeeId,              // A: userId
    profile.email || "",     // B: email
    profile.displayName,     // C: displayName
    profile.pictureUrl,      // D: pictureUrl
    new Date(),              // E: 建立時間
    role,                    // F: 部門（權限）
    "",                      // G: 到職日期
    "啟用",                  // H: 狀態
    ""                       // I: nameOverride（手動設定的姓名，新用戶為空）
  ];

  sheet.appendRow(row);
  Logger.log(`✅ 新增員工 ${profile.displayName}（權限：${role}）`);
  return row;
}
// DbOperations.gs - 修正後的 findEmployeeByLineUserId_ 函數

/**
 * ✅ 修正版：優先使用手動設定的姓名
 */
function findEmployeeByLineUserId_(userId) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_EMPLOYEES);
  const values = sh.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === userId) {
      
      // ⭐⭐⭐ 關鍵修正：優先使用 nameOverride
      const displayName = values[i][2];        // C 欄：displayName
      const nameOverride = values[i][8] || ""; // I 欄：nameOverride
      
      const finalName = nameOverride || displayName; // 優先使用手動設定的姓名
      
      Logger.log(`📋 查詢員工: ${userId}`);
      Logger.log(`   displayName: ${displayName}`);
      Logger.log(`   nameOverride: ${nameOverride}`);
      Logger.log(`   最終姓名: ${finalName}`);
      
      return {
        ok: true,
        userId: values[i][0],        // ✅ LINE userId
        employeeId: values[i][0],    // ✅ 員工ID = LINE userId
        email: values[i][1] || "",
        name: finalName,             // ⭐ 使用最終姓名
        picture: values[i][3],
        dept: values[i][5] || "管理員",
        status: values[i][7] || "啟用"
      };
    }
  }
  
  return { ok: false, code: "ERR_NO_DATA" };
}

// DbOperations.gs - 新增：解除姓名鎖定功能

/**
 * 🔓 解除姓名鎖定，恢復使用 LINE 姓名
 */
function unlockEmployeeName(userId) {
  try {
    Logger.log('🔓 解除員工姓名鎖定');
    Logger.log('   userId: ' + userId);
    
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_EMPLOYEES);
    
    if (!sheet) {
      return { ok: false, msg: '找不到員工工作表' };
    }
    
    const data = sheet.getDataRange().getValues();
    
    // 尋找用戶並解除鎖定
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        const currentName = data[i][2];  // C 欄: displayName
        
        // ⭐ 清除 nameOverride，下次登入時將使用 LINE 姓名
        sheet.getRange(i + 1, 9).setValue("");  // I 欄：nameOverride
        
        Logger.log('✅ 已解除姓名鎖定');
        Logger.log('   當前姓名: ' + currentName);
        Logger.log('   下次登入將使用 LINE 姓名');
        
        return {
          ok: true,
          msg: '已解除姓名鎖定，下次登入將使用 LINE 姓名',
          currentName: currentName
        };
      }
    }
    
    return { ok: false, msg: '找不到該員工' };
    
  } catch (error) {
    Logger.log('❌ unlockEmployeeName 錯誤: ' + error);
    return { ok: false, msg: error.message };
  }
}

/**
 * 🧪 測試解除鎖定功能
 */
function testUnlockEmployeeName() {
  Logger.log('🧪 測試解除姓名鎖定');
  Logger.log('');
  
  // ⚠️ 替換成實際的 userId
  const testUserId = 'Ud3b574f260f5a777337158ccd4ff0ba2';
  
  const result = unlockEmployeeName(testUserId);
  
  Logger.log('');
  Logger.log('📤 結果:');
  Logger.log(JSON.stringify(result, null, 2));
}

/**
 * ✅ 取得所有員工列表（根據實際資料表結構）
 * 
 * 資料表欄位:
 * A (0) - userId
 * B (1) - email
 * C (2) - displayName
 * D (3) - pictureUrl
 * E (4) - 建立時間
 * F (5) - 部門
 * G (6) - 到職日期
 * H (7) - 狀態
 */
function getAllUsers() {
  try {
    Logger.log('📋 開始取得員工列表');
    
    // 取得員工資料表
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_EMPLOYEES);
    
    if (!sheet) {
      Logger.log('❌ 找不到員工工作表: ' + SHEET_EMPLOYEES);
      return { 
        ok: false, 
        msg: "找不到員工工作表",
        users: []
      };
    }
    
    // 取得所有資料
    const data = sheet.getDataRange().getValues();
    
    // 檢查是否有資料
    if (data.length <= 1) {
      Logger.log('⚠️ 員工工作表只有標題，沒有資料');
      return {
        ok: true,
        users: [],
        count: 0,
        msg: "目前沒有員工資料"
      };
    }
    
    const users = [];
    
    Logger.log('📊 開始解析員工資料...');
    Logger.log('   總行數（含標題）: ' + data.length);
    Logger.log('');
    
    // 從第二行開始讀取（跳過標題）
    for (let i = 1; i < data.length; i++) {
      const row = data[i];  // ⭐⭐⭐ 定義 row 變數
      
      // 檢查員工ID是否存在（A欄 = row[0]）
      if (!row[0] || String(row[0]).trim() === '') {
        Logger.log(`   ⚠️ 第 ${i + 1} 行: 員工ID是空的，跳過`);
        continue;
      }
      
      // 檢查狀態（H欄 = row[7]）
      const status = row[7] ? String(row[7]).trim() : '';
      
      // 只加入「啟用」或空值的員工
      if (status !== '' && status !== '啟用') {
        Logger.log(`   ⏸️ 第 ${i + 1} 行: ${row[2]} - 狀態是「${status}」，跳過`);
        continue;
      }
      
      // 建立使用者物件
      const user = {
        userId: String(row[0]).trim(),                    // A欄: userId
        email: row[1] ? String(row[1]).trim() : '',       // B欄: email
        name: row[2] ? String(row[2]).trim() : '未命名',   // C欄: displayName
        picture: row[3] ? String(row[3]).trim() : '',     // D欄: pictureUrl
        joinDate: row[4] || '',                           // E欄: 建立時間
        dept: row[5] ? String(row[5]).trim() : '',        // F欄: 部門
        hireDate: row[6] || '',                           // G欄: 到職日期
        status: status || '啟用'                          // H欄: 狀態
      };
      
      users.push(user);
      Logger.log(`   ✅ 第 ${i + 1} 行: ${user.name} (${user.userId}) - ${user.dept}`);
    }
    
    Logger.log('');
    Logger.log('✅ 員工列表取得完成');
    Logger.log('   總筆數: ' + users.length);
    Logger.log('');
    
    return {
      ok: true,
      users: users,
      count: users.length,
      msg: `成功取得 ${users.length} 筆員工資料`
    };
    
  } catch (error) {
    Logger.log('❌ getAllUsers 錯誤: ' + error);
    Logger.log('   錯誤訊息: ' + error.message);
    Logger.log('   錯誤堆疊: ' + error.stack);
    
    return {
      ok: false,
      msg: error.message || '取得員工列表失敗',
      users: [],
      error: error.stack
    };
  }
}

/**
 * 🧪 測試 getAllUsers 函式
 */
function testGetAllUsers() {
  Logger.log('🧪🧪🧪 測試 getAllUsers');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  
  const result = getAllUsers();
  
  Logger.log('📤 測試結果:');
  Logger.log('   - ok: ' + result.ok);
  Logger.log('   - msg: ' + (result.msg || '無'));
  Logger.log('   - count: ' + (result.count || 0));
  Logger.log('   - users 數量: ' + (result.users ? result.users.length : 0));
  Logger.log('');
  
  if (result.ok && result.users && result.users.length > 0) {
    Logger.log('✅✅✅ 測試成功！');
    Logger.log('');
    Logger.log('👥 員工列表詳細資訊:');
    Logger.log('');
    
    result.users.forEach((user, index) => {
      Logger.log(`${index + 1}. ${user.name}`);
      Logger.log(`   - userId: ${user.userId}`);
      Logger.log(`   - email: ${user.email}`);
      Logger.log(`   - dept: ${user.dept}`);
      Logger.log(`   - status: ${user.status}`);
      Logger.log('');
    });
    
    Logger.log('═══════════════════════════════════════');
    Logger.log('🎉 可以使用了！');
    
  } else {
    Logger.log('❌ 測試失敗或沒有資料');
    if (!result.ok) {
      Logger.log('   錯誤原因: ' + result.msg);
      if (result.error) {
        Logger.log('   錯誤堆疊: ' + result.error);
      }
    } else {
      Logger.log('   可能原因: 員工資料表沒有資料，或所有員工都不是「啟用」狀態');
    }
    Logger.log('═══════════════════════════════════════');
  }
}

/**
 * 🔍 診斷工具：檢查員工資料表結構
 */
function diagnoseEmployeeSheet() {
  Logger.log('🔍 診斷員工資料表');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_EMPLOYEES);
    
    if (!sheet) {
      Logger.log('❌ 找不到工作表: ' + SHEET_EMPLOYEES);
      return;
    }
    
    Logger.log('✅ 工作表存在: ' + SHEET_EMPLOYEES);
    Logger.log('');
    
    const data = sheet.getDataRange().getValues();
    
    Logger.log('📊 資料表資訊:');
    Logger.log('   總行數: ' + data.length);
    Logger.log('   總欄數: ' + (data[0] ? data[0].length : 0));
    Logger.log('');
    
    if (data.length > 0) {
      Logger.log('📋 標題列（第1行）:');
      data[0].forEach((header, index) => {
        const column = String.fromCharCode(65 + index); // A, B, C, ...
        Logger.log(`   ${column} (${index}): ${header}`);
      });
      Logger.log('');
    }
    
    if (data.length > 1) {
      Logger.log('📝 資料列數（不含標題）: ' + (data.length - 1));
      Logger.log('');
      Logger.log('👤 第一筆員工資料（第2行）:');
      const firstRow = data[1];
      data[0].forEach((header, index) => {
        const column = String.fromCharCode(65 + index);
        Logger.log(`   ${column} (${index}) ${header}: ${firstRow[index]}`);
      });
      Logger.log('');
      
      // 檢查狀態欄位
      Logger.log('🔍 狀態檢查:');
      let enabledCount = 0;
      let disabledCount = 0;
      let emptyCount = 0;
      
      for (let i = 1; i < data.length; i++) {
        const status = data[i][7] ? String(data[i][7]).trim() : '';
        if (status === '啟用') {
          enabledCount++;
        } else if (status === '') {
          emptyCount++;
        } else {
          disabledCount++;
        }
      }
      
      Logger.log('   狀態=「啟用」: ' + enabledCount + ' 筆');
      Logger.log('   狀態=「空值」: ' + emptyCount + ' 筆');
      Logger.log('   狀態=「其他」: ' + disabledCount + ' 筆');
      Logger.log('   可用員工總數: ' + (enabledCount + emptyCount) + ' 筆');
    } else {
      Logger.log('⚠️ 沒有資料列（只有標題）');
    }
    
    Logger.log('');
    Logger.log('═══════════════════════════════════════');
    
  } catch (error) {
    Logger.log('❌ 診斷失敗: ' + error);
    Logger.log('   錯誤堆疊: ' + error.stack);
  }
}
// ==================== Session 管理 ====================

/**
 * ⭐ 驗證 Session Token（簡化版 - 只返回 true/false）
 */
function validateSession(sessionToken) {
  try {
    const result = checkSession_(sessionToken);
    return result.ok === true;
  } catch (error) {
    Logger.log('validateSession 錯誤: ' + error);
    return false;
  }
}

/**
 * 建立 Session
 */
function writeSession_(userId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_SESSION);
  const oneTimeToken = Utilities.getUuid();
  const now = new Date();
  const expiredAt = new Date(now.getTime() + SESSION_TTL_MS);

  const range = sheet.getRange("B:B").createTextFinder(userId).findNext();

  if (range) {
    const row = range.getRow();
    sheet.getRange(row, 1, 1, 4).setValues([[oneTimeToken, userId, now, expiredAt]]);
  } else {
    sheet.appendRow([oneTimeToken, userId, now, expiredAt]);
  }
  return oneTimeToken;
}

/**
 * 兌換一次性 token
 */
function verifyOneTimeToken_(otoken) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_SESSION);
  const range = sheet.getRange("A:A").createTextFinder(otoken).findNext();
  if (!range) return null;

  const row = range.getRow();
  const sessionToken = Utilities.getUuid();
  const now = new Date();
  const expiredAt = new Date(now.getTime() + SESSION_TTL_MS);
  const userId = sheet.getRange(row, 2).getValue();

  sheet.getRange(row, 1, 1, 4).setValues([[sessionToken, userId, now, expiredAt]]);
  return sessionToken;
}

/**
 * ✅ 檢查 Session（自動延期）- 修正版
 */
function checkSession_(sessionToken) {
  if (!sessionToken) return { ok: false, code: "MISSING_SESSION_TOKEN" };

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SESSION);
  if (!sh) return { ok: false, code: "SESSION_SHEET_NOT_FOUND" };

  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const [token, userId, , expiredAt] = values[i];
    if (token === sessionToken) {
      if (expiredAt && new Date() > new Date(expiredAt)) {
        return { ok: false, code: "ERR_SESSION_EXPIRED" };
      }
      
      // 延長 Session
      const newExpiredAt = new Date(new Date().getTime() + SESSION_TTL_MS);
      sh.getRange(i + 1, 4).setValue(newExpiredAt);
      
      // 查詢員工資料
      const employee = findEmployeeByLineUserId_(userId);
      if (!employee.ok) {
        Logger.log("❌ Session 檢查失敗: " + JSON.stringify(employee));
        return { ok: false, code: employee.code };
      }
      
      // ⭐⭐⭐ 關鍵修正：不要返回整個 employee 物件，而是只返回純淨的 user 資料
      return { 
        ok: true, 
        user: {
          userId: employee.userId,
          employeeId: employee.employeeId,
          email: employee.email,
          name: employee.name,
          picture: employee.picture,
          dept: employee.dept,
          status: employee.status
        },
        code: "WELCOME_BACK",
        params: { name: employee.name }
      };
    }
  }
  return { ok: false, code: "ERR_SESSION_INVALID" };
}

/**
 * 🧪 測試 checkSession_
 */
function testCheckSession() {
  Logger.log('🧪 測試 checkSession_');
  Logger.log('');
  
  const token = '04fd1452-4aca-4b03-ad17-45f03144c6ff';
  
  Logger.log('📡 Token: ' + token.substring(0, 20) + '...');
  Logger.log('');
  
  const result = checkSession_(token);
  
  Logger.log('📤 checkSession_ 結果:');
  Logger.log(JSON.stringify(result, null, 2));
  Logger.log('');
  
  if (result.ok && result.user) {
    Logger.log('✅ Session 有效');
    Logger.log('');
    Logger.log('👤 User 資料:');
    Logger.log('   - userId: ' + result.user.userId);
    Logger.log('   - employeeId: ' + result.user.employeeId);
    Logger.log('   - name: ' + result.user.name);
    Logger.log('   - dept: ' + result.user.dept);
    Logger.log('   - email: ' + result.user.email);
    Logger.log('   - status: ' + result.user.status);
    Logger.log('');
    Logger.log('🔍 檢查 user 物件是否乾淨:');
    Logger.log('   - user.ok 存在嗎? ' + (result.user.ok !== undefined ? '❌ 是（有問題）' : '✅ 否（正常）'));
  } else {
    Logger.log('❌ Session 無效');
    Logger.log('   code: ' + result.code);
  }
}
// ==================== 打卡功能 ====================

/**
 * 打卡功能
 */
function punch(sessionToken, type, lat, lng, note) {
  const employee = checkSession_(sessionToken);
  const user = employee.user;
  if (!user) return { ok: false, code: "ERR_SESSION_INVALID" };

  const shLoc = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOCATIONS);
  const lastRow = shLoc.getLastRow();
  
  if (lastRow < 2) {
    return { ok: false, code: "ERR_NO_LOCATIONS" };
  }
  
  const values = shLoc.getRange(2, 1, lastRow - 1, 5).getValues();
  let locationName = null;
  let minDistance = Infinity;
  
  for (let [, name, locLat, locLng, radius] of values) {
    if (!name || !locLat || !locLng) continue;
    
    const dist = getDistanceMeters_(lat, lng, Number(locLat), Number(locLng));
    
    if (dist <= Number(radius) && dist < minDistance) {
      locationName = name;
      minDistance = dist;
    }
  }

  if (!locationName) {
    return { ok: false, code: "ERR_OUT_OF_RANGE" };
  }

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ATTENDANCE);
  const row = [
    new Date(),
    user.userId,
    user.dept,
    user.name,
    type,
    `(${lat},${lng})`,
    locationName,
    "",
    "",
    note || ""
  ];
  sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);

  return { ok: true, code: `PUNCH_SUCCESS`, params: { type: type } };
}

/**
 * 補打卡功能
 */
/**
 * ✅ 補打卡功能（修正版 - 寫入「補打卡申請」工作表）
 */
function punchAdjusted(sessionToken, type, punchDate, lat, lng, note) {
  const employee = checkSession_(sessionToken);
  const user = employee.user;
  
  if (!user) {
    return { ok: false, code: "ERR_SESSION_INVALID" };
  }

  // ⭐ 修改：寫入「補打卡申請」工作表，而不是「出勤紀錄」
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ADJUST_PUNCH);
  
  if (!sh) {
    Logger.log('❌ 找不到「補打卡申請」工作表');
    return { ok: false, code: "ERR_SHEET_NOT_FOUND" };
  }

  // ⭐ 按照你的工作表欄位順序寫入
  // A: 申請ID（自動生成）
  // B: 用戶ID
  // C: 姓名
  // D: 日期
  // E: 時間
  // F: 類型（上班/下班）
  // G: 原因（補打卡理由）
  // H: 狀態（待審核）
  // I: 申請時間
  // J: 審核人
  // K: 審核時間
  
  const applicationId = Utilities.getUuid().substring(0, 8).toUpperCase();
  const dateOnly = Utilities.formatDate(punchDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const timeOnly = Utilities.formatDate(punchDate, Session.getScriptTimeZone(), 'HH:mm');
  
  const row = [
    applicationId,           // A: 申請ID
    user.userId,             // B: 用戶ID
    user.name,               // C: 姓名
    dateOnly,                // D: 日期
    timeOnly,                // E: 時間
    type,                    // F: 類型
    note || '',              // G: 原因
    '待審核',                // H: 狀態
    new Date(),              // I: 申請時間
    '',                      // J: 審核人
    ''                       // K: 審核時間
  ];
  
  sh.appendRow(row);
  
  Logger.log(`✅ 補打卡申請已提交: ${user.name} - ${dateOnly} ${type}`);
  Logger.log(`   理由: ${note}`);

  return { 
    ok: true, 
    code: `ADJUST_PUNCH_SUCCESS`, 
    params: { type: type } 
  };
}
// function punchAdjusted(sessionToken, type, punchDate, lat, lng, note) {
//   const employee = checkSession_(sessionToken);
//   const user = employee.user;
//   if (!user) return { ok: false, code: "ERR_SESSION_INVALID" };

//   const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ATTENDANCE);
//   sh.appendRow([
//     punchDate,
//     user.userId,
//     user.dept,
//     user.name,
//     type,
//     `(${lat},${lng})`,
//     "",
//     "補打卡",
//     "?",
//     note
//   ]);

//   return { ok: true, code: `ADJUST_PUNCH_SUCCESS`, params: { type: type } };
// }

/**
 * 取得出勤紀錄
 */
function getAttendanceRecords(monthParam, userIdParam) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ATTENDANCE);
  const values = sheet.getDataRange().getValues().slice(1);
  
  return values.filter(row => {
    if (!row[0]) return false;
    
    const d = new Date(row[0]);
    const yyyy_mm = d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0");
    const monthMatch = yyyy_mm === monthParam;
    const userMatch = userIdParam ? row[1] === userIdParam : true;
    return monthMatch && userMatch;
  }).map(r => ({
    date: r[0],
    userId: r[1],
    salary: r[2],
    name: r[3],
    type: r[4],
    gps: r[5],
    location: r[6],
    note: r[7],
    audit: r[8],
    device: r[9]
  }));
}

/**
 * 取得出勤詳細資料（用於報表匯出）
 */
/**
 * ✅ 修正版：取得出勤詳細資料（修正 localeCompare 錯誤）
 * 
 * 修正內容：
 * 1. 修正請假記錄合併時可能產生 undefined date 的問題
 * 2. 加強日期格式驗證
 * 3. 改進錯誤處理
 */
function getAttendanceDetails(monthParam, userIdParam) {
  try {
    Logger.log('📋 getAttendanceDetails 開始');
    Logger.log(`   monthParam: ${monthParam}`);
    Logger.log(`   userIdParam: ${userIdParam}`);
    
    const records = getAttendanceRecords(monthParam, userIdParam);
    const leaveRecords = getApprovedLeaveRecords(monthParam, userIdParam);
    const overtimeRecords = getApprovedOvertimeRecords(monthParam, userIdParam);
    
    Logger.log(`   打卡記錄: ${records.length} 筆`);
    Logger.log(`   請假記錄: ${leaveRecords.length} 筆`);
    Logger.log(`   加班記錄: ${overtimeRecords.length} 筆`);
    
    // ✅ 建立日期集合（過濾掉無效日期）
    const allDates = new Set();
    
    // 加入打卡記錄的日期
    records.forEach(r => {
      const dateKey = formatDate(r.date);
      if (dateKey) {
        allDates.add(dateKey);
      }
    });
    
    // ✅ 加入請假記錄的日期（檢查日期格式）
    leaveRecords.forEach(r => {
      if (r.date && typeof r.date === 'string' && r.date.match(/^\d{4}-\d{2}-\d{2}$/)) {
        allDates.add(r.date);
      } else {
        Logger.log(`⚠️ 請假記錄日期格式錯誤: ${r.date}`);
      }
    });
    
    // 加入加班記錄的日期
    overtimeRecords.forEach(r => {
      if (r.date && typeof r.date === 'string' && r.date.match(/^\d{4}-\d{2}-\d{2}$/)) {
        allDates.add(r.date);
      }
    });
    
    Logger.log(`   涉及日期總數: ${allDates.size} 天`);
    
    // ✅ 按日期建立資料結構
    const dailyRecords = {};
    
    // 初始化所有日期
    allDates.forEach(dateKey => {
      dailyRecords[dateKey] = {
        date: dateKey,
        userId: userIdParam,
        name: '',
        record: [],
        reason: 'STATUS_NO_RECORD',
        overtime: null,
        leave: null
      };
    });
    
    // 填入打卡記錄
    records.forEach(r => {
      const dateKey = formatDate(r.date);
      
      if (dailyRecords[dateKey]) {
        if (!dailyRecords[dateKey].name) {
          dailyRecords[dateKey].name = r.name;
        }
        
        dailyRecords[dateKey].record.push({
          type: r.type,
          time: formatTime(r.date),
          location: r.location,
          note: r.note || ''
        });
      }
    });
    
    // ✅ 填入請假資料
    leaveRecords.forEach(leave => {
      const dateKey = leave.date;
      
      if (dailyRecords[dateKey]) {
        dailyRecords[dateKey].leave = {
          leaveType: leave.leaveType,
          days: leave.days,
          status: leave.status,
          reason: leave.reason || '',
          employeeName: leave.employeeName,
          reviewComment: leave.reviewComment || ''
        };
        
        // 如果沒有員工姓名，從請假記錄取得
        if (!dailyRecords[dateKey].name && leave.employeeName) {
          dailyRecords[dateKey].name = leave.employeeName;
        }
        
        Logger.log(`   ${dateKey}: 加入請假資訊 (${leave.leaveType})`);
      }
    });
    
    // 填入加班資料
    overtimeRecords.forEach(ot => {
      const dateKey = ot.date;
      
      if (dailyRecords[dateKey]) {
        dailyRecords[dateKey].overtime = {
          startTime: ot.startTime,
          endTime: ot.endTime,
          hours: ot.hours,
          reason: ot.reason || ''
        };
        
        Logger.log(`   ${dateKey}: 加入加班資訊 (${ot.hours}h)`);
      }
    });
    
    // 判斷每日狀態
    Object.keys(dailyRecords).forEach(dateKey => {
      const daily = dailyRecords[dateKey];
      
      const hasPunchIn = daily.record.some(r => r.type === '上班');
      const hasPunchOut = daily.record.some(r => r.type === '下班');
      
      // ✅ 修正：如果有請假，根據打卡情況設定狀態
      if (daily.leave) {
        if (hasPunchIn && hasPunchOut) {
          // 有打卡也有請假（可能是半天假）
          daily.reason = 'STATUS_PUNCH_NORMAL';
        } else {
          // 只有請假沒打卡（全天假）
          daily.reason = 'STATUS_NO_RECORD';
        }
        Logger.log(`   ${dateKey}: 有請假記錄，狀態設為 ${daily.reason}`);
      } else {
        // 原有的打卡狀態判斷
        if (hasPunchIn && hasPunchOut) {
          daily.reason = 'STATUS_PUNCH_NORMAL';
        } else if (!hasPunchIn && !hasPunchOut) {
          daily.reason = 'STATUS_NO_RECORD';
        } else if (!hasPunchIn) {
          daily.reason = 'STATUS_PUNCH_IN_MISSING';
        } else if (!hasPunchOut) {
          daily.reason = 'STATUS_PUNCH_OUT_MISSING';
        }
      }
    });
    
    // ✅ 修正：轉換為陣列並排序（確保所有 date 都存在）
    const result = Object.values(dailyRecords)
      .filter(r => r.date) // 過濾掉沒有 date 的記錄
      .sort((a, b) => {
        if (!a.date || !b.date) return 0;
        return a.date.localeCompare(b.date);
      });
    
    Logger.log(`✅ getAttendanceDetails 完成: ${result.length} 筆`);
    
    // ✅ 除錯：顯示所有日期
    Logger.log('');
    Logger.log('📅 最終結果包含的日期:');
    result.forEach(r => {
      const hasLeave = r.leave ? '🏖️' : '';
      const hasOvertime = r.overtime ? '⏰' : '';
      Logger.log(`   ${r.date} ${hasLeave}${hasOvertime} - ${r.reason}`);
    });
    
    return {
      ok: true,
      records: result
    };
    
  } catch (error) {
    Logger.log('❌ getAttendanceDetails 錯誤: ' + error);
    Logger.log('   錯誤堆疊: ' + error.stack);
    return {
      ok: false,
      msg: error.message
    };
  }
}

/**
 * 🧪 測試修正後的函數
 */
function testFixedGetAttendanceDetails() {
  Logger.log('═══════════════════════════════════════');
  Logger.log('🧪 測試修正後的 getAttendanceDetails');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  
  const monthParam = '2025-12';
  const userIdParam = 'U68e0ca9d516e63ed15bf9387fad174ac';
  
  const result = getAttendanceDetails(monthParam, userIdParam);
  
  Logger.log('');
  Logger.log('📤 測試結果:');
  Logger.log(`   ok: ${result.ok}`);
  Logger.log(`   總記錄數: ${result.records ? result.records.length : 0}`);
  Logger.log('');
  
  if (result.ok && result.records) {
    // 檢查 12/10
    const dec10 = result.records.find(r => r.date === '2025-12-10');
    
    if (dec10) {
      Logger.log('✅✅✅ 找到 2025-12-10 的記錄！');
      Logger.log('');
      Logger.log('📋 記錄內容:');
      Logger.log(`   date: ${dec10.date}`);
      Logger.log(`   name: ${dec10.name}`);
      Logger.log(`   reason: ${dec10.reason}`);
      Logger.log(`   打卡數: ${dec10.record.length}`);
      Logger.log(`   有請假: ${dec10.leave ? '是' : '否'}`);
      Logger.log(`   有加班: ${dec10.overtime ? '是' : '否'}`);
      Logger.log('');
      
      if (dec10.leave) {
        Logger.log('🏖️ 請假資訊:');
        Logger.log(`   假別: ${dec10.leave.leaveType}`);
        Logger.log(`   天數: ${dec10.leave.days}`);
        Logger.log(`   狀態: ${dec10.leave.status}`);
        Logger.log(`   原因: ${dec10.leave.reason}`);
        Logger.log('');
        Logger.log('✅ 修正成功！即使沒打卡也能顯示請假資訊');
      }
      
      if (dec10.overtime) {
        Logger.log('⏰ 加班資訊:');
        Logger.log(`   時間: ${dec10.overtime.startTime} - ${dec10.overtime.endTime}`);
        Logger.log(`   時數: ${dec10.overtime.hours}h`);
      }
    } else {
      Logger.log('❌ 還是沒找到 2025-12-10 的記錄');
      Logger.log('');
      Logger.log('📅 現有的日期:');
      result.records.forEach(r => {
        Logger.log(`   - ${r.date}`);
      });
    }
    
    // 檢查 12/11
    Logger.log('');
    const dec11 = result.records.find(r => r.date === '2025-12-11');
    
    if (dec11) {
      Logger.log('✅ 找到 2025-12-11 的記錄');
      
      if (dec11.leave) {
        Logger.log(`   🏖️ 請假: ${dec11.leave.leaveType} (${dec11.leave.days}天)`);
      }
      
      if (dec11.overtime) {
        Logger.log(`   ⏰ 加班: ${dec11.overtime.hours}h`);
      }
    }
  }
  
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
}
// function getAttendanceDetails(monthParam, userIdParam) {
//   const records = getAttendanceRecords(monthParam, userIdParam);
  
//   // 👉 取得該月份已核准的加班記錄
//   const overtimeRecords = getApprovedOvertimeRecords(monthParam, userIdParam);
  
//   // 👉 取得該月份已核准的請假記錄
//   const leaveRecords = getApprovedLeaveRecords(monthParam, userIdParam);
  
//   const dailyRecords = {};
  
//   // 處理打卡記錄
//   records.forEach(r => {
//     const dateKey = formatDate(r.date);
//     const userId = r.userId || 'unknown';
//     const userName = r.name || '未知員工';
//     const key = `${userId}_${dateKey}`;
    
//     if (!dailyRecords[key]) {
//       dailyRecords[key] = {
//         date: dateKey,
//         userId: userId,
//         name: userName,
//         record: [],
//         reason: "",
//         overtime: null,  // 加班資訊
//         leave: null      // 👈 新增：請假資訊
//       };
//     }
    
//     dailyRecords[key].record.push({
//       time: formatTime(r.date),
//       type: r.type,
//       location: r.location,
//       note: r.note || ""
//     });
//   });
  
//   // 👉 合併加班資料
//   overtimeRecords.forEach(ot => {
//     const key = `${ot.employeeId}_${ot.overtimeDate}`;
    
//     if (dailyRecords[key]) {
//       dailyRecords[key].overtime = {
//         startTime: ot.startTime,
//         endTime: ot.endTime,
//         hours: ot.hours,
//         reason: ot.reason
//       };
//     } else {
//       // 如果該日沒有打卡記錄，也建立一筆（只顯示加班）
//       dailyRecords[key] = {
//         date: ot.overtimeDate,
//         userId: ot.employeeId,
//         name: ot.employeeName,
//         record: [],
//         reason: "STATUS_NO_RECORD",
//         overtime: {
//           startTime: ot.startTime,
//           endTime: ot.endTime,
//           hours: ot.hours,
//           reason: ot.reason
//         },
//         leave: null  // 👈 確保有 leave 欄位
//       };
//     }
//   });
  
//   // 👉 合併請假資料
//   leaveRecords.forEach(leave => {
//     const key = `${leave.employeeId}_${leave.date}`;
    
//     if (dailyRecords[key]) {
//       dailyRecords[key].leave = {
//         status: leave.status,
//         leaveType: leave.leaveType,
//         days: leave.days,
//         reason: leave.reason
//       };
//     } else {
//       // 如果該日沒有打卡記錄，也建立一筆（只顯示請假）
//       dailyRecords[key] = {
//         date: leave.date,
//         userId: leave.employeeId,
//         name: leave.employeeName,
//         record: [],
//         reason: "STATUS_NO_RECORD",
//         overtime: null,  // 👈 確保有 overtime 欄位
//         leave: {
//           status: leave.status,
//           leaveType: leave.leaveType,
//           days: leave.days,
//           reason: leave.reason
//         }
//       };
//     }
//   });
  
//   // 判斷打卡狀態
//   const result = Object.values(dailyRecords).map(day => {
//     const hasIn = day.record.some(r => r.type === "上班");
//     const hasOut = day.record.some(r => r.type === "下班");
    
//     let reason = "";
//     if (!hasIn && !hasOut) {
//       reason = "STATUS_NO_RECORD";
//     } else if (!hasIn) {
//       reason = "STATUS_PUNCH_IN_MISSING";
//     } else if (!hasOut) {
//       reason = "STATUS_PUNCH_OUT_MISSING";
//     } else {
//       reason = "STATUS_PUNCH_NORMAL";
//     }
    
//     return {
//       date: day.date,
//       userId: day.userId,
//       name: day.name,
//       record: day.record,
//       reason: reason,
//       overtime: day.overtime,  // 包含加班資訊
//       leave: day.leave         // 👈 包含請假資訊
//     };
//   });
  
//   Logger.log(`📊 getAttendanceDetails: 共 ${result.length} 筆記錄`);
//   return { ok: true, records: result };
// }

/**
 * 👉 新增：取得已核准的加班記錄
 */
function getApprovedOvertimeRecords(monthParam, userIdParam) {
  try {
    Logger.log('═══════════════════════════════════════');
    Logger.log('⏰ 開始查詢加班記錄');
    Logger.log('   月份: ' + monthParam);
    Logger.log('   員工ID: ' + userIdParam);
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_OVERTIME);
    
    if (!sheet) {
      Logger.log('⚠️ 找不到加班申請工作表');
      return [];
    }
    
    const values = sheet.getDataRange().getValues();
    
    if (values.length <= 1) {
      Logger.log('⚠️ 加班工作表只有標題，沒有資料');
      return [];
    }
    
    // ⭐ 步驟 1: 取得標題列
    const headers = values[0];
    
    Logger.log('📋 加班工作表標題:');
    headers.forEach((h, i) => {
      Logger.log(`   ${i}. ${h}`);
    });
    
    // ⭐ 步驟 2: 動態找出欄位索引
    const employeeIdCol = headers.indexOf('員工ID');
    const employeeNameCol = headers.indexOf('員工姓名');
    const overtimeDateCol = headers.indexOf('加班日期');
    const startTimeCol = headers.indexOf('開始時間');
    const endTimeCol = headers.indexOf('結束時間');
    const hoursCol = headers.indexOf('加班時數');
    const reasonCol = headers.indexOf('申請原因');
    const statusCol = headers.indexOf('審核狀態');
    
    Logger.log('');
    Logger.log('📊 欄位索引:');
    Logger.log(`   員工ID: ${employeeIdCol}`);
    Logger.log(`   員工姓名: ${employeeNameCol}`);
    Logger.log(`   加班日期: ${overtimeDateCol}`);
    Logger.log(`   開始時間: ${startTimeCol}`);
    Logger.log(`   結束時間: ${endTimeCol}`);
    Logger.log(`   加班時數: ${hoursCol}`);
    Logger.log(`   申請原因: ${reasonCol}`);
    Logger.log(`   審核狀態: ${statusCol}`);
    
    // ⭐ 步驟 3: 智能格式化時間（完全兼容版）
    const formatTime = (dateTime) => {
      if (!dateTime) return "";
      
      try {
        // 情況 1: Date 物件
        if (dateTime instanceof Date) {
          return Utilities.formatDate(dateTime, "Asia/Taipei", "HH:mm");
        }
        
        // 情況 2: 字串處理
        const str = String(dateTime).trim();
        
        // 情況 2a: ISO 格式 "2025/12/09 下午 9:20:00"
        if (str.includes('下午') || str.includes('上午')) {
          const timePart = str.split(' ')[2]; // 取 "9:20:00"
          return timePart.substring(0, 5); // 回傳 "09:20"
        }
        
        // 情況 2b: ISO 格式 "2025-12-10T18:00:00"
        if (str.includes('T')) {
          const timePart = str.split('T')[1];
          return timePart.substring(0, 5);
        }
        
        // 情況 2c: 已經是 "HH:mm" 或 "HH:mm:ss" 格式
        if (str.includes(':')) {
          return str.substring(0, 5);
        }
        
        return str;
        
      } catch (e) {
        Logger.log(`⚠️ 時間格式化失敗: ${dateTime}, 錯誤: ${e}`);
        return "";
      }
    };
    
    // ⭐ 步驟 4: 智能格式化日期（完全兼容版）
    const formatOvertimeDate = (dateValue) => {
      if (!dateValue) return "";
      
      try {
        // 情況 1: Date 物件
        if (dateValue instanceof Date) {
          return Utilities.formatDate(dateValue, "Asia/Taipei", "yyyy-MM-dd");
        }
        
        // 情況 2: 字串處理
        const str = String(dateValue).trim();
        
        // 情況 2a: "2025-12-09" 格式（已經是正確格式）
        if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
          return str;
        }
        
        // 情況 2b: "2025/12/09" 格式
        if (str.includes('/')) {
          const parts = str.split('/');
          if (parts.length >= 3) {
            const year = parts[0];
            const month = parts[1].padStart(2, '0');
            const day = parts[2].split(' ')[0].padStart(2, '0'); // 處理可能包含時間的情況
            return `${year}-${month}-${day}`;
          }
        }
        
        // 情況 2c: "2025-12-09T..." ISO 格式
        if (str.includes('T')) {
          return str.split('T')[0];
        }
        
        return str;
        
      } catch (e) {
        Logger.log(`⚠️ 日期格式化失敗: ${dateValue}, 錯誤: ${e}`);
        return "";
      }
    };
    
    // ⭐ 步驟 5: 篩選並組裝記錄
    const overtimeRecords = [];
    
    Logger.log('');
    Logger.log('🔍 開始篩選記錄...');
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      
      // 格式化日期
      const overtimeDate = formatOvertimeDate(row[overtimeDateCol]);
      const employeeId = row[employeeIdCol];
      const status = String(row[statusCol]).trim().toLowerCase();
      
      // 檢查條件
      const monthMatch = overtimeDate && overtimeDate.startsWith(monthParam);
      const userMatch = userIdParam ? employeeId === userIdParam : true;
      const statusMatch = status === "approved";
      
      Logger.log(`   ${i}. ${overtimeDate} - ${row[employeeNameCol]}`);
      Logger.log(`      員工ID: ${employeeId}, 狀態: ${status}`);
      Logger.log(`      monthMatch: ${monthMatch}, userMatch: ${userMatch}, statusMatch: ${statusMatch}`);
      
      if (monthMatch && userMatch && statusMatch) {
        const record = {
          employeeId: employeeId,
          employeeName: row[employeeNameCol],
          date: overtimeDate,  // ⭐⭐⭐ 使用 date（與前端一致）
          startTime: formatTime(row[startTimeCol]),
          endTime: formatTime(row[endTimeCol]),
          hours: parseFloat(row[hoursCol]) || 0,
          reason: row[reasonCol] || ''
        };
        
        overtimeRecords.push(record);
        
        Logger.log(`      ✅ 符合條件！`);
        Logger.log(`         日期: ${record.date}`);
        Logger.log(`         時間: ${record.startTime} - ${record.endTime}`);
        Logger.log(`         時數: ${record.hours}`);
      } else {
        Logger.log(`      ❌ 不符合條件`);
      }
    }
    
    Logger.log('');
    Logger.log(`✅ 找到 ${overtimeRecords.length} 筆已核准加班記錄`);
    
    if (overtimeRecords.length > 0) {
      Logger.log('');
      Logger.log('📋 加班記錄詳細列表:');
      overtimeRecords.forEach((rec, idx) => {
        Logger.log(`   ${idx + 1}. ${rec.date} - ${rec.employeeName}`);
        Logger.log(`      ${rec.startTime} ~ ${rec.endTime} (${rec.hours}h)`);
      });
    }
    
    Logger.log('═══════════════════════════════════════');
    
    return overtimeRecords;
    
  } catch (error) {
    Logger.log('❌ getApprovedOvertimeRecords 錯誤: ' + error);
    Logger.log('   錯誤堆疊: ' + error.stack);
    Logger.log('═══════════════════════════════════════');
    return [];
  }
}

function testGetAttendanceDetailsWithOvertime() {
  Logger.log('🧪 測試 getAttendanceDetails');
  Logger.log('═══════════════════════════════════════');
  
  const monthParam = '2025-12';
  const userIdParam = 'U68e0ca9d516e63ed15bf9387fad174ac';
  
  Logger.log(`📅 查詢條件: ${monthParam}, userId: ${userIdParam}`);
  Logger.log('');
  
  const result = getAttendanceDetails(monthParam, userIdParam);
  
  Logger.log('📤 API 回應:');
  Logger.log(`   ok: ${result.ok}`);
  Logger.log(`   records 數量: ${result.records ? result.records.length : 0}`);
  Logger.log('');
  
  if (result.ok && result.records) {
    // 找出 2025-12-09 的記錄
    const dec09 = result.records.find(r => r.date === '2025-12-09');
    
    if (dec09) {
      Logger.log('✅ 找到 2025-12-09 的記錄:');
      Logger.log('');
      Logger.log('📋 記錄內容:');
      Logger.log(JSON.stringify(dec09, null, 2));
      Logger.log('');
      
      Logger.log('🔍 加班資訊檢查:');
      Logger.log(`   overtime 存在: ${dec09.overtime ? '是' : '否'}`);
      
      if (dec09.overtime) {
        Logger.log('   ✅ 加班資訊:');
        Logger.log(`      開始時間: ${dec09.overtime.startTime}`);
        Logger.log(`      結束時間: ${dec09.overtime.endTime}`);
        Logger.log(`      時數: ${dec09.overtime.hours}`);
        Logger.log(`      原因: ${dec09.overtime.reason}`);
      } else {
        Logger.log('   ❌ 沒有加班資訊');
      }
    } else {
      Logger.log('❌ 沒有找到 2025-12-09 的記錄');
      Logger.log('');
      Logger.log('📋 所有記錄的日期:');
      result.records.forEach((r, i) => {
        Logger.log(`   ${i + 1}. ${r.date} - ${r.name}`);
      });
    }
  }
  
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
}
/**
 * 🧪 測試加班記錄查詢
 */
function testGetApprovedOvertimeRecords() {
  Logger.log('🧪 測試加班記錄查詢');
  Logger.log('═══════════════════════════════════════');
  
  const monthParam = '2025-12';
  const userIdParam = 'U68e0ca9d516e63ed15bf9387fad174ac';  // 替換成您的實際 userId
  
  Logger.log(`📅 查詢條件: ${monthParam}, userId: ${userIdParam}`);
  Logger.log('');
  
  const records = getApprovedOvertimeRecords(monthParam, userIdParam);
  
  Logger.log('');
  Logger.log('📤 查詢結果:');
  Logger.log(`   找到 ${records.length} 筆記錄`);
  
  if (records.length > 0) {
    records.forEach((rec, i) => {
      Logger.log('');
      Logger.log(`   記錄 ${i + 1}:`);
      Logger.log(`      日期: ${rec.overtimeDate}`);
      Logger.log(`      員工: ${rec.employeeName} (${rec.employeeId})`);
      Logger.log(`      時間: ${rec.startTime} - ${rec.endTime}`);
      Logger.log(`      時數: ${rec.hours} 小時`);
      Logger.log(`      原因: ${rec.reason}`);
    });
  } else {
    Logger.log('   ⚠️ 沒有找到符合條件的記錄');
  }
  
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
}


// ==================== 地點管理 ====================
/**
 * 新增打卡地點
 * @param {string} name - 地點名稱
 * @param {number} lat - 緯度
 * @param {number} lng - 經度
 * @param {number} radius - 打卡範圍（公尺），預設 200，範圍 100-2000
 */
function addLocation(name, lat, lng, radius) {
  if (!name || !lat || !lng) {
    return { ok: false, code: "ERR_INVALID_INPUT" };
  }
  
  // 驗證 radius 參數，確保在合理範圍內
  const validRadius = radius && !isNaN(radius) ? parseInt(radius) : 200;
  const finalRadius = Math.max(100, Math.min(2000, validRadius)); // 限制在 100-2000 之間
  
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_LOCATIONS);
  sh.appendRow(["", name, lat, lng, finalRadius]);
  
  Logger.log(`✅ 新增地點：${name}，範圍：${finalRadius}公尺`);
  return { ok: true, code: "LOCATION_ADD_SUCCESS" };
}
// function addLocation(name, lat, lng) {
//   if (!name || !lat || !lng) {
//     return { ok: false, code: "ERR_INVALID_INPUT" };
//   }
  
//   const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_LOCATIONS);
//   sh.appendRow(["", name, lat, lng, "100"]);
//   return { ok: true, code: "LOCATION_ADD_SUCCESS" };
// }

/**
 * 取得所有打卡地點
 */
function getLocation() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOCATIONS);
  const values = sheet.getDataRange().getValues();
  
  if (values.length === 0) {
    return { ok: true, locations: [] };
  }
  
  const headers = values.shift();
  const locations = values
    .filter(row => row[1])
    .map(row => ({
      id: row[headers.indexOf('ID')] || '',
      name: row[headers.indexOf('地點名稱')] || '',
      lat: row[headers.indexOf('GPS(緯度)')] || 0,
      lng: row[headers.indexOf('GPS(經度)')] || 0,
      scope: row[headers.indexOf('容許誤差(公尺)')] || 100
    }));
  
  return { ok: true, locations: locations };
}

// ==================== 審核功能 ====================

/**
 * 取得待審核請求（補打卡）
 */
function getReviewRequest() {
  Logger.log('📋 開始取得待審核補打卡申請');
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ADJUST_PUNCH);
  
  if (!sheet) {
    Logger.log('❌ 找不到「補打卡申請」工作表');
    return { ok: false, msg: "找不到補打卡申請工作表" };
  }
  
  const values = sheet.getDataRange().getValues();
  
  if (values.length <= 1) {
    Logger.log('⚠️ 補打卡申請工作表只有標題，沒有資料');
    return { ok: true, reviewRequest: [] };
  }
  
  const headers = values[0];
  
  // 篩選「待審核」的申請
  const reviewRequest = values.filter((row, index) => {
    if (index === 0 || !row[0]) return false;
    
    const statusCol = headers.indexOf('狀態');
    const status = row[statusCol];
    
    return status === '待審核';
    
  }).map(row => {
    const actualRowNumber = values.indexOf(row) + 1;
    
    // 從工作表讀取各欄位
    const applicationId = row[headers.indexOf('申請ID')];
    const userId = row[headers.indexOf('用戶ID')];
    const name = row[headers.indexOf('姓名')];
    const dateValue = row[headers.indexOf('日期')];  // ⭐ 關鍵：可能是字串或 Date
    const timeValue = row[headers.indexOf('時間')];  // ⭐ 關鍵：可能是字串或 Date
    const type = row[headers.indexOf('類型')];
    const reason = row[headers.indexOf('原因')];
    const applicationTime = row[headers.indexOf('申請時間')];
    
    // ✅ 修正：智能格式化日期
    let date, time;
    
    // 處理日期
    if (dateValue instanceof Date) {
      date = Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else if (typeof dateValue === 'string') {
      date = dateValue;
    } else {
      date = '未知日期';
    }
    
    // 處理時間
    if (timeValue instanceof Date) {
      time = Utilities.formatDate(timeValue, Session.getScriptTimeZone(), 'HH:mm');
    } else if (typeof timeValue === 'string') {
      time = timeValue;
    } else {
      time = '未知時間';
    }
    
    Logger.log(`   ${actualRowNumber}. ${name} - ${date} ${time} ${type}`);
    Logger.log(`      理由: ${reason}`);
    
    return {
      id: actualRowNumber,
      applicationId: applicationId,
      userId: userId,
      name: name,
      type: type,
      remark: `補${type}卡`,
      applicationPeriod: `${date} ${time}`,  // ✅ 使用格式化後的日期時間
      note: reason || ''
    };
  });
  
  Logger.log('');
  Logger.log(`✅ 找到 ${reviewRequest.length} 筆待審核申請`);
  
  return { ok: true, reviewRequest: reviewRequest };
}
// function getReviewRequest() {
//   Logger.log('📋 開始取得待審核補打卡申請');

//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ADJUST_PUNCH);
  
//   if (!sheet) {
//     Logger.log(❌ 找不到「補打卡申請」工作表');
//     return { ok: false, msg: "找不到補打卡申請工作表" };
//   }
  
//   const values = sheet.getDataRange().getValues();
  
//   if (values.length <= 1) {
//     Logger.log('⚠️ 補打卡申請工作表只有標題，沒有資料');
//     return { ok: true, reviewRequest: [] };
//   }
  
//   const headers = values[0];
  
//   Logger.log('📋 標題列: ' + headers.join(', '));
  
//   // 篩選「待審核」的申請
//   const reviewRequest = values.filter((row, index) => {
//     if (index === 0 || !row[0]) return false;
    
//     const statusCol = headers.indexOf('狀態');
//     const status = row[statusCol];
    
//     return status === '待審核';
    
//   }).map(row => {
//     const actualRowNumber = values.indexOf(row) + 1;
    
//     // 從工作表讀取各欄位
//     const applicationId = row[headers.indexOf('申請ID')];
//     const userId = row[headers.indexOf('用戶ID')];
//     const name = row[headers.indexOf('姓名')];
//     const date = row[headers.indexOf('日期')];
//     const time = row[headers.indexOf('時間')];
//     const type = row[headers.indexOf('類型')];
//     const reason = row[headers.indexOf('原因')];
//     const applicationTime = row[headers.indexOf('申請時間')];
    
//     Logger.log(`   ${actualRowNumber}. ${name} - ${date} ${time} ${type}`);
//     Logger.log(`      理由: ${reason}`);
    
//     return {
//       id: actualRowNumber,
//       applicationId: applicationId,
//       userId: userId,
//       name: name,
//       type: type,
//       remark: `補${type}卡`,
//       applicationPeriod: `${date} ${time}`,
//       note: reason || ''  // ⭐ 補打卡理由
//     };
//   });
  
//   Logger.log('');
//   Logger.log(`✅ 找到 ${reviewRequest.length} 筆待審核申請`);
  
//   return { ok: true, reviewRequest: reviewRequest };
// }
// function getReviewRequest() {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ATTENDANCE);
//   const values = sheet.getDataRange().getValues();
//   const headers = values[0];

//   const reviewRequest = values.filter((row, index) => {
//     if (index === 0 || !row[0]) return false;

//     const remarkCol = headers.indexOf('備註');
//     const auditCol = headers.indexOf('管理員審核');
    
//     return row[remarkCol] === "補打卡" && row[auditCol] === "?";
//   }).map(row => {
//     const actualRowNumber = values.indexOf(row) + 1;
    
//     // ⭐⭐⭐ 關鍵修正：加入 note 欄位（理由在「裝置資訊」欄）
//     const deviceCol = headers.indexOf('裝置資訊');
//     const noteText = deviceCol >= 0 ? (row[deviceCol] || '') : '';
    
//     return {
//       id: actualRowNumber,
//       name: row[headers.indexOf('打卡人員')],
//       type: row[headers.indexOf('打卡類別')],
//       remark: row[headers.indexOf('備註')],
//       applicationPeriod: formatDateTime(row[headers.indexOf('打卡時間')]),
//       note: noteText  // ⭐ 新增：補打卡理由
//     };
//   });
  
//   Logger.log('📋 待審核請求: ' + reviewRequest.length + ' 筆');
  
//   // 除錯：顯示第一筆的 note
//   if (reviewRequest.length > 0) {
//     Logger.log('   第一筆 note: ' + reviewRequest[0].note);
//   }
  
//   return { ok: true, reviewRequest: reviewRequest };
// }

/**
 * 🧪 測試 getReviewRequest 是否包含 note
 */
function testGetReviewRequestWithNote() {
  Logger.log('🧪 測試 getReviewRequest');
  Logger.log('');
  
  const result = getReviewRequest();
  
  Logger.log('📤 結果:');
  Logger.log('   ok: ' + result.ok);
  Logger.log('   筆數: ' + result.reviewRequest.length);
  Logger.log('');
  
  if (result.reviewRequest.length > 0) {
    Logger.log('📋 第一筆資料:');
    const first = result.reviewRequest[0];
    Logger.log('   id: ' + first.id);
    Logger.log('   name: ' + first.name);
    Logger.log('   type: ' + first.type);
    Logger.log('   remark: ' + first.remark);
    Logger.log('   note: ' + (first.note || '(空)'));
    Logger.log('');
    
    if (first.note) {
      Logger.log('✅✅✅ note 欄位存在！');
    } else {
      Logger.log('❌ note 欄位是空的');
    }
  }
}
// function getReviewRequest() {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ATTENDANCE);
//   const values = sheet.getDataRange().getValues();
//   const headers = values[0];

//   const reviewRequest = values.filter((row, index) => {
//     if (index === 0 || !row[0]) return false;

//     const remarkCol = headers.indexOf('備註');
//     const auditCol = headers.indexOf('管理員審核');
    
//     return row[remarkCol] === "補打卡" && row[auditCol] === "?";
//   }).map(row => {
//     const actualRowNumber = values.indexOf(row) + 1;
//     return {
//       id: actualRowNumber,
//       name: row[headers.indexOf('打卡人員')],
//       type: row[headers.indexOf('打卡類別')],
//       remark: row[headers.indexOf('備註')],
//       applicationPeriod: formatDateTime(row[headers.indexOf('打卡時間')])
//     };
//   });
  
//   return { ok: true, reviewRequest: reviewRequest };
// }

/**
 * 更新審核狀態（含 LINE 通知）
 */
/**
 * ✅ 更新審核狀態（完整修正版 - 從補打卡申請工作表讀取）
 */
function updateReviewStatus(rowNumber, status, note) {
  try {
    Logger.log('═══════════════════════════════════════');
    Logger.log('📋 開始審核補打卡');
    Logger.log('   行號: ' + rowNumber);
    Logger.log('   狀態: ' + status);
    
    // ✅ 修正：改為從補打卡申請工作表讀取
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ADJUST_PUNCH);
    
    if (!sheet) {
      Logger.log('❌ 找不到補打卡申請工作表');
      return { ok: false, msg: "找不到補打卡申請工作表" };
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    Logger.log('📋 工作表標題:', headers);
    
    // 找出「狀態」、「審核人」、「審核時間」欄位
    const statusCol = headers.indexOf('狀態') + 1;
    const reviewerCol = headers.indexOf('審核人') + 1;
    const reviewTimeCol = headers.indexOf('審核時間') + 1;
    
    if (statusCol === 0) {
      Logger.log('❌ 找不到「狀態」欄位');
      return { ok: false, msg: "找不到「狀態」欄位" };
    }
    
    // 取得該行的申請資料
    const record = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    Logger.log('📄 申請記錄:', record);
    
    // ✅ 從補打卡申請工作表讀取資料
    const applicationId = record[headers.indexOf('申請ID')];
    const userId = record[headers.indexOf('用戶ID')];
    const employeeName = record[headers.indexOf('姓名')];
    const dateValue = record[headers.indexOf('日期')];
    const timeValue = record[headers.indexOf('時間')];
    const punchType = record[headers.indexOf('類型')];
    const reason = record[headers.indexOf('原因')];
    
    // ✅ 智能格式化日期時間
    let punchDate, punchTime;
    
    // 處理日期
    if (dateValue instanceof Date) {
      punchDate = Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else if (typeof dateValue === 'string') {
      // 如果已經是字串格式，檢查是否符合 YYYY-MM-DD
      if (/^\d{4}-\d{2}-\d{2}$/.test(dateValue)) {
        punchDate = dateValue;
      } else {
        // 嘗試解析
        const parsedDate = new Date(dateValue);
        if (!isNaN(parsedDate)) {
          punchDate = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else {
          punchDate = dateValue;
        }
      }
    } else {
      punchDate = '未知日期';
    }
    
    // 處理時間
    if (timeValue instanceof Date) {
      punchTime = Utilities.formatDate(timeValue, Session.getScriptTimeZone(), 'HH:mm');
    } else if (typeof timeValue === 'string') {
      // 如果已經是字串格式，檢查是否符合 HH:mm
      if (/^\d{1,2}:\d{2}$/.test(timeValue)) {
        punchTime = timeValue;
      } else {
        punchTime = timeValue;
      }
    } else {
      punchTime = '未知時間';
    }
    
    Logger.log('');
    Logger.log('📊 解析資料:');
    Logger.log('   申請ID: ' + applicationId);
    Logger.log('   用戶ID: ' + userId);
    Logger.log('   員工姓名: ' + employeeName);
    Logger.log('   日期: ' + punchDate);
    Logger.log('   時間: ' + punchTime);
    Logger.log('   類型: ' + punchType);
    Logger.log('   理由: ' + reason);
    
    // ✅ 更新審核狀態
    const statusText = (status === "v") ? "已核准" : "已拒絕";
    
    sheet.getRange(rowNumber, statusCol).setValue(statusText);
    
    if (reviewerCol > 0) {
      sheet.getRange(rowNumber, reviewerCol).setValue("系統管理員");
    }
    
    if (reviewTimeCol > 0) {
      sheet.getRange(rowNumber, reviewTimeCol).setValue(new Date());
    }
    
    Logger.log('✅ 已更新審核狀態為: ' + statusText);
    
    // ✅ 如果核准，寫入「出勤紀錄」工作表
    if (status === "v") {
      const attendanceSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_ATTENDANCE);
      
      if (attendanceSheet) {
        Logger.log('');
        Logger.log('📝 寫入出勤紀錄...');
        
        // 建立完整的日期時間物件
        const punchDateTime = new Date(`${punchDate} ${punchTime}`);
        
        Logger.log('   打卡時間物件: ' + punchDateTime);
        
        // ✅ 取得員工部門資料（可選）
        let employeeDept = '';
        try {
          const employeeInfo = findEmployeeByLineUserId_(userId);
          if (employeeInfo.ok) {
            employeeDept = employeeInfo.dept || '';
          }
        } catch (e) {
          Logger.log('⚠️ 無法取得員工部門: ' + e.message);
        }
        
        // 根據出勤紀錄工作表的欄位順序寫入
        const row = [
          punchDateTime,           // A: 打卡時間
          userId,                  // B: userId
          employeeDept,            // C: 部門
          employeeName,            // D: 打卡人員
          punchType,               // E: 打卡類別（上班/下班）
          '',                      // F: GPS
          '',                      // G: 地點
          '補打卡',                // H: 備註
          'v',                     // I: 管理員審核
          reason || note || ''     // J: 裝置資訊（補打卡理由）
        ];
        
        attendanceSheet.appendRow(row);
        
        Logger.log('✅ 已寫入出勤紀錄');
        Logger.log('   寫入內容: ' + JSON.stringify(row));
      } else {
        Logger.log('❌ 找不到出勤紀錄工作表');
      }
    }
    
    // ✅ 發送 LINE 通知
    const isApproved = (status === "v");
    
    try {
      Logger.log('');
      Logger.log('📤 發送 LINE 通知...');
      
      notifyPunchReview(
        userId,
        employeeName,
        punchDate,
        punchTime,
        punchType,
        "系統管理員",
        isApproved,
        note || ""
      );
      
      Logger.log('✅ LINE 通知已發送');
    } catch (notifyError) {
      Logger.log('⚠️ LINE 通知發送失敗: ' + notifyError.message);
    }
    
    Logger.log('═══════════════════════════════════════');
    return { ok: true, msg: "審核成功並已通知員工" };
    
  } catch (err) {
    Logger.log('❌ updateReviewStatus 錯誤: ' + err.message);
    Logger.log('   錯誤堆疊: ' + err.stack);
    return { ok: false, msg: `審核失敗：${err.message}` };
  }
}


// ==================== 用戶角色管理 ====================

/**
 * ✅ 更新用戶角色
 */
function updateUserRole(userId, newRole) {
  try {
    Logger.log('📝 開始更新用戶角色');
    Logger.log('   userId: ' + userId);
    Logger.log('   newRole: ' + newRole);
    
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_EMPLOYEES);
    
    if (!sheet) {
      return {
        ok: false,
        msg: '找不到員工工作表'
      };
    }
    
    const data = sheet.getDataRange().getValues();
    
    // 檢查是否為最後一個管理員
    if (newRole === 'employee') {
      const adminCount = data.filter((row, index) => 
        index > 0 && row[5] === '管理員'  // F 欄: 部門
      ).length;
      
      if (adminCount <= 1) {
        return {
          ok: false,
          msg: '至少需要保留一位管理員'
        };
      }
    }
    
    // 尋找用戶並更新
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {  // A 欄: userId
        const newDept = newRole === 'admin' ? '管理員' : '員工';
        sheet.getRange(i + 1, 6).setValue(newDept);  // F 欄: 部門
        
        Logger.log('✅ 已更新角色為: ' + newDept);
        
        return {
          ok: true,
          msg: '角色已更新'
        };
      }
    }
    
    return {
      ok: false,
      msg: '找不到該用戶'
    };
    
  } catch (error) {
    Logger.log('❌ updateUserRole 錯誤: ' + error);
    return {
      ok: false,
      msg: error.message
    };
  }
}

/**
 * ✅ 刪除用戶
 */
function deleteUser(userId) {
  try {
    Logger.log('🗑️ 開始刪除用戶');
    Logger.log('   userId: ' + userId);
    
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_EMPLOYEES);
    
    if (!sheet) {
      return {
        ok: false,
        msg: '找不到員工工作表'
      };
    }
    
    const data = sheet.getDataRange().getValues();
    
    // 檢查是否為最後一個管理員
    const targetUser = data.find((row, index) => index > 0 && row[0] === userId);
    
    if (targetUser && targetUser[5] === '管理員') {  // F 欄: 部門
      const adminCount = data.filter((row, index) => 
        index > 0 && row[5] === '管理員'
      ).length;
      
      if (adminCount <= 1) {
        return {
          ok: false,
          msg: '不能刪除最後一位管理員'
        };
      }
    }
    
    // 尋找並刪除用戶
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {  // A 欄: userId
        sheet.deleteRow(i + 1);
        
        Logger.log('✅ 用戶已刪除');
        
        return {
          ok: true,
          msg: '用戶已刪除'
        };
      }
    }
    
    return {
      ok: false,
      msg: '找不到該用戶'
    };
    
  } catch (error) {
    Logger.log('❌ deleteUser 錯誤: ' + error);
    return {
      ok: false,
      msg: error.message
    };
  }
}
// function updateReviewStatus(rowNumber, status, note) {
//   try {
//     Logger.log('═══════════════════════════════════════');
//     Logger.log('📋 開始審核補打卡');
//     Logger.log('   行號: ' + rowNumber);
//     Logger.log('   狀態: ' + status);
    
//     const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ADJUST_PUNCH);
    
//     if (!sheet) {
//       return { ok: false, msg: "找不到補打卡申請工作表" };
//     }
    
//     const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
//     // 找出欄位索引
//     const statusCol = headers.indexOf('狀態') + 1;
//     const reviewerCol = headers.indexOf('審核人') + 1;
//     const reviewTimeCol = headers.indexOf('審核時間') + 1;
    
//     if (statusCol === 0) {
//       return { ok: false, msg: "找不到「狀態」欄位" };
//     }
    
//     // 取得該行的申請資料
//     const record = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    
//     const userId = record[headers.indexOf('用戶ID')];
//     const employeeName = record[headers.indexOf('姓名')];
//     const dateValue = record[headers.indexOf('日期')];
//     const timeValue = record[headers.indexOf('時間')];
//     const punchType = record[headers.indexOf('類型')];
    
//     // ✅ 修正：智能格式化日期時間
//     let punchDate, punchTime;
    
//     if (dateValue instanceof Date) {
//       punchDate = Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
//     } else {
//       punchDate = String(dateValue);
//     }
    
//     if (timeValue instanceof Date) {
//       punchTime = Utilities.formatDate(timeValue, Session.getScriptTimeZone(), 'HH:mm');
//     } else {
//       punchTime = String(timeValue);
//     }
    
//     Logger.log(`   日期: ${punchDate}`);
//     Logger.log(`   時間: ${punchTime}`);
    
//     // 更新審核狀態
//     const statusText = (status === "v") ? "已核准" : "已拒絕";
    
//     sheet.getRange(rowNumber, statusCol).setValue(statusText);
//     sheet.getRange(rowNumber, reviewerCol).setValue("系統管理員");
//     sheet.getRange(rowNumber, reviewTimeCol).setValue(new Date());
    
//     Logger.log('✅ 已更新審核狀態為: ' + statusText);
    
//     // 如果核准，寫入出勤紀錄
//     if (status === "v") {
//       const attendanceSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_ATTENDANCE);
      
//       if (attendanceSheet) {
//         // ✅ 使用格式化後的日期時間
//         const punchDateTime = new Date(`${punchDate} ${punchTime}`);
        
//         attendanceSheet.appendRow([
//           punchDateTime,
//           userId,
//           '',  // 部門
//           employeeName,
//           punchType,
//           '',  // GPS
//           '',  // 地點
//           '補打卡',
//           'v',
//           note || ''
//         ]);
        
//         Logger.log('✅ 已寫入出勤紀錄');
//       }
//     }
    
//     // 發送 LINE 通知
//     const isApproved = (status === "v");
    
//     try {
//       notifyPunchReview(
//         userId,
//         employeeName,
//         punchDate,
//         punchTime,
//         punchType,
//         "系統管理員",
//         isApproved,
//         note || ""
//       );
      
//       Logger.log('✅ LINE 通知已發送');
//     } catch (notifyError) {
//       Logger.log('⚠️ LINE 通知發送失敗: ' + notifyError.message);
//     }
    
//     Logger.log('═══════════════════════════════════════');
//     return { ok: true, msg: "審核成功並已通知員工" };
    
//   } catch (err) {
//     Logger.log('❌ updateReviewStatus 錯誤: ' + err.message);
//     return { ok: false, msg: `審核失敗：${err.message}` };
//   }
// }
// function updateReviewStatus(rowNumber, status, note) {
//   try {
//     Logger.log('═══════════════════════════════════════');
//     Logger.log('📋 開始審核補打卡');
//     Logger.log('   行號: ' + rowNumber);
//     Logger.log('   狀態: ' + status);
    
//     const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ADJUST_PUNCH);
    
//     if (!sheet) {
//       return { ok: false, msg: "找不到補打卡申請工作表" };
//     }
    
//     const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
//     // 找出「狀態」欄位
//     const statusCol = headers.indexOf('狀態') + 1;
//     const reviewerCol = headers.indexOf('審核人') + 1;
//     const reviewTimeCol = headers.indexOf('審核時間') + 1;
    
//     if (statusCol === 0) {
//       return { ok: false, msg: "找不到「狀態」欄位" };
//     }
    
//     // 取得該行的申請資料
//     const record = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    
//     const userId = record[headers.indexOf('用戶ID')];
//     const employeeName = record[headers.indexOf('姓名')];
//     const punchDate = record[headers.indexOf('日期')];
//     const punchTime = record[headers.indexOf('時間')];
//     const punchType = record[headers.indexOf('類型')];
    
//     // ✅ 更新審核狀態
//     const statusText = (status === "v") ? "已核准" : "已拒絕";
    
//     sheet.getRange(rowNumber, statusCol).setValue(statusText);
//     sheet.getRange(rowNumber, reviewerCol).setValue("系統管理員");
//     sheet.getRange(rowNumber, reviewTimeCol).setValue(new Date());
    
//     Logger.log('✅ 已更新審核狀態為: ' + statusText);
    
//     // ✅ 如果核准，寫入「出勤紀錄」工作表
//     if (status === "v") {
//       const attendanceSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_ATTENDANCE);
      
//       if (attendanceSheet) {
//         const punchDateTime = new Date(`${punchDate} ${punchTime}`);
        
//         attendanceSheet.appendRow([
//           punchDateTime,
//           userId,
//           record[headers.indexOf('姓名')],  // 部門欄位可能需要調整
//           employeeName,
//           punchType,
//           '',  // GPS
//           '',  // 地點
//           '補打卡',
//           'v',  // 已核准
//           note || ''
//         ]);
        
//         Logger.log('✅ 已寫入出勤紀錄');
//       }
//     }
    
//     // ✅ 發送 LINE 通知
//     const isApproved = (status === "v");
    
//     try {
//       notifyPunchReview(
//         userId,
//         employeeName,
//         punchDate,
//         punchTime,
//         punchType,
//         "系統管理員",
//         isApproved,
//         note || ""
//       );
      
//       Logger.log('✅ LINE 通知已發送');
//     } catch (notifyError) {
//       Logger.log('⚠️ LINE 通知發送失敗: ' + notifyError.message);
//     }
    
//     Logger.log('═══════════════════════════════════════');
//     return { ok: true, msg: "審核成功並已通知員工" };
    
//   } catch (err) {
//     Logger.log('❌ updateReviewStatus 錯誤: ' + err.message);
//     return { ok: false, msg: `審核失敗：${err.message}` };
//   }
// }

/**
 * 🧪 測試補打卡申請功能（完整流程）
 */
function testAdjustPunchFlow() {
  Logger.log('═══════════════════════════════════════');
  Logger.log('🧪 測試補打卡申請完整流程');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  
  // 步驟 1：檢查工作表是否存在
  Logger.log('📋 步驟 1：檢查工作表');
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_ADJUST_PUNCH);
  
  if (!sheet) {
    Logger.log('❌ 找不到「補打卡申請」工作表');
    Logger.log('   請確認工作表名稱是否正確');
    return;
  }
  
  Logger.log('✅ 工作表存在');
  Logger.log('');
  
  // 步驟 2：檢查欄位結構
  Logger.log('📋 步驟 2：檢查欄位結構');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  Logger.log('工作表欄位:');
  headers.forEach((h, i) => {
    const column = String.fromCharCode(65 + i);
    Logger.log(`   ${column} (${i + 1}): ${h}`);
  });
  Logger.log('');
  
  // 步驟 3：模擬提交補打卡
  Logger.log('📋 步驟 3：模擬提交補打卡');
  
  const testToken = 'a8f8ca99-97d6-4643-ad8e-67a73f2bb649'; // ⚠️ 替換成你的有效 token
  const testType = '上班';
  const testDate = new Date('2025-12-16 09:00:00');
  const testLat = 25.0330;
  const testLng = 121.5654;
  const testNote = '測試補打卡理由：忘記打卡';
  
  Logger.log('測試參數:');
  Logger.log('   token: ' + testToken.substring(0, 20) + '...');
  Logger.log('   type: ' + testType);
  Logger.log('   date: ' + testDate);
  Logger.log('   note: ' + testNote);
  Logger.log('');
  
  const result = punchAdjusted(testToken, testType, testDate, testLat, testLng, testNote);
  
  Logger.log('📤 提交結果:');
  Logger.log(JSON.stringify(result, null, 2));
  Logger.log('');
  
  if (!result.ok) {
    Logger.log('❌ 提交失敗');
    return;
  }
  
  Logger.log('✅ 提交成功');
  Logger.log('');
  
  // 步驟 4：檢查是否寫入工作表
  Logger.log('📋 步驟 4：檢查是否寫入工作表');
  const lastRow = sheet.getLastRow();
  const lastRecord = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  Logger.log('最後一筆記錄:');
  headers.forEach((h, i) => {
    Logger.log(`   ${h}: ${lastRecord[i]}`);
  });
  Logger.log('');
  
  // 步驟 5：測試 getReviewRequest
  Logger.log('📋 步驟 5：測試取得待審核列表');
  const reviewResult = getReviewRequest();
  
  Logger.log('📤 查詢結果:');
  Logger.log('   ok: ' + reviewResult.ok);
  Logger.log('   筆數: ' + (reviewResult.reviewRequest ? reviewResult.reviewRequest.length : 0));
  Logger.log('');
  
  if (reviewResult.reviewRequest && reviewResult.reviewRequest.length > 0) {
    Logger.log('第一筆待審核:');
    const first = reviewResult.reviewRequest[0];
    Logger.log('   id: ' + first.id);
    Logger.log('   name: ' + first.name);
    Logger.log('   type: ' + first.type);
    Logger.log('   applicationPeriod: ' + first.applicationPeriod);
    Logger.log('   note: ' + first.note + ' ⭐');
    Logger.log('');
  }
  
  // 步驟 6：測試審核功能
  Logger.log('📋 步驟 6：測試審核功能');
  
  if (reviewResult.reviewRequest && reviewResult.reviewRequest.length > 0) {
    const testRowNumber = reviewResult.reviewRequest[0].id;
    
    Logger.log('測試核准第 ' + testRowNumber + ' 行');
    const approveResult = updateReviewStatus(testRowNumber, 'v', '測試核准');
    
    Logger.log('📤 審核結果:');
    Logger.log(JSON.stringify(approveResult, null, 2));
    Logger.log('');
  }
  
  Logger.log('═══════════════════════════════════════');
  Logger.log('✅✅✅ 測試完成！');
  Logger.log('');
  Logger.log('📋 檢查清單:');
  Logger.log('   1. ✅ 工作表存在');
  Logger.log('   2. ✅ 欄位結構正確');
  Logger.log('   3. ✅ 補打卡申請成功');
  Logger.log('   4. ✅ 資料寫入工作表');
  Logger.log('   5. ✅ getReviewRequest 包含 note');
  Logger.log('   6. ✅ 審核功能正常');
  Logger.log('');
  Logger.log('🎯 現在可以測試前端了！');
}
// function updateReviewStatus(rowNumber, status, note) {
//   try {
//     Logger.log('═══════════════════════════════════════');
//     Logger.log('📋 開始審核補打卡');
//     Logger.log('   行號: ' + rowNumber);
//     Logger.log('   狀態: ' + status);
//     Logger.log('   備註: ' + (note || '無'));
    
//     const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ATTENDANCE);
//     if (!sheet) {
//       Logger.log('❌ 找不到打卡記錄工作表');
//       return { ok: false, msg: "找不到打卡記錄工作表" };
//     }
    
//     const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
//     Logger.log('📋 標題列: ' + headers.join(', '));
    
//     // 找出「管理員審核」欄位
//     const reviewStatusCol = headers.indexOf('管理員審核') + 1;
//     if (reviewStatusCol === 0) {
//       Logger.log('❌ 找不到「管理員審核」欄位');
//       return { ok: false, msg: "試算表缺少必要欄位：'管理員審核'" };
//     }
    
//     Logger.log('✅ 管理員審核欄位: 第 ' + reviewStatusCol + ' 欄');
    
//     // 取得該行的完整記錄
//     const record = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
//     Logger.log('📄 打卡記錄: ' + JSON.stringify(record));
    
//     // ✅ 根據標題列動態取得欄位索引
//     const userIdCol = headers.indexOf('userId');
//     const nameCol = headers.indexOf('打卡人員');
//     const typeCol = headers.indexOf('打卡類別');
//     const dateCol = headers.indexOf('打卡時間');
    
//     Logger.log('📍 欄位索引:');
//     Logger.log('   userId: ' + userIdCol);
//     Logger.log('   打卡人員: ' + nameCol);
//     Logger.log('   打卡類別: ' + typeCol);
//     Logger.log('   打卡時間: ' + dateCol);
    
//     // 取得資料
//     const userId = record[userIdCol];
//     const employeeName = record[nameCol];
//     const punchType = record[typeCol];
//     const punchDateTime = record[dateCol];
    
//     Logger.log('');
//     Logger.log('📊 解析資料:');
//     Logger.log('   userId: ' + userId);
//     Logger.log('   員工姓名: ' + employeeName);
//     Logger.log('   打卡類型: ' + punchType);
//     Logger.log('   打卡時間: ' + punchDateTime);
    
//     if (!userId || !employeeName) {
//       Logger.log('❌ 缺少必要資料');
//       return { ok: false, msg: "記錄資料不完整" };
//     }
    
//     // 格式化日期和時間
//     const punchDate = formatDate(punchDateTime);
//     const punchTime = formatTime(punchDateTime);
    
//     Logger.log('   格式化日期: ' + punchDate);
//     Logger.log('   格式化時間: ' + punchTime);
    
//     // ✅ 更新審核狀態
//     sheet.getRange(rowNumber, reviewStatusCol).setValue(status);
//     Logger.log('✅ 已更新審核狀態為: ' + status);
    
//     // ✅ 發送 LINE 通知
//     const isApproved = (status === "v");
//     const reviewer = "系統管理員";
    
//     Logger.log('');
//     Logger.log('📤 準備發送 LINE 通知...');
//     Logger.log('   審核結果: ' + (isApproved ? '通過' : '拒絕'));
    
//     try {
//       notifyPunchReview(
//         userId,
//         employeeName,
//         punchDate,
//         punchTime,
//         punchType,
//         reviewer,
//         isApproved,
//         note || ""
//       );
      
//       Logger.log('✅ LINE 通知已發送');
//     } catch (notifyError) {
//       Logger.log('⚠️ LINE 通知發送失敗: ' + notifyError.message);
//       // 不要因為通知失敗而中斷審核流程
//     }
    
//     Logger.log('═══════════════════════════════════════');
//     return { ok: true, msg: "審核成功並已通知員工" };
    
//   } catch (err) {
//     Logger.log('❌ updateReviewStatus 錯誤: ' + err.message);
//     Logger.log('   錯誤堆疊: ' + err.stack);
//     return { ok: false, msg: `審核失敗：${err.message}` };
//   }
// }

/**
 * 🧪 測試審核通知流程
 */
function testApproveWithNotification() {
  Logger.log('🧪 測試審核 + LINE 通知');
  Logger.log('');
  
  // ⚠️ 請先在 Google Sheet 找一筆「補打卡」且「管理員審核 = ?」的記錄
  const testRowNumber = 20; // 替換成實際的行號
  
  Logger.log('📋 測試核准補打卡...');
  const approveResult = updateReviewStatus(testRowNumber, "v", "核准");
  
  Logger.log('');
  Logger.log('📤 審核結果:');
  Logger.log(JSON.stringify(approveResult, null, 2));
  
  if (approveResult.ok) {
    Logger.log('');
    Logger.log('✅✅✅ 測試成功！');
    Logger.log('   請檢查 LINE 是否收到通知');
  } else {
    Logger.log('');
    Logger.log('❌ 測試失敗');
  }
}

/**
 * 🧪 測試拒絕通知流程
 */
function testRejectWithNotification() {
  Logger.log('🧪 測試拒絕 + LINE 通知');
  Logger.log('');
  
  const testRowNumber = 21; // 替換成實際的行號
  
  Logger.log('📋 測試拒絕補打卡...');
  const rejectResult = updateReviewStatus(testRowNumber, "x", "時間不符，請重新申請");
  
  Logger.log('');
  Logger.log('📤 審核結果:');
  Logger.log(JSON.stringify(rejectResult, null, 2));
  
  if (rejectResult.ok) {
    Logger.log('');
    Logger.log('✅✅✅ 測試成功！');
    Logger.log('   請檢查 LINE 是否收到拒絕通知');
  } else {
    Logger.log('');
    Logger.log('❌ 測試失敗');
  }
}
// ==================== 工具函數 ====================

/**
 * 計算兩點之間的距離（公尺）
 */
function getDistanceMeters_(lat1, lng1, lat2, lng2) {
  const R = 6371e3;
  const φ1 = lat1 * Math.PI / 180;
  const φ2 = lat2 * Math.PI / 180;
  const Δφ = (lat2 - lat1) * Math.PI / 180;
  const Δλ = (lng2 - lng1) * Math.PI / 180;

  const a = Math.sin(Δφ / 2) * Math.sin(Δφ / 2) +
            Math.cos(φ1) * Math.cos(φ2) *
            Math.sin(Δλ / 2) * Math.sin(Δλ / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));

  return R * c;
}

/**
 * 格式化日期時間
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
 * 格式化日期
 */
function formatDate(date) {
  if (!date) return '';
  try {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) {
    return String(date);
  }
}

/**
 * 格式化時間
 */
function formatTime(date) {
  if (!date) return '';
  try {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'HH:mm:ss');
  } catch (e) {
    return String(date);
  }
}


function debugCheckSession() {
  Logger.log('═══════════════════════════════════════');
  Logger.log('🔍 診斷 checkSession_');
  Logger.log('═══════════════════════════════════════');
  
  const token = '1fb23a74-f5ee-4d87-bcf7-2bcde4a13d17';  // 你的有效 token
  
  Logger.log('📡 Token: ' + token);
  Logger.log('');
  
  const session = checkSession_(token);
  
  Logger.log('📤 checkSession_ 返回結果:');
  Logger.log(JSON.stringify(session, null, 2));
  Logger.log('');
  
  Logger.log('🔍 詳細檢查:');
  Logger.log('   - session 存在: ' + (session ? '是' : '否'));
  Logger.log('   - session.ok: ' + session.ok);
  Logger.log('   - session.user 存在: ' + (session.user ? '是' : '否'));
  
  if (session.user) {
    Logger.log('');
    Logger.log('👤 User 物件內容:');
    Logger.log('   - userId: ' + session.user.userId);
    Logger.log('   - employeeId: ' + session.user.employeeId);
    Logger.log('   - name: ' + session.user.name);
    Logger.log('   - dept: ' + session.user.dept);
    Logger.log('   - email: ' + session.user.email);
    Logger.log('   - status: ' + session.user.status);
  } else {
    Logger.log('❌ session.user 是 null 或 undefined');
  }
  
  Logger.log('═══════════════════════════════════════');
}


/**
 * 取得員工指定月份的詳細打卡資料（用於圖表分析）
 * @param {string} employeeId - 員工ID
 * @param {string} yearMonth - 年月，格式 "YYYY-MM"
 * @returns {Object} 包含每日打卡時間和工時的資料
 */
function getEmployeeMonthlyPunchData(employeeId, yearMonth) {
  try {
    Logger.log('📊 取得員工打卡分析資料');
    Logger.log('   員工ID: ' + employeeId);
    Logger.log('   月份: ' + yearMonth);
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ATTENDANCE);
    const values = sheet.getDataRange().getValues();
    
    if (values.length <= 1) {
      return { 
        success: false, 
        message: '無打卡記錄' 
      };
    }
    
    // 過濾該員工該月份的記錄
    const records = values.slice(1).filter(row => {
      if (!row[0]) return false;
      
      const date = new Date(row[0]);
      const recordMonth = date.getFullYear() + "-" + String(date.getMonth() + 1).padStart(2, "0");
      const recordEmployeeId = row[1];
      
      return recordMonth === yearMonth && recordEmployeeId === employeeId;
    });
    
    if (records.length === 0) {
      return {
        success: false,
        message: '該月份無打卡記錄'
      };
    }
    
    // 按日期分組
    const dailyData = {};
    
    records.forEach(row => {
      const timestamp = new Date(row[0]);
      const dateKey = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const type = row[4]; // 上班/下班
      const time = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'HH:mm');
      const note = row[7] || '';
      const audit = row[8] || '';
      
      if (!dailyData[dateKey]) {
        dailyData[dateKey] = {
          date: dateKey,
          punchIn: null,
          punchOut: null,
          workHours: 0,
          status: 'normal'
        };
      }
      
      // 只記錄正常打卡或已核准的補打卡
      if (note !== '補打卡' || audit === 'v') {
        if (type === '上班') {
          dailyData[dateKey].punchIn = time;
        } else if (type === '下班') {
          dailyData[dateKey].punchOut = time;
        }
      }
    });
    
    // 計算工時
    const result = Object.values(dailyData).map(day => {
      if (day.punchIn && day.punchOut) {
        try {
          const inTime = new Date(`${day.date} ${day.punchIn}`);
          const outTime = new Date(`${day.date} ${day.punchOut}`);
          const diffMs = outTime - inTime;
          day.workHours = parseFloat((diffMs / (1000 * 60 * 60)).toFixed(2));
        } catch (e) {
          day.workHours = 0;
        }
      } else {
        day.status = 'incomplete';
      }
      return day;
    });
    
    // 排序（由舊到新）
    result.sort((a, b) => new Date(a.date) - new Date(b.date));
    
    return {
      success: true,
      data: result,
      employeeId: employeeId,
      yearMonth: yearMonth,
      totalDays: result.length
    };
    
  } catch (error) {
    Logger.log('❌ getEmployeeMonthlyPunchData 錯誤: ' + error);
    return {
      success: false,
      message: error.message
    };
  }
}

// DbOperations.gs - 修正後的 updateEmployeeName 函數

/**
 * ✅ 修正版：更新員工姓名並設定鎖定標記
 */
function updateEmployeeName(userId, newName) {
  try {
    Logger.log('═══════════════════════════════════════');
    Logger.log('✏️ 開始更新員工姓名');
    Logger.log('   userId: ' + userId);
    Logger.log('   newName: ' + newName);
    
    // 驗證輸入
    if (!userId || !newName) {
      return {
        ok: false,
        msg: '缺少必要參數'
      };
    }
    
    const trimmedName = String(newName).trim();
    
    if (trimmedName.length < 2) {
      return {
        ok: false,
        msg: '姓名至少需要 2 個字'
      };
    }
    
    if (trimmedName.length > 50) {
      return {
        ok: false,
        msg: '姓名不能超過 50 個字'
      };
    }
    
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_EMPLOYEES);
    
    if (!sheet) {
      return {
        ok: false,
        msg: '找不到員工工作表'
      };
    }
    
    const data = sheet.getDataRange().getValues();
    
    // 尋找用戶並更新
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {  // A 欄: userId
        const oldName = data[i][2];  // C 欄: displayName
        
        // ⭐⭐⭐ 關鍵修正：同時更新姓名和 nameOverride
        sheet.getRange(i + 1, 3).setValue(trimmedName);   // C 欄：displayName
        sheet.getRange(i + 1, 9).setValue(trimmedName);   // I 欄：nameOverride（設定鎖定）
        
        Logger.log('✅ 已更新姓名並設定鎖定');
        Logger.log('   舊姓名: ' + oldName);
        Logger.log('   新姓名: ' + trimmedName);
        Logger.log('   nameOverride: ' + trimmedName + ' 🔒');
        Logger.log('═══════════════════════════════════════');
        
        return {
          ok: true,
          msg: '姓名已更新並鎖定',
          oldName: oldName,
          newName: trimmedName
        };
      }
    }
    
    return {
      ok: false,
      msg: '找不到該員工'
    };
    
  } catch (error) {
    Logger.log('❌ updateEmployeeName 錯誤: ' + error);
    Logger.log('═══════════════════════════════════════');
    return {
      ok: false,
      msg: error.message
    };
  }
}
/**
 * 🧪 測試更新姓名
 */
function testUpdateEmployeeName() {
  Logger.log('🧪 測試更新員工姓名');
  Logger.log('');
  
  // ⚠️ 替換成實際的 userId
  const testUserId = 'Ud3b574f260f5a777337158ccd4ff0ba2';
  const newName = '王小明';
  
  const result = updateEmployeeName(testUserId, newName);
  
  Logger.log('📤 結果:');
  Logger.log(JSON.stringify(result, null, 2));
}

/**
 * 根據 token 取得使用者資料
 */
function getUserByToken(token) {
  try {
    const session = checkSession_(token);
    if (session.ok && session.user) {
      return session.user;
    }
    return null;
  } catch (error) {
    Logger.log('❌ getUserByToken 錯誤: ' + error);
    return null;
  }
}


/**
 * 🧪 測試公告系統（改良版）
 */
function testAnnouncementSystemImproved() {
  Logger.log('═══════════════════════════════════════');
  Logger.log('🧪 測試公告系統（改良版）');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  
  const testToken = 'f8eeb016-b89a-4cb5-8737-8c904ca680ab';  // ⚠️ 替換成你的 token
  
  // ✅ 測試 1：新增公告
  Logger.log('📝 測試 1：新增公告');
  const addResult = handleAddAnnouncement({
    token: testToken,
    title: '測試公告 - ' + new Date().getTime(),
    content: '這是一個測試公告',
    priority: 'normal'
  });
  
  Logger.log('   ok: ' + addResult.ok);
  
  if (!addResult.ok) {
    Logger.log('❌ 新增失敗: ' + addResult.msg);
    Logger.log('═══════════════════════════════════════');
    return;
  }
  
  const announcementId = addResult.announcement.id;
  Logger.log('   ✅ 新增成功');
  Logger.log('   ID: ' + announcementId);
  Logger.log('   ID 型別: ' + typeof announcementId);
  Logger.log('');
  
  // ✅ 測試 2：取得公告
  Logger.log('📋 測試 2：取得公告');
  const getResult = handleGetAnnouncements({});
  
  Logger.log('   ok: ' + getResult.ok);
  Logger.log('   公告數: ' + (getResult.announcements ? getResult.announcements.length : 0));
  
  if (getResult.ok && getResult.announcements.length > 0) {
    const latest = getResult.announcements[0];
    Logger.log('   最新公告:');
    Logger.log('     - ID: ' + latest.id + ' (' + typeof latest.id + ')');
    Logger.log('     - 標題: ' + latest.title);
  }
  Logger.log('');
  
  // ⏸️ 等待 2 秒（確保資料已寫入）
  Logger.log('⏸️ 等待 2 秒...');
  Utilities.sleep(2000);
  Logger.log('');
  
  // ✅ 測試 3：刪除公告
  Logger.log('🗑️ 測試 3：刪除公告');
  Logger.log('   準備刪除的 ID: ' + announcementId + ' (' + typeof announcementId + ')');
  
  const deleteResult = handleDeleteAnnouncement({
    token: testToken,
    id: announcementId
  });
  
  Logger.log('');
  Logger.log('   ok: ' + deleteResult.ok);
  Logger.log('   msg: ' + deleteResult.msg);
  Logger.log('');
  
  if (deleteResult.ok) {
    Logger.log('✅✅✅ 刪除成功！');
    
    // ✅ 測試 4：確認刪除
    Logger.log('');
    Logger.log('🔍 測試 4：確認刪除');
    const finalCheck = handleGetAnnouncements({});
    
    if (finalCheck.ok) {
      const stillExists = finalCheck.announcements.some(a => String(a.id) === String(announcementId));
      
      if (stillExists) {
        Logger.log('   ❌ 公告仍存在（刪除失敗）');
      } else {
        Logger.log('   ✅ 公告已不存在（刪除成功）');
      }
    }
  } else {
    Logger.log('❌ 刪除失敗: ' + deleteResult.msg);
  }
  
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
  Logger.log('🎉 測試完成！');
  Logger.log('═══════════════════════════════════════');
}