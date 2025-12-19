// DailySalary.gs - 日薪管理系統（修正版 - 接受前端參數）

/**
 * ✅ 計算日薪（修正版 - 接受前端手動輸入的參數）
 * 
 * @param {string} employeeId - 員工ID
 * @param {string} yearMonth - 年月 (格式: YYYY-MM)
 * @param {object} manualInputs - 前端手動輸入的參數（可選）
 * @param {number} manualInputs.workDays - 上班天數
 * @param {number} manualInputs.overtimeHours - 加班時數
 * @param {number} manualInputs.leaveDeduction - 請假扣款
 * @param {number} manualInputs.advancePayment - 預支金額
 * @param {number} manualInputs.agencyDeduction - 代辦6小時扣款
 * @param {number} manualInputs.otherDeduction - 其他代扣
 * @param {number} manualInputs.fineDeduction - 罰單均分
 */
function calculateDailySalary(employeeId, yearMonth, manualInputs) {
  try {
    Logger.log('💰 計算日薪: ' + employeeId + ', ' + yearMonth);
    
    // 1. 取得員工資料
    const employeeResult = getDailyEmployee(employeeId);
    if (!employeeResult.success) {
      return employeeResult;
    }
    
    const employee = employeeResult.data;
    
    // 2. 決定使用自動計算還是手動輸入
    let workDays = 0;
    let totalOvertimeHours = 0;
    let leaveDeduction = 0;
    let advance = 0;
    let sixHourDeduction = 0;
    let otherDeductions = 0;
    let fineShare = 0;
    
    if (manualInputs && typeof manualInputs === 'object') {
      // ✅ 使用前端傳來的手動輸入
      Logger.log('📝 使用手動輸入的參數');
      
      workDays = parseFloat(manualInputs.workDays) || 0;
      totalOvertimeHours = parseFloat(manualInputs.overtimeHours) || 0;
      leaveDeduction = parseFloat(manualInputs.leaveDeduction) || 0;
      advance = parseFloat(manualInputs.advancePayment) || 0;
      sixHourDeduction = parseFloat(manualInputs.agencyDeduction) || 0;
      otherDeductions = parseFloat(manualInputs.otherDeduction) || 0;
      fineShare = parseFloat(manualInputs.fineDeduction) || 0;
      
      Logger.log('   上班天數: ' + workDays);
      Logger.log('   加班時數: ' + totalOvertimeHours);
      
    } else {
      // ✅ 自動從打卡記錄計算
      Logger.log('🤖 自動計算（從打卡記錄）');
      
      // 取得當月打卡記錄（計算上班天數）
      const attendanceRecords = getAttendanceRecords(yearMonth, employeeId);
      for (let record of attendanceRecords) {
        if (record['上班時間']) {
          workDays++;
        }
      }
      
      Logger.log('📅 本月上班天數: ' + workDays);
      
      // 取得加班記錄
      const overtimeRecords = getEmployeeOvertimeForMonth(employeeId, yearMonth);
      for (let ot of overtimeRecords) {
        if (ot['狀態'] === '已核准') {
          totalOvertimeHours += parseFloat(ot['加班時數']) || 0;
        }
      }
      
      Logger.log('⏰ 本月加班時數: ' + totalOvertimeHours);
      
      // 取得請假扣款（從請假記錄計算）
      const dailySalary = parseFloat(employee['日薪']) || 0;
      const leaveRecords = getEmployeeLeaveForMonth(employeeId, yearMonth);
      for (let leave of leaveRecords) {
        if (leave['狀態'] === '已核准' && 
            (leave['假別'] === '事假' || leave['假別'] === '病假')) {
          const leaveDays = parseFloat(leave['天數']) || 0;
          leaveDeduction += dailySalary * leaveDays;
        }
      }
    }
    
    // 3. 計算應發項目
    const dailySalary = parseFloat(employee['日薪']) || 0;
    const overtimeHourlyRate = parseFloat(employee['加班時薪']) || 0;
    const mealAllowancePerDay = parseFloat(employee['伙食津貼（天）']) || 0;
    
    const basePay = dailySalary * workDays;
    const overtimePay = overtimeHourlyRate * totalOvertimeHours;
    const mealAllowance = mealAllowancePerDay * workDays;
    const drivingAllowance = parseFloat(employee['開車津貼']) || 0;
    const positionAllowance = parseFloat(employee['職務津貼']) || 0;
    const housingAllowance = parseFloat(employee['租屋津貼']) || 0;
    const otherPayments = 0;
    
    const grossSalary = basePay + overtimePay + mealAllowance + 
                       drivingAllowance + positionAllowance + 
                       housingAllowance + otherPayments;
    
    // 4. 計算扣款項目
    const laborFee = parseFloat(employee['勞保費']) || 0;
    const healthFee = parseFloat(employee['健保費']) || 0;
    const dependentHealthFee = parseFloat(employee['眷屬健保費']) || 0;
    
    const totalDeductions = laborFee + healthFee + dependentHealthFee + 
                           leaveDeduction + advance + sixHourDeduction + 
                           otherDeductions + fineShare;
    
    const netSalary = grossSalary - totalDeductions;
    
    // 5. 組成計算結果
    const result = {
      employeeId: employeeId,
      employeeName: employee['員工姓名'],
      yearMonth: yearMonth,
      workDays: workDays,
      baseDailySalary: dailySalary,
      basePay: basePay,
      overtimeHours: totalOvertimeHours,
      overtimePay: overtimePay,
      mealAllowance: mealAllowance,
      drivingAllowance: drivingAllowance,
      positionAllowance: positionAllowance,
      housingAllowance: housingAllowance,
      otherPayments: otherPayments,
      grossSalary: grossSalary,
      laborFee: laborFee,
      healthFee: healthFee,
      dependentHealthFee: dependentHealthFee,
      leaveDeduction: leaveDeduction,
      advancePayment: advance,
      agencyDeduction: sixHourDeduction,
      otherDeduction: otherDeductions,
      fineDeduction: fineShare,
      totalDeductions: totalDeductions,
      netSalary: netSalary,
      bankCode: employee['銀行代碼'],
      bankAccount: employee['銀行帳號']
    };
    
    Logger.log('✅ 日薪計算完成');
    Logger.log('   應發: $' + grossSalary);
    Logger.log('   扣款: $' + totalDeductions);
    Logger.log('   實發: $' + netSalary);
    
    return {
      success: true,
      data: result,
      message: '計算完成'
    };
    
  } catch (error) {
    Logger.log('❌ 計算失敗: ' + error);
    return {
      success: false,
      message: '計算失敗: ' + error.message
    };
  }
}

// ==================== 其他函數保持不變 ====================

/**
 * ✅ 取得日薪 Sheet
 */
function getDailySalarySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('日薪員工');
  
  if (!sheet) {
    Logger.log('📝 建立日薪員工工作表');
    sheet = ss.insertSheet('日薪員工');
    
    const headers = [
      '員工ID', '員工姓名', '血型', '手機', '出生年月日',
      '緊急聯絡人', '緊急聯絡人電話', '通訊地址',
      '日薪', '加班時薪', '伙食津貼（天）',
      '開車津貼', '職務津貼', '租屋津貼',
      '勞保費', '健保費', '眷屬健保費',
      '銀行代碼', '銀行帳號', '建立日期', '更新日期', '備註'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

/**
 * ✅ 取得日薪員工資料
 */
function getDailyEmployee(employeeId) {
  try {
    const sheet = getDailySalarySheet();
    const allData = sheet.getDataRange().getValues();
    
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === employeeId) {
        const row = allData[i];
        
        return {
          success: true,
          data: {
            '員工ID': row[0],
            '員工姓名': row[1],
            '血型': row[2],
            '手機': row[3],
            '出生年月日': row[4],
            '緊急聯絡人': row[5],
            '緊急聯絡人電話': row[6],
            '通訊地址': row[7],
            '日薪': row[8],
            '加班時薪': row[9],
            '伙食津貼（天）': row[10],
            '開車津貼': row[11],
            '職務津貼': row[12],
            '租屋津貼': row[13],
            '勞保費': row[14],
            '健保費': row[15],
            '眷屬健保費': row[16],
            '銀行代碼': row[17],
            '銀行帳號': row[18],
            '建立日期': row[19],
            '更新日期': row[20],
            '備註': row[21]
          }
        };
      }
    }
    
    return {
      success: false,
      message: '找不到該員工資料'
    };
    
  } catch (error) {
    return {
      success: false,
      message: '查詢失敗: ' + error.message
    };
  }
}

/**
 * ✅ 設定日薪員工資料
 */
function setDailyEmployee(data) {
  try {
    const sheet = getDailySalarySheet();
    const allData = sheet.getDataRange().getValues();
    
    let targetRow = -1;
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === data.employeeId) {
        targetRow = i + 1;
        break;
      }
    }
    
    const now = new Date();
    const rowData = [
      data.employeeId,
      data.employeeName || '',
      data.bloodType || '',
      data.phone || '',
      data.birthDate || '',
      data.emergencyContact || '',
      data.emergencyPhone || '',
      data.address || '',
      parseFloat(data.dailySalary) || 0,
      parseFloat(data.overtimeHourlyRate) || 0,
      parseFloat(data.mealAllowancePerDay) || 0,
      parseFloat(data.drivingAllowance) || 0,
      parseFloat(data.positionAllowance) || 0,
      parseFloat(data.housingAllowance) || 0,
      parseFloat(data.laborFee) || 0,
      parseFloat(data.healthFee) || 0,
      parseFloat(data.dependentHealthFee) || 0,
      data.bankCode || '',
      data.bankAccount || '',
      targetRow === -1 ? now : allData[targetRow - 1][19],
      now,
      data.note || ''
    ];
    
    if (targetRow !== -1) {
      sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    
    return {
      success: true,
      message: '日薪員工設定成功',
      employeeId: data.employeeId
    };
    
  } catch (error) {
    return {
      success: false,
      message: '設定失敗: ' + error.message
    };
  }
}

/**
 * ✅ 儲存日薪計算記錄
 */
function saveDailySalaryRecord(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('日薪計算記錄');
    
    if (!sheet) {
      sheet = ss.insertSheet('日薪計算記錄');
      const headers = [
        '計算ID', '員工ID', '員工姓名', '年月', '上班天數',
        '基本日薪', '加班時數', '加班費', '伙食津貼',
        '開車津貼', '職務津貼', '租屋津貼', '其他代付',
        '應發總額', '勞保費', '健保費', '眷屬健保費',
        '請假扣款', '預支', '代辦6小時', '其他代扣',
        '罰單（均分）', '扣款總額', '實發金額',
        '銀行代碼', '銀行帳號', '計算日期', '狀態', '備註'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    
    const calculationId = 'DS-' + data.yearMonth + '-' + data.employeeId + '-' + Date.now();
    const now = new Date();
    
    const rowData = [
      calculationId,
      data.employeeId,
      data.employeeName,
      data.yearMonth,
      data.workDays,
      data.baseDailySalary,
      data.overtimeHours,
      data.overtimePay,
      data.mealAllowance,
      data.drivingAllowance,
      data.positionAllowance,
      data.housingAllowance,
      data.otherPayments || 0,
      data.grossSalary,
      data.laborFee,
      data.healthFee,
      data.dependentHealthFee,
      data.leaveDeduction,
      data.advancePayment || 0,
      data.agencyDeduction || 0,
      data.otherDeduction || 0,
      data.fineDeduction || 0,
      data.totalDeductions,
      data.netSalary,
      data.bankCode || '',
      data.bankAccount || '',
      now,
      '已計算',
      data.note || ''
    ];
    
    sheet.appendRow(rowData);
    
    return {
      success: true,
      message: '儲存成功',
      calculationId: calculationId
    };
    
  } catch (error) {
    return {
      success: false,
      message: '儲存失敗: ' + error.message
    };
  }
}

/**
 * ✅ 取得員工的加班記錄（指定月份）
 */
function getEmployeeOvertimeForMonth(employeeId, yearMonth) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('加班申請');
    
    if (!sheet) return [];
    
    const allData = sheet.getDataRange().getValues();
    const records = [];
    
    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      const recordEmployeeId = row[0];
      const overtimeDate = row[2];
      
      if (recordEmployeeId === employeeId && 
          overtimeDate && 
          overtimeDate.toString().startsWith(yearMonth)) {
        records.push({
          '員工ID': row[0],
          '員工姓名': row[1],
          '加班日期': row[2],
          '開始時間': row[3],
          '結束時間': row[4],
          '加班時數': row[5],
          '申請原因': row[6],
          '狀態': row[7]
        });
      }
    }
    
    return records;
    
  } catch (error) {
    Logger.log('❌ 取得加班記錄失敗: ' + error);
    return [];
  }
}

/**
 * ✅ 取得員工的請假記錄（指定月份）
 */
function getEmployeeLeaveForMonth(employeeId, yearMonth) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('請假申請');
    
    if (!sheet) return [];
    
    const allData = sheet.getDataRange().getValues();
    const records = [];
    
    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      const recordEmployeeId = row[0];
      const startDate = row[3];
      
      if (recordEmployeeId === employeeId && 
          startDate && 
          startDate.toString().startsWith(yearMonth)) {
        records.push({
          '員工ID': row[0],
          '員工姓名': row[1],
          '假別': row[2],
          '開始日期': row[3],
          '結束日期': row[4],
          '天數': row[5],
          '原因': row[6],
          '狀態': row[7]
        });
      }
    }
    
    return records;
    
  } catch (error) {
    Logger.log('❌ 取得請假記錄失敗: ' + error);
    return [];
  }
}

/**
 * ✅ 取得所有日薪員工列表
 */
function getAllDailyEmployees() {
  try {
    const sheet = getDailySalarySheet();
    const allData = sheet.getDataRange().getValues();
    const employees = [];
    
    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      employees.push({
        '員工ID': row[0],
        '員工姓名': row[1],
        '血型': row[2],
        '手機': row[3],
        '日薪': row[8],
        '加班時薪': row[9],
        '銀行代碼': row[17],
        '銀行帳號': row[18]
      });
    }
    
    return {
      success: true,
      data: employees,
      total: employees.length
    };
    
  } catch (error) {
    return {
      success: false,
      message: '取得失敗: ' + error.message
    };
  }
}

/**
 * ✅ 取得日薪計算記錄（指定年月）
 */
function getDailySalaryRecords(yearMonth) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('日薪計算記錄');
    
    if (!sheet) {
      return {
        success: true,
        data: [],
        total: 0
      };
    }
    
    const allData = sheet.getDataRange().getValues();
    const records = [];
    
    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      
      if (!yearMonth || row[3] === yearMonth) {
        records.push({
          '計算ID': row[0],
          '員工ID': row[1],
          '員工姓名': row[2],
          '年月': row[3],
          '上班天數': row[4],
          '實發金額': row[23],
          '狀態': row[27],
          '計算日期': row[26]
        });
      }
    }
    
    return {
      success: true,
      data: records,
      total: records.length
    };
    
  } catch (error) {
    return {
      success: false,
      message: '取得失敗: ' + error.message
    };
  }
}

function testDailySalaryCalculation() {
  const params = {
    token: '3b419320-57b1-4cd0-861a-23a48b132a5c',
    employeeId: 'D001',
    yearMonth: '2025-11',
    workDays: 20,
    overtimeHours: 10,
    leaveDeduction: 500,
    advancePayment: 1000,
    agencyDeduction: 200,
    otherDeduction: 100,
    fineDeduction: 50
  };
  
  const result = handleCalculateDailySalary(params);
  Logger.log(JSON.stringify(result, null, 2));
}