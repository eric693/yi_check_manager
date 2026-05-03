// ComplianceCheck.gs - 勞基法合規自動警示

// ==================== 2026 年台灣法定標準 ====================
const LABOR_LAW = {
  MIN_MONTHLY_SALARY:   29500,  // 最低月薪（元）
  MIN_HOURLY_SALARY:    196,    // 最低時薪（元）
  MAX_OT_MONTHLY:       46,     // 單月加班上限（含休息日 54，一般 46）
  MAX_OT_MONTHLY_TOTAL: 54,     // 含休息日加班上限
  MAX_CONSECUTIVE_DAYS: 6,      // 最多連續出勤天數（每7天至少1日例假）
  MIN_REST_BETWEEN_SHIFTS: 8,   // 兩班次間最少休息（小時）
  MAX_DAILY_WORK_HOURS: 12,     // 單日工時上限（正常8+加班4）
};

// ==================== 主要函數 ====================

/**
 * 執行合規檢查，回傳違規清單
 */
function runComplianceCheck(params) {
  try {
    const session = handleCheckSession(params.token);
    if (!session.ok) return { ok: false, msg: 'SESSION_INVALID' };
    if (session.user.dept !== '管理員') return { ok: false, msg: '僅限管理員' };

    const yearMonth = params.yearMonth || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM');
    Logger.log(`🔍 開始合規檢查：${yearMonth}`);

    const violations = [];
    const warnings  = [];

    // 取得所有員工
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const empSheet = ss.getSheetByName(SHEET_EMPLOYEES);
    if (!empSheet) return { ok: false, msg: '找不到員工名單' };

    const empData = empSheet.getDataRange().getValues();
    const employees = [];
    for (let i = 1; i < empData.length; i++) {
      const row = empData[i];
      if (String(row[EMPLOYEE_COL.STATUS] || '') === '停用') continue;
      employees.push({
        userId: String(row[EMPLOYEE_COL.USER_ID]),
        name:   String(row[EMPLOYEE_COL.NAME]),
        dept:   String(row[EMPLOYEE_COL.DEPT] || '')
      });
    }

    employees.forEach(emp => {
      const empViolations = checkEmployee(emp, yearMonth);
      empViolations.forEach(v => {
        v.employeeId   = emp.userId;
        v.employeeName = emp.name;
        if (v.level === 'red') violations.push(v);
        else warnings.push(v);
      });
    });

    // 全域檢查（最低薪資）
    const salaryViolations = checkMinimumSalary(yearMonth);
    salaryViolations.forEach(v => {
      if (v.level === 'red') violations.push(v);
      else warnings.push(v);
    });

    Logger.log(`✅ 合規檢查完成：${violations.length} 項違規，${warnings.length} 項警告`);
    return {
      ok: true,
      yearMonth,
      violations,
      warnings,
      summary: {
        totalEmployees: employees.length,
        redCount:    violations.length,
        yellowCount: warnings.length
      }
    };
  } catch (e) {
    Logger.log('❌ runComplianceCheck: ' + e);
    return { ok: false, msg: e.toString() };
  }
}

// ==================== 個別員工檢查 ====================

function checkEmployee(emp, yearMonth) {
  const results = [];

  try { results.push(...checkOvertimeLimit(emp.userId, emp.name, yearMonth)); } catch(e) {}
  try { results.push(...checkConsecutiveDays(emp.userId, emp.name, yearMonth)); } catch(e) {}
  try { results.push(...checkDailyWorkHours(emp.userId, emp.name, yearMonth)); } catch(e) {}
  try { results.push(...checkRestBetweenShifts(emp.userId, emp.name, yearMonth)); } catch(e) {}

  return results;
}

// ==================== 各項合規檢查 ====================

/**
 * 1. 加班時數上限檢查
 */
function checkOvertimeLimit(userId, name, yearMonth) {
  const results = [];
  const overtimeRecords = getEmployeeMonthlyOvertime(userId, yearMonth);

  let totalOT = 0;
  overtimeRecords.forEach(r => { totalOT += parseFloat(r.hours) || 0; });

  if (totalOT > LABOR_LAW.MAX_OT_MONTHLY_TOTAL) {
    results.push({
      level: 'red',
      type: 'OVERTIME_EXCEED',
      message: `加班時數超過每月上限 ${LABOR_LAW.MAX_OT_MONTHLY_TOTAL} 小時`,
      detail: `本月加班 ${totalOT.toFixed(1)} 小時（上限 ${LABOR_LAW.MAX_OT_MONTHLY_TOTAL}h）`
    });
  } else if (totalOT > LABOR_LAW.MAX_OT_MONTHLY) {
    results.push({
      level: 'yellow',
      type: 'OVERTIME_WARNING',
      message: `加班時數超過一般上限 ${LABOR_LAW.MAX_OT_MONTHLY} 小時`,
      detail: `本月加班 ${totalOT.toFixed(1)} 小時（一般上限 ${LABOR_LAW.MAX_OT_MONTHLY}h，含休息日上限 ${LABOR_LAW.MAX_OT_MONTHLY_TOTAL}h）`
    });
  }

  return results;
}

/**
 * 2. 連續出勤天數檢查（勞基法第36條：每7日至少1日例假）
 */
function checkConsecutiveDays(userId, name, yearMonth) {
  const results = [];
  const attendanceRecords = getEmployeeMonthlyAttendanceInternal(userId, yearMonth);

  if (!attendanceRecords.length) return results;

  // 取出有出勤的日期並排序
  const workedDates = attendanceRecords
    .filter(r => r.workHours > 0 || r.punchIn)
    .map(r => r.date)
    .sort();

  let consecutive = 1;
  let maxConsecutive = 1;
  let consecutiveStart = workedDates[0];

  for (let i = 1; i < workedDates.length; i++) {
    const prev = new Date(workedDates[i - 1]);
    const curr = new Date(workedDates[i]);
    const diffDays = (curr - prev) / (1000 * 60 * 60 * 24);

    if (diffDays === 1) {
      consecutive++;
      if (consecutive > maxConsecutive) {
        maxConsecutive = consecutive;
      }
    } else {
      consecutive = 1;
      consecutiveStart = workedDates[i];
    }
  }

  if (maxConsecutive > LABOR_LAW.MAX_CONSECUTIVE_DAYS) {
    results.push({
      level: 'red',
      type: 'CONSECUTIVE_DAYS',
      message: `連續出勤超過 ${LABOR_LAW.MAX_CONSECUTIVE_DAYS} 天（違反勞基法第36條）`,
      detail: `本月最長連續出勤 ${maxConsecutive} 天，勞基法規定每7日至少1日例假`
    });
  } else if (maxConsecutive >= LABOR_LAW.MAX_CONSECUTIVE_DAYS) {
    results.push({
      level: 'yellow',
      type: 'CONSECUTIVE_DAYS_WARNING',
      message: `連續出勤已達 ${maxConsecutive} 天，接近法定上限`,
      detail: `已連續出勤 ${maxConsecutive} 天，法定上限為連續 ${LABOR_LAW.MAX_CONSECUTIVE_DAYS} 天`
    });
  }

  return results;
}

/**
 * 3. 單日工時上限檢查（正常工時8h + 加班上限4h = 12h）
 */
function checkDailyWorkHours(userId, name, yearMonth) {
  const results = [];
  const attendanceRecords = getEmployeeMonthlyAttendanceInternal(userId, yearMonth);
  const overtimeRecords   = getEmployeeMonthlyOvertime(userId, yearMonth);

  // 按日期合計工時（出勤 + 加班）
  const dailyHours = {};
  attendanceRecords.forEach(r => {
    if (r.workHours > 0) dailyHours[r.date] = (dailyHours[r.date] || 0) + r.workHours;
  });
  overtimeRecords.forEach(r => {
    if (r.date) dailyHours[r.date] = (dailyHours[r.date] || 0) + (parseFloat(r.hours) || 0);
  });

  const violations = [];
  Object.keys(dailyHours).forEach(date => {
    if (dailyHours[date] > LABOR_LAW.MAX_DAILY_WORK_HOURS) {
      violations.push(`${date}（${dailyHours[date].toFixed(1)}h）`);
    }
  });

  if (violations.length) {
    results.push({
      level: 'red',
      type: 'DAILY_HOURS_EXCEED',
      message: `單日工時超過 ${LABOR_LAW.MAX_DAILY_WORK_HOURS} 小時`,
      detail: `超時日期：${violations.join('、')}`
    });
  }

  return results;
}

/**
 * 4. 班次間隔不足檢查（勞基法第35條：工作逾4小時至少休息30分鐘；兩班次間至少8小時）
 */
function checkRestBetweenShifts(userId, name, yearMonth) {
  const results = [];
  const attendanceRecords = getEmployeeMonthlyAttendanceInternal(userId, yearMonth);

  // 排序打卡記錄
  const sorted = attendanceRecords
    .filter(r => r.punchIn && r.punchOut)
    .sort((a, b) => a.date.localeCompare(b.date));

  const shortRests = [];

  for (let i = 1; i < sorted.length; i++) {
    try {
      const prevOut  = new Date(sorted[i-1].date + 'T' + sorted[i-1].punchOut);
      const currIn   = new Date(sorted[i].date   + 'T' + sorted[i].punchIn);
      const restHrs  = (currIn - prevOut) / (1000 * 60 * 60);
      if (restHrs > 0 && restHrs < LABOR_LAW.MIN_REST_BETWEEN_SHIFTS) {
        shortRests.push(`${sorted[i-1].date}→${sorted[i].date}（休息${restHrs.toFixed(1)}h）`);
      }
    } catch(e) {}
  }

  if (shortRests.length) {
    results.push({
      level: 'yellow',
      type: 'SHORT_REST',
      message: `班次間隔不足 ${LABOR_LAW.MIN_REST_BETWEEN_SHIFTS} 小時`,
      detail: `短休情況：${shortRests.join('、')}`
    });
  }

  return results;
}

/**
 * 5. 最低薪資檢查（讀已計算的月薪資記錄）
 */
function checkMinimumSalary(yearMonth) {
  const results = [];
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('月薪資記錄');
    if (!sheet) return results;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // col E=薪資類型, col B=員工ID, col C=員工姓名, col D=年月, col I=基本薪資
      if (String(row[3]) !== yearMonth) continue;
      const salaryType = String(row[4]);
      const baseSalary = parseFloat(row[8]) || 0;
      const netSalary  = parseFloat(row[40]) || 0; // 實發金額

      if (salaryType === '月薪' && baseSalary > 0 && baseSalary < LABOR_LAW.MIN_MONTHLY_SALARY) {
        results.push({
          level: 'red',
          type: 'BELOW_MIN_WAGE',
          employeeId:   String(row[1]),
          employeeName: String(row[2]),
          message: `基本薪資低於最低工資（$${LABOR_LAW.MIN_MONTHLY_SALARY}）`,
          detail: `基本薪資 $${baseSalary.toLocaleString()}，低於法定最低月薪 $${LABOR_LAW.MIN_MONTHLY_SALARY.toLocaleString()}`
        });
      }
    }

    // 時薪員工從薪資設定表讀取
    const salaryConfigSheet = ss.getSheetByName('員工薪資設定');
    if (salaryConfigSheet) {
      const cfgData = salaryConfigSheet.getDataRange().getValues();
      for (let i = 1; i < cfgData.length; i++) {
        const row = cfgData[i];
        // col E=薪資類型, col F=基本薪資
        if (String(row[4]) !== '時薪') continue;
        if (String(row[20] || 'active') === '停用') continue;
        const hourlyRate = parseFloat(row[5]) || 0;
        if (hourlyRate > 0 && hourlyRate < LABOR_LAW.MIN_HOURLY_SALARY) {
          results.push({
            level: 'red',
            type: 'BELOW_MIN_HOURLY',
            employeeId:   String(row[0]),
            employeeName: String(row[1]),
            message: `時薪低於最低工資（$${LABOR_LAW.MIN_HOURLY_SALARY}/h）`,
            detail: `時薪 $${hourlyRate}，低於法定最低時薪 $${LABOR_LAW.MIN_HOURLY_SALARY}`
          });
        }
      }
    }
  } catch(e) {
    Logger.log('⚠️ checkMinimumSalary: ' + e);
  }
  return results;
}
