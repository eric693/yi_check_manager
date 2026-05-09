// AttendanceExport.gs - 考勤報表 Excel 匯出

const ATTENDANCE_EXPORT_FOLDER = '考勤報表';

// ==================== 對外 Handler ====================

/**
 * 匯出月考勤報表 Excel
 * params: { token, yearMonth, employeeId(選填，管理員專用) }
 *
 * 員工：只能匯出自己
 * 管理員：不傳 employeeId 則匯出全員
 */
function handleExportAttendanceExcel(params) {
  try {
    const session = handleCheckSession(params.token);
    if (!session.ok) return { ok: false, msg: 'SESSION_INVALID' };

    const userId   = session.user.userId;
    const isAdmin  = (session.user.dept === '管理員');
    const yearMonth = params.yearMonth || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM');

    // 決定要匯出誰
    let targetEmployees; // [{ id, name }]
    if (isAdmin && !params.employeeId) {
      targetEmployees = getAllActiveEmployees_();
    } else {
      const targetId = (isAdmin && params.employeeId) ? params.employeeId : userId;
      const name     = getEmployeeNameById_(targetId) || targetId;
      targetEmployees = [{ id: targetId, name }];
    }

    if (targetEmployees.length === 0) return { ok: false, msg: '找不到員工資料' };

    // 建立 Excel
    const url = createAttendanceExcel_(targetEmployees, yearMonth, isAdmin);
    return { ok: true, url };

  } catch (err) {
    Logger.log('❌ handleExportAttendanceExcel: ' + err.message);
    return { ok: false, msg: err.message };
  }
}

// ==================== 核心邏輯 ====================

function createAttendanceExcel_(employees, yearMonth, isAdmin) {
  const ssTitle = `考勤報表_${yearMonth}`;
  const ss = SpreadsheetApp.create(ssTitle);

  // 工作表 1：出勤明細
  const detailSheet = ss.getActiveSheet();
  detailSheet.setName('出勤明細');
  buildDetailSheet_(detailSheet, employees, yearMonth);

  // 工作表 2：出勤統計摘要
  const summarySheet = ss.insertSheet('出勤統計');
  buildSummarySheet_(summarySheet, employees, yearMonth);

  // 若只有單一員工，加第三頁薪資摘要
  if (employees.length === 1) {
    const salarySheet = ss.insertSheet('薪資摘要');
    buildSalarySheet_(salarySheet, employees[0], yearMonth);
  }

  ss.setActiveSheet(detailSheet);

  // 匯出為 Excel
  const xlsxBlob = DriveApp.getFileById(ss.getId())
    .getAs('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    .setName(`考勤報表_${yearMonth}.xlsx`);

  // 刪除暫時試算表
  DriveApp.getFileById(ss.getId()).setTrashed(true);

  // 儲存到 Drive
  const folder   = getOrCreateExportFolder_();
  const xlsxFile = folder.createFile(xlsxBlob);
  xlsxFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  xlsxFile.setDescription(`ym:${yearMonth}|exported:${new Date().toISOString()}`);

  Logger.log('✅ 考勤報表 Excel 已建立: ' + xlsxFile.getDownloadUrl());
  return xlsxFile.getDownloadUrl();
}

// ==================== 工作表 1：出勤明細 ====================

function buildDetailSheet_(sheet, employees, yearMonth) {
  const headers = [
    '日期', '星期', '員工姓名', '上班時間', '下班時間',
    '工作時數', '狀態', '請假類型', '請假時數', '加班時數', '備註'
  ];

  const allRows = [];

  employees.forEach(emp => {
    const result = getAttendanceDetails(yearMonth, emp.id);
    if (!result.ok || !result.records) return;

    result.records.forEach(day => {
      const punchIn  = getPunchTime_(day.record, '上班');
      const punchOut = getPunchTime_(day.record, '下班');
      const workHours = calcWorkHours_(punchIn, punchOut, day.date);
      const status   = translateStatus_(day.reason, day.leave);
      const leaveType  = day.leave ? day.leave.leaveType : '';
      const leaveDays  = day.leave ? (day.leave.days || '') : '';
      const otHours    = day.overtime ? (day.overtime.hours || '') : '';
      const weekday  = getWeekdayLabel_(day.date);
      const note = buildNoteText_(day);

      allRows.push([
        day.date,
        weekday,
        day.name || emp.name,
        punchIn,
        punchOut,
        workHours,
        status,
        leaveType,
        leaveDays,
        otHours,
        note
      ]);
    });
  });

  if (allRows.length === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(2, 1).setValue('（本月無出勤記錄）');
    return;
  }

  // 寫入 headers + 資料
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, allRows.length, headers.length).setValues(allRows);

  // ── 樣式 ──
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#10b981').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  // 依狀態上色
  for (let i = 0; i < allRows.length; i++) {
    const rowNum = i + 2;
    const status = allRows[i][6];
    let bgColor = null;
    if (status === '正常') bgColor = null;
    else if (status === '請假') bgColor = '#e0f2fe';
    else if (status.includes('異常') || status.includes('缺卡')) bgColor = '#fee2e2';
    else if (status === '加班') bgColor = '#fef9c3';

    if (bgColor) sheet.getRange(rowNum, 1, 1, headers.length).setBackground(bgColor);
  }

  // 自動調整欄寬
  sheet.autoResizeColumns(1, headers.length);

  // 數字欄置中
  sheet.getRange(2, 6, allRows.length, 1).setHorizontalAlignment('center'); // 工作時數
  sheet.getRange(2, 9, allRows.length, 2).setHorizontalAlignment('center'); // 請假/加班時數
}

// ==================== 工作表 2：出勤統計 ====================

function buildSummarySheet_(sheet, employees, yearMonth) {
  const headers = [
    '員工姓名', '出勤天數', '總工時(h)', '遲到次數',
    '缺卡次數', '請假天數', '加班總時數', '加班次數'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const summaryRows = [];

  employees.forEach(emp => {
    const result = getAttendanceDetails(yearMonth, emp.id);
    if (!result.ok || !result.records) {
      summaryRows.push([emp.name, 0, 0, 0, 0, 0, 0, 0]);
      return;
    }

    let workDays = 0, totalHours = 0, lateCount = 0;
    let missingPunch = 0, leaveDays = 0, otHours = 0, otCount = 0;

    result.records.forEach(day => {
      const punchIn  = getPunchTime_(day.record, '上班');
      const punchOut = getPunchTime_(day.record, '下班');
      const wh = calcWorkHours_(punchIn, punchOut, day.date);

      if (day.reason === 'STATUS_PUNCH_NORMAL') {
        workDays++;
        totalHours += wh;
      }
      if (day.reason === 'STATUS_PUNCH_IN_MISSING' || day.reason === 'STATUS_PUNCH_OUT_MISSING') {
        missingPunch++;
      }
      if (day.leave) leaveDays += parseFloat(day.leave.days) || 1;
      if (day.overtime) {
        otHours += parseFloat(day.overtime.hours) || 0;
        otCount++;
      }
    });

    summaryRows.push([
      emp.name,
      workDays,
      Math.round(totalHours * 10) / 10,
      lateCount,
      missingPunch,
      leaveDays,
      Math.round(otHours * 10) / 10,
      otCount
    ]);
  });

  if (summaryRows.length > 0) {
    sheet.getRange(2, 1, summaryRows.length, headers.length).setValues(summaryRows);
  }

  // 樣式
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#6366f1').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);

  // 數字欄置中
  if (summaryRows.length > 0) {
    sheet.getRange(2, 2, summaryRows.length, headers.length - 1).setHorizontalAlignment('center');
  }

  // 月份標題
  sheet.getRange(summaryRows.length + 3, 1).setValue(`統計月份：${yearMonth}`)
       .setFontColor('#6b7280').setFontSize(10);
}

// ==================== 工作表 3：薪資摘要（單人時附加）====================

function buildSalarySheet_(sheet, emp, yearMonth) {
  const salaryResult = getMySalary(emp.id, yearMonth);
  if (!salaryResult.success) {
    sheet.getRange(1, 1).setValue('本月無薪資資料');
    return;
  }

  const s = salaryResult.data;
  const n = v => parseFloat(v) || 0;

  const rows = [
    ['薪資摘要', yearMonth],
    ['員工姓名', s['員工姓名'] || emp.name],
    ['薪資類型', s['薪資類型'] || '月薪'],
    ['', ''],
    ['── 應發項目 ──', ''],
    ['基本薪資',          n(s['基本薪資'])],
    ['職務加給',          n(s['職務加給'])],
    ['伙食費',            n(s['伙食費'])],
    ['交通補助',          n(s['交通補助'])],
    ['全勤獎金',          n(s['全勤獎金'])],
    ['業績獎金',          n(s['業績獎金'])],
    ['其他津貼',          n(s['其他津貼'])],
    ['平日加班費',        n(s['平日加班費'])],
    ['休息日加班費',      n(s['休息日加班費'])],
    ['例假日加班費',      n(s['例假日加班費'])],
    ['國定假日加班費',    n(s['國定假日加班費'])],
    ['', ''],
    ['── 扣款項目 ──', ''],
    ['勞保費',    n(s['勞保費'])],
    ['健保費',    n(s['健保費'])],
    ['就業保險費', n(s['就業保險費'])],
    ['勞退自提',  n(s['勞退自提'])],
    ['所得稅',    n(s['所得稅'])],
    ['請假扣款',  n(s['請假扣款'])],
    ['早退扣款',  n(s['早退扣款'])],
    ['其他扣款',  n(s['其他扣款'])],
    ['', ''],
    ['應發總額',  n(s['應發總額'])],
    ['實發金額',  n(s['實發金額'])],
  ];

  sheet.getRange(1, 1, rows.length, 2).setValues(rows);

  // 標題
  sheet.getRange(1, 1, 1, 2).setBackground('#6366f1').setFontColor('#ffffff').setFontWeight('bold');
  // 小節標題
  [5, 18].forEach(r => sheet.getRange(r, 1, 1, 2).setBackground('#f0fdf4').setFontWeight('bold'));
  // 合計列
  sheet.getRange(rows.length - 1, 1, 2, 2).setBackground('#e0f2fe').setFontWeight('bold');
  sheet.getRange(rows.length, 1, 1, 2).setBackground('#10b981').setFontColor('#ffffff').setFontWeight('bold').setFontSize(12);

  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 120);
  sheet.getRange(1, 2, rows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right');
}

// ==================== 輔助函式 ====================

function getPunchTime_(records, type) {
  if (!records || records.length === 0) return '';
  const found = records.find(r => r.type === type);
  return found ? found.time : '';
}

function calcWorkHours_(punchIn, punchOut, dateStr) {
  if (!punchIn || !punchOut) return 0;
  try {
    const inT  = new Date(`${dateStr} ${punchIn}`);
    const outT = new Date(`${dateStr} ${punchOut}`);
    const diff = (outT - inT) / (1000 * 60 * 60);
    return diff > 0 ? Math.round(diff * 10) / 10 : 0;
  } catch (e) { return 0; }
}

function translateStatus_(reason, leave) {
  if (leave) return '請假';
  switch (reason) {
    case 'STATUS_PUNCH_NORMAL':       return '正常';
    case 'STATUS_PUNCH_IN_MISSING':   return '異常-缺上班卡';
    case 'STATUS_PUNCH_OUT_MISSING':  return '異常-缺下班卡';
    case 'STATUS_NO_RECORD':          return '缺卡';
    default: return reason || '';
  }
}

function getWeekdayLabel_(dateStr) {
  const days = ['日', '一', '二', '三', '四', '五', '六'];
  try {
    return '週' + days[new Date(dateStr).getDay()];
  } catch (e) { return ''; }
}

function buildNoteText_(day) {
  const parts = [];
  if (day.overtime) parts.push(`加班 ${day.overtime.hours}h`);
  if (day.leave && day.leave.reason) parts.push(`請假原因：${day.leave.reason}`);
  return parts.join(' / ');
}

function getAllActiveEmployees_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EMPLOYEES);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const id     = String(data[i][EMPLOYEE_COL.USER_ID] || '').trim();
    const name   = String(data[i][EMPLOYEE_COL.NAME] || '').trim();
    const status = String(data[i][EMPLOYEE_COL.STATUS] || '');
    if (id && status !== '停用') result.push({ id, name });
  }
  return result;
}

function getEmployeeNameById_(userId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EMPLOYEES);
    if (!sheet) return null;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][EMPLOYEE_COL.USER_ID]) === userId) {
        return String(data[i][EMPLOYEE_COL.NAME]);
      }
    }
  } catch (e) {}
  return null;
}

function getOrCreateExportFolder_() {
  const folders = DriveApp.getFoldersByName(ATTENDANCE_EXPORT_FOLDER);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(ATTENDANCE_EXPORT_FOLDER);
}
