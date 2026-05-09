// PayslipPdf.gs - PDF 薪資單產生系統

const PAYSLIP_FOLDER_NAME = '薪資單PDF';

// ==================== 對外 Handler ====================

/**
 * 產生並下載 PDF 薪資單
 * params: { token, yearMonth }
 * 員工查自己；管理員可傳 employeeId 查任意人
 */
function handleGeneratePayslipPdf(params) {
  try {
    const session = handleCheckSession(params.token);
    if (!session.ok) return { ok: false, msg: 'SESSION_INVALID' };

    const userId    = session.user.userId;
    const isAdmin   = (session.user.dept === '管理員');
    const targetId  = (isAdmin && params.employeeId) ? params.employeeId : userId;
    const yearMonth = params.yearMonth || getCurrentYearMonth_();

    // 取得薪資資料
    const salaryResult = getMySalary(targetId, yearMonth);
    if (!salaryResult.success) {
      return { ok: false, msg: salaryResult.message || '查無薪資記錄' };
    }

    const salary = salaryResult.data;
    const employeeName = salary['員工姓名'] || targetId;

    // 嘗試取得已存在的 PDF（避免重複建立）
    const existingUrl = findExistingPayslipPdf_(targetId, yearMonth);
    if (existingUrl) {
      return { ok: true, url: existingUrl, cached: true };
    }

    // 建立 PDF
    const pdfUrl = createPayslipPdf_(salary, employeeName, yearMonth);
    return { ok: true, url: pdfUrl, cached: false };

  } catch (err) {
    Logger.log('❌ handleGeneratePayslipPdf: ' + err.message);
    return { ok: false, msg: err.message };
  }
}

/**
 * 管理員批次產生整月所有人的薪資單 PDF
 * params: { token, yearMonth }
 */
function handleBatchGeneratePayslipPdf(params) {
  try {
    const session = handleCheckSession(params.token);
    if (!session.ok) return { ok: false, msg: 'SESSION_INVALID' };
    if (session.user.dept !== '管理員') return { ok: false, msg: '僅限管理員' };

    const yearMonth = params.yearMonth || getCurrentYearMonth_();
    const sheet     = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EMPLOYEES);
    if (!sheet) return { ok: false, msg: '找不到員工資料表' };

    const employees = sheet.getDataRange().getValues();
    const results   = [];

    for (let i = 1; i < employees.length; i++) {
      const empId = String(employees[i][EMPLOYEE_COL.USER_ID]);
      const name  = String(employees[i][EMPLOYEE_COL.NAME]);
      const status = String(employees[i][EMPLOYEE_COL.STATUS] || '');
      if (status === '停用' || !empId) continue;

      try {
        const salaryResult = getMySalary(empId, yearMonth);
        if (!salaryResult.success) {
          results.push({ name, ok: false, msg: '無薪資記錄' });
          continue;
        }
        const existing = findExistingPayslipPdf_(empId, yearMonth);
        const url = existing || createPayslipPdf_(salaryResult.data, name, yearMonth);
        results.push({ name, ok: true, url });
        Utilities.sleep(300);
      } catch (e) {
        results.push({ name, ok: false, msg: e.message });
      }
    }

    const success = results.filter(r => r.ok).length;
    return { ok: true, msg: `已產生 ${success} 份薪資單 PDF`, results };

  } catch (err) {
    Logger.log('❌ handleBatchGeneratePayslipPdf: ' + err.message);
    return { ok: false, msg: err.message };
  }
}

// ==================== 核心邏輯 ====================

/**
 * 建立薪資單 PDF，存入 Drive，回傳下載連結
 */
function createPayslipPdf_(salary, employeeName, yearMonth) {
  // 1. 建立暫時 Google Doc
  const docTitle = `薪資單_${employeeName}_${yearMonth}`;
  const doc  = DocumentApp.create(docTitle);
  const body = doc.getBody();

  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(54).setMarginRight(54);

  // 2. 填入內容
  buildPayslipDocument_(body, salary, employeeName, yearMonth);
  doc.saveAndClose();

  // 3. 匯出為 PDF blob
  const pdfBlob = DriveApp.getFileById(doc.getId())
    .getAs('application/pdf')
    .setName(`薪資單_${employeeName}_${yearMonth}.pdf`);

  // 4. 刪除暫時 Doc
  DriveApp.getFileById(doc.getId()).setTrashed(true);

  // 5. 儲存 PDF 到 Drive 資料夾
  const folder  = getOrCreatePayslipFolder_();
  const pdfFile = folder.createFile(pdfBlob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // 6. 在檔案描述中記錄 employeeId & yearMonth（供後續查找）
  pdfFile.setDescription(`empId:${salary['員工ID']}|ym:${yearMonth}`);

  Logger.log(`✅ 薪資單 PDF 已建立: ${pdfFile.getDownloadUrl()}`);
  return pdfFile.getDownloadUrl();
}

/**
 * 在 Google Doc body 中建立薪資單版面
 */
function buildPayslipDocument_(body, salary, employeeName, yearMonth) {
  const n = v => {
    const num = parseFloat(v) || 0;
    return num === 0 ? '-' : 'NT$ ' + num.toLocaleString();
  };
  const num = v => parseFloat(v) || 0;

  // ── 標題 ──
  const title = body.appendParagraph('薪  資  單');
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1)
       .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
       .setSpacingBefore(0).setSpacingAfter(4);

  const subTitle = body.appendParagraph(yearMonth + ' 月份');
  subTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
          .setSpacingAfter(12);

  body.appendHorizontalRule();

  // ── 員工基本資料 ──
  const infoTable = body.appendTable([
    ['員工姓名', employeeName, '年月', yearMonth],
    ['員工ID',   String(salary['員工ID'] || ''), '薪資類型', String(salary['薪資類型'] || '月薪')],
    ['工作時數', (num(salary['工作時數']) || '-') + ' h', '總加班時數', (num(salary['總加班時數']) || '-') + ' h'],
  ]);
  styleInfoTable_(infoTable);
  body.appendParagraph('').setSpacingAfter(4);

  // ── 應發項目 & 扣款項目（並排兩欄） ──
  const incomeItems = [
    ['應發項目', '金額'],
    ['基本薪資',          n(salary['基本薪資'])],
    ['職務加給',          n(salary['職務加給'])],
    ['伙食費',            n(salary['伙食費'])],
    ['交通補助',          n(salary['交通補助'])],
    ['全勤獎金',          n(salary['全勤獎金'])],
    ['業績獎金',          n(salary['業績獎金'])],
    ['其他津貼',          n(salary['其他津貼'])],
    ['平日加班費',        n(salary['平日加班費'])],
    ['休息日加班費',      n(salary['休息日加班費'])],
    ['例假日加班費',      n(salary['例假日加班費'])],
    ['國定假日加班費',    n(salary['國定假日加班費'])],
    ['國定假日出勤薪資',  n(salary['國定假日出勤薪資'])],
  ];

  const deductItems = [
    ['扣款項目', '金額'],
    ['勞保費',    n(salary['勞保費'])],
    ['健保費',    n(salary['健保費'])],
    ['就業保險費', n(salary['就業保險費'])],
    ['勞退自提',  n(salary['勞退自提'])],
    ['所得稅',    n(salary['所得稅'])],
    ['請假扣款',  n(salary['請假扣款'])],
    ['早退扣款',  n(salary['早退扣款'])],
    ['病假扣款',  n(salary['病假扣款']) + (num(salary['病假時數']) > 0 ? ` (${salary['病假時數']}h)` : '')],
    ['事假扣款',  n(salary['事假扣款']) + (num(salary['事假時數']) > 0 ? ` (${salary['事假時數']}h)` : '')],
    ['福利金扣款', n(salary['福利金扣款'])],
    ['宿舍費用',  n(salary['宿舍費用'])],
    ['其他扣款',  n(salary['其他扣款'])],
  ];

  // 補齊行數使兩欄等長
  while (incomeItems.length < deductItems.length) incomeItems.push(['', '']);
  while (deductItems.length < incomeItems.length) deductItems.push(['', '']);

  // 合併成四欄表格
  const combinedRows = incomeItems.map((row, i) => [
    row[0], row[1], deductItems[i][0], deductItems[i][1]
  ]);
  const mainTable = body.appendTable(combinedRows);
  styleMainTable_(mainTable);
  body.appendParagraph('').setSpacingAfter(4);

  // ── 合計區 ──
  const grossSalary = num(salary['應發總額']);
  const netSalary   = num(salary['實發金額']);
  const totalDeduct = grossSalary - netSalary;

  const summaryTable = body.appendTable([
    ['應發總額', 'NT$ ' + grossSalary.toLocaleString(),
     '扣款總額', 'NT$ ' + totalDeduct.toLocaleString()],
    ['實  發  金  額', '', 'NT$ ' + netSalary.toLocaleString(), ''],
  ]);
  styleSummaryTable_(summaryTable, netSalary);
  body.appendParagraph('').setSpacingAfter(4);

  // ── 備註 / 請假明細 ──
  if (num(salary['病假時數']) > 0 || num(salary['事假時數']) > 0) {
    const noteP = body.appendParagraph(
      `備註：病假 ${num(salary['病假時數'])} 小時 / 事假 ${num(salary['事假時數'])} 小時`
    );
    noteP.setFontSize(9).setForegroundColor('#666666').setSpacingAfter(4);
  }

  // ── 頁尾 ──
  body.appendHorizontalRule();
  const footer = body.appendParagraph(
    `本薪資單由系統自動產生，如有疑問請洽人資。列印日期：${Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd')}`
  );
  footer.setFontSize(8).setForegroundColor('#999999')
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
}

// ==================== 表格樣式 ====================

function styleInfoTable_(table) {
  table.setBorderWidth(1);
  for (let r = 0; r < table.getNumRows(); r++) {
    for (let c = 0; c < table.getRow(r).getNumCells(); c++) {
      const cell = table.getRow(r).getCell(c);
      cell.setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(6).setPaddingRight(6);
      if (c % 2 === 0) {
        cell.setBackgroundColor('#f0fdf4');
        cell.editAsText().setBold(true);
      }
    }
  }
}

function styleMainTable_(table) {
  table.setBorderWidth(1);
  const totalRows = table.getNumRows();
  for (let r = 0; r < totalRows; r++) {
    for (let c = 0; c < 4; c++) {
      const cell = table.getRow(r).getCell(c);
      cell.setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(6).setPaddingRight(6);
      if (r === 0) {
        // header row
        cell.setBackgroundColor('#10b981');
        cell.editAsText().setBold(true).setForegroundColor('#ffffff');
      } else if (c === 0 || c === 2) {
        cell.setBackgroundColor('#f8fafc');
        cell.editAsText().setBold(false).setForegroundColor('#374151');
      } else {
        cell.editAsText().setForegroundColor('#1f2937');
      }
    }
  }
}

function styleSummaryTable_(table, netSalary) {
  table.setBorderWidth(1);
  // Row 0: 應發 / 扣款
  for (let c = 0; c < 4; c++) {
    const cell = table.getRow(0).getCell(c);
    cell.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(8).setPaddingRight(8);
    if (c % 2 === 0) {
      cell.setBackgroundColor('#e0f2fe');
      cell.editAsText().setBold(true);
    }
  }
  // Row 1: 實發金額（合計列，跨欄只用文字表現）
  for (let c = 0; c < 4; c++) {
    const cell = table.getRow(1).getCell(c);
    cell.setBackgroundColor('#10b981');
    cell.setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(8).setPaddingRight(8);
    cell.editAsText().setBold(true).setForegroundColor('#ffffff').setFontSize(13);
  }
}

// ==================== 輔助函式 ====================

function getOrCreatePayslipFolder_() {
  const folders = DriveApp.getFoldersByName(PAYSLIP_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  const folder = DriveApp.createFolder(PAYSLIP_FOLDER_NAME);
  Logger.log('📁 建立薪資單 Drive 資料夾: ' + folder.getId());
  return folder;
}

/**
 * 在 Drive 資料夾中尋找已存在的薪資單 PDF
 */
function findExistingPayslipPdf_(employeeId, yearMonth) {
  try {
    const folder = getOrCreatePayslipFolder_();
    const files  = folder.getFiles();
    const tag    = `empId:${employeeId}|ym:${yearMonth}`;
    while (files.hasNext()) {
      const file = files.next();
      if (file.getDescription() === tag) {
        return file.getDownloadUrl();
      }
    }
  } catch (e) {}
  return null;
}

/**
 * 刪除特定員工特定月份的舊 PDF（重新產生時使用）
 */
function deleteExistingPayslipPdf_(employeeId, yearMonth) {
  try {
    const folder = getOrCreatePayslipFolder_();
    const files  = folder.getFiles();
    const tag    = `empId:${employeeId}|ym:${yearMonth}`;
    while (files.hasNext()) {
      const file = files.next();
      if (file.getDescription() === tag) {
        file.setTrashed(true);
      }
    }
  } catch (e) {}
}

function getCurrentYearMonth_() {
  return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM');
}
