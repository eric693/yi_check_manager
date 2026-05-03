// ExpenseManagement.gs - 費用申報與報銷模組

const SHEET_EXPENSE = '費用申報';

const EXPENSE_CATEGORIES = ['餐費', '交通費', '耗材採購', '住宿費', '娛樂費', '其他'];

const EXPENSE_STATUS = {
  PENDING:  '待審核',
  APPROVED: '核准',
  REJECTED: '拒絕'
};

// ==================== Sheet 初始化 ====================

function getExpenseSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_EXPENSE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_EXPENSE);
    const headers = [
      '費用ID', '員工ID', '員工姓名', '申請時間',
      '費用類別', '金額', '費用日期', '說明',
      '審核狀態', '審核人ID', '審核人姓名', '審核時間', '審核意見',
      '入薪狀態', '入薪月份'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
         .setFontWeight('bold').setBackground('#10b981').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    Logger.log('✅ 建立費用申報工作表');
  }
  return sheet;
}

// ==================== 核心函數 ====================

function submitExpenseClaim(params) {
  try {
    const session = handleCheckSession(params.token);
    if (!session.ok) return { ok: false, msg: 'SESSION_INVALID' };

    const { category, amount, expenseDate, description } = params;
    if (!category || !amount || !expenseDate) return { ok: false, msg: '請填寫費用類別、金額與費用日期' };

    const parsedAmount = parseFloat(amount);
    if (isNaN(parsedAmount) || parsedAmount <= 0) return { ok: false, msg: '金額必須大於 0' };

    const sheet = getExpenseSheet();
    const expenseId = 'EXP-' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 1000);
    const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');

    sheet.appendRow([
      expenseId,
      session.user.userId,
      session.user.name,
      now,
      category,
      parsedAmount,
      expenseDate,
      description || '',
      EXPENSE_STATUS.PENDING,
      '', '', '', '',
      0, ''
    ]);

    // 通知管理員
    try {
      notifyAdminsNewExpense(session.user.name, category, parsedAmount, expenseDate, description || '');
    } catch (e) {
      Logger.log('⚠️ 通知管理員失敗: ' + e);
    }

    Logger.log(`✅ 費用申報成功: ${expenseId}, 員工: ${session.user.name}, 金額: ${parsedAmount}`);
    return { ok: true, msg: '費用申報已送出，等待審核', expenseId };
  } catch (e) {
    Logger.log('❌ submitExpenseClaim: ' + e);
    return { ok: false, msg: e.toString() };
  }
}

function getMyExpenseClaims(params) {
  try {
    const session = handleCheckSession(params.token);
    if (!session.ok) return { ok: false, msg: 'SESSION_INVALID' };

    const sheet = getExpenseSheet();
    const data = sheet.getDataRange().getValues();
    const records = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[1]) !== session.user.userId) continue;
      if (params.yearMonth && !String(row[3]).startsWith(params.yearMonth)) continue;
      records.push({
        expenseId:    row[0],
        employeeId:   row[1],
        employeeName: row[2],
        appliedAt:    row[3],
        category:     row[4],
        amount:       row[5],
        expenseDate:  row[6],
        description:  row[7],
        status:       row[8],
        reviewerId:   row[9],
        reviewerName: row[10],
        reviewedAt:   row[11],
        reviewComment:row[12],
        salaryIncluded: row[13],
        salaryMonth:  row[14],
        rowIndex: i + 1
      });
    }

    return { ok: true, records: records.reverse() };
  } catch (e) {
    Logger.log('❌ getMyExpenseClaims: ' + e);
    return { ok: false, msg: e.toString() };
  }
}

function getPendingExpenseClaims(params) {
  try {
    const session = handleCheckSession(params.token);
    if (!session.ok) return { ok: false, msg: 'SESSION_INVALID' };
    if (session.user.dept !== '管理員') return { ok: false, msg: '僅限管理員' };

    const sheet = getExpenseSheet();
    const data = sheet.getDataRange().getValues();
    const records = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const statusFilter = params.status || EXPENSE_STATUS.PENDING;
      if (String(row[8]) !== statusFilter) continue;
      records.push({
        expenseId:    row[0],
        employeeId:   row[1],
        employeeName: row[2],
        appliedAt:    row[3],
        category:     row[4],
        amount:       row[5],
        expenseDate:  row[6],
        description:  row[7],
        status:       row[8],
        salaryIncluded: row[13],
        salaryMonth:  row[14],
        rowIndex: i + 1
      });
    }

    return { ok: true, records: records.reverse() };
  } catch (e) {
    Logger.log('❌ getPendingExpenseClaims: ' + e);
    return { ok: false, msg: e.toString() };
  }
}

function reviewExpenseClaim(params) {
  try {
    const session = handleCheckSession(params.token);
    if (!session.ok) return { ok: false, msg: 'SESSION_INVALID' };
    if (session.user.dept !== '管理員') return { ok: false, msg: '僅限管理員' };

    const { expenseId, action, comment } = params;
    if (!expenseId || !action) return { ok: false, msg: '缺少參數' };

    const newStatus = action === 'approve' ? EXPENSE_STATUS.APPROVED : EXPENSE_STATUS.REJECTED;

    const sheet = getExpenseSheet();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== expenseId) continue;

      const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
      sheet.getRange(i + 1, 9).setValue(newStatus);
      sheet.getRange(i + 1, 10).setValue(session.user.userId);
      sheet.getRange(i + 1, 11).setValue(session.user.name);
      sheet.getRange(i + 1, 12).setValue(now);
      sheet.getRange(i + 1, 13).setValue(comment || '');

      // 通知申請人
      try {
        const employeeId = String(data[i][1]);
        const employeeName = String(data[i][2]);
        const category = String(data[i][4]);
        const amount = parseFloat(data[i][5]);
        notifyExpenseResult(employeeId, employeeName, category, amount, action === 'approve', comment || '');
      } catch (e) {
        Logger.log('⚠️ 通知員工失敗: ' + e);
      }

      return { ok: true, msg: `已${newStatus}該費用申報` };
    }

    return { ok: false, msg: '找不到該筆費用申報' };
  } catch (e) {
    Logger.log('❌ reviewExpenseClaim: ' + e);
    return { ok: false, msg: e.toString() };
  }
}

/**
 * 取得指定月份已核准且尚未入薪的費用，用於薪資計算
 */
function getApprovedExpensesForSalary(employeeId, yearMonth) {
  try {
    const sheet = getExpenseSheet();
    const data = sheet.getDataRange().getValues();
    let total = 0;
    const rows = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[1]) !== employeeId) continue;
      if (String(row[8]) !== EXPENSE_STATUS.APPROVED) continue;
      if (parseInt(row[13]) === 1) continue; // 已入薪
      const expenseDate = String(row[6]);
      if (!expenseDate.startsWith(yearMonth)) continue;
      total += parseFloat(row[5]) || 0;
      rows.push(i + 1);
    }

    return { total, rows };
  } catch (e) {
    Logger.log('❌ getApprovedExpensesForSalary: ' + e);
    return { total: 0, rows: [] };
  }
}

/**
 * 標記費用已入薪
 */
function markExpensesAsSalaried(rows, yearMonth) {
  const sheet = getExpenseSheet();
  rows.forEach(rowIdx => {
    sheet.getRange(rowIdx, 14).setValue(1);
    sheet.getRange(rowIdx, 15).setValue(yearMonth);
  });
}

// ==================== Handlers ====================

function handleSubmitExpense(params) {
  return submitExpenseClaim(params);
}

function handleGetMyExpenses(params) {
  return getMyExpenseClaims(params);
}

function handleGetPendingExpenses(params) {
  return getPendingExpenseClaims(params);
}

function handleReviewExpense(params) {
  return reviewExpenseClaim(params);
}

function handleGetAllExpenses(params) {
  try {
    const session = handleCheckSession(params.token);
    if (!session.ok) return { ok: false, msg: 'SESSION_INVALID' };
    if (session.user.dept !== '管理員') return { ok: false, msg: '僅限管理員' };

    const sheet = getExpenseSheet();
    const data = sheet.getDataRange().getValues();
    const records = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (params.yearMonth && !String(row[3]).startsWith(params.yearMonth)) continue;
      records.push({
        expenseId: row[0], employeeId: row[1], employeeName: row[2],
        appliedAt: row[3], category: row[4], amount: row[5],
        expenseDate: row[6], description: row[7], status: row[8],
        reviewerName: row[10], reviewedAt: row[11], reviewComment: row[12],
        salaryIncluded: row[13], salaryMonth: row[14], rowIndex: i + 1
      });
    }

    return { ok: true, records: records.reverse() };
  } catch (e) {
    return { ok: false, msg: e.toString() };
  }
}
