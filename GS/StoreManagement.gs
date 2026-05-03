// StoreManagement.gs - 多門市 / 多部門管理後端

// ── 門市設定工作表 ─────────────────────────────────────────────
function getStoreSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_STORES);
  if (!sh) {
    sh = ss.insertSheet(SHEET_STORES);
    sh.getRange(1, 1, 1, 7).setValues([[
      '門市ID', '門市名稱', '地址', '電話', '負責人', '狀態', '建立時間'
    ]]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// 讀取所有門市列表（內部用）
function getStoreList_() {
  const sh = getStoreSheet_();
  const data = sh.getDataRange().getValues();
  const stores = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    stores.push({
      storeId: String(row[0]).trim(),
      name:    String(row[1] || '').trim(),
      address: String(row[2] || '').trim(),
      phone:   String(row[3] || '').trim(),
      manager: String(row[4] || '').trim(),
      status:  String(row[5] || '啟用').trim(),
      created: row[6] ? String(row[6]).trim() : ''
    });
  }
  return stores;
}

// ── Handlers ──────────────────────────────────────────────────

function handleGetStoreList(params) {
  const session = handleCheckSession(params.token);
  if (!session.ok) return session;
  return { ok: true, stores: getStoreList_() };
}

function handleCreateStore(params) {
  const session = handleCheckSession(params.token);
  if (!session.ok) return session;
  if (session.user.dept !== '管理員') return { ok: false, error: '權限不足' };
  if (!params.name) return { ok: false, error: '請填寫門市名稱' };

  const sh = getStoreSheet_();
  const storeId = 'STORE-' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMddHHmmss');
  sh.appendRow([
    storeId,
    params.name,
    params.address  || '',
    params.phone    || '',
    params.manager  || '',
    '啟用',
    Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss')
  ]);
  return { ok: true, storeId, message: '門市已新增' };
}

function handleUpdateStore(params) {
  const session = handleCheckSession(params.token);
  if (!session.ok) return session;
  if (session.user.dept !== '管理員') return { ok: false, error: '權限不足' };
  if (!params.storeId) return { ok: false, error: '缺少門市ID' };

  const sh = getStoreSheet_();
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== String(params.storeId).trim()) continue;
    if (params.name    !== undefined) sh.getRange(i + 1, 2).setValue(params.name);
    if (params.address !== undefined) sh.getRange(i + 1, 3).setValue(params.address);
    if (params.phone   !== undefined) sh.getRange(i + 1, 4).setValue(params.phone);
    if (params.manager !== undefined) sh.getRange(i + 1, 5).setValue(params.manager);
    if (params.status  !== undefined) sh.getRange(i + 1, 6).setValue(params.status);
    return { ok: true, message: '門市已更新' };
  }
  return { ok: false, error: '找不到該門市' };
}

// 指派員工到門市（寫入員工名單 欄 I = index 8）
function handleAssignEmployeeStore(params) {
  const session = handleCheckSession(params.token);
  if (!session.ok) return session;
  if (session.user.dept !== '管理員') return { ok: false, error: '權限不足' };
  if (!params.employeeId) return { ok: false, error: '缺少員工ID' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_EMPLOYEES);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][EMPLOYEE_COL.USER_ID]).trim() !== String(params.employeeId).trim()) continue;
    sh.getRange(i + 1, EMPLOYEE_COL.STORE + 1).setValue(params.storeId || '');
    return { ok: true, message: '已指派門市' };
  }
  return { ok: false, error: '找不到該員工' };
}

// 跨門市儀表板：各門市當月人數、薪資成本、加班時數、出勤天數
function handleGetStoreDashboard(params) {
  const session = handleCheckSession(params.token);
  if (!session.ok) return session;
  if (session.user.dept !== '管理員') return { ok: false, error: '權限不足' };

  const yearMonth = params.yearMonth || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM');
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 員工 → 門市、姓名
  const empSh   = ss.getSheetByName(SHEET_EMPLOYEES);
  const empData = empSh.getDataRange().getValues();
  const empStoreMap = {};
  const empNameMap  = {};
  for (let i = 1; i < empData.length; i++) {
    const uid = String(empData[i][EMPLOYEE_COL.USER_ID] || '').trim();
    if (!uid) continue;
    empStoreMap[uid] = String(empData[i][EMPLOYEE_COL.STORE] || '').trim();
    empNameMap[uid]  = String(empData[i][EMPLOYEE_COL.NAME]  || '').trim();
  }

  // 當月薪資（月薪資記錄優先，無則用設定基本薪）
  const payrollSh   = ss.getSheetByName('月薪資記錄');
  const payByEmp    = {};
  if (payrollSh) {
    const pr = payrollSh.getDataRange().getValues();
    for (let i = 1; i < pr.length; i++) {
      if (String(pr[i][1] || '').trim() !== yearMonth) continue;
      const uid = String(pr[i][0] || '').trim();
      if (uid) payByEmp[uid] = parseFloat(pr[i][14]) || 0;
    }
  }
  const salarySh = ss.getSheetByName('員工薪資設定');
  const baseByEmp = {};
  if (salarySh) {
    const sd = salarySh.getDataRange().getValues();
    for (let i = 1; i < sd.length; i++) {
      const uid = String(sd[i][0] || '').trim();
      if (uid) baseByEmp[uid] = parseFloat(sd[i][2]) || 0;
    }
  }

  // 出勤天數
  const attSh   = ss.getSheetByName(SHEET_ATTENDANCE);
  const attData = attSh ? attSh.getDataRange().getValues() : [];
  const attByEmp = {};
  for (let i = 1; i < attData.length; i++) {
    const dateStr = attData[i][1]
      ? Utilities.formatDate(new Date(attData[i][1]), 'Asia/Taipei', 'yyyy-MM') : '';
    if (dateStr !== yearMonth) continue;
    const uid = String(attData[i][0] || '').trim();
    if (!attByEmp[uid]) attByEmp[uid] = new Set();
    const dayStr = Utilities.formatDate(new Date(attData[i][1]), 'Asia/Taipei', 'yyyy-MM-dd');
    attByEmp[uid].add(dayStr);
  }

  // 加班時數 — 欄位: [0]=申請ID [1]=員工ID [2]=姓名 [3]=日期 [6]=時數 [9]=狀態
  const otSh   = ss.getSheetByName(SHEET_OVERTIME);
  const otData = otSh ? otSh.getDataRange().getValues() : [];
  const otByEmp = {};
  for (let i = 1; i < otData.length; i++) {
    if (String(otData[i][9] || '').toLowerCase() !== 'approved') continue;
    const dateStr = otData[i][3]
      ? Utilities.formatDate(new Date(otData[i][3]), 'Asia/Taipei', 'yyyy-MM') : '';
    if (dateStr !== yearMonth) continue;
    const uid = String(otData[i][1] || '').trim();
    if (!otByEmp[uid]) otByEmp[uid] = 0;
    otByEmp[uid] += parseFloat(otData[i][6]) || 0;
  }

  // 彙整門市
  const stores = getStoreList_();
  const buckets = {};
  stores.forEach(s => {
    buckets[s.storeId] = {
      storeId: s.storeId, name: s.name, status: s.status,
      employeeCount: 0, totalSalaryCost: 0,
      totalOvertimeHours: 0, totalAttendanceDays: 0
    };
  });
  buckets['__none__'] = {
    storeId: '__none__', name: '未指派門市', status: '啟用',
    employeeCount: 0, totalSalaryCost: 0,
    totalOvertimeHours: 0, totalAttendanceDays: 0
  };

  Object.keys(empStoreMap).forEach(uid => {
    const sid    = empStoreMap[uid] || '__none__';
    const bucket = buckets[sid] || buckets['__none__'];
    bucket.employeeCount++;
    bucket.totalSalaryCost     += payByEmp[uid] || baseByEmp[uid] || 0;
    bucket.totalOvertimeHours  += otByEmp[uid]  || 0;
    bucket.totalAttendanceDays += attByEmp[uid] ? attByEmp[uid].size : 0;
  });

  const result = Object.values(buckets)
    .filter(b => b.employeeCount > 0 || b.storeId !== '__none__');
  return { ok: true, yearMonth, stores: result };
}

// 指定門市的員工列表
function handleGetStoreEmployees(params) {
  const session = handleCheckSession(params.token);
  if (!session.ok) return session;
  if (session.user.dept !== '管理員') return { ok: false, error: '權限不足' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_EMPLOYEES);
  const data = sh.getDataRange().getValues();
  const employees = [];

  for (let i = 1; i < data.length; i++) {
    const uid = String(data[i][EMPLOYEE_COL.USER_ID] || '').trim();
    if (!uid) continue;
    const storeId = String(data[i][EMPLOYEE_COL.STORE] || '').trim();
    if (params.storeId && params.storeId !== '__none__' && storeId !== params.storeId) continue;
    if (params.storeId === '__none__' && storeId) continue;
    employees.push({
      uid,
      name:    String(data[i][EMPLOYEE_COL.NAME]   || '').trim(),
      dept:    String(data[i][EMPLOYEE_COL.DEPT]   || '').trim(),
      storeId: storeId,
      status:  String(data[i][EMPLOYEE_COL.STATUS] || '').trim()
    });
  }
  return { ok: true, employees };
}
