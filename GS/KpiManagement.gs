// KpiManagement.gs - 績效獎金連動後端

// ── KPI 類型定義 ──────────────────────────────────────────────
const KPI_TYPE_DEFS = {
  sales_target:    { label: '業績目標',  autoCalc: false, lowerIsBetter: false },
  attendance_rate: { label: '出勤率',    autoCalc: true,  lowerIsBetter: false },
  complaint_count: { label: '客訴次數',  autoCalc: false, lowerIsBetter: true  }
};

// ── KPI 工作表 ────────────────────────────────────────────────
function getKpiSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_KPI);
  if (!sh) {
    sh = ss.insertSheet(SHEET_KPI);
    sh.getRange(1, 1, 1, 12).setValues([[
      'KPI_ID', '員工ID', '員工姓名', '年月', 'KPI項目',
      '目標值', '實際值', '達成率(%)', '獎金上限', '計算獎金', '設定人', '建立時間'
    ]]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// ── 達成率 & 獎金計算 ─────────────────────────────────────────
function calcKpiBonus_(kpiType, target, actual, bonusCap) {
  const def = KPI_TYPE_DEFS[kpiType];
  const t = parseFloat(target)   || 0;
  const a = parseFloat(actual)   || 0;
  const c = parseFloat(bonusCap) || 0;
  let rate = 0;
  let bonus = 0;

  if (def && def.lowerIsBetter) {
    // 客訴次數：實際 ≤ 目標 → 100%，超過則線性扣減
    rate  = t <= 0 ? 100 : Math.min(150, Math.max(0, Math.round((1 - (a - t) / Math.max(t, 1)) * 100)));
    if (a <= t) {
      bonus = c;
    } else {
      bonus = Math.max(0, c * (1 - (a - t) / Math.max(t, 1)));
    }
  } else {
    // 一般指標：達成率 = 實際 / 目標 × 100%，獎金比例等比
    rate  = t <= 0 ? 0 : Math.min(150, Math.round((a / t) * 100));
    bonus = c * Math.min(rate, 100) / 100;
  }
  return { achievementRate: rate, calcBonus: Math.round(bonus) };
}

// ── Handlers ──────────────────────────────────────────────────

function handleGetKpiList(params) {
  const session = handleCheckSession(params.token);
  if (!session.ok) return session;

  const yearMonth = params.yearMonth || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM');
  const sh   = getKpiSheet_();
  const data = sh.getDataRange().getValues();
  const list = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] || String(row[3]) !== yearMonth) continue;
    if (session.user.dept !== '管理員' && String(row[1]) !== session.user.userId) continue;
    if (params.employeeId && String(row[1]) !== String(params.employeeId)) continue;
    list.push({
      kpiId:           String(row[0]),
      employeeId:      String(row[1]),
      employeeName:    String(row[2] || ''),
      yearMonth:       String(row[3]),
      kpiType:         String(row[4]),
      target:          parseFloat(row[5]) || 0,
      actual:          parseFloat(row[6]) || 0,
      achievementRate: parseFloat(row[7]) || 0,
      bonusCap:        parseFloat(row[8]) || 0,
      calcBonus:       parseFloat(row[9]) || 0,
      setBy:           String(row[10] || ''),
      created:         String(row[11] || '')
    });
  }
  return { ok: true, yearMonth, list };
}

function handleSaveKpi(params) {
  const session = handleCheckSession(params.token);
  if (!session.ok) return session;
  if (session.user.dept !== '管理員') return { ok: false, error: '權限不足' };

  const required = ['employeeId', 'employeeName', 'yearMonth', 'kpiType', 'bonusCap'];
  for (const f of required) {
    if (params[f] === undefined || params[f] === '') return { ok: false, error: `缺少欄位: ${f}` };
  }

  const target   = parseFloat(params.target)   || 0;
  const actual   = parseFloat(params.actual)   || 0;
  const bonusCap = parseFloat(params.bonusCap) || 0;
  const { achievementRate, calcBonus } = calcKpiBonus_(params.kpiType, target, actual, bonusCap);

  const sh   = getKpiSheet_();
  const data = sh.getDataRange().getValues();
  const now  = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
  const setBy = (session.user && session.user.name) ? session.user.name : (session.user && session.user.userId ? session.user.userId : '');

  // 更新現有 kpiId
  if (params.kpiId) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== String(params.kpiId)) continue;
      sh.getRange(i + 1, 1, 1, 12).setValues([[
        params.kpiId, params.employeeId, params.employeeName,
        params.yearMonth, params.kpiType,
        target, actual, achievementRate, bonusCap, calcBonus, setBy, now
      ]]);
      return { ok: true, kpiId: params.kpiId, calcBonus, achievementRate, message: 'KPI已更新' };
    }
  }

  // 同員工+月份+KPI類型 → 覆蓋更新
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(params.employeeId) &&
        String(data[i][3]) === String(params.yearMonth)  &&
        String(data[i][4]) === String(params.kpiType)) {
      sh.getRange(i + 1, 6, 1, 7).setValues([[
        target, actual, achievementRate, bonusCap, calcBonus, setBy, now
      ]]);
      return { ok: true, kpiId: String(data[i][0]), calcBonus, achievementRate, message: 'KPI已更新' };
    }
  }

  // 新增
  const kpiId = 'KPI-' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMddHHmmss') +
                '-' + Math.floor(Math.random() * 900 + 100);
  sh.appendRow([
    kpiId, params.employeeId, params.employeeName,
    params.yearMonth, params.kpiType,
    target, actual, achievementRate, bonusCap, calcBonus, setBy, now
  ]);
  return { ok: true, kpiId, calcBonus, achievementRate, message: 'KPI已新增' };
}

function handleDeleteKpi(params) {
  const session = handleCheckSession(params.token);
  if (!session.ok) return session;
  if (session.user.dept !== '管理員') return { ok: false, error: '權限不足' };
  if (!params.kpiId) return { ok: false, error: '缺少kpiId' };

  const sh   = getKpiSheet_();
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(params.kpiId)) {
      sh.deleteRow(i + 1);
      return { ok: true, message: 'KPI已刪除' };
    }
  }
  return { ok: false, error: '找不到該KPI' };
}

// 自動計算所有員工本月出勤率 KPI
function handleAutoCalcAttendanceKpi(params) {
  const session = handleCheckSession(params.token);
  if (!session.ok) return session;
  if (session.user.dept !== '管理員') return { ok: false, error: '權限不足' };

  const yearMonth = params.yearMonth || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM');
  const [yr, mo]  = yearMonth.split('-').map(Number);
  const daysInMonth = new Date(yr, mo, 0).getDate();

  // 計算當月工作日（週一～週五）
  let workdays = 0;
  for (let d = 1; d <= daysInMonth; d++) {
    const dow = new Date(yr, mo - 1, d).getDay();
    if (dow !== 0 && dow !== 6) workdays++;
  }

  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const empSh  = ss.getSheetByName(SHEET_EMPLOYEES);
  const empData = empSh.getDataRange().getValues();

  // 出勤天數（有打卡記錄的去重日期數）
  const attSh   = ss.getSheetByName(SHEET_ATTENDANCE);
  const attData = attSh ? attSh.getDataRange().getValues() : [];
  const attByEmp = {};
  for (let i = 1; i < attData.length; i++) {
    if (!attData[i][1]) continue;
    const ym = Utilities.formatDate(new Date(attData[i][1]), 'Asia/Taipei', 'yyyy-MM');
    if (ym !== yearMonth) continue;
    const uid = String(attData[i][0] || '').trim();
    if (!attByEmp[uid]) attByEmp[uid] = new Set();
    attByEmp[uid].add(Utilities.formatDate(new Date(attData[i][1]), 'Asia/Taipei', 'yyyy-MM-dd'));
  }

  const bonusCap = parseFloat(params.attendanceBonusCap) || 1000;
  let updated = 0;

  for (let i = 1; i < empData.length; i++) {
    const uid    = String(empData[i][EMPLOYEE_COL.USER_ID] || '').trim();
    const name   = String(empData[i][EMPLOYEE_COL.NAME]    || '').trim();
    const status = String(empData[i][EMPLOYEE_COL.STATUS]  || '').trim();
    if (!uid || status === '離職') continue;

    const attended = attByEmp[uid] ? attByEmp[uid].size : 0;
    const rate     = workdays > 0 ? Math.round((attended / workdays) * 100) : 0;

    handleSaveKpi({
      token:         params.token,
      employeeId:    uid,
      employeeName:  name,
      yearMonth,
      kpiType:       'attendance_rate',
      target:        100,
      actual:        rate,
      bonusCap
    });
    updated++;
  }
  return { ok: true, updatedCount: updated, message: `已自動計算 ${updated} 位員工出勤率 KPI（工作日基準: ${workdays} 天）` };
}

// 各員工 KPI 總獎金彙整（用於薪資列表顯示）
function handleGetKpiSummary(params) {
  const session = handleCheckSession(params.token);
  if (!session.ok) return session;
  if (session.user.dept !== '管理員') return { ok: false, error: '權限不足' };

  const yearMonth = params.yearMonth || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM');
  const sh   = getKpiSheet_();
  const data = sh.getDataRange().getValues();
  const byEmp = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] || String(row[3]) !== yearMonth) continue;
    const uid = String(row[1]);
    if (!byEmp[uid]) byEmp[uid] = { employeeId: uid, employeeName: String(row[2]), totalBonus: 0, items: [] };
    byEmp[uid].totalBonus += parseFloat(row[9]) || 0;
    byEmp[uid].items.push({
      kpiType:        String(row[4]),
      target:         parseFloat(row[5]) || 0,
      actual:         parseFloat(row[6]) || 0,
      achievementRate:parseFloat(row[7]) || 0,
      bonusCap:       parseFloat(row[8]) || 0,
      calcBonus:      parseFloat(row[9]) || 0
    });
  }
  return { ok: true, yearMonth, summary: Object.values(byEmp) };
}

// 供薪資計算模組呼叫：取得員工當月 KPI 獎金總額
function getKpiBonusForEmployee(employeeId, yearMonth) {
  try {
    const sh   = getKpiSheet_();
    const data = sh.getDataRange().getValues();
    let total  = 0;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(employeeId) && String(data[i][3]) === yearMonth) {
        total += parseFloat(data[i][9]) || 0;
      }
    }
    return total;
  } catch (e) {
    return 0;
  }
}
