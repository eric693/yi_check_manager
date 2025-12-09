// Constants.gs - 系統常數定義（完整版）

// ==================== LINE 登入設定 ====================
const LINE_CHANNEL_ID     = PropertiesService.getScriptProperties().getProperty("LINE_CHANNEL_ID");
const LINE_CHANNEL_SECRET = PropertiesService.getScriptProperties().getProperty("LINE_CHANNEL_SECRET");
const LINE_REDIRECT_URL   = "https://eric693.github.io/check_manager_plus/";

// ==================== Session 設定 ====================
const SESSION_TTL_MS = 1000 * 60 * 60 * 24; // 1 天
const TOKEN_LENGTH   = 36;

// ==================== 工作表名稱 ====================
// 基礎系統
const SHEET_EMPLOYEES  = '員工名單';
const SHEET_ATTENDANCE = '打卡紀錄';
const SHEET_SESSION    = 'Session';
const SHEET_LOCATIONS  = '打卡地點表';

// 加班系統（如果有）
const SHEET_OVERTIME_RECORDS = '加班紀錄';

// ⭐ 請假系統（新增）
const SHEET_LEAVE_RECORDS = '請假紀錄';
const SHEET_LEAVE_BALANCE = '員工假期額度';

// ==================== 員工資料表欄位索引 ====================
// 根據您的實際表格結構
// A-userId, B-email, C-displayName, D-pictureUrl, E-建立時間, F-部門, G-到職日期, H-狀態
const EMPLOYEE_COL = {
  USER_ID: 0,      // A 欄：userId
  EMAIL: 1,        // B 欄：email
  NAME: 2,         // C 欄：displayName
  PICTURE: 3,      // D 欄：pictureUrl
  CREATED: 4,      // E 欄：建立時間
  DEPT: 5,         // F 欄：部門
  HIRE_DATE: 6,    // G 欄：到職日期 ⭐
  STATUS: 7        // H 欄：狀態
};

// ==================== 假期類型定義 ====================
const LEAVE_TYPES = {
  ANNUAL_LEAVE: {
    code: 'ANNUAL_LEAVE',
    name: '特休假',
    maxDays: null,  // 根據年資計算
    isPaid: true,
    requiresProof: false
  },
  SICK_LEAVE: {
    code: 'SICK_LEAVE',
    name: '病假',
    maxDays: 30,
    isPaid: true,   // 半薪（30天內）
    requiresProof: true
  },
  PERSONAL_LEAVE: {
    code: 'PERSONAL_LEAVE',
    name: '事假',
    maxDays: 14,
    isPaid: false,
    requiresProof: false
  },
  MARRIAGE_LEAVE: {
    code: 'MARRIAGE_LEAVE',
    name: '婚假',
    maxDays: 8,
    isPaid: true,
    requiresProof: true
  },
  BEREAVEMENT_LEAVE: {
    code: 'BEREAVEMENT_LEAVE',
    name: '喪假',
    maxDays: 0,  // 依親屬關係而定
    isPaid: true,
    requiresProof: true
  },
  MATERNITY_LEAVE: {
    code: 'MATERNITY_LEAVE',
    name: '產假',
    maxDays: 56,
    isPaid: true,
    requiresProof: true
  },
  PATERNITY_LEAVE: {
    code: 'PATERNITY_LEAVE',
    name: '陪產假',
    maxDays: 7,
    isPaid: true,
    requiresProof: true
  },
  FAMILY_CARE_LEAVE: {
    code: 'FAMILY_CARE_LEAVE',
    name: '家庭照顧假',
    maxDays: 7,
    isPaid: false,
    requiresProof: false
  },
  MENSTRUAL_LEAVE: {
    code: 'MENSTRUAL_LEAVE',
    name: '生理假',
    maxDays: 12,  // 每月1天
    isPaid: true,  // 半薪
    requiresProof: false
  }
};

// ==================== 喪假天數對照表 ====================
const BEREAVEMENT_DAYS = {
  '父母': 8,
  '配偶': 8,
  '養父母': 8,
  '繼父母': 8,
  '子女': 6,
  '祖父母': 6,
  '配偶父母': 6,
  '配偶養父母': 6,
  '配偶繼父母': 6,
  '曾祖父母': 3,
  '兄弟姊妹': 3,
  '配偶祖父母': 3
};

// ==================== 特休假計算規則（台灣勞基法）====================
const ANNUAL_LEAVE_RULES = [
  { minMonths: 0, maxMonths: 6, days: 0, description: '未滿6個月' },
  { minMonths: 6, maxMonths: 12, days: 3, description: '6個月以上未滿1年' },
  { minYears: 1, maxYears: 2, days: 7, description: '1年以上未滿2年' },
  { minYears: 2, maxYears: 3, days: 10, description: '2年以上未滿3年' },
  { minYears: 3, maxYears: 5, days: 14, description: '3年以上未滿5年' },
  { minYears: 5, maxYears: 10, days: 15, description: '5年以上未滿10年' },
  { minYears: 10, maxYears: 999, days: 15, extraPerYear: 1, maxDays: 30, description: '10年以上' }
];

// ==================== 請假狀態 ====================
const LEAVE_STATUS = {
  PENDING: 'PENDING',      // 待審核
  APPROVED: 'APPROVED',    // 已核准
  REJECTED: 'REJECTED',    // 已拒絕
  CANCELLED: 'CANCELLED'   // 已取消（預留）
};

// ==================== 系統設定 ====================
const LEAVE_SYSTEM_CONFIG = {
  // 是否允許員工查看其他人的請假記錄
  allowViewOthersLeave: false,
  
  // 是否允許申請未來日期的請假
  allowFutureLeave: true,
  
  // 可以提前幾天申請請假
  advanceNoticeDays: 0,
  
  // 是否允許申請半天假
  allowHalfDay: true,
  
  // 是否在假期不足時仍可申請（變成負數）
  allowNegativeBalance: false
};

// ==================== 工具函數 ====================

/**
 * 取得假期類型資訊
 */
function getLeaveTypeInfo(leaveTypeCode) {
  return LEAVE_TYPES[leaveTypeCode] || null;
}

/**
 * 取得當前年度
 */
function getCurrentLeaveYear() {
  return new Date().getFullYear();
}

/**
 * 檢查是否為管理員
 */
function isAdmin(user) {
  return user && user.dept === '管理員';
}

/**
 * 格式化日期
 */
function formatDate(date) {
  const d = new Date(date);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}


// ==================== 系統版本資訊 ====================
const SYSTEM_VERSION = {
  version: '0.7.0',
  buildDate: '2025-10-11',
  features: [
    '打卡系統',
    '補打卡審核',
    '打卡紀錄',
    '地點管理',
    '加班管理',
    '請假管理'
  ]
};