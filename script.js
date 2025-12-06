// 使用 CDN 或絕對路徑來載入 JSON 檔案
// 注意：本檔案需要依賴 config.js，請確保它在腳本之前被載入。

let currentLang = localStorage.getItem("lang");
let currentMonthDate = new Date();
let translations = {};
let monthDataCache = {}; // 新增：用於快取月份打卡資料
let userId = localStorage.getItem("sessionUserId");
let todayShiftCache = null; // 快取今日排班
let weekShiftCache = null;  // 快取本週排班
// 載入語系檔
async function loadTranslations(lang) {
    try {
        const res = await fetch(`https://eric693.github.io/New_Attendance-System/i18n/${lang}.json`);
        if (!res.ok) {
            throw new Error(`HTTP 錯誤: ${res.status}`);
        }
        translations = await res.json();
        currentLang = lang;
        localStorage.setItem("lang", lang);
        renderTranslations();
    } catch (err) {
        console.error("載入語系失敗:", err);
    }
}

// 翻譯函式
function t(code, params = {}) {
    let text = translations[code] || code;
    
    // 檢查並替換參數中的變數
    for (const key in params) {
        // 在替換之前，先翻譯參數的值
        let paramValue = params[key];
        if (paramValue in translations) {
            paramValue = translations[paramValue];
        }
        
        text = text.replace(`{${key}}`, paramValue);
    }
    return text;
}

// renderTranslations 可接受一個容器參數
function renderTranslations(container = document) {
    // 翻譯網頁標題（只在整頁翻譯時執行）
    if (container === document) {
        document.title = t("APP_TITLE");
    }

    // 處理靜態內容：[data-i18n]
    const elementsToTranslate = container.querySelectorAll('[data-i18n]');
    elementsToTranslate.forEach(element => {
        const key = element.getAttribute('data-i18n');
        const translatedText = t(key);
        
        // 檢查翻譯結果是否為空字串，或是否回傳了原始鍵值
        if (translatedText !== key) {
            if (element.tagName === 'INPUT') {
                element.placeholder = translatedText;
            } else {
                element.textContent = translatedText;
            }
        }
    });

    // ✨ 新增邏輯：處理動態內容的翻譯，使用 [data-i18n-key]
    const dynamicElements = container.querySelectorAll('[data-i18n-key]');
    dynamicElements.forEach(element => {
        const key = element.getAttribute('data-i18n-key');
        if (key) {
             const translatedText = t(key);
             
             // 只有當翻譯結果不是原始鍵值時才進行更新
             if (translatedText !== key) {
                 element.textContent = translatedText;
             }
        }
    });
}

/**
 * 透過 fetch API 呼叫後端 API。
 * @param {string} action - API 的動作名稱。
 * @param {string} [loadingId="loading"] - 顯示 loading 狀態的 DOM 元素 ID。
 * @returns {Promise<object>} - 回傳一個包含 API 回應資料的 Promise。
 */
async function callApifetch(action, loadingId = "loading") {
    const token = localStorage.getItem("sessionToken");
    const url = `${API_CONFIG.apiUrl}?action=${action}&token=${token}`;
    
    const loadingEl = document.getElementById(loadingId);
    if (loadingEl) loadingEl.style.display = "block";
    
    try {
        const response = await fetch(url);
        
        if (!response.ok) {
            throw new Error(`HTTP 錯誤: ${response.status}`);
        }
        
        const data = await response.json();
        
        // ✅✅✅ 雙向格式統一（關鍵修正）
        // 1. 如果後端回傳 success，轉換為 ok
        if (data.success !== undefined && data.ok === undefined) {
            data.ok = data.success;
        }
        
        // 2. 如果後端回傳 ok，轉換為 success
        if (data.ok !== undefined && data.success === undefined) {
            data.success = data.ok;
        }
        
        // 3. 如果後端回傳 data，轉換為 records
        if (data.data !== undefined && data.records === undefined) {
            data.records = data.data;
        }
        
        // 4. 如果後端回傳 records，轉換為 data
        if (data.records !== undefined && data.data === undefined) {
            data.data = data.records;
        }
        
        return data;
    } catch (error) {
        showNotification(t("CONNECTION_FAILED"), "error");
        console.error("API 呼叫失敗:", error);
        throw error;
    } finally {
        if (loadingEl) loadingEl.style.display = "none";
    }
}

// ==================== 📊 管理員匯出所有員工報表功能 ====================

/**
 * 管理員匯出所有員工的出勤報表
 * @param {string} monthKey - 月份，格式: "YYYY-MM"
 */
async function exportAllEmployeesReport(monthKey) {
    const exportBtn = document.getElementById('admin-export-all-btn');
    const loadingText = t('EXPORT_LOADING') || '正在準備報表...';
    
    showNotification(loadingText, 'warning');
    
    if (exportBtn) {
        generalButtonState(exportBtn, 'processing', loadingText);
    }
    
    try {
        // 呼叫 API 取得所有員工的出勤資料（不傳 userId）
        const res = await callApifetch(`getAttendanceDetails&month=${monthKey}`);
        
        if (!res.ok || !res.records || res.records.length === 0) {
            showNotification(t('EXPORT_NO_DATA') || '本月沒有出勤記錄', 'warning');
            return;
        }
        
        // 👇 修正：先檢查資料結構
        console.log('API 回傳的資料:', res.records[0]); // 除錯用
        
        // 按員工分組
        const employeeData = {};
        
        res.records.forEach(record => {
            // 👇 修正：確保正確讀取 userId 和 name
            const userId = record.userId || 'unknown';
            const userName = record.name || '未知員工';
            
            if (!employeeData[userId]) {
                employeeData[userId] = {
                    name: userName,
                    records: []
                };
            }
            
            // 找出上班和下班的記錄
            const punchIn = record.record ? record.record.find(r => r.type === '上班') : null;
            const punchOut = record.record ? record.record.find(r => r.type === '下班') : null;
            
            // 計算工時
            let workHours = '-';
            if (punchIn && punchOut) {
                try {
                    const inTime = new Date(`${record.date} ${punchIn.time}`);
                    const outTime = new Date(`${record.date} ${punchOut.time}`);
                    const diffMs = outTime - inTime;
                    const diffHours = (diffMs / (1000 * 60 * 60)).toFixed(2);
                    workHours = diffHours > 0 ? diffHours : '-';
                } catch (e) {
                    console.error('計算工時失敗:', e);
                    workHours = '-';
                }
            }
            
            const statusText = t(record.reason) || record.reason;
            
            const notes = record.record
                ? record.record
                    .filter(r => r.note && r.note !== '系統虛擬卡')
                    .map(r => r.note)
                    .join('; ')
                : '';
            
            employeeData[userId].records.push({
                '日期': record.date,
                '上班時間': punchIn?.time || '-',
                '上班地點': punchIn?.location || '-',
                '下班時間': punchOut?.time || '-',
                '下班地點': punchOut?.location || '-',
                '工作時數': workHours,
                '狀態': statusText,
                '備註': notes || '-'
            });
        });
        
        // 建立工作簿
        const wb = XLSX.utils.book_new();
        
        // 為每位員工建立一個工作表
        for (const userId in employeeData) {
            const employee = employeeData[userId];
            const ws = XLSX.utils.json_to_sheet(employee.records);
            
            const wscols = [
                { wch: 12 },  // 日期
                { wch: 10 },  // 上班時間
                { wch: 20 },  // 上班地點
                { wch: 10 },  // 下班時間
                { wch: 20 },  // 下班地點
                { wch: 10 },  // 工作時數
                { wch: 15 },  // 狀態
                { wch: 30 }   // 備註
            ];
            ws['!cols'] = wscols;
            
            const sheetName = employee.name.substring(0, 31);
            XLSX.utils.book_append_sheet(wb, ws, sheetName);
        }
        
        const [year, month] = monthKey.split('-');
        const fileName = `所有員工出勤記錄_${year}年${month}月.xlsx`;
        XLSX.writeFile(wb, fileName);
        
        showNotification(t('EXPORT_SUCCESS') || '報表已成功匯出！', 'success');
        
    } catch (error) {
        console.error('匯出失敗:', error);
        showNotification(t('EXPORT_FAILED') || '匯出失敗，請稍後再試', 'error');
        
    } finally {
        if (exportBtn) {
            generalButtonState(exportBtn, 'idle');
        }
    }
}

// ==================== 📊 管理員匯出功能結束 ====================

// ==================== 📊 匯出出勤報表功能 ====================

/**
 * 匯出指定月份的出勤報表為 Excel 檔案
 * @param {Date} date - 要匯出的月份日期物件
 */
async function exportAttendanceReport(date) {
    const year = date.getFullYear();
    const month = date.getMonth() + 1;
    const monthKey = `${year}-${String(month).padStart(2, '0')}`;
    const userId = localStorage.getItem("sessionUserId");
    
    // 取得匯出按鈕
    const exportBtn = document.getElementById('export-attendance-btn');
    const loadingText = t('EXPORT_LOADING') || '正在準備報表...';
    
    // 顯示載入提示
    showNotification(loadingText, 'warning');
    
    // 按鈕進入處理中狀態
    if (exportBtn) {
        generalButtonState(exportBtn, 'processing', loadingText);
    }
    
    try {
        // 呼叫 API 取得出勤資料
        const res = await callApifetch(`getAttendanceDetails&month=${monthKey}&userId=${userId}`);
        
        if (!res.ok || !res.records || res.records.length === 0) {
            showNotification(t('EXPORT_NO_DATA') || '本月沒有出勤記錄', 'warning');
            return;
        }
        
        // 整理資料為 Excel 格式
        const exportData = [];
        
        res.records.forEach(record => {
            // 找出上班和下班的記錄
            const punchIn = record.record.find(r => r.type === '上班');
            const punchOut = record.record.find(r => r.type === '下班');
            
            // 計算工時
            let workHours = '-';
            if (punchIn && punchOut) {
                try {
                    const inTime = new Date(`${record.date} ${punchIn.time}`);
                    const outTime = new Date(`${record.date} ${punchOut.time}`);
                    const diffMs = outTime - inTime;
                    const diffHours = (diffMs / (1000 * 60 * 60)).toFixed(2);
                    workHours = diffHours > 0 ? diffHours : '-';
                } catch (e) {
                    console.error('計算工時失敗:', e);
                    workHours = '-';
                }
            }
            
            // 翻譯狀態
            const statusText = t(record.reason) || record.reason;
            
            // 處理備註
            const notes = record.record
                .filter(r => r.note && r.note !== '系統虛擬卡')
                .map(r => r.note)
                .join('; ');
            
            exportData.push({
                '日期': record.date,
                '上班時間': punchIn?.time || '-',
                '上班地點': punchIn?.location || '-',
                '下班時間': punchOut?.time || '-',
                '下班地點': punchOut?.location || '-',
                '工作時數': workHours,
                '狀態': statusText,
                '備註': notes || '-'
            });
        });
        
        // 使用 SheetJS 建立 Excel 檔案
        const ws = XLSX.utils.json_to_sheet(exportData);
        
        // 設定欄位寬度
        const wscols = [
            { wch: 12 },  // 日期
            { wch: 10 },  // 上班時間
            { wch: 20 },  // 上班地點
            { wch: 10 },  // 下班時間
            { wch: 20 },  // 下班地點
            { wch: 10 },  // 工作時數
            { wch: 15 },  // 狀態
            { wch: 30 }   // 備註
        ];
        ws['!cols'] = wscols;
        
        // 建立工作簿
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, `${month}月出勤記錄`);
        
        // 下載檔案
        const fileName = `出勤記錄_${year}年${month}月.xlsx`;
        XLSX.writeFile(wb, fileName);
        
        showNotification(t('EXPORT_SUCCESS') || '報表已成功匯出！', 'success');
        
    } catch (error) {
        console.error('匯出失敗:', error);
        showNotification(t('EXPORT_FAILED') || '匯出失敗，請稍後再試', 'error');
        
    } finally {
        // 恢復按鈕狀態
        if (exportBtn) {
            generalButtonState(exportBtn, 'idle');
        }
    }
}

// ==================== 📊 匯出功能結束 ====================

/* ===== 共用訊息顯示 ===== */
const showNotification = (message, type = 'success') => {
    const notification = document.getElementById('notification');
    const notificationMessage = document.getElementById('notification-message');
    notificationMessage.textContent = message;
    notification.className = 'notification'; // reset classes
    if (type === 'success') {
        notification.classList.add('bg-green-500', 'text-white');
    } else if (type === 'warning') {
        notification.classList.add('bg-yellow-500', 'text-white');
    } else {
        notification.classList.add('bg-red-500', 'text-white');
    }
    notification.classList.add('show');
    setTimeout(() => {
        notification.classList.remove('show');
    }, 3000);
};

// 確保登入
// script.js - 完整替換 ensureLogin 函數

/**
 * ⭐ 確保登入（優化版 - 使用合併 API）
 */
async function ensureLogin() {
    return new Promise(async (resolve) => {
        if (localStorage.getItem("sessionToken")) {
            document.getElementById("status").textContent = t("CHECKING_LOGIN");
            
            try {
                // ⭐⭐⭐ 關鍵修改：改用合併的 initApp API
                const res = await callApifetch("initApp");
                
                if (res.ok) {
                    console.log('✅ initApp 成功', res);
                    
                    // 檢查是否為管理員
                    if (res.user.dept === "管理員") {
                        console.log('👑 管理員身份:', res.user.dept);
                        document.getElementById('tab-admin-btn').style.display = 'block';
                    }
                    
                    // 設定使用者資訊
                    document.getElementById("user-name").textContent = res.user.name;
                    document.getElementById("profile-img").src = res.user.picture || res.user.rate;
                    
                    // 儲存使用者 ID
                    localStorage.setItem("sessionUserId", res.user.userId);
                    
                    // 顯示成功訊息
                    showNotification(t("LOGIN_SUCCESS"));
                    
                    // 切換介面
                    document.getElementById('login-section').style.display = 'none';
                    document.getElementById('user-header').style.display = 'flex';
                    document.getElementById('main-app').style.display = 'block';
                    
                    // ⭐⭐⭐ 直接渲染異常記錄，不需要再呼叫 checkAbnormal()
                    renderAbnormalRecords(res.abnormalRecords);
                    
                    resolve(true);
                    
                } else {
                    // 登入失敗
                    console.error('❌ initApp 失敗:', res);
                    
                    const errorMsg = t(res.code || "UNKNOWN_ERROR");
                    showNotification(`❌ ${errorMsg}`, "error");
                    
                    document.getElementById("status").textContent = t("PLEASE_RELOGIN");
                    document.getElementById('login-btn').style.display = 'block';
                    document.getElementById('user-header').style.display = 'none';
                    document.getElementById('main-app').style.display = 'none';
                    
                    resolve(false);
                }
                
            } catch (err) {
                console.error('❌ ensureLogin 錯誤:', err);
                
                document.getElementById('login-btn').style.display = 'block';
                document.getElementById('user-header').style.display = 'none';
                document.getElementById('main-app').style.display = 'none';
                document.getElementById("status").textContent = t("PLEASE_RELOGIN");
                
                resolve(false);
            }
            
        } else {
            // 未登入
            document.getElementById('login-btn').style.display = 'block';
            document.getElementById('user-header').style.display = 'none';
            document.getElementById('main-app').style.display = 'none';
            document.getElementById("status").textContent = t("SUBTITLE_LOGIN");
            
            resolve(false);
        }
    });
}
// script.js - 在 checkAbnormal 函數附近加入
/**
 * ⭐ 渲染異常記錄（從 initApp 返回的資料）
 */
function renderAbnormalRecords(records) {
    console.log('📋 renderAbnormalRecords 開始', records);
    
    const recordsLoading = document.getElementById("abnormal-records-loading");
    const abnormalRecordsSection = document.getElementById("abnormal-records-section");
    const abnormalList = document.getElementById("abnormal-list");
    const recordsEmpty = document.getElementById("abnormal-records-empty");
    
    if (!recordsLoading || !abnormalRecordsSection || !abnormalList || !recordsEmpty) {
        console.error('❌ 找不到必要的 DOM 元素');
        return;
    }
    
    recordsLoading.style.display = 'none';
    abnormalRecordsSection.style.display = 'block';
    
    if (records && records.length > 0) {
        console.log(`✅ 有 ${records.length} 筆異常記錄`);
        
        recordsEmpty.style.display = 'none';
        abnormalList.innerHTML = '';
        
        const sortedRecords = records.sort((a, b) => {
            return new Date(b.date) - new Date(a.date);
        });
        
        sortedRecords.forEach((record, index) => {
            console.log(`   ${index + 1}. ${record.date} - ${record.reason}`);
            
            let reasonClass, displayReason, buttonHtml;
            
            // ⭐⭐⭐ 新增翻譯映射函數
            function translatePunchTypes(punchTypes) {
                if (!punchTypes) return '';
                
                const translations = {
                    '補上班審核中': t('STATUS_REPAIR_PENDING_IN') || 'Punch In Review Pending',
                    '補下班審核中': t('STATUS_REPAIR_PENDING_OUT') || 'Punch Out Review Pending',
                    '補上班通過': t('STATUS_REPAIR_APPROVED_IN') || 'Punch In Approved',
                    '補下班通過': t('STATUS_REPAIR_APPROVED_OUT') || 'Punch Out Approved',
                    '補上班被拒絕': t('STATUS_REPAIR_REJECTED_IN') || 'Punch In Rejected',
                    '補下班被拒絕': t('STATUS_REPAIR_REJECTED_OUT') || 'Punch Out Rejected'
                };
                
                return translations[punchTypes] || punchTypes;
            }
            
            switch(record.reason) {
                case 'STATUS_REPAIR_PENDING':
                    reasonClass = 'text-yellow-600 dark:text-yellow-400';
                    displayReason = translatePunchTypes(record.punchTypes);
                    buttonHtml = `
                        <span class="text-sm font-semibold text-yellow-600 dark:text-yellow-400">
                            ⏳ ${translatePunchTypes(record.punchTypes)}
                        </span>
                    `;
                    break;
                    
                case 'STATUS_REPAIR_APPROVED':
                    reasonClass = 'text-green-600 dark:text-green-400';
                    displayReason = translatePunchTypes(record.punchTypes);
                    buttonHtml = `
                        <span class="text-sm font-semibold text-green-600 dark:text-green-400">
                            ✓ ${translatePunchTypes(record.punchTypes)}
                        </span>
                    `;
                    break;
                
                case 'STATUS_REPAIR_REJECTED':
                    reasonClass = 'text-orange-600 dark:text-orange-400';
                    displayReason = translatePunchTypes(record.punchTypes);
                    
                    // ⭐ 判斷是上班還是下班
                    const isIn = record.punchTypes && record.punchTypes.includes('上班');
                    const punchType = isIn ? '上班' : '下班';
                    
                    buttonHtml = `
                        <button data-date="${record.date}" 
                                data-type="${punchType}"
                                class="adjust-btn px-4 py-2 text-sm font-semibold text-white bg-orange-600 dark:bg-orange-500 rounded-md hover:bg-orange-700 dark:hover:bg-orange-600 transition-colors">
                            ${t('REAPPLY') || 'Reapply'}
                        </button>
                    `;
                    break;
                    
                case 'STATUS_PUNCH_IN_MISSING':
                    reasonClass = 'text-red-600 dark:text-red-400';
                    displayReason = t('STATUS_PUNCH_IN_MISSING');
                    buttonHtml = `
                        <button data-date="${record.date}" 
                                data-type="上班"
                                class="adjust-btn px-4 py-2 text-sm font-semibold text-white bg-indigo-600 dark:bg-indigo-500 rounded-md hover:bg-indigo-700 dark:hover:bg-indigo-600 transition-colors">
                            ${t('BTN_ADJUST_IN')}
                        </button>
                    `;
                    break;
                    
                case 'STATUS_PUNCH_OUT_MISSING':
                    reasonClass = 'text-red-600 dark:text-red-400';
                    displayReason = t('STATUS_PUNCH_OUT_MISSING');
                    buttonHtml = `
                        <button data-date="${record.date}" 
                                data-type="下班"
                                class="adjust-btn px-4 py-2 text-sm font-semibold text-white bg-purple-600 dark:bg-purple-500 rounded-md hover:bg-purple-700 dark:hover:bg-purple-600 transition-colors">
                            ${t('BTN_ADJUST_OUT')}
                        </button>
                    `;
                    break;
                    
                default:
                    reasonClass = 'text-gray-600 dark:text-gray-400';
                    displayReason = t(record.reason) || record.reason;
                    buttonHtml = '';
            }
            
            const li = document.createElement('li');
            li.className = 'p-3 bg-gray-50 rounded-lg flex justify-between items-center dark:bg-gray-700';
            
            li.innerHTML = `
                <div>
                    <p class="font-medium text-gray-800 dark:text-white">${record.date}</p>
                    <p class="text-sm ${reasonClass}">
                        ${displayReason}
                    </p>
                </div>
                ${buttonHtml}
            `;
            
            abnormalList.appendChild(li);
        });
        
        console.log('✅ 渲染完成');
        
    } else {
        console.log('ℹ️  沒有異常記錄');
        recordsEmpty.style.display = 'block';
        abnormalList.innerHTML = '';
    }
}
/**
/**
 * ✅ 檢查本月打卡異常（完整修正版 - 支援多語言）
 */
async function checkAbnormal() {
    const now = new Date();
    const month = now.getFullYear() + "-" + String(now.getMonth() + 1).padStart(2, "0");
    const userId = localStorage.getItem("sessionUserId");
    
    console.log('🔍 開始檢查異常記錄');
    
    const recordsLoading = document.getElementById("abnormal-records-loading");
    const abnormalRecordsSection = document.getElementById("abnormal-records-section");
    const abnormalList = document.getElementById("abnormal-list");
    const recordsEmpty = document.getElementById("abnormal-records-empty");
    
    if (!recordsLoading || !abnormalRecordsSection || !abnormalList || !recordsEmpty) {
        console.error('❌ 找不到必要的 DOM 元素');
        return;
    }
    
    recordsLoading.style.display = 'block';
    abnormalRecordsSection.style.display = 'none';
    
    // ⭐⭐⭐ 翻譯映射函數
    function translatePunchTypes(punchTypes) {
        if (!punchTypes) return '';
        
        const translations = {
            '補上班審核中': t('STATUS_REPAIR_PENDING_IN') || 'Punch In Review Pending',
            '補下班審核中': t('STATUS_REPAIR_PENDING_OUT') || 'Punch Out Review Pending',
            '補上班通過': t('STATUS_REPAIR_APPROVED_IN') || 'Punch In Approved',
            '補下班通過': t('STATUS_REPAIR_APPROVED_OUT') || 'Punch Out Approved',
            '補上班被拒絕': t('STATUS_REPAIR_REJECTED_IN') || 'Punch In Rejected',
            '補下班被拒絕': t('STATUS_REPAIR_REJECTED_OUT') || 'Punch Out Rejected'
        };
        
        return translations[punchTypes] || punchTypes;
    }
    
    try {
        const res = await callApifetch(`getAbnormalRecords&month=${month}&userId=${userId}`);
        
        console.log('📤 API 回傳結果:', res);
        console.log('   記錄數量:', res.records?.length || 0);
        
        recordsLoading.style.display = 'none';
        
        if (res.ok) {
            abnormalRecordsSection.style.display = 'block';
            
            if (res.records && res.records.length > 0) {
                console.log('✅ 有異常記錄，開始渲染');
                
                recordsEmpty.style.display = 'none';
                abnormalList.innerHTML = '';
                
                // ✅ 按日期排序（由新到舊）
                const sortedRecords = res.records.sort((a, b) => {
                    return new Date(b.date) - new Date(a.date);
                });
                
                sortedRecords.forEach((record, index) => {
                    console.log(`   渲染第 ${index + 1} 筆: ${record.date} - ${record.reason}`);
                    
                    let reasonClass, displayReason, buttonHtml;
                    
                    switch(record.reason) {
                        case 'STATUS_REPAIR_PENDING':
                            // 審核中 - 黃色，按鈕禁用
                            reasonClass = 'text-yellow-600 dark:text-yellow-400';
                            displayReason = translatePunchTypes(record.punchTypes);
                            buttonHtml = `
                                <span class="text-sm font-semibold text-yellow-600 dark:text-yellow-400">
                                    ⏳ ${translatePunchTypes(record.punchTypes)}
                                </span>
                            `;
                            break;
                            
                        case 'STATUS_REPAIR_APPROVED':
                            // 已通過 - 綠色，按鈕禁用
                            reasonClass = 'text-green-600 dark:text-green-400';
                            displayReason = translatePunchTypes(record.punchTypes);
                            buttonHtml = `
                                <span class="text-sm font-semibold text-green-600 dark:text-green-400">
                                    ✓ ${translatePunchTypes(record.punchTypes)}
                                </span>
                            `;
                            break;
                            
                        case 'STATUS_PUNCH_IN_MISSING':
                            // 缺上班卡 - 紅色，可補打卡
                            reasonClass = 'text-red-600 dark:text-red-400';
                            displayReason = t('STATUS_PUNCH_IN_MISSING');
                            buttonHtml = `
                                <button data-date="${record.date}" 
                                        data-type="上班"
                                        class="adjust-btn px-4 py-2 text-sm font-semibold text-white bg-indigo-600 dark:bg-indigo-500 rounded-md hover:bg-indigo-700 dark:hover:bg-indigo-600 transition-colors">
                                    ${t('BTN_ADJUST_IN')}
                                </button>
                            `;
                            break;
                            
                        case 'STATUS_PUNCH_OUT_MISSING':
                            // 缺下班卡 - 紅色，可補打卡
                            reasonClass = 'text-red-600 dark:text-red-400';
                            displayReason = t('STATUS_PUNCH_OUT_MISSING');
                            buttonHtml = `
                                <button data-date="${record.date}" 
                                        data-type="下班"
                                        class="adjust-btn px-4 py-2 text-sm font-semibold text-white bg-purple-600 dark:bg-purple-500 rounded-md hover:bg-purple-700 dark:hover:bg-purple-600 transition-colors">
                                    ${t('BTN_ADJUST_OUT')}
                                </button>
                            `;
                            break;

                        case 'STATUS_REPAIR_REJECTED':
                            // ❌ 被拒絕 - 橘色，可重新申請
                            reasonClass = 'text-orange-600 dark:text-orange-400';
                            displayReason = translatePunchTypes(record.punchTypes);
                            
                            // ⭐ 判斷是上班還是下班
                            const isIn = record.punchTypes && record.punchTypes.includes('上班');
                            const punchType = isIn ? '上班' : '下班';
                            
                            buttonHtml = `
                                <button data-date="${record.date}" 
                                        data-type="${punchType}"
                                        class="adjust-btn px-4 py-2 text-sm font-semibold text-white bg-orange-600 dark:bg-orange-500 rounded-md hover:bg-orange-700 dark:hover:bg-orange-600 transition-colors">
                                    ${t('REAPPLY') || 'Reapply'}
                                </button>
                            `;
                            break;
                            
                        default:
                            reasonClass = 'text-gray-600 dark:text-gray-400';
                            displayReason = t(record.reason) || record.reason;
                            buttonHtml = '';
                    }
                    
                    const li = document.createElement('li');
                    li.className = 'p-3 bg-gray-50 rounded-lg flex justify-between items-center dark:bg-gray-700';
                    
                    li.innerHTML = `
                        <div>
                            <p class="font-medium text-gray-800 dark:text-white">${record.date}</p>
                            <p class="text-sm ${reasonClass}">
                                ${displayReason}
                            </p>
                        </div>
                        ${buttonHtml}
                    `;
                    
                    abnormalList.appendChild(li);
                });
                
                console.log('✅ 渲染完成');
                
            } else {
                console.log('ℹ️  沒有異常記錄');
                recordsEmpty.style.display = 'block';
                abnormalList.innerHTML = '';
            }
        } else {
            console.error("❌ API 返回失敗:", res.msg || res.code);
            showNotification(t("ERROR_FETCH_RECORDS") || "無法取得記錄", "error");
        }
    } catch (err) {
        console.error('❌ 發生錯誤:', err);
        recordsLoading.style.display = 'none';
        showNotification(t("ERROR_FETCH_RECORDS") || "無法取得記錄", "error");
    }
}
// 渲染日曆的函式
async function renderCalendar(date) {
    const monthTitle = document.getElementById('month-title');
    const calendarGrid = document.getElementById('calendar-grid');
    const year = date.getFullYear();
    const month = date.getMonth();
    const today = new Date();
    
    // 生成 monthKey
    const monthkey = currentMonthDate.getFullYear() + "-" + String(currentMonthDate.getMonth() + 1).padStart(2, "0");
    
    // 檢查快取中是否已有該月份資料
    if (monthDataCache[monthkey]) {
        // 如果有，直接從快取讀取資料並渲染
        const records = monthDataCache[monthkey];
        renderCalendarWithData(year, month, today, records, calendarGrid, monthTitle);
    } else {
        // 如果沒有，才發送 API 請求
        // 清空日曆，顯示載入狀態，並確保置中
        calendarGrid.innerHTML = '<div data-i18n="LOADING" class="col-span-full text-center text-gray-500 dark:text-gray-400 py-4">正在載入...</div>';
        renderTranslations(calendarGrid);
        try {
            const res = await callApifetch(`getAttendanceDetails&month=${monthkey}&userId=${userId}`);
            if (res.ok) {
                // 將資料存入快取
                monthDataCache[monthkey] = res.records;
                
                // 收到資料後，清空載入訊息
                calendarGrid.innerHTML = '';
                
                // 從快取取得本月資料
                const records = monthDataCache[monthkey] || [];
                renderCalendarWithData(year, month, today, records, calendarGrid, monthTitle);
            } else {
                console.error("Failed to fetch attendance records:", res.msg);
                showNotification(t("ERROR_FETCH_RECORDS"), "error");
            }
        } catch (err) {
            console.error(err);
        }
    }
}

async function submitAdjustPunch(date, type, note) {
    try {
        showNotification("正在提交補打卡...", "info");
        
        const sessionToken = localStorage.getItem("sessionToken");
        
        // 取得當前位置
        const position = await new Promise((resolve, reject) => {
            navigator.geolocation.getCurrentPosition(resolve, reject);
        });
        
        const lat = position.coords.latitude;
        const lng = position.coords.longitude;
        
        // 設定預設時間
        const datetime = `${date}T${type === '上班' ? '09:00:00' : '18:00:00'}`;
        
        const params = new URLSearchParams({
            token: sessionToken,
            type: type,
            lat: lat,
            lng: lng,
            datetime: datetime,
            note: note || `補打卡 - ${type}`
        });
        
        const res = await callApifetch(`adjustPunch&${params.toString()}`);
        
        if (res.ok) {
            showNotification("補打卡申請成功！等待管理員審核", "success");
            
            // ⭐⭐⭐ 關鍵：補打卡成功後，重新檢查異常記錄
            await checkAbnormal();
            
            // 關閉對話框
            closeAdjustDialog();
        } else {
            showNotification(t(res.code) || "補打卡失敗", "error");
        }
    } catch (err) {
        console.error('補打卡錯誤:', err);
        showNotification("補打卡失敗", "error");
    }
}

// 新增一個獨立的渲染函式，以便從快取或 API 回應中調用
function renderCalendarWithData(year, month, today, records, calendarGrid, monthTitle) {
    // 確保日曆網格在每次渲染前被清空
    calendarGrid.innerHTML = '';
    monthTitle.textContent = t("MONTH_YEAR_TEMPLATE", {
        year: year,
        month: month+1
    });
    
    // 取得該月第一天是星期幾
    const firstDayOfMonth = new Date(year, month, 1).getDay();
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    
    // 填補月初的空白格子
    for (let i = 0; i < firstDayOfMonth; i++) {
        const emptyCell = document.createElement('div');
        emptyCell.className = 'day-cell';
        calendarGrid.appendChild(emptyCell);
    }
    
    // 根據資料渲染每一天的顏色
    for (let i = 1; i <= daysInMonth; i++) {
        const dayCell = document.createElement('div');
        const cellDate = new Date(year, month, i);
        dayCell.textContent = i;
        let dateKey = `${year}-${String(month + 1).padStart(2, '0')}-${String(i).padStart(2, '0')}`;
        let dateClass = 'normal-day';
        
        const todayRecords = records.filter(r => r.date === dateKey);
        
        if (todayRecords.length > 0) {
            const reason = todayRecords[0].reason;
            switch (reason) {
                case "STATUS_PUNCH_IN_MISSING":
                    dateClass = 'abnormal-day';
                    break;
                case "STATUS_PUNCH_OUT_MISSING":
                    dateClass = 'abnormal-day';
                    break;
                case "STATUS_PUNCH_NORMAL":
                    dateClass = 'day-off';
                    break;
                case "STATUS_REPAIR_PENDING":
                    dateClass = 'pending-virtual';
                    break;
                case "STATUS_REPAIR_APPROVED":
                    dateClass = 'approved-virtual';
                    break;
                default:
                    if (reason && reason !== "") {
                        dateClass = 'pending-adjustment'; // 假設所有有備註的都算 pending
                    }
                    break;
            }
        }
        
        const isToday = (year === today.getFullYear() && month === today.getMonth() && i === today.getDate());
        if (isToday) {
            dayCell.classList.add('today');
        } else if (cellDate > today) {
            dayCell.classList.add('future-day');
            dayCell.style.pointerEvents = 'none'; // 未來日期不可點擊
        } else {
            dayCell.classList.add(dateClass);
        }
        
        dayCell.classList.add('day-cell');
        dayCell.dataset.date = dateKey;
        dayCell.dataset.records = JSON.stringify(todayRecords); // 儲存當天資料
        calendarGrid.appendChild(dayCell);
    }
}

async function renderDailyRecords(dateKey) {
    // 1. 取得所有需要的 DOM 元素
    const dailyRecordsCard = document.getElementById('daily-records-card');
    const dailyRecordsTitle = document.getElementById('daily-records-title');
    const dailyRecordsList = document.getElementById('daily-records-list');
    const dailyRecordsEmpty = document.getElementById('daily-records-empty');
    const recordsLoading = document.getElementById("daily-records-loading");
    const adjustmentFormContainer = document.getElementById('daily-adjustment-form-container');
    
    // 2. ✅ 檢查必要元素是否存在
    if (!dailyRecordsCard || !dailyRecordsTitle || !dailyRecordsList || !dailyRecordsEmpty) {
        console.error('❌ renderDailyRecords: 找不到必要的 DOM 元素');
        console.log('元素檢查結果:', {
            'daily-records-card': !!dailyRecordsCard,
            'daily-records-title': !!dailyRecordsTitle,
            'daily-records-list': !!dailyRecordsList,
            'daily-records-empty': !!dailyRecordsEmpty,
            'daily-records-loading': !!recordsLoading,
            'daily-adjustment-form-container': !!adjustmentFormContainer
        });
        
        showNotification('介面元素載入失敗，請重新整理頁面', 'error');
        return;
    }
    
    // 3. 安全地設置內容
    dailyRecordsTitle.textContent = t("DAILY_RECORDS_TITLE", {
        dateKey: dateKey
    });
    
    dailyRecordsList.innerHTML = '';
    dailyRecordsEmpty.style.display = 'none';
    
    if (adjustmentFormContainer) {
        adjustmentFormContainer.innerHTML = '';
    }
    
    if (recordsLoading) {
        recordsLoading.style.display = 'block';
    }
    
    // 4. 繼續原有邏輯
    const dateObject = new Date(dateKey);
    const month = dateObject.getFullYear() + "-" + String(dateObject.getMonth() + 1).padStart(2, "0");
    const userId = localStorage.getItem("sessionUserId");
    
    if (monthDataCache[month]) {
        renderRecords(monthDataCache[month]);
        if (recordsLoading) {
            recordsLoading.style.display = 'none';
        }
    } else {
        try {
            const res = await callApifetch(`getAttendanceDetails&month=${month}&userId=${userId}`);
            if (recordsLoading) {
                recordsLoading.style.display = 'none';
            }
            if (res.ok) {
                monthDataCache[month] = res.records;
                renderRecords(res.records);
            } else {
                console.error("Failed to fetch attendance records:", res.msg);
                showNotification(t("ERROR_FETCH_RECORDS"), "error");
            }
        } catch (err) {
            console.error(err);
            if (recordsLoading) {
                recordsLoading.style.display = 'none';
            }
        }
    }
    
    // 5. renderRecords 函數（保持不變）
    function renderRecords(records) {
        const dailyRecords = records.filter(record => record.date === dateKey);
        
        if (dailyRecords.length > 0) {
            dailyRecordsEmpty.style.display = 'none';
            dailyRecords.forEach(recordData => {
                const li = document.createElement('li');
                li.className = 'p-3 bg-gray-50 dark:bg-gray-700 rounded-lg';
                
                const recordHtml = recordData.record.map(r => {
                    const typeKey = r.type === '上班' ? 'PUNCH_IN' : 'PUNCH_OUT';
                    return `
                        <p class="font-medium text-gray-800 dark:text-white">${r.time} - ${t(typeKey)}</p>
                        <p class="text-sm text-gray-500 dark:text-gray-400">${r.location}</p>
                        <p data-i18n="RECORD_NOTE_PREFIX" class="text-sm text-gray-500 dark:text-gray-400">備註：${r.note}</p>
                    `;
                }).join("");
                
                li.innerHTML = `
                    ${recordHtml}
                    <p class="text-sm text-gray-500 dark:text-gray-400">
                        <span data-i18n="RECORD_REASON_PREFIX">系統判斷：</span>
                        ${t(recordData.reason)}
                    </p>
                `;
                dailyRecordsList.appendChild(li);
                renderTranslations(li);
            });
            
            // 檢查是否需要顯示補打卡按鈕
            
        } else {
            dailyRecordsEmpty.style.display = 'block';
        }
        
        dailyRecordsCard.style.display = 'block';
    }
}

document.addEventListener('DOMContentLoaded', async () => {
    
    const loginBtn = document.getElementById('login-btn');
    const logoutBtn = document.getElementById('logout-btn');
    const punchInBtn = document.getElementById('punch-in-btn');
    const punchOutBtn = document.getElementById('punch-out-btn');
    const tabDashboardBtn = document.getElementById('tab-dashboard-btn');
    const tabMonthlyBtn = document.getElementById('tab-monthly-btn');
    const tabLocationBtn = document.getElementById('tab-location-btn');
    const tabAdminBtn = document.getElementById('tab-admin-btn');
    const tabOvertimeBtn = document.getElementById('tab-overtime-btn');
    const tabLeaveBtn = document.getElementById('tab-leave-btn'); // 👈 新增請假按鈕
    const tabSalaryBtn = document.getElementById('tab-salary-btn'); // 👈 新增
    const abnormalList = document.getElementById('abnormal-list');
    const adjustmentFormContainer = document.getElementById('adjustment-form-container');
    const calendarGrid = document.getElementById('calendar-grid');
    // 取得當前位置按鈕事件
    const getLocationBtn = document.getElementById('get-location-btn');
    const locationLatInput = document.getElementById('location-lat');
    const locationLngInput = document.getElementById('location-lng');
    const addLocationBtn = document.getElementById('add-location-btn');
    
    let pendingRequests = []; // 新增：用於快取待審核的請求
    
    // 全域變數，用於儲存地圖實例
    let mapInstance = null;
    let mapLoadingText = null;
    let currentCoords = null;
    let marker = null;
    let circle = null;
    /**
     * 從後端取得所有打卡地點，並將它們顯示在地圖上。
     */
    // 全域變數，用於儲存地點標記和圓形
    let locationMarkers = L.layerGroup();
    let locationCircles = L.layerGroup();
    
    /**
     * 取得並渲染所有待審核的請求。
     */
    async function fetchAndRenderReviewRequests() {
        const loadingEl = document.getElementById('requests-loading');
        const emptyEl = document.getElementById('requests-empty');
        const listEl = document.getElementById('pending-requests-list');
        
        loadingEl.style.display = 'block';
        emptyEl.style.display = 'none';
        listEl.innerHTML = '';
        
        try {
            const res = await callApifetch("getReviewRequest");
            
            if (res.ok && Array.isArray(res.reviewRequest)) {
                pendingRequests = res.reviewRequest; // 快取所有請求
                
                if (pendingRequests.length === 0) {
                    emptyEl.style.display = 'block';
                } else {
                    renderReviewRequests(pendingRequests);
                }
            } else {
                showNotification("取得待審核請求失敗：" + res.msg, "error");
                emptyEl.style.display = 'block';
            }
        } catch (error) {
            showNotification("取得待審核請求失敗，請檢查網路。", "error");
            emptyEl.style.display = 'block';
            console.error("Failed to fetch review requests:", error);
        } finally {
            loadingEl.style.display = 'none';
        }
    }
    
    /**
     * 根據資料渲染待審核列表。
     * @param {Array<Object>} requests - 請求資料陣列。
     */
    function renderReviewRequests(requests) {
        const listEl = document.getElementById('pending-requests-list');
        listEl.innerHTML = '';
        
        requests.forEach((req, index) => {
            const li = document.createElement('li');
            li.className = 'p-4 bg-gray-50 rounded-lg shadow-sm flex flex-col space-y-2 dark:bg-gray-700';
            li.innerHTML = `
             <div class="flex flex-col space-y-1">

                        <div class="flex items-center justify-between w-full">
                            <p class="text-sm font-semibold text-gray-800 dark:text-white">${req.name} - ${req.remark}</p>
                            <span class="text-xs text-gray-500 dark:text-gray-400">${req.applicationPeriod}</span>
                        </div>
                    </div>
                    
                <div class="flex items-center justify-between w-full mt-2">
                    <p 
                        data-i18n-key="${req.type}" 
                        class="text-sm text-indigo-600 dark:text-indigo-400 font-medium">
                    </p> 
                    
                    <div class="flex space-x-2"> 
                        <button data-i18n="ADMIN_APPROVE_BUTTON" data-index="${index}" class="approve-btn px-3 py-1 rounded-md text-sm font-bold btn-primary">核准</button>
                        <button data-i18n="ADMIN_REJECT_BUTTON" data-index="${index}" class="reject-btn px-3 py-1 rounded-md text-sm font-bold btn-warning">拒絕</button>
                    </div>
                </div>
            `;
            listEl.appendChild(li);
            renderTranslations(li);
        });
        
        listEl.querySelectorAll('.approve-btn').forEach(button => {
            button.addEventListener('click', (e) => handleReviewAction(e.currentTarget, e.currentTarget.dataset.index, 'approve'));
        });
        
        listEl.querySelectorAll('.reject-btn').forEach(button => {
            button.addEventListener('click', (e) => handleReviewAction(e.currentTarget, e.currentTarget.dataset.index, 'reject'));
        });
    }
    
    /**
     * 處理審核動作（核准或拒絕）。
     * @param {HTMLElement} button - 被點擊的按鈕元素。
     * @param {number} index - 請求在陣列中的索引。
     * @param {string} action - 'approve' 或 'reject'。
     */
    async function handleReviewAction(button, index, action) {
        const request = pendingRequests[index];
        if (!request) {
            showNotification("找不到請求資料。", "error");
            return;
        }

        const recordId = request.id;
        const endpoint = action === 'approve' ? 'approveReview' : 'rejectReview';
        const loadingText = t('LOADING') || '處理中...';
        
        // A. 進入處理中狀態
        generalButtonState(button, 'processing', loadingText);
        
        try {
            const res = await callApifetch(`${endpoint}&id=${recordId}`);
            
            if (res.ok) {
                const translationKey = action === 'approve' ? 'REQUEST_APPROVED' : 'REQUEST_REJECTED';
                showNotification(t(translationKey), "success");
                
                // 由於成功後列表會被重新整理，這裡可以不立即恢復按鈕狀態
                // 但是為了保險起見，我們仍然在 finally 中恢復。
                
                // 延遲執行，讓按鈕的禁用狀態能被看到
                await new Promise(resolve => setTimeout(resolve, 500));
                
                // 列表重新整理會渲染新按鈕，覆蓋舊的按鈕
                fetchAndRenderReviewRequests();
            } else {
                showNotification(t('REVIEW_FAILED', { msg: res.msg }), "error");
            }
            
        } catch (err) {
            showNotification(t("REVIEW_NETWORK_ERROR"), "error");
            console.error(err);
            
        } finally {
            // B. 無論成功或失敗，都需要將按鈕恢復到可點擊狀態
            // 只有在列表沒有被重新整理時，這個恢復才有意義
            generalButtonState(button, 'idle');
        }
    }
    /**
     * 從後端取得所有打卡地點，並將它們顯示在地圖上。
     */
    async function fetchAndRenderLocationsOnMap() {
        try {
            const res = await callApifetch("getLocations");
            
            // 清除舊的地點標記和圓形
            locationMarkers.clearLayers();
            locationCircles.clearLayers();
            
            if (res.ok && Array.isArray(res.locations)) {
                // 遍歷所有地點並在地圖上放置標記和圓形
                res.locations.forEach(loc => {
                    // 如果沒有容許誤差，則預設為 50 公尺
                    const punchInRadius = loc.scope || 50;
                    
                    // 加入圓形範圍
                    const locationCircle = L.circle([loc.lat, loc.lng], {
                        color: 'red',
                        fillColor: '#f03',
                        fillOpacity: 0.2,
                        radius: punchInRadius
                    });
                    locationCircle.bindPopup(`<b>${loc.name}</b><br>可打卡範圍：${punchInRadius}公尺`);
                    locationCircles.addLayer(locationCircle);
                });
                
                // 將所有地點標記和圓形一次性加到地圖上
                locationMarkers.addTo(mapInstance);
                locationCircles.addTo(mapInstance);
                
                console.log("地點標記和範圍已成功載入地圖。");
            } else {
                showNotification("取得地點清單失敗：" + res.msg, "error");
                console.error("Failed to fetch locations:", res.msg);
            }
        } catch (error) {
            showNotification("取得地點清單失敗，請檢查網路。", "error");
            console.error("Failed to fetch locations:", error);
        }
    }
    // 初始化地圖並取得使用者位置
    function initLocationMap(forceReload = false){
        const mapContainer = document.getElementById('map-container');
        const statusEl = document.getElementById('location-status');
        const coordsEl = document.getElementById('location-coords');
        console.log(mapInstance && !forceReload);
        // 取得載入文字元素
        if (!mapLoadingText) {
            mapLoadingText = document.getElementById('map-loading-text');
        }
        // 檢查地圖實例是否已存在
        if (mapInstance) {
            // 如果已經存在，並且沒有被要求強制重新載入，則直接返回
            if (!forceReload) {
                mapInstance.invalidateSize();
                return;
            }
            
            // 如果被要求強制重新載入，則先徹底銷毀舊的地圖實例
            mapInstance.remove();
            mapInstance = null;
        }
        
        
        // 顯示載入中的文字
        mapLoadingText.style.display = 'block'; // 或 'block'，根據你的樣式決定
        
        // 建立地圖
        mapInstance = L.map('map-container', {
            center: [25.0330, 121.5654], // 預設中心點為台北市
            zoom: 13
        });
        
        // 加入 OpenStreetMap 圖層
        L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
            attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
        }).addTo(mapInstance);
        
        // 讓地圖在完成載入後隱藏載入中的文字
        mapInstance.whenReady(() => {
            mapLoadingText.style.display = 'none';
            // 確保地圖的尺寸正確
            mapInstance.invalidateSize();
        });
        
        // 顯示載入狀態
        //mapContainer.innerHTML = t("MAP_LOADING");
        statusEl.textContent = t('DETECTING_LOCATION');
        coordsEl.textContent = t('UNKNOWN_LOCATION');
        
        // 取得使用者地理位置
        if (navigator.geolocation) {
            navigator.geolocation.getCurrentPosition(
                                                     (position) => {
                                                         const { latitude, longitude } = position.coords;
                                                         currentCoords = [latitude, longitude];
                                                         
                                                         // 更新狀態顯示
                                                         statusEl.textContent = t('DETECTION_SUCCESS');
                                                         coordsEl.textContent = `${latitude.toFixed(6)}, ${longitude.toFixed(6)}`;
                                                         
                                                         // 設定地圖視圖
                                                         mapInstance.setView(currentCoords, 18);
                                                         
                                                         // 在地圖上放置標記
                                                         if (marker) mapInstance.removeLayer(marker);
                                                         marker = L.marker(currentCoords).addTo(mapInstance)
                                                         .bindPopup(t('CURRENT_LOCATION'))
                                                         .openPopup();
                                                         
                                                         
                                                     },
                                                     (error) => {
                                                         // 處理定位失敗
                                                         statusEl.textContent = t('ERROR_GEOLOCATION_PERMISSION_DENIED');
                                                         console.error("Geolocation failed:", error);
                                                         
                                                         let message;
                                                         switch(error.code) {
                                                             case error.PERMISSION_DENIED:
                                                                 message = t('ERROR_GEOLOCATION_PERMISSION_DENIED');
                                                                 break;
                                                             case error.POSITION_UNAVAILABLE:
                                                                 message = t('ERROR_GEOLOCATION_UNAVAILABLE');
                                                                 break;
                                                             case error.TIMEOUT:
                                                                 message = t('ERROR_GEOLOCATION_TIMEOUT');
                                                                 break;
                                                             case error.UNKNOWN_ERROR:
                                                                 message = t('ERROR_GEOLOCATION_UNKNOWN');
                                                                 break;
                                                         }
                                                         showNotification(`定位失敗：${message}`, "error");
                                                     }
                                                     );
            // 成功取得使用者位置後，載入所有打卡地點
            fetchAndRenderLocationsOnMap();
        } else {
            showNotification(t('ERROR_BROWSER_NOT_SUPPORTED'), "error");
            statusEl.textContent = '不支援定位';
        }
    }
    
    
    // 處理 API 測試按鈕事件
    document.getElementById('test-api-btn').addEventListener('click', async () => {
        // 這裡替換成您想要測試的 API action 名稱
        const testAction = "testEndpoint";
        
        try {
            // 使用 await 等待 API 呼叫完成並取得回應
            const res = await callApifetch(testAction);
            
            // 檢查 API 回應中的 'ok' 屬性
            if (res && res.ok) {
                showNotification("API 測試成功！回應：" + JSON.stringify(res), "success");
            } else {
                // 如果 res.ok 為 false，表示後端處理失敗
                showNotification("API 測試失敗：" + (res ? res.msg : "無回應資料"), "error");
            }
        } catch (error) {
            // 捕捉任何在 callApifetch 函式中拋出的錯誤（例如網路連線問題）
            console.error("API 呼叫發生錯誤:", error);
            showNotification("API 呼叫失敗，請檢查網路連線或後端服務。", "error");
        }
    });
    
    getLocationBtn.addEventListener('click', () => {
        if (!navigator.geolocation) {
            showNotification(t("ERROR_GEOLOCATION", { msg: t('ERROR_BROWSER_NOT_SUPPORTED') }), "error");
            return;
        }
        
        getLocationBtn.textContent = '取得中...';
        getLocationBtn.disabled = true;
        
        navigator.geolocation.getCurrentPosition((pos) => {
            locationLatInput.value = pos.coords.latitude;
            locationLngInput.value = pos.coords.longitude;
            getLocationBtn.textContent = '已取得';
            addLocationBtn.disabled = false;
            showNotification("位置已成功取得！", "success");
        }, (err) => {
            showNotification(t("ERROR_GEOLOCATION", { msg: err.message }), "error");
            getLocationBtn.textContent = '取得當前位置';
            getLocationBtn.disabled = false;
        });
    });
    // 處理新增打卡地點
    document.getElementById('add-location-btn').addEventListener('click', async () => {
        const name = document.getElementById('location-name').value;
        const lat = document.getElementById('location-lat').value;
        const lng = document.getElementById('location-lng').value;
        
        if (!name || !lat || !lng) {
            showNotification("請填寫所有欄位並取得位置", "error");
            return;
        }
        
        try {
            const res = await callApifetch(`addLocation&name=${encodeURIComponent(name)}&lat=${encodeURIComponent(lat)}&lng=${encodeURIComponent(lng)}`);
            if (res.ok) {
                showNotification("地點新增成功！", "success");
                // 清空輸入欄位
                document.getElementById('location-name').value = '';
                document.getElementById('location-lat').value = '';
                document.getElementById('location-lng').value = '';
                // 重設按鈕狀態
                getLocationBtn.textContent = '取得當前位置';
                getLocationBtn.disabled = false;
                addLocationBtn.disabled = true;
            } else {
                showNotification("新增地點失敗：" + res.msg, "error");
            }
        } catch (err) {
            console.error(err);
        }
    });
    // UI切換邏輯
    const switchTab = (tabId) => {
        // 修改這一行，加入 'shift-view'
        const tabs = ['dashboard-view', 'monthly-view', 'location-view', 'shift-view', 'admin-view', 'overtime-view', 'leave-view', 'salary-view'];
        
        // 修改這一行，加入 'tab-shift-btn'
        const btns = ['tab-dashboard-btn', 'tab-monthly-btn', 'tab-location-btn', 'tab-shift-btn', 'tab-admin-btn', 'tab-overtime-btn', 'tab-leave-btn', 'tab-salary-btn'];
    
        // 1. 移除舊的 active 類別和 CSS 屬性
        tabs.forEach(id => {
            const tabElement = document.getElementById(id);
            tabElement.style.display = 'none';
            tabElement.classList.remove('active');
        });
        
        // 2. 移除按鈕的選中狀態
        btns.forEach(id => {
            const btnElement = document.getElementById(id);
            if (btnElement) {
                btnElement.classList.replace('bg-indigo-600', 'bg-gray-200');
                btnElement.classList.replace('text-white', 'text-gray-600');
                btnElement.classList.add('dark:text-gray-300', 'dark:bg-gray-700');
            }
        });
        
        // 3. 顯示新頁籤並新增 active 類別
        const newTabElement = document.getElementById(tabId);
        newTabElement.style.display = 'block';
        newTabElement.classList.add('active');
        
        // 4. 設定新頁籤按鈕的選中狀態
        const newBtnElement = document.getElementById(`tab-${tabId.replace('-view', '-btn')}`);
        if (newBtnElement) {
            newBtnElement.classList.replace('bg-gray-200', 'bg-indigo-600');
            newBtnElement.classList.replace('text-gray-600', 'text-white');
            newBtnElement.classList.remove('dark:text-gray-300', 'dark:bg-gray-700');
            newBtnElement.classList.add('dark:bg-indigo-500');
        }
        
        // 5. 根據頁籤 ID 執行特定動作
        if (tabId === 'monthly-view') {
            renderCalendar(currentMonthDate);
        } else if (tabId === 'location-view') {
            initLocationMap();
        } else if (tabId === 'shift-view') { // 新增：排班分頁初始化
            initShiftTab();
        } else if (tabId === 'admin-view') {
            fetchAndRenderReviewRequests();
            loadPendingOvertimeRequests();
            loadPendingLeaveRequests();
        } else if (tabId === 'overtime-view') {
            initOvertimeTab();
        } else if (tabId === 'leave-view') {
            initLeaveTab();
        } else if (tabId === 'salary-view') { // 👈 新增
            initSalaryTab();
        }
        
    };
    
    // 語系初始化
    let currentLang = localStorage.getItem("lang"); // 先從 localStorage 讀取上次的設定
    
    // 如果 localStorage 沒有紀錄，才根據瀏覽器設定判斷
    if (!currentLang) {
        const browserLang = navigator.language || navigator.userLanguage;
        if (browserLang.startsWith("zh")) {
            currentLang = "zh-TW";
        } else if (browserLang.startsWith("ja")) {
            currentLang = "ja"; // 建議使用 ja.json，所以這裡可以只用 'ja'
        } else if (browserLang.startsWith("vi")) {
            currentLang = "vi";
        } else if (browserLang.startsWith("id")) {
            currentLang = "id";
        } else if (browserLang.startsWith("ko")) {
            currentLang = "ko";
        } else if (browserLang.startsWith("th")) {
            currentLang = "th";
        } else {
            currentLang = "en-US";
        }
    }
    // 在這裡設定語言切換器的值
    document.getElementById('language-switcher').value = currentLang;
    // 將最終確定的語言存入 localStorage 並載入翻譯
    localStorage.setItem("lang", currentLang);
    await loadTranslations(currentLang);
    
    
    
    const params = new URLSearchParams(window.location.search);
    const otoken = params.get('code');
    
    if (otoken) {
        try {
            const res = await callApifetch(`getProfile&otoken=${otoken}`);
            if (res.ok && res.sToken) {
                localStorage.setItem("sessionToken", res.sToken);
                history.replaceState({}, '', window.location.pathname);
                ensureLogin();
            } else {
                showNotification(t("ERROR_LOGIN_FAILED", { msg: res.msg || t("UNKNOWN_ERROR") }), "error");
                loginBtn.style.display = 'block';
            }
        } catch (err) {
            console.error(err);
            loginBtn.style.display = 'block';
        }
    } else {
        ensureLogin();
    }
    
    // 綁定按鈕事件
    loginBtn.onclick = async () => {
        const res = await callApifetch("getLoginUrl");
        if (res.url) window.location.href = res.url;
    };
    
    logoutBtn.onclick = () => {
        localStorage.removeItem("sessionToken");
        window.location.href = "/New_Attendance-System"
    };
    
    /* ===== 打卡功能 ===== */
    function generalButtonState(button, state, loadingText = '處理中...') {
        if (!button) return;
        const loadingClasses = 'opacity-50 cursor-not-allowed';

        if (state === 'processing') {
            // --- 進入處理中狀態 ---
            
            // 1. 儲存原始文本 (用於恢復)
            button.dataset.originalText = button.textContent;
            
            // 2. 儲存原始類別 (用於恢復樣式)
            // 這是為了在恢復時移除我們為了禁用而添加的類別
            button.dataset.loadingClasses = 'opacity-50 cursor-not-allowed';

            // 3. 禁用並設置處理中文字
            button.disabled = true;
            button.textContent = loadingText; // 使用傳入的 loadingText
            
            // 4. 添加視覺反饋 (禁用時的樣式)
            button.classList.add(...loadingClasses.split(' '));
            
            // 可選：移除 hover 效果，防止滑鼠移動時顏色變化
            // 假設您的按鈕有 hover:opacity-100 之類的類別，這裡需要調整
            
        } else {
            // --- 恢復到原始狀態 ---
            
            // 1. 移除視覺反饋
            if (button.dataset.loadingClasses) {
                button.classList.remove(...button.dataset.loadingClasses.split(' '));
            }

            // 2. 恢復禁用狀態
            button.disabled = false;
            
            // 3. 恢復原始文本
            if (button.dataset.originalText) {
                button.textContent = button.dataset.originalText;
                delete button.dataset.originalText; // 清除儲存，讓它在下一次點擊時再次儲存
            }
        }
    }

        /**
     * 輔助函數：計算時間差（分鐘）
     * @param {string} time1 - 時間 1，格式 "HH:MM"
     * @param {string} time2 - 時間 2，格式 "HH:MM"
     * @returns {number} - 時間差（分鐘），正數表示 time1 晚於 time2
     */
    function getTimeDifference(time1, time2) {
        const [h1, m1] = time1.split(':').map(Number);
        const [h2, m2] = time2.split(':').map(Number);
        
        const minutes1 = h1 * 60 + m1;
        const minutes2 = h2 * 60 + m2;
        
        return minutes1 - minutes2;
    }

    async function doPunch(type) {
        const punchButtonId = type === '上班' ? 'punch-in-btn' : 'punch-out-btn';
        
        // 獲取按鈕元素
        const button = document.getElementById(punchButtonId);
        const loadingText = t('LOADING') || '處理中...';
    
        // 檢查按鈕是否存在，若不存在則直接返回
        if (!button) return;
    
        // A. 進入處理中狀態
        generalButtonState(button, 'processing', loadingText);
        
        // ==================== 新增：上班打卡前檢查排班 ====================
        if (type === '上班') {
            try {
                const userId = localStorage.getItem('sessionUserId');
                const today = new Date().toISOString().split('T')[0];
                
                // 呼叫排班 API
                const shiftRes = await callApifetch(`getEmployeeShiftForDate&employeeId=${userId}&date=${today}`);
                
                if (shiftRes.ok && shiftRes.hasShift) {
                    const shift = shiftRes.data;
                    
                    // 顯示排班資訊提示
                    showNotification(
                        t('SHIFT_INFO_NOTIFICATION', {
                            shiftType: shift.shiftType,
                            startTime: shift.startTime,
                            endTime: shift.endTime
                        }) || `今日排班：${shift.shiftType} (${shift.startTime}-${shift.endTime})`,
                        'info'
                    );
                    
                    // 可選：檢查打卡時間是否合理
                    const now = new Date();
                    const currentTime = `${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`;
                    
                    if (shift.startTime) {
                        const timeDiff = getTimeDifference(currentTime, shift.startTime);
                        
                        // 如果提前超過 30 分鐘打卡，給予提醒
                        if (timeDiff < -30) {
                            showNotification(
                                t('EARLY_PUNCH_WARNING') || `注意：您的排班時間是 ${shift.startTime}，目前提前超過 30 分鐘打卡。`,
                                'warning'
                            );
                        }
                        // 如果遲到超過 30 分鐘，給予提醒
                        else if (timeDiff > 30) {
                            showNotification(
                                t('LATE_PUNCH_WARNING') || `注意：您的排班時間是 ${shift.startTime}，目前已遲到超過 30 分鐘。`,
                                'warning'
                            );
                        }
                    }
                } else {
                    // 今日沒有排班，可選擇是否提醒
                    // showNotification(t('NO_SHIFT_TODAY') || '今日無排班記錄', 'info');
                }
            } catch (error) {
                console.error('檢查排班失敗:', error);
                // 排班檢查失敗不影響打卡流程，繼續執行
            }
        }
        // ==================== 排班檢查結束 ====================
        
        // 檢查瀏覽器是否支援定位
        if (!navigator.geolocation) {
            showNotification(t("ERROR_GEOLOCATION", { msg: "您的瀏覽器不支援地理位置功能。" }), "error");
            
            // B. 退出點 1: 不支援定位，恢復按鈕狀態
            generalButtonState(button, 'idle');
            return;
        }
        
        // C. 處理地理位置的異步回呼
        navigator.geolocation.getCurrentPosition(async (pos) => {
            // --- 定位成功：執行 API 請求 ---
            const lat = pos.coords.latitude;
            const lng = pos.coords.longitude;
            const action = `punch&type=${encodeURIComponent(type)}&lat=${lat}&lng=${lng}&note=${encodeURIComponent(navigator.userAgent)}`;
            
            try {
                const res = await callApifetch(action);
                const msg = t(res.code || "UNKNOWN_ERROR", res.params || {});
                showNotification(msg, res.ok ? "success" : "error");
                
                // 打卡成功後，清除排班快取（以便下次載入最新資料）
                if (res.ok && type === '上班') {
                    clearShiftCache();
                }
                
                // D. 退出點 2: API 成功，恢復按鈕狀態
                generalButtonState(button, 'idle');
            } catch (err) {
                console.error(err);
                
                // E. 退出點 3: API 失敗，恢復按鈕狀態
                generalButtonState(button, 'idle');
            }
            
        }, (err) => {
            // --- 定位失敗：處理權限錯誤等 ---
            showNotification(t("ERROR_GEOLOCATION", { msg: err.message }), "error");
            
            // F. 退出點 4: 定位回呼失敗，恢復按鈕狀態
            generalButtonState(button, 'idle');
        });
    }
    
    punchInBtn.addEventListener('click', () => doPunch("上班"));
    punchOutBtn.addEventListener('click', () => doPunch("下班"));

    // 處理補打卡表單
    // ✅ 修正：使用事件委派處理補打卡按鈕
    abnormalList.addEventListener('click', (e) => {
        // ⭐ 關鍵：使用 closest 找到被點擊的按鈕
        const button = e.target.closest('.adjust-btn');
        
        if (button) {
            const date = button.dataset.date;
            const type = button.dataset.type;
            
            console.log(`點擊補打卡: ${date} - ${type}`);
            
            const formHtml = `
                <div class="p-4 border-t border-gray-200 dark:border-gray-600 fade-in">
                    <p class="font-semibold mb-2 dark:text-white">
                        補打卡：<span class="text-indigo-600 dark:text-indigo-400">${date}</span>
                        <span class="ml-2 text-sm ${type === '上班' ? 'text-indigo-600' : 'text-purple-600'}">
                            (${type})
                        </span>
                    </p>
                    <div class="form-group mb-3">
                        <label for="adjustDateTime" class="block text-sm font-medium text-gray-700 mb-1 dark:text-gray-300">
                            選擇${type}時間：
                        </label>
                        <input id="adjustDateTime" 
                            type="datetime-local" 
                            class="w-full p-2 border border-gray-300 dark:border-gray-600 rounded-md shadow-sm dark:bg-gray-700 dark:text-white focus:ring-indigo-500 focus:border-indigo-500">
                    </div>
                    <div class="grid grid-cols-2 gap-2">
                        <button id="cancel-adjust-btn" 
                                class="py-2 px-4 rounded-lg font-bold bg-gray-300 dark:bg-gray-600 text-gray-700 dark:text-gray-200 hover:bg-gray-400 dark:hover:bg-gray-500">
                            取消
                        </button>
                        <button id="submit-adjust-btn" 
                                data-type="${type}"
                                data-date="${date}"
                                class="py-2 px-4 rounded-lg font-bold btn-primary">
                            提交補${type}卡
                        </button>
                    </div>
                </div>
            `;
            
            adjustmentFormContainer.innerHTML = formHtml;
            
            const adjustDateTimeInput = document.getElementById("adjustDateTime");
            const defaultTime = type === '上班' ? '09:00' : '18:00';
            adjustDateTimeInput.value = `${date}T${defaultTime}`;
            
            // ⭐ 綁定取消按鈕
            document.getElementById('cancel-adjust-btn').addEventListener('click', () => {
                adjustmentFormContainer.innerHTML = '';
            });
        }
    });
    
    function validateAdjustTime(value) {
        const selected = new Date(value);
        const now = new Date();
        const yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
        const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
        const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
        if (selected < monthStart) {
            showNotification(t("ERR_BEFORE_MONTH_START"), "error");
            return false;
        }
        // 不允許選今天以後
        if (selected > today) {
            showNotification(t("ERR_AFTER_TODAY"), "error");
            return false;
        }
        return true;
    }


    adjustmentFormContainer.addEventListener('click', async (e) => {
        const button = e.target.closest('#submit-adjust-btn');
        
        if (button) {
            const loadingText = t('LOADING') || '處理中...';
            
            const datetime = document.getElementById("adjustDateTime").value;
            const type = button.dataset.type;
            const date = button.dataset.date;
            
            if (!datetime) {
                showNotification("請選擇補打卡日期時間", "error");
                return;
            }
            
            if (!validateAdjustTime(datetime)) return;
            
            generalButtonState(button, 'processing', loadingText);
            
            try {
                const sessionToken = localStorage.getItem("sessionToken");
                
                const position = await new Promise((resolve, reject) => {
                    navigator.geolocation.getCurrentPosition(resolve, reject);
                });
                
                const lat = position.coords.latitude;
                const lng = position.coords.longitude;
                
                // ✅ 修正：note 改為裝置資訊
                const params = new URLSearchParams({
                    token: sessionToken,
                    type: type,
                    lat: lat,
                    lng: lng,
                    datetime: datetime,
                    note: navigator.userAgent  // ⭐ 改成裝置資訊
                });
                
                const res = await callApifetch(`adjustPunch&${params.toString()}`);
                
                if (res.ok) {
                    showNotification("補打卡申請成功！等待管理員審核", "success");
                    await checkAbnormal();
                    adjustmentFormContainer.innerHTML = '';
                } else {
                    showNotification(t(res.code) || "補打卡失敗", "error");
                }
                
            } catch (err) {
                console.error('補打卡錯誤:', err);
                showNotification("補打卡失敗", "error");
                
            } finally {
                if (adjustmentFormContainer.innerHTML !== '') {
                    generalButtonState(button, 'idle');
                }
            }
        }
    });
    

    // 頁面切換事件
    const tabShiftBtn = document.getElementById('tab-shift-btn');

    // 在現有的分頁按鈕事件後面加入：
    tabShiftBtn.addEventListener('click', () => {switchTab('shift-view');});


    tabSalaryBtn.addEventListener('click', () => {switchTab('salary-view');});
    tabDashboardBtn.addEventListener('click', () => switchTab('dashboard-view'));
    
    tabLocationBtn.addEventListener('click', () => switchTab('location-view'));
    tabMonthlyBtn.addEventListener('click', () => switchTab('monthly-view'));
    tabOvertimeBtn.addEventListener('click', () => {
        switchTab('overtime-view');
        initOvertimeTab();
    });

    // 👈 新增請假按鈕事件
    tabLeaveBtn.addEventListener('click', () => {
        switchTab('leave-view');
        initLeaveTab();
    });

    // tabAdminBtn.addEventListener('click', () => {
    //     // 簡化版：直接切換分頁
    //     // 按鈕只有管理員能看到，無需再次驗證
    //     switchTab('admin-view');
    // });
    tabAdminBtn.addEventListener('click', async () => {
    
        // 獲取按鈕元素和處理中文字
        const button = tabAdminBtn;
        const loadingText = t('CHECKING') || '檢查中...';
        
        // A. 進入處理中狀態
        generalButtonState(button, 'processing', loadingText);
        
        try {
            // ✅ 修正：改用 initApp（與 ensureLogin 一致）
            const res = await callApifetch("initApp");
            
            console.log('🔍 管理員權限檢查:', res);
            console.log('   - ok:', res.ok);
            console.log('   - user:', res.user);
            console.log('   - dept:', res.user?.dept);
            
            // 檢查回傳的結果和權限
            if (res.ok && res.user && res.user.dept === "管理員") {
                console.log('✅ 管理員權限驗證通過');
                // 如果 Session 有效且是管理員，執行頁籤切換
                switchTab('admin-view');
            } else {
                console.log('❌ 權限驗證失敗');
                console.log('   實際部門:', res.user?.dept);
                // 如果權限不足或 Session 無效，給予錯誤提示
                showNotification(t("ERR_NO_PERMISSION") || "您沒有權限執行此操作", "error");
            }
            
        } catch (err) {
            // 處理網路錯誤或 API 呼叫失敗
            console.error('❌ API 呼叫錯誤:', err);
            showNotification(t("NETWORK_ERROR") || '網絡錯誤', "error");
            
        } finally {
            // B. 無論 API 成功、失敗或網路錯誤，都要恢復按鈕狀態
            generalButtonState(button, 'idle');
        }
    });
    // 月曆按鈕事件
    document.getElementById('prev-month').addEventListener('click', () => {
        currentMonthDate.setMonth(currentMonthDate.getMonth() - 1);
        renderCalendar(currentMonthDate);
    });
    
    document.getElementById('next-month').addEventListener('click', () => {
        currentMonthDate.setMonth(currentMonthDate.getMonth() + 1);
        renderCalendar(currentMonthDate);
    });

    const exportAttendanceBtn = document.getElementById('export-attendance-btn');
    if (exportAttendanceBtn) {
        exportAttendanceBtn.addEventListener('click', () => {
            exportAttendanceReport(currentMonthDate);
        });
    }

    const adminExportAllBtn = document.getElementById('admin-export-all-btn');
    const adminExportMonthInput = document.getElementById('admin-export-month');

    if (adminExportAllBtn && adminExportMonthInput) {
        // 設定預設月份為當月
        const now = new Date();
        const defaultMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
        adminExportMonthInput.value = defaultMonth;
        
        // 綁定按鈕點擊事件
        adminExportAllBtn.addEventListener('click', () => {
            const selectedMonth = adminExportMonthInput.value;
            
            if (!selectedMonth) {
                showNotification('請選擇要匯出的月份', 'error');
                return;
            }
            
            exportAllEmployeesReport(selectedMonth);
        });
    }
    // 語系切換事件
    document.getElementById('language-switcher').addEventListener('change', (e) => {
        const newLang = e.target.value;
        loadTranslations(newLang);
        // 取得當前顯示的標籤頁ID
        const currentTab = document.querySelector('.active');
        const currentTabId = currentTab ? currentTab.id : null;
        console.log(currentTabId);
        // 如果當前頁面是「地圖」頁籤，則重新載入地圖
        if (currentTabId === 'location-view') {
            initLocationMap(true); // 重新載入地圖
        }
    });
    // 點擊日曆日期的事件監聽器
    calendarGrid.addEventListener('click', (e) => {
        if (e.target.classList.contains('day-cell') && e.target.dataset.date) {
            const date = e.target.dataset.date;
            renderDailyRecords(date);
        }
    });
});

/**
 * 初始化排班分頁
 */
function initShiftTab() {
    loadTodayShift();
    loadWeekShift();
}

/**
 * 載入今日排班
 */
async function loadTodayShift() {
    const loadingEl = document.getElementById('today-shift-loading');
    const emptyEl = document.getElementById('today-shift-empty');
    const infoEl = document.getElementById('today-shift-info');
    
    // 如果有快取，直接使用
    if (todayShiftCache !== null) {
        displayTodayShift(todayShiftCache);
        return;
    }
    
    try {
        loadingEl.style.display = 'block';
        emptyEl.style.display = 'none';
        infoEl.style.display = 'none';
        
        const userId = localStorage.getItem('sessionUserId');
        const today = new Date().toISOString().split('T')[0];
        
        const res = await callApifetch(`getEmployeeShiftForDate&employeeId=${userId}&date=${today}`);
        
        loadingEl.style.display = 'none';
        
        // 快取結果
        todayShiftCache = res;
        displayTodayShift(res);
        
    } catch (error) {
        console.error('載入今日排班失敗:', error);
        loadingEl.style.display = 'none';
        emptyEl.style.display = 'block';
    }
}

/**
 * 顯示今日排班
 */
function displayTodayShift(res) {
    const emptyEl = document.getElementById('today-shift-empty');
    const infoEl = document.getElementById('today-shift-info');
    
    if (res.ok && res.hasShift) {
        document.getElementById('shift-type').textContent = res.data.shiftType;
        document.getElementById('shift-time').textContent = 
            `${res.data.startTime} - ${res.data.endTime}`;
        document.getElementById('shift-location').textContent = res.data.location;
        infoEl.style.display = 'block';
    } else {
        emptyEl.style.display = 'block';
    }
}

/**
 * ✅ 載入未來 7 天排班（完全修正版 - 強制清除舊快取）
 */
async function loadWeekShift() {
    const loadingEl = document.getElementById('week-shift-loading');
    const emptyEl = document.getElementById('week-shift-empty');
    const listEl = document.getElementById('week-shift-list');
    
    // ✅ 步驟 1: 計算「今天到未來 7 天」的範圍
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    
    const startOfWeek = today;
    const endOfWeek = new Date(today);
    endOfWeek.setDate(today.getDate() + 7);
    
    const startDateStr = startOfWeek.toISOString().split('T')[0];
    const endDateStr = endOfWeek.toISOString().split('T')[0];
    
    console.log('📅 未來排班範圍:', {
        today: today.toISOString().split('T')[0],
        startOfWeek: startDateStr,
        endOfWeek: endDateStr
    });
    
    // ✅ 步驟 2: 生成快取鍵值
    const cacheKey = `${startDateStr}_${endDateStr}`;
    
    // ✅ 步驟 3: 檢查快取（但只有在「分頁初次載入」時才使用）
    // 如果快取存在且日期範圍相同，才使用快取
    if (weekShiftCache !== null && 
        weekShiftCache.cacheKey === cacheKey &&
        Date.now() - weekShiftCache.timestamp < 60000) { // 快取 1 分鐘有效
        
        console.log('✅ 使用有效快取（1 分鐘內）');
        displayWeekShift(weekShiftCache.data);
        return;
    }
    
    // ✅ 步驟 4: 清除舊快取，強制重新載入
    console.log('🗑️ 清除舊快取，重新載入');
    weekShiftCache = null;
    
    try {
        loadingEl.style.display = 'block';
        emptyEl.style.display = 'none';
        listEl.innerHTML = '';
        
        const userId = localStorage.getItem('sessionUserId');
        
        const filters = {
            employeeId: userId,
            startDate: startDateStr,
            endDate: endDateStr
        };
        
        console.log('📡 呼叫 API，篩選條件:', filters);
        
        const res = await callApifetch(`getShifts&filters=${encodeURIComponent(JSON.stringify(filters))}`);
        
        console.log('📤 API 回應:', res);
        
        loadingEl.style.display = 'none';
        
        // ✅ 步驟 5: 快取新資料
        weekShiftCache = {
            cacheKey: cacheKey,
            data: res,
            timestamp: Date.now()
        };
        
        console.log('💾 已快取新資料:', weekShiftCache);
        
        // ✅ 步驟 6: 顯示資料
        displayWeekShift(res);
        
    } catch (error) {
        console.error('❌ 載入未來排班失敗:', error);
        loadingEl.style.display = 'none';
        emptyEl.style.display = 'block';
    }
}
/**
 * 顯示本週排班
 */
function displayWeekShift(res) {
    const emptyEl = document.getElementById('week-shift-empty');
    const listEl = document.getElementById('week-shift-list');
    
    console.log('📋 displayWeekShift 收到的資料:', res);
    
    if (res.ok && res.data && res.data.length > 0) {
        listEl.innerHTML = '';
        
        console.log('✅ 開始渲染', res.data.length, '筆排班');
        
        res.data.forEach((shift, index) => {
            console.log(`   ${index + 1}. ${shift.date} - ${shift.shiftType}`);
            
            const item = document.createElement('div');
            item.className = 'flex justify-between items-center text-sm bg-white dark:bg-gray-800 p-2 rounded-md';
            item.innerHTML = `
                <div>
                    <span class="font-semibold text-purple-900 dark:text-purple-200">
                        ${formatShiftDate(shift.date)}
                    </span>
                    <span class="text-purple-700 dark:text-purple-400 ml-2">
                        ${shift.shiftType}
                    </span>
                </div>
                <div class="text-purple-700 dark:text-purple-400">
                    ${shift.startTime} - ${shift.endTime}
                </div>
            `;
            listEl.appendChild(item);
        });
        
        emptyEl.style.display = 'none';
    } else {
        console.log('⚠️ 沒有排班資料或資料格式錯誤');
        emptyEl.style.display = 'block';
        listEl.innerHTML = '';
    }
}

/**
 * 格式化排班日期
 */
function formatShiftDate(dateString) {
    const date = new Date(dateString);
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const weekdays = ['日', '一', '二', '三', '四', '五', '六'];
    const weekday = weekdays[date.getDay()];
    
    return `${month}/${day} (${weekday})`;
}

/**
 * 清除排班快取（當有更新時使用）
 */
function clearShiftCache() {
    todayShiftCache = null;
    weekShiftCache = null;
}