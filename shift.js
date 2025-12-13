/**
 * 排班管理前端邏輯 - 完整版(含月曆功能)
 * 功能: 查看/新增/編輯/刪除排班、批量上傳、月曆顯示、統計分析
 */

// ========== 全域變數 ==========
let currentShifts = [];
let allEmployees = [];
let allLocations = [];
let batchData = [];
let translations = {}; // 👈 新增：翻譯物件
let currentLang = localStorage.getItem("lang") || "zh-TW"; // 👈 新增：當前語言

// 月曆專用全域變數
let currentYear = new Date().getFullYear();
let currentMonth = new Date().getMonth(); // 0-11
let allMonthShifts = [];

// 👇 新增：翻譯函式
function t(code, params = {}) {
    let text = translations[code] || code;
    
    for (const key in params) {
        let paramValue = params[key];
        if (paramValue in translations) {
            paramValue = translations[paramValue];
        }
        text = text.replace(`{${key}}`, paramValue);
    }
    return text;
}

// 👇 新增：載入翻譯檔案
async function loadTranslations(lang) {
    try {
        const res = await fetch(`https://eric693.github.io/check_manager_plus/i18n/${lang}.json`);
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

// 👇 新增：渲染翻譯
function renderTranslations(container = document) {
    if (container === document) {
        document.title = t("SHIFT_PAGE_TITLE");
    }

    const elementsToTranslate = container.querySelectorAll('[data-i18n]');
    elementsToTranslate.forEach(element => {
        const key = element.getAttribute('data-i18n');
        const translatedText = t(key);
        
        if (translatedText !== key) {
            if (element.tagName === 'INPUT' || element.tagName === 'TEXTAREA') {
                element.placeholder = translatedText;
            } else {
                element.textContent = translatedText;
            }
        }
    });

    const selectElements = container.querySelectorAll('select');
    selectElements.forEach(select => {
        const options = select.querySelectorAll('option[data-i18n-option]');
        options.forEach(option => {
            const key = option.getAttribute('data-i18n-option');
            if (key) {
                const translatedText = t(key);
                if (translatedText !== key) {
                    option.textContent = translatedText;
                }
            }
        });
    });
}

// ========== 初始化 ==========
document.addEventListener('DOMContentLoaded', async function() { 
    await loadTranslations(currentLang);
    initializeTabs();
    loadEmployees();
    loadLocations();
    loadShifts();
    setupEventListeners();
    setupBatchUpload();
    
    // 設定預設日期為今天
    const today = new Date().toISOString().split('T')[0];
    const shiftDateEl = document.getElementById('shift-date');
    if (shiftDateEl) shiftDateEl.value = today;
    
    // 設定篩選日期為本週
    const startOfWeek = new Date();
    startOfWeek.setDate(startOfWeek.getDate() - startOfWeek.getDay());
    const endOfWeek = new Date(startOfWeek);
    endOfWeek.setDate(endOfWeek.getDate() + 6);
    
    const filterStartEl = document.getElementById('filter-start-date');
    const filterEndEl = document.getElementById('filter-end-date');
    if (filterStartEl) filterStartEl.value = startOfWeek.toISOString().split('T')[0];
    if (filterEndEl) filterEndEl.value = endOfWeek.toISOString().split('T')[0];
});

// ========== 分頁管理 ==========

function initializeTabs() {
    const tabs = document.querySelectorAll('.shift-tab');
    tabs.forEach(tab => {
        tab.addEventListener('click', function() {
            const tabName = this.getAttribute('data-tab');
            switchTab(tabName);
        });
    });
}

function switchTab(tabName) {
    document.querySelectorAll('.shift-tab').forEach(tab => {
        tab.classList.remove('active');
    });
    document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');
    
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
    });
    document.getElementById(`${tabName}-tab`).classList.add('active');
    
    // 載入對應資料
    if (tabName === 'view') {
        loadShifts();
    } else if (tabName === 'stats') {
        loadStats();
    }
}

// ========== 事件監聽器 ==========

function setupEventListeners() {
    const addForm = document.getElementById('add-shift-form');
    if (addForm) {
        addForm.addEventListener('submit', function(e) {
            e.preventDefault();
            addShift();
        });
    }
    
    const shiftTypeEl = document.getElementById('shift-type');
    if (shiftTypeEl) {
        shiftTypeEl.addEventListener('change', function() {
            autoFillShiftTime(this.value);
        });
    }
}

function autoFillShiftTime(shiftType) {
    const times = {
        '早班': ['08:00', '16:00'],
        '中班': ['12:00', '20:00'],
        '晚班': ['16:00', '00:00'],
        '全日班': ['09:00', '18:00'],
        '排休': ['00:00', '00:00'] 
    };
    
    const startTimeInput = document.getElementById('start-time');
    const endTimeInput = document.getElementById('end-time');
    
    if (shiftType === '排休') {
        // 排休時禁用時間選擇
        startTimeInput.value = '00:00';
        endTimeInput.value = '00:00';
        startTimeInput.disabled = true;
        endTimeInput.disabled = true;
    } else if (shiftType === '自訂') {
        // 選擇「自訂」時,清空時間並啟用輸入
        startTimeInput.value = '';
        endTimeInput.value = '';
        startTimeInput.disabled = false;
        endTimeInput.disabled = false;
        startTimeInput.focus();
    } else if (times[shiftType]) {
        // 選擇預設班別時,自動填入時間並啟用
        startTimeInput.value = times[shiftType][0];
        endTimeInput.value = times[shiftType][1];
        startTimeInput.disabled = false;
        endTimeInput.disabled = false;
    }
}

// ==================== 員工載入函式（完整除錯版） ====================

/**
 * ✅ 載入員工列表（加強除錯版）
 */
async function loadEmployees() {
    try {
        const token = localStorage.getItem('sessionToken');
        
        // ✅ 步驟 1: 檢查 token
        if (!token) {
            console.error('❌ 沒有 session token');
            showMessage('請先登入', 'error');
            return;
        }
        
        console.log('═══════════════════════════════════════');
        console.log('📋 載入員工列表');
        console.log('═══════════════════════════════════════');
        console.log('📡 Token:', token.substring(0, 20) + '...');
        console.log('📡 API URL:', apiUrl);
        console.log('');
        
        // ✅ 步驟 2: 呼叫 API
        const url = `${apiUrl}?action=getAllUsers&token=${token}`;
        console.log('📡 完整 URL:', url);
        console.log('📡 開始呼叫 API...');
        
        const response = await fetch(url);
        
        // ✅ 步驟 3: 檢查 HTTP 狀態
        console.log('📤 HTTP 狀態:', response.status, response.statusText);
        
        if (!response.ok) {
            throw new Error(`HTTP 錯誤: ${response.status} ${response.statusText}`);
        }
        
        // ✅ 步驟 4: 解析 JSON
        const data = await response.json();
        
        console.log('');
        console.log('📤 API 回應:');
        console.log('   - ok:', data.ok);
        console.log('   - msg:', data.msg || '無');
        console.log('   - count:', data.count || '無');
        console.log('   - users 存在:', data.users ? '是' : '否');
        console.log('   - users 型別:', typeof data.users);
        console.log('   - users 長度:', data.users ? data.users.length : 'null');
        console.log('');
        
        // ✅ 步驟 5: 檢查回應
        if (data.ok) {
            allEmployees = data.users || [];
            
            console.log('✅ API 回傳成功');
            console.log('   員工數量:', allEmployees.length);
            
            if (allEmployees.length === 0) {
                console.warn('⚠️ 員工列表是空的');
                console.warn('   可能原因:');
                console.warn('   1. 員工工作表沒有資料');
                console.warn('   2. 所有員工都不是「啟用」狀態');
                console.warn('   3. 資料格式不正確');
                showMessage('目前沒有員工資料', 'warning');
            } else {
                console.log('✅ 員工列表預覽（前 5 筆）:');
                allEmployees.slice(0, 5).forEach((emp, index) => {
                    console.log(`   ${index + 1}. ${emp.name} (${emp.userId}) - ${emp.dept}`);
                });
                
                if (allEmployees.length > 5) {
                    console.log(`   ... 還有 ${allEmployees.length - 5} 筆`);
                }
            }
            
            // ✅ 步驟 6: 填入下拉選單
            console.log('');
            console.log('📝 開始填入員工下拉選單...');
            populateEmployeeSelect();
            
        } else {
            console.error('❌ API 回傳失敗');
            console.error('   原因:', data.msg || '未知錯誤');
            showMessage(data.msg || '載入員工列表失敗', 'error');
        }
        
        console.log('═══════════════════════════════════════');
        
    } catch (error) {
        console.error('');
        console.error('❌❌❌ 載入員工列表失敗');
        console.error('錯誤訊息:', error.message);
        console.error('錯誤堆疊:', error.stack);
        console.error('═══════════════════════════════════════');
        
        showMessage('載入員工列表失敗: ' + error.message, 'error');
    }
}

/**
 * ✅ 填入員工下拉選單（加強除錯版）
 */
function populateEmployeeSelect() {
    console.log('');
    console.log('📝 populateEmployeeSelect 開始');
    console.log('───────────────────────────────────────');
    
    const select = document.getElementById('employee-select');
    
    // ✅ 檢查元素是否存在
    if (!select) {
        console.error('❌ 找不到 employee-select 元素');
        console.error('   請檢查 HTML 中是否有:');
        console.error('   <select id="employee-select">');
        return;
    }
    
    console.log('✅ 找到 employee-select 元素');
    console.log('   當前選項數量:', select.options.length);
    
    // ✅ 檢查員工列表
    if (!allEmployees) {
        console.error('❌ allEmployees 是 undefined 或 null');
        return;
    }
    
    if (!Array.isArray(allEmployees)) {
        console.error('❌ allEmployees 不是陣列');
        console.error('   型別:', typeof allEmployees);
        console.error('   內容:', allEmployees);
        return;
    }
    
    console.log('✅ allEmployees 驗證通過');
    console.log('   員工數量:', allEmployees.length);
    
    // ✅ 清空並重設為預設選項
    select.innerHTML = '<option value="">請選擇員工</option>';
    console.log('✅ 已重設為預設選項');
    
    if (allEmployees.length === 0) {
        console.warn('⚠️ 沒有員工可以填入');
        select.innerHTML = '<option value="">目前沒有員工資料</option>';
        return;
    }
    
    // ✅ 填入員工選項
    console.log('📝 開始逐筆填入...');
    
    let successCount = 0;
    let failCount = 0;
    
    allEmployees.forEach((emp, index) => {
        try {
            // 驗證必要欄位
            if (!emp.userId) {
                console.warn(`   ⚠️ 第 ${index + 1} 筆: 缺少 userId，跳過`);
                failCount++;
                return;
            }
            
            if (!emp.name) {
                console.warn(`   ⚠️ 第 ${index + 1} 筆: 缺少 name，跳過`);
                failCount++;
                return;
            }
            
            const option = document.createElement('option');
            option.value = emp.userId;
            option.textContent = `${emp.name} (${emp.userId})`;
            option.dataset.name = emp.name;
            
            if (emp.dept) {
                option.textContent += ` - ${emp.dept}`;
            }
            
            select.appendChild(option);
            
            successCount++;
            
            // 只顯示前 5 筆的詳細資訊
            if (index < 5) {
                console.log(`   ✅ ${index + 1}. ${emp.name} (${emp.userId})`);
            }
            
        } catch (error) {
            console.error(`   ❌ 第 ${index + 1} 筆失敗:`, error.message);
            failCount++;
        }
    });
    
    if (allEmployees.length > 5) {
        console.log(`   ... 還有 ${allEmployees.length - 5} 筆（已略過顯示）`);
    }
    
    console.log('');
    console.log('📊 填入結果:');
    console.log('   成功:', successCount, '筆');
    console.log('   失敗:', failCount, '筆');
    console.log('   總計:', allEmployees.length, '筆');
    console.log('   最終選項數量:', select.options.length, '個（含預設選項）');
    console.log('───────────────────────────────────────');
    console.log('✅ populateEmployeeSelect 完成');
    console.log('');
}

// ==================== 除錯工具函式 ====================

/**
 * 🧪 手動測試載入員工
 * 在瀏覽器 Console 中執行: testLoadEmployees()
 */
async function testLoadEmployees() {
    console.log('🧪 手動測試載入員工');
    console.log('');
    
    // 檢查 apiUrl
    console.log('1️⃣ 檢查 apiUrl:');
    console.log('   apiUrl:', typeof apiUrl !== 'undefined' ? apiUrl : '❌ undefined');
    console.log('');
    
    // 檢查 token
    console.log('2️⃣ 檢查 token:');
    const token = localStorage.getItem('sessionToken');
    console.log('   token 存在:', token ? '✅ 是' : '❌ 否');
    if (token) {
        console.log('   token 預覽:', token.substring(0, 20) + '...');
    }
    console.log('');
    
    // 檢查 HTML 元素
    console.log('3️⃣ 檢查 HTML 元素:');
    const select = document.getElementById('employee-select');
    console.log('   employee-select 存在:', select ? '✅ 是' : '❌ 否');
    if (select) {
        console.log('   當前選項數量:', select.options.length);
    }
    console.log('');
    
    // 執行載入
    console.log('4️⃣ 開始載入員工列表...');
    console.log('');
    
    await loadEmployees();
    
    console.log('');
    console.log('5️⃣ 檢查結果:');
    console.log('   allEmployees 存在:', typeof allEmployees !== 'undefined' ? '✅ 是' : '❌ 否');
    if (typeof allEmployees !== 'undefined') {
        console.log('   allEmployees 長度:', allEmployees.length);
    }
    if (select) {
        console.log('   下拉選單選項數量:', select.options.length);
    }
}

async function loadLocations() {
    try {
        const token = localStorage.getItem('sessionToken');
        const response = await fetch(`${apiUrl}?action=getLocations&token=${token}`);
        const data = await response.json();
        
        console.log('✅ 地點列表回應:', data);
        
        if (data.ok) {
            allLocations = data.locations || [];
            populateLocationSelects();
        }
    } catch (error) {
        console.error('載入地點列表失敗:', error);
    }
}

function populateLocationSelects() {
    const selects = ['shift-location', 'filter-location'];
    
    selects.forEach(id => {
        const select = document.getElementById(id);
        if (!select) return;
        
        const currentValue = select.value;
        select.innerHTML = id === 'filter-location' ? 
            '<option value="">全部</option>' : 
            '<option value="">請選擇地點</option>';
        
        allLocations.forEach(loc => {
            const option = document.createElement('option');
            option.value = loc.name;
            option.textContent = loc.name;
            select.appendChild(option);
        });
        
        if (currentValue) select.value = currentValue;
    });
}

async function loadShifts(filters = {}) {
    const listContainer = document.getElementById('shift-list');
    if (!listContainer) return;
    
    listContainer.innerHTML = '<div class="loading">載入中</div>';
    
    try {
        const token = localStorage.getItem('sessionToken');
        
        // 使用預設日期範圍
        if (!filters.startDate && !filters.endDate) {
            const startDateEl = document.getElementById('filter-start-date');
            const endDateEl = document.getElementById('filter-end-date');
            if (startDateEl && startDateEl.value) filters.startDate = startDateEl.value;
            if (endDateEl && endDateEl.value) filters.endDate = endDateEl.value;
        }
        
        const queryParams = new URLSearchParams({
            action: 'getShifts',
            token: token
        });
        
        if (filters.employeeId) queryParams.append('employeeId', filters.employeeId);
        if (filters.startDate) queryParams.append('startDate', filters.startDate);
        if (filters.endDate) queryParams.append('endDate', filters.endDate);
        if (filters.shiftType) queryParams.append('shiftType', filters.shiftType);
        if (filters.location) queryParams.append('location', filters.location);
        
        const response = await fetch(`${apiUrl}?${queryParams}`);
        const data = await response.json();
        
        console.log('✅ 排班回應:', data);
        
        if (data.ok) {
            currentShifts = data.data || [];
            displayShifts(currentShifts);
        } else {
            listContainer.innerHTML = `<div class="empty-state"><div class="empty-state-icon">📋</div><p>載入失敗: ${data.msg}</p></div>`;
        }
    } catch (error) {
        console.error('❌ 載入排班失敗:', error);
        listContainer.innerHTML = '<div class="empty-state"><div class="empty-state-icon">❌</div><p>載入失敗</p></div>';
    }
}

function displayShifts(shifts) {
    const listContainer = document.getElementById('shift-list');
    if (!listContainer) return;
    
    if (shifts.length === 0) {
        listContainer.innerHTML = `
            <div class="empty-state">
                <div class="empty-state-icon">📅</div>
                <p>目前沒有排班資料</p>
            </div>
        `;
        return;
    }
    
    listContainer.innerHTML = '';
    
    shifts.forEach(shift => {
        const shiftItem = createShiftItem(shift);
        listContainer.appendChild(shiftItem);
    });
}

function createShiftItem(shift) {
    const div = document.createElement('div');
    div.className = 'shift-item';
    
    const shiftTypeBadge = getShiftTypeBadge(shift.shiftType);
    
    // 格式化時間
    const startTime = formatTimeOnly(shift.startTime);
    const endTime = formatTimeOnly(shift.endTime);
    
    div.innerHTML = `
        <div class="shift-info">
            <h3>${shift.employeeName} ${shiftTypeBadge}</h3>
            <p>日期: ${formatDate(shift.date)}</p>
            <p>時間: ${startTime} - ${endTime}</p>
            <p>地點: ${shift.location}</p>
            ${shift.note ? `<p>備註: ${shift.note}</p>` : ''}
        </div>
        <div class="shift-actions">
            <button class="btn-icon" onclick="editShift('${shift.shiftId}')">編輯</button>
            <button class="btn-icon btn-danger" onclick="deleteShift('${shift.shiftId}')">刪除</button>
        </div>
    `;
    
    return div;
}

function getShiftTypeBadge(shiftType) {
    const badgeClass = {
        '早班': 'badge-morning',
        '中班': 'badge-afternoon',
        '晚班': 'badge-night',
        '全日班': 'badge-full',
        '排休': 'badge-dayoff',
        '自訂': 'badge-custom'
    }[shiftType] || 'badge-morning';
    
    return `<span class="badge ${badgeClass}">${shiftType}</span>`;
}

function formatDate(dateString) {
    const date = new Date(dateString);
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const weekdays = ['日', '一', '二', '三', '四', '五', '六'];
    const weekday = weekdays[date.getDay()];
    
    return `${year}/${month}/${day} (${weekday})`;
}

async function addShift() {
    const employeeSelect = document.getElementById('employee-select');
    const selectedOption = employeeSelect.selectedOptions[0];
    
    if (!selectedOption || !selectedOption.value) {
        showMessage('請選擇員工', 'error');
        return;
    }
    
    const shiftType = document.getElementById('shift-type').value;
    const startTime = document.getElementById('start-time').value;
    const endTime = document.getElementById('end-time').value;
    
    // 驗證時間欄位
    if (!startTime || !endTime) {
        showMessage('請填寫上班時間和下班時間', 'error');
        return;
    }
    
    // 驗證時間邏輯(結束時間應該晚於開始時間,除非是跨日班)
    if (startTime >= endTime && endTime !== '00:00') {
        const confirmCrossDay = confirm('下班時間早於上班時間,是否為跨日班別?');
        if (!confirmCrossDay) {
            return;
        }
    }
    
    const token = localStorage.getItem('sessionToken');
    const shiftNoteEl = document.getElementById('shift-note');
    
    const shiftData = {
        action: 'addShift',
        token: token,
        employeeId: selectedOption.value,
        employeeName: selectedOption.dataset.name || selectedOption.textContent.split('(')[0].trim(),
        date: document.getElementById('shift-date').value,
        shiftType: shiftType,
        startTime: startTime,
        endTime: endTime,
        location: document.getElementById('shift-location').value,
        note: shiftNoteEl ? shiftNoteEl.value : ''
    };
    
    console.log('📝 新增排班:', shiftData);
    
    try {
        const queryParams = new URLSearchParams(shiftData);
        const response = await fetch(`${apiUrl}?${queryParams}`);
        const data = await response.json();
        
        console.log('✅ 新增回應:', data);
        
        if (data.ok) {
            showMessage('排班新增成功!', 'success');
            resetForm();
            switchTab('view');
            loadShifts();
        } else {
            showMessage(data.msg || '新增失敗', 'error');
        }
    } catch (error) {
        console.error('❌ 新增排班失敗:', error);
        showMessage('新增排班失敗', 'error');
    }
}

async function editShift(shiftId) {
    const shift = currentShifts.find(s => s.shiftId === shiftId);
    if (!shift) return;
    
    switchTab('add');
    
    document.querySelector('#add-tab h2').textContent = '編輯排班';
    document.getElementById('employee-select').value = shift.employeeId;
    document.getElementById('shift-date').value = shift.date;
    document.getElementById('shift-type').value = shift.shiftType;
    
    // 格式化時間
    document.getElementById('start-time').value = formatTimeOnly(shift.startTime);
    document.getElementById('end-time').value = formatTimeOnly(shift.endTime);
    
    document.getElementById('shift-location').value = shift.location;
    
    const shiftNoteEl = document.getElementById('shift-note');
    if (shiftNoteEl) shiftNoteEl.value = shift.note || '';
    
    const submitBtn = document.querySelector('#add-shift-form button[type="submit"]');
    submitBtn.textContent = '更新排班';
    submitBtn.onclick = function(e) {
        e.preventDefault();
        updateShift(shiftId);
    };
}

async function updateShift(shiftId) {
    const employeeSelect = document.getElementById('employee-select');
    const selectedOption = employeeSelect.selectedOptions[0];
    
    if (!selectedOption || !selectedOption.value) {
        showMessage('請選擇員工', 'error');
        return;
    }
    
    const token = localStorage.getItem('sessionToken');
    const shiftNoteEl = document.getElementById('shift-note');
    
    const shiftData = {
        action: 'updateShift',
        token: token,
        shiftId: shiftId,
        employeeId: selectedOption.value,
        employeeName: selectedOption.dataset.name || selectedOption.textContent.split('(')[0].trim(),
        date: document.getElementById('shift-date').value,
        shiftType: document.getElementById('shift-type').value,
        startTime: document.getElementById('start-time').value,
        endTime: document.getElementById('end-time').value,
        location: document.getElementById('shift-location').value,
        note: shiftNoteEl ? shiftNoteEl.value : ''
    };
    
    try {
        const queryParams = new URLSearchParams(shiftData);
        const response = await fetch(`${apiUrl}?${queryParams}`);
        const data = await response.json();
        
        if (data.ok) {
            showMessage('排班更新成功!', 'success');
            resetForm();
            switchTab('view');
            loadShifts();
        } else {
            showMessage(data.msg || '更新失敗', 'error');
        }
    } catch (error) {
        console.error('❌ 更新排班失敗:', error);
        showMessage('更新排班失敗', 'error');
    }
}

async function deleteShift(shiftId) {
    if (!confirm('確定要刪除這個排班嗎?')) return;
    
    try {
        const token = localStorage.getItem('sessionToken');
        const url = `${apiUrl}?action=deleteShift&token=${token}&shiftId=${shiftId}`;
        
        const response = await fetch(url);
        const data = await response.json();
        
        if (data.ok) {
            showMessage('排班已刪除', 'success');
            loadShifts();
        } else {
            showMessage(data.msg || '刪除失敗', 'error');
        }
    } catch (error) {
        console.error('❌ 刪除排班失敗:', error);
        showMessage('刪除失敗', 'error');
    }
}

function filterShifts() {
    const filters = {};
    
    const employeeEl = document.getElementById('filter-employee');
    const startDateEl = document.getElementById('filter-start-date');
    const endDateEl = document.getElementById('filter-end-date');
    const shiftTypeEl = document.getElementById('filter-shift-type');
    const locationEl = document.getElementById('filter-location');
    
    if (employeeEl && employeeEl.value) filters.employeeId = employeeEl.value;
    if (startDateEl && startDateEl.value) filters.startDate = startDateEl.value;
    if (endDateEl && endDateEl.value) filters.endDate = endDateEl.value;
    if (shiftTypeEl && shiftTypeEl.value) filters.shiftType = shiftTypeEl.value;
    if (locationEl && locationEl.value) filters.location = locationEl.value;
    
    console.log('🔍 篩選條件:', filters);
    loadShifts(filters);
}

function clearFilters() {
    const employeeEl = document.getElementById('filter-employee');
    const shiftTypeEl = document.getElementById('filter-shift-type');
    const locationEl = document.getElementById('filter-location');
    
    if (employeeEl) employeeEl.value = '';
    if (shiftTypeEl) shiftTypeEl.value = '';
    if (locationEl) locationEl.value = '';
    
    // 重設為本週
    const startOfWeek = new Date();
    startOfWeek.setDate(startOfWeek.getDate() - startOfWeek.getDay());
    const endOfWeek = new Date(startOfWeek);
    endOfWeek.setDate(endOfWeek.getDate() + 6);
    
    document.getElementById('filter-start-date').value = startOfWeek.toISOString().split('T')[0];
    document.getElementById('filter-end-date').value = endOfWeek.toISOString().split('T')[0];
    
    loadShifts();
}

function exportShifts() {
    if (currentShifts.length === 0) {
        showMessage('目前沒有可匯出的資料', 'error');
        return;
    }
    
    const csv = convertToCSV(currentShifts);
    const filename = `排班表_${new Date().toISOString().split('T')[0]}.csv`;
    downloadCSV(csv, filename);
    showMessage('匯出成功', 'success');
}

function convertToCSV(data) {
    const headers = ['排班ID', '員工ID', '員工姓名', '日期', '班別', '上班時間', '下班時間', '地點', '備註'];
    const rows = data.map(shift => [
        shift.shiftId,
        shift.employeeId,
        shift.employeeName,
        shift.date,
        shift.shiftType,
        shift.startTime,
        shift.endTime,
        shift.location,
        shift.note || ''
    ]);
    
    const csvContent = [headers, ...rows]
        .map(row => row.map(cell => `"${cell}"`).join(','))
        .join('\n');
    
    return '\ufeff' + csvContent;
}

function downloadCSV(csv, filename) {
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    link.click();
}

function resetForm() {
    const form = document.getElementById('add-shift-form');
    if (form) form.reset();
    
    document.querySelector('#add-tab h2').textContent = '新增排班';
    
    const submitBtn = document.querySelector('#add-shift-form button[type="submit"]');
    if (submitBtn) {
        submitBtn.textContent = '新增排班';
        submitBtn.onclick = null;
    }
    
    const today = new Date().toISOString().split('T')[0];
    const shiftDateEl = document.getElementById('shift-date');
    if (shiftDateEl) shiftDateEl.value = today;
}

// ========== 批量上傳 ==========

function setupBatchUpload() {
    const uploadArea = document.getElementById('upload-area');
    const fileInput = document.getElementById('batch-file-input');
    
    if (!uploadArea || !fileInput) return;
    
    uploadArea.addEventListener('dragover', function(e) {
        e.preventDefault();
        this.classList.add('drag-over');
    });
    
    uploadArea.addEventListener('dragleave', function() {
        this.classList.remove('drag-over');
    });
    
    uploadArea.addEventListener('drop', function(e) {
        e.preventDefault();
        this.classList.remove('drag-over');
        
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleBatchFile(files[0]);
        }
    });
    
    fileInput.addEventListener('change', function() {
        if (this.files.length > 0) {
            handleBatchFile(this.files[0]);
        }
    });
}

function handleBatchFile(file) {
    const reader = new FileReader();
    
    reader.onload = function(e) {
        const content = e.target.result;
        parseBatchData(content, file.name);
    };
    
    if (file.name.endsWith('.csv')) {
        reader.readAsText(file, 'UTF-8');
    } else {
        showMessage('目前只支援 CSV 格式', 'error');
    }
}

function parseBatchData(content, filename) {
    // 移除 BOM (如果有的話)
    content = content.replace(/^\ufeff/, '');
    
    const lines = content.split('\n');
    const data = [];
    
    // 從第二行開始(跳過標題)
    for (let i = 1; i < lines.length; i++) {
        const line = lines[i].trim();
        if (!line) continue;
        
        // ⭐ 正確處理 CSV 引號
        const values = parseCSVLine(line);
        
        // 檢查是否有足夠的欄位(至少 7 個)
        if (values.length >= 7) {
            // 跳過排班ID欄位(第一個),從員工ID開始
            const shift = {
                employeeId: values[1],      // 第 2 欄: 員工ID
                employeeName: values[2],    // 第 3 欄: 員工姓名
                date: values[3],            // 第 4 欄: 日期
                shiftType: values[4],       // 第 5 欄: 班別
                startTime: values[5],       // 第 6 欄: 上班時間
                endTime: values[6],         // 第 7 欄: 下班時間
                location: values[7] || '',  // 第 8 欄: 地點
                note: values[8] || ''       // 第 9 欄: 備註
            };
            
            // 驗證必填欄位
            if (shift.employeeId && shift.date && shift.shiftType) {
                data.push(shift);
            } else {
                console.warn('第 ' + (i+1) + ' 行資料不完整,已略過');
            }
        }
    }
    
    if (data.length === 0) {
        showMessage('檔案中沒有有效資料', 'error');
        return;
    }
    
    batchData = data;
    displayBatchPreview(data);
}

/**
 * ⭐ 正確解析 CSV 行(處理引號)
 */
function parseCSVLine(line) {
    const values = [];
    let currentValue = '';
    let insideQuotes = false;
    
    for (let i = 0; i < line.length; i++) {
        const char = line[i];
        const nextChar = line[i + 1];
        
        if (char === '"') {
            if (insideQuotes && nextChar === '"') {
                // 兩個連續的引號 = 一個引號字元
                currentValue += '"';
                i++; // 跳過下一個引號
            } else {
                // 切換引號狀態
                insideQuotes = !insideQuotes;
            }
        } else if (char === ',' && !insideQuotes) {
            // 在引號外的逗號 = 欄位分隔符
            values.push(currentValue.trim());
            currentValue = '';
        } else {
            currentValue += char;
        }
    }
    
    // 加入最後一個欄位
    values.push(currentValue.trim());
    
    return values;
}

function displayBatchPreview(data) {
    const previewDiv = document.getElementById('batch-preview');
    const tableDiv = document.getElementById('preview-table');
    
    if (!previewDiv || !tableDiv) return;
    
    let html = '<table style="width: 100%; border-collapse: collapse;">';
    html += '<tr style="background: #f5f5f5;">';
    html += '<th>員工ID</th><th>員工姓名</th><th>日期</th><th>班別</th><th>上班時間</th><th>下班時間</th><th>地點</th>';
    html += '</tr>';
    
    data.slice(0, 10).forEach(row => {
        html += '<tr style="border-bottom: 1px solid #eee;">';
        html += `<td>${row.employeeId}</td>`;
        html += `<td>${row.employeeName}</td>`;
        html += `<td>${row.date}</td>`;
        html += `<td>${row.shiftType}</td>`;
        html += `<td>${row.startTime}</td>`;
        html += `<td>${row.endTime}</td>`;
        html += `<td>${row.location}</td>`;
        html += '</tr>';
    });
    
    if (data.length > 10) {
        html += `<tr><td colspan="7" style="text-align: center; padding: 10px; color: #666;">還有 ${data.length - 10} 筆資料...</td></tr>`;
    }
    
    html += '</table>';
    
    tableDiv.innerHTML = html;
    previewDiv.style.display = 'block';
    document.getElementById('upload-area').style.display = 'none';
}

async function confirmBatchUpload() {
    if (batchData.length === 0) return;
    
    try {
        const token = localStorage.getItem('sessionToken');
        
        console.log('📤 準備上傳批量資料:', batchData.length, '筆');
        
        // ⭐ 改用 GET 請求避免 CORS 問題
        // 將資料轉成 JSON 字串並編碼
        const shiftsJson = encodeURIComponent(JSON.stringify(batchData));
        
        const url = `${apiUrl}?action=batchAddShifts&token=${token}&shiftsArray=${shiftsJson}`;
        
        // 使用 JSONP 方式呼叫
        const callbackName = 'batchUploadCallback_' + Date.now();
        
        return new Promise((resolve, reject) => {
            // 建立回調函數
            window[callbackName] = function(data) {
                console.log('📥 批量上傳回應:', data);
                
                // 清理
                delete window[callbackName];
                document.body.removeChild(script);
                
                if (data.ok) {
                    showMessage(data.msg || data.message || '批量上傳成功', 'success');
                    cancelBatchUpload();
                    switchTab('view');
                    loadShifts();
                    resolve(data);
                } else {
                    showMessage(data.msg || data.message || '批量上傳失敗', 'error');
                    reject(new Error(data.msg));
                }
            };
            
            // 建立 script 標籤
            const script = document.createElement('script');
            script.src = url + `&callback=${callbackName}`;
            script.onerror = function() {
                console.error('❌ 批量上傳失敗: 無法載入腳本');
                delete window[callbackName];
                document.body.removeChild(script);
                showMessage('批量上傳失敗: 網路錯誤', 'error');
                reject(new Error('Network error'));
            };
            
            document.body.appendChild(script);
        });
        
    } catch (error) {
        console.error('❌ 批量上傳失敗:', error);
        showMessage('批量上傳失敗: ' + error.message, 'error');
    }
}

function cancelBatchUpload() {
    batchData = [];
    const previewDiv = document.getElementById('batch-preview');
    const uploadArea = document.getElementById('upload-area');
    const fileInput = document.getElementById('batch-file-input');
    
    if (previewDiv) previewDiv.style.display = 'none';
    if (uploadArea) uploadArea.style.display = 'block';
    if (fileInput) fileInput.value = '';
}

function downloadTemplate() {
    const template = '員工ID,員工姓名,日期,班別,上班時間,下班時間,地點,備註\n' +
                    'EMP001,張三,2025-10-25,早班,08:00,16:00,總公司,\n' +
                    'EMP002,李四,2025-10-25,中班,12:00,20:00,分公司,';
    
    downloadCSV(template, '排班範本.csv');
}

// ========== 月曆功能 ==========

async function loadStats() {
    try {
        currentYear = new Date().getFullYear();
        currentMonth = new Date().getMonth();
        
        updateMonthDisplay();
        await loadMonthlyStats();
        await loadMonthlyShifts();
        await loadShiftDistribution();
    } catch (error) {
        console.error('載入統計失敗:', error);
    }
}

function updateMonthDisplay() {
    const monthNames = ['1月', '2月', '3月', '4月', '5月', '6月', 
                        '7月', '8月', '9月', '10月', '11月', '12月'];
    const displayText = `${currentYear}年${monthNames[currentMonth]}`;
    const monthEl = document.getElementById('current-month');
    if (monthEl) monthEl.textContent = displayText;
}

function previousMonth() {
    currentMonth--;
    if (currentMonth < 0) {
        currentMonth = 11;
        currentYear--;
    }
    updateMonthDisplay();
    loadMonthlyStats();
    loadMonthlyShifts();
    loadShiftDistribution();
}

function nextMonth() {
    currentMonth++;
    if (currentMonth > 11) {
        currentMonth = 0;
        currentYear++;
    }
    updateMonthDisplay();
    loadMonthlyStats();
    loadMonthlyShifts();
    loadShiftDistribution();
}

function goToToday() {
    const today = new Date();
    currentYear = today.getFullYear();
    currentMonth = today.getMonth();
    updateMonthDisplay();
    loadMonthlyStats();
    loadMonthlyShifts();
    loadShiftDistribution();
}

async function loadMonthlyStats() {
    try {
        const token = localStorage.getItem('sessionToken');
        const startDate = new Date(currentYear, currentMonth, 1);
        const endDate = new Date(currentYear, currentMonth + 1, 0);
        
        const queryParams = new URLSearchParams({
            action: 'getShifts',
            token: token,
            startDate: formatDateYMD(startDate),
            endDate: formatDateYMD(endDate)
        });
        
        const response = await fetch(`${apiUrl}?${queryParams}`);
        const data = await response.json();
        
        console.log('📊 月度統計:', data);
        
        if (data.ok && data.data) {
            allMonthShifts = data.data;
            displayMonthlyStats(data.data);
        }
    } catch (error) {
        console.error('載入月度統計失敗:', error);
    }
}

function displayMonthlyStats(shifts) {
    const statsGrid = document.getElementById('stats-grid');
    if (!statsGrid) return;
    
    const stats = {
        total: shifts.length,
        morning: 0,
        afternoon: 0,
        night: 0,
        full: 0,
        dayoff: 0,  // 新增
        custom: 0
    };
    
    shifts.forEach(shift => {
        switch(shift.shiftType) {
            case '早班': stats.morning++; break;
            case '中班': stats.afternoon++; break;
            case '晚班': stats.night++; break;
            case '全日班': stats.full++; break;
            case '排休': stats.dayoff++; break;  // 新增
            case '自訂': stats.custom++; break;
        }
    });
    
    const html = `
        <div class="stat-card">
            <div class="stat-label">本月總排班</div>
            <div class="stat-value">${stats.total}</div>
        </div>
        <div class="stat-card">
            <div class="stat-label">早班</div>
            <div class="stat-value" style="color: #ff9800;">${stats.morning}</div>
        </div>
        <div class="stat-card">
            <div class="stat-label">中班</div>
            <div class="stat-value" style="color: #2196f3;">${stats.afternoon}</div>
        </div>
        <div class="stat-card">
            <div class="stat-label">晚班</div>
            <div class="stat-value" style="color: #9c27b0;">${stats.night}</div>
        </div>
        <div class="stat-card">
            <div class="stat-label">排休</div>
            <div class="stat-value" style="color: #757575;">${stats.dayoff}</div>
        </div>
        ${stats.custom > 0 ? `
        <div class="stat-card">
            <div class="stat-label">自訂班別</div>
            <div class="stat-value" style="color: #fbc02d;">${stats.custom}</div>
        </div>
        ` : ''}
    `;
    
    statsGrid.innerHTML = html;
}

async function loadMonthlyShifts() {
    const calendarGrid = document.getElementById('calendar-grid');
    if (!calendarGrid) return;
    
    calendarGrid.innerHTML = '<div class="loading">載入月曆中</div>';
    
    try {
        displayMonthCalendar(allMonthShifts);
    } catch (error) {
        console.error('載入月曆失敗:', error);
        calendarGrid.innerHTML = '<div class="loading">載入失敗</div>';
    }
}

function displayMonthCalendar(shifts) {
    const calendarGrid = document.getElementById('calendar-grid');
    if (!calendarGrid) return;
    
    const firstDay = new Date(currentYear, currentMonth, 1);
    const lastDay = new Date(currentYear, currentMonth + 1, 0);
    const daysInMonth = lastDay.getDate();
    const startingDayOfWeek = firstDay.getDay();
    const prevMonthLastDay = new Date(currentYear, currentMonth, 0).getDate();
    
    const today = new Date();
    const todayStr = formatDateYMD(today);
    
    let html = '';
    let dayCounter = 1;
    let nextMonthDay = 1;
    
    const totalCells = Math.ceil((daysInMonth + startingDayOfWeek) / 7) * 7;
    
    for (let i = 0; i < totalCells; i++) {
        let dateStr = '';
        let dayNumber = '';
        let otherMonthClass = '';
        let isToday = false;
        
        if (i < startingDayOfWeek) {
            dayNumber = prevMonthLastDay - startingDayOfWeek + i + 1;
            otherMonthClass = 'other-month';
            const prevMonth = currentMonth === 0 ? 11 : currentMonth - 1;
            const prevYear = currentMonth === 0 ? currentYear - 1 : currentYear;
            dateStr = `${prevYear}-${String(prevMonth + 1).padStart(2, '0')}-${String(dayNumber).padStart(2, '0')}`;
        } else if (dayCounter <= daysInMonth) {
            dayNumber = dayCounter;
            dateStr = `${currentYear}-${String(currentMonth + 1).padStart(2, '0')}-${String(dayNumber).padStart(2, '0')}`;
            isToday = dateStr === todayStr;
            dayCounter++;
        } else {
            dayNumber = nextMonthDay;
            otherMonthClass = 'other-month';
            const nextMonth = currentMonth === 11 ? 0 : currentMonth + 1;
            const nextYear = currentMonth === 11 ? currentYear + 1 : currentYear;
            dateStr = `${nextYear}-${String(nextMonth + 1).padStart(2, '0')}-${String(dayNumber).padStart(2, '0')}`;
            nextMonthDay++;
        }
        
        const dayShifts = shifts.filter(shift => {
            const shiftDate = new Date(shift.date);
            return formatDateYMD(shiftDate) === dateStr;
        });
        
        const hasShifts = dayShifts.length > 0;
        const todayClass = isToday ? 'today' : '';
        const hasShiftsClass = hasShifts ? 'has-shifts' : '';
        
        html += `
            <div class="calendar-day ${otherMonthClass} ${todayClass} ${hasShiftsClass}">
                <div class="day-number">${dayNumber}</div>
                <div class="day-shifts">
                    ${dayShifts.slice(0, 10).map(shift => {
                        const startTime = formatTimeOnly(shift.startTime);
                        const endTime = formatTimeOnly(shift.endTime);
                        return `
                        <div class="shift-item-mini ${getShiftClass(shift.shiftType)}" 
                             onclick="showShiftDetail('${shift.shiftId}')"
                             title="${shift.employeeName} - ${shift.shiftType} (${startTime}-${endTime})">
                            <div class="shift-item-name">${shift.employeeName}</div>
                            <div class="shift-item-time">${startTime}-${endTime}</div>
                        </div>
                    `}).join('')}
                </div>
                ${dayShifts.length > 3 ? `<div class="shift-count">+${dayShifts.length - 3}</div>` : ''}
            </div>
        `;
    }
    
    calendarGrid.innerHTML = html;
}

function getShiftClass(shiftType) {
    const classMap = {
        '早班': 'shift-morning',
        '中班': 'shift-afternoon',
        '晚班': 'shift-night',
        '全日班': 'shift-full',
        '排休': 'shift-dayoff',
        '自訂': 'shift-custom'
    };
    return classMap[shiftType] || 'shift-morning';
}

function showShiftDetail(shiftId) {
    const shift = allMonthShifts.find(s => s.shiftId === shiftId);
    if (shift) {
        const startTime = formatTimeOnly(shift.startTime);
        const endTime = formatTimeOnly(shift.endTime);
        
        const detail = `排班詳情:\n\n` +
              `員工: ${shift.employeeName}\n` +
              `日期: ${shift.date}\n` +
              `班別: ${shift.shiftType}\n` +
              `時間: ${startTime} - ${endTime}\n` +
              `地點: ${shift.location}\n` +
              `備註: ${shift.note || '無'}`;
        
        alert(detail);
    }
}

async function loadShiftDistribution() {
    const distributionContainer = document.getElementById('shift-distribution');
    if (!distributionContainer) return;
    
    displayShiftDistribution(allMonthShifts);
}

function displayShiftDistribution(shifts) {
    const distributionContainer = document.getElementById('shift-distribution');
    if (!distributionContainer || shifts.length === 0) {
        if (distributionContainer) distributionContainer.innerHTML = '';
        return;
    }
    
    const employeeStats = {};
    const shiftTypeStats = { 
        '早班': 0, 
        '中班': 0, 
        '晚班': 0, 
        '全日班': 0,
        '排休': 0,
        '自訂': 0
    };
    
    shifts.forEach(shift => {
        if (!employeeStats[shift.employeeName]) {
            employeeStats[shift.employeeName] = 0;
        }
        employeeStats[shift.employeeName]++;
        
        if (shiftTypeStats[shift.shiftType] !== undefined) {
            shiftTypeStats[shift.shiftType]++;
        }
    });
    
    const maxCount = Math.max(...Object.values(employeeStats), 1);
    
    let html = '<div class="distribution-section">';
    html += '<h3 class="distribution-title">📊 本月員工排班分布</h3>';
    html += '<div class="distribution-bars">';
    
    const sortedEmployees = Object.entries(employeeStats)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 15);
    
    sortedEmployees.forEach(([name, count]) => {
        const percentage = (count / maxCount * 100).toFixed(0);
        html += `
            <div class="distribution-bar-item">
                <div class="distribution-bar-label">${name}</div>
                <div class="distribution-bar-container">
                    <div class="distribution-bar" style="width: ${percentage}%">
                        <div class="distribution-bar-value">${count} 班</div>
                    </div>
                </div>
            </div>
        `;
    });
    
    html += '</div></div>';
    
    html += '<div class="distribution-section">';
    html += '<h3 class="distribution-title">🎨 本月班別分布</h3>';
    html += '<div class="shift-type-distribution">';
    
    const totalShifts = Object.values(shiftTypeStats).reduce((a, b) => a + b, 0);
    const shiftTypeColors = {
        '早班': '#ff9800',
        '中班': '#2196f3',
        '晚班': '#9c27b0',
        '全日班': '#4caf50',
        '排休': '#9e9e9e',
        '自訂': '#fbc02d'
    };
    
    Object.entries(shiftTypeStats).forEach(([type, count]) => {
        const percentage = totalShifts > 0 ? (count / totalShifts * 100).toFixed(1) : 0;
        const color = shiftTypeColors[type];
        
        html += `
            <div class="shift-type-stat">
                <div class="shift-type-stat-header">
                    <span class="shift-type-label ${getShiftClass(type)}">${type}</span>
                    <span class="shift-type-count">${count}</span>
                </div>
                <div class="shift-type-bar-container">
                    <div class="shift-type-bar" style="width: ${percentage}%; background: ${color};"></div>
                </div>
                <div class="shift-type-percentage">${percentage}%</div>
            </div>
        `;
    });
    
    html += '</div></div>';
    
    distributionContainer.innerHTML = html;
}

function formatDateYMD(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

/**
 * 格式化時間為 HH:MM 格式
 * 支援多種輸入格式:
 * - "08:00" → "08:00"
 * - "1899-12-30T01:00:00.000Z" → "09:00" (UTC+8)
 * - Date 物件 → "HH:MM"
 */
function formatTimeOnly(timeValue) {
    if (!timeValue) return '--:--';
    
    // 如果已經是 HH:MM 格式,直接返回
    if (typeof timeValue === 'string' && /^\d{2}:\d{2}$/.test(timeValue)) {
        return timeValue;
    }
    
    // 如果是 ISO 格式字串
    if (typeof timeValue === 'string' && timeValue.includes('T')) {
        try {
            const date = new Date(timeValue);
            // 轉換為台灣時間 (UTC+8)
            const hours = String(date.getUTCHours() + 8).padStart(2, '0');
            const minutes = String(date.getUTCMinutes()).padStart(2, '0');
            return `${hours}:${minutes}`;
        } catch (e) {
            console.error('時間格式錯誤:', timeValue);
            return '--:--';
        }
    }
    
    // 如果是 Date 物件
    if (timeValue instanceof Date) {
        const hours = String(timeValue.getHours()).padStart(2, '0');
        const minutes = String(timeValue.getMinutes()).padStart(2, '0');
        return `${hours}:${minutes}`;
    }
    
    // 其他情況直接返回
    return String(timeValue);
}

// ========== 工具函數 ==========

function showMessage(message, type = 'info') {
    const messageDiv = document.createElement('div');
    messageDiv.className = `message message-${type}`;
    messageDiv.textContent = message;
    messageDiv.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 16px 24px;
        background: ${type === 'success' ? '#4caf50' : type === 'error' ? '#f44336' : '#2196f3'};
        color: white;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        z-index: 10000;
        animation: slideIn 0.3s ease;
    `;
    
    document.body.appendChild(messageDiv);
    
    setTimeout(() => {
        messageDiv.style.animation = 'slideOut 0.3s ease';
        setTimeout(() => {
            if (messageDiv.parentNode) {
                document.body.removeChild(messageDiv);
            }
        }, 300);
    }, 3000);
}

function goBack() {
    window.history.back();
}

console.log('✅ 排班管理系統(含月曆功能)已載入');