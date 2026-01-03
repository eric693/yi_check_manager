// leave.js - 請假系統前端邏輯（完全修正版）

// ⭐ 添加全域標記
let leaveTabInitialized = false;
let leaveEventsBound = false;

/**
 * 初始化請假頁籤（防重複版）
 */
async function initLeaveTab() {
    console.log('📋 initLeaveTab 被調用');
    
    // ⭐ 檢查是否已初始化
    if (leaveTabInitialized) {
        console.log('⏭️ 請假系統已初始化，跳過');
        return;
    }
    
    console.log('🔄 開始初始化請假系統...');
    
    // 標記為正在初始化（防止並發調用）
    leaveTabInitialized = true;
    
    try {
        // 載入假期餘額
        await loadLeaveBalance();
        
        // 載入請假紀錄
        await loadLeaveRecords();
        
        // 綁定事件監聽器
        bindLeaveEventListeners();
        
        console.log('✅ 請假系統初始化完成');
    } catch (error) {
        console.error('❌ 請假系統初始化失敗:', error);
        // 如果初始化失敗，重置標記以便重試
        leaveTabInitialized = false;
    }
}

/**
 * 綁定請假相關事件（防重複版）
 */
function bindLeaveEventListeners() {
    // ⭐ 檢查是否已綁定
    if (leaveEventsBound) {
        console.log('⏭️ 事件監聽器已綁定，跳過');
        return;
    }
    
    console.log('🔗 綁定事件監聽器...');
    
    // 請假類型改變
    const leaveTypeSelect = document.getElementById('leave-type');
    if (leaveTypeSelect) {
        leaveTypeSelect.addEventListener('change', handleLeaveTypeChange);
    }
    
    // 時間輸入改變
    const startInput = document.getElementById('leave-start-datetime');
    const endInput = document.getElementById('leave-end-datetime');
    
    if (startInput) {
        startInput.addEventListener('change', updateWorkHoursPreview);
        console.log('✅ 已綁定開始時間事件');
    }
    
    if (endInput) {
        endInput.addEventListener('change', updateWorkHoursPreview);
        console.log('✅ 已綁定結束時間事件');
    }
    
    // ⭐ 標記為已綁定
    leaveEventsBound = true;
}

/**
 * 重置初始化狀態（用於手動刷新）
 */
function resetLeaveTab() {
    console.log('🔄 重置請假系統狀態');
    leaveTabInitialized = false;
    leaveEventsBound = false;
}

/**
 * 手動刷新請假數據（不重新綁定事件）
 */
async function refreshLeaveData() {
    console.log('🔄 手動刷新請假數據...');
    await loadLeaveBalance();
    await loadLeaveRecords();
}

/**
 * ✅ 完全修正版：計算工作時數（09:00-18:00，扣除 12:00-13:00）
 */
function calculateWorkHours(startTime, endTime) {
    if (!startTime || !endTime) {
        return 0;
    }
    
    const start = new Date(startTime);
    const end = new Date(endTime);
    
    // 檢查日期是否有效
    if (isNaN(start.getTime()) || isNaN(end.getTime())) {
        console.error('❌ 無效的日期格式');
        return 0;
    }
    
    // 檢查結束時間是否早於開始時間
    if (end <= start) {
        console.error('❌ 結束時間必須晚於開始時間');
        return 0;
    }
    
    console.log('📊 開始計算工時:', {
        start: start.toISOString(),
        end: end.toISOString()
    });
    
    // ⭐ 工作時間設定
    const WORK_START_HOUR = 9;      // 上班 09:00
    const WORK_END_HOUR = 18;       // 下班 18:00
    const LUNCH_START = 12;         // 午休開始 12:00
    const LUNCH_END = 13;           // 午休結束 13:00
    const DAILY_WORK_HOURS = 8;     // 每日工作時數（已扣午休）
    
    // 判斷是否同一天
    const startDate = new Date(start.getFullYear(), start.getMonth(), start.getDate());
    const endDate = new Date(end.getFullYear(), end.getMonth(), end.getDate());
    const isSameDay = startDate.getTime() === endDate.getTime();
    
    // 1️⃣ 同一天請假
    if (isSameDay) {
        console.log('   ℹ️ 同日請假');
        
        // ⭐ 限制在工作時間內
        const startHour = Math.max(
            start.getHours() + start.getMinutes() / 60,
            WORK_START_HOUR
        );
        const endHour = Math.min(
            end.getHours() + end.getMinutes() / 60,
            WORK_END_HOUR
        );
        
        if (startHour >= endHour) {
            console.log('   ⚠️ 請假時間不在工作時段內');
            return 0;
        }
        
        let workHours = endHour - startHour;
        
        // 扣除午休時間（如果跨越）
        if (startHour < LUNCH_END && endHour > LUNCH_START) {
            const lunchOverlapStart = Math.max(startHour, LUNCH_START);
            const lunchOverlapEnd = Math.min(endHour, LUNCH_END);
            const lunchOverlap = Math.max(0, lunchOverlapEnd - lunchOverlapStart);
            workHours -= lunchOverlap;
            
            console.log(`   🍱 扣除午休時間: ${lunchOverlap.toFixed(2)} 小時`);
        }
        
        workHours = Math.max(0, workHours);
        const finalHours = Math.round(workHours * 100) / 100;
        
        console.log(`   ✅ 同日請假工時: ${finalHours} 小時`);
        
        return finalHours;
    }
    
    // 2️⃣ 跨日請假 - 分段精確計算
    else {
        console.log('   ℹ️ 跨日請假');
        
        let totalWorkHours = 0;
        
        // 🔹 第一天：從請假開始到當天下班
        // ⭐ 修正：限制開始時間不早於上班時間
        const firstDayStartHour = Math.max(
            start.getHours() + start.getMinutes() / 60,
            WORK_START_HOUR
        );
        
        // ⭐ 修正：確保不晚於下班時間
        const firstDayEndHour = WORK_END_HOUR;
        
        let firstDayHours = Math.max(0, firstDayEndHour - firstDayStartHour);
        
        // 扣除第一天的午休
        if (firstDayStartHour < LUNCH_END && firstDayEndHour > LUNCH_START) {
            const lunchStart = Math.max(firstDayStartHour, LUNCH_START);
            const lunchEnd = Math.min(firstDayEndHour, LUNCH_END);
            const lunchOverlap = Math.max(0, lunchEnd - lunchStart);
            firstDayHours -= lunchOverlap;
        }
        
        firstDayHours = Math.max(0, firstDayHours);
        totalWorkHours += firstDayHours;
        
        console.log(`   📅 第一天 (${start.toLocaleDateString()}): ${firstDayHours.toFixed(2)} 小時`);
        
        // 🔹 中間完整工作日
        const daysDiff = Math.floor((endDate - startDate) / (1000 * 60 * 60 * 24));
        
        if (daysDiff > 1) {
            const middleDays = daysDiff - 1;
            const middleHours = middleDays * DAILY_WORK_HOURS;
            totalWorkHours += middleHours;
            
            console.log(`   📅 中間完整工作日 (${middleDays} 天): ${middleHours} 小時`);
        }
        
        // 🔹 最後一天：從上班到請假結束
        // ⭐ 修正：限制結束時間不早於上班時間、不晚於下班時間
        const lastDayEndHour = Math.min(
            Math.max(
                end.getHours() + end.getMinutes() / 60,
                WORK_START_HOUR  // ⭐ 不早於上班時間
            ),
            WORK_END_HOUR
        );
        
        const lastDayStartHour = WORK_START_HOUR;
        
        let lastDayHours = Math.max(0, lastDayEndHour - lastDayStartHour);
        
        // 扣除最後一天的午休
        if (lastDayStartHour < LUNCH_END && lastDayEndHour > LUNCH_START) {
            const lunchStart = Math.max(lastDayStartHour, LUNCH_START);
            const lunchEnd = Math.min(lastDayEndHour, LUNCH_END);
            const lunchOverlap = Math.max(0, lunchEnd - lunchStart);
            lastDayHours -= lunchOverlap;
        }
        
        lastDayHours = Math.max(0, lastDayHours);
        totalWorkHours += lastDayHours;
        
        console.log(`   📅 最後一天 (${end.toLocaleDateString()}): ${lastDayHours.toFixed(2)} 小時`);
        
        // 四捨五入到小數點後 2 位
        const finalHours = Math.round(totalWorkHours * 100) / 100;
        
        console.log(`   ✅ 跨日請假總工時: ${finalHours} 小時`);
        
        return finalHours;
    }
}
/**
 * 更新工時預覽（即時顯示）
 */
function updateWorkHoursPreview() {
    console.log('🔄 updateWorkHoursPreview 被觸發');
    
    const startTime = document.getElementById('leave-start-datetime').value;
    const endTime = document.getElementById('leave-end-datetime').value;
    const previewEl = document.getElementById('work-hours-preview');
    const hoursEl = document.getElementById('calculated-hours');
    const warningEl = document.getElementById('work-hours-warning');
    
    console.log('📥 輸入值:', {
        startTime: startTime,
        endTime: endTime
    });
    
    // 如果沒有輸入，隱藏預覽
    if (!startTime || !endTime) {
        console.log('⚠️ 開始或結束時間為空');
        if (previewEl) previewEl.classList.add('hidden');
        return;
    }
    
    // 計算工時
    const workHours = calculateWorkHours(startTime, endTime);
    
    console.log('💡 計算結果:', workHours, '小時');
    
    // 顯示預覽區塊
    if (previewEl) previewEl.classList.remove('hidden');
    if (hoursEl) hoursEl.textContent = workHours;
    
    // 清除之前的警告
    if (warningEl) {
        warningEl.classList.add('hidden');
        warningEl.textContent = '';
    }
    
    // 檢查各種錯誤情況
    let hasError = false;
    let errorMsg = '';
    
    if (workHours <= 0) {
        hasError = true;
        errorMsg = '❌ 結束時間必須晚於開始時間';
    } else if (!Number.isInteger(workHours)) {
        hasError = true;
        errorMsg = `❌ 請假時數必須是整數小時（目前為 ${workHours} 小時）\n請調整時間使其為整數小時`;
    } else {
        const start = new Date(startTime);
        const end = new Date(endTime);
        
        // 檢查單日上限
        if (start.toDateString() === end.toDateString() && workHours > 8) {
            hasError = true;
            errorMsg = '❌ 單日請假不能超過 8 小時（已扣除午休）';
        }
    }
    
    // 顯示警告訊息
    if (hasError) {
        if (warningEl) {
            warningEl.classList.remove('hidden');
            warningEl.textContent = errorMsg;
        }
        if (hoursEl) {
            hoursEl.classList.add('text-red-600', 'dark:text-red-400');
            hoursEl.classList.remove('text-blue-800', 'dark:text-blue-300');
        }
    } else {
        // 顯示成功狀態
        if (hoursEl) {
            hoursEl.classList.remove('text-red-600', 'dark:text-red-400');
            hoursEl.classList.add('text-blue-800', 'dark:text-blue-300');
        }
        
        // 顯示成功提示
        if (warningEl) {
            warningEl.classList.remove('hidden');
            warningEl.classList.remove('text-red-600', 'dark:text-red-400');
            warningEl.classList.add('text-green-600', 'dark:text-green-400');
            warningEl.textContent = '✅ 時數計算正確，可以提交申請';
        }
    }
}

/**
 * 快速選擇時段
 */
function quickSelectTimeRange(type) {
    console.log('🎯 快速選擇:', type);
    
    const now = new Date();
    const today = now.toISOString().split('T')[0]; // YYYY-MM-DD
    
    let startTime, endTime;
    
    switch(type) {
        case '1h':
            startTime = `${today}T09:00`;
            endTime = `${today}T10:00`;
            break;
            
        case '2h':
            startTime = `${today}T09:00`;
            endTime = `${today}T11:00`;
            break;
            
        case '4h':
            startTime = `${today}T13:00`;
            endTime = `${today}T17:00`;
            break;
            
        case '8h':
            startTime = `${today}T09:00`;
            endTime = `${today}T18:00`;
            break;
            
        default:
            console.error('❌ 未知的時段類型:', type);
            return;
    }
    
    console.log('📅 設定時間:', { startTime, endTime });
    
    // 設定時間
    const startInput = document.getElementById('leave-start-datetime');
    const endInput = document.getElementById('leave-end-datetime');
    
    if (startInput) startInput.value = startTime;
    if (endInput) endInput.value = endTime;
    
    // 更新工時預覽
    updateWorkHoursPreview();
}

async function submitLeaveApplication() {
    console.log('📤 開始提交請假申請');
    
    if (!validateLeaveForm()) {
        console.error('❌ 表單驗證失敗');
        return;
    }
    
    const leaveType = document.getElementById('leave-type').value;
    const startTime = document.getElementById('leave-start-datetime').value;
    const endTime = document.getElementById('leave-end-datetime').value;
    const reason = document.getElementById('leave-reason').value;
    const workHours = calculateWorkHours(startTime, endTime);
    
    // ⭐⭐⭐ 新增：計算請假天數
    const days = workHours / 8;
    
    console.log('📋 提交資料:', {
        leaveType,
        startTime,
        endTime,
        workHours,
        days,
        reason
    });
    
    // ⭐⭐⭐ 新增：檢查假期餘額
    try {
        const balanceRes = await callApifetch('getLeaveBalance');
        
        if (balanceRes.ok && balanceRes.balance) {
            const availableDays = balanceRes.balance[leaveType] || 0;
            
            console.log(`💰 假期餘額檢查:`, {
                假別: leaveType,
                可用天數: availableDays,
                申請天數: days
            });
            
            if (days > availableDays) {
                showNotification(
                    `餘額不足！${t(leaveType)} 剩餘 ${availableDays * 8} 小時（${availableDays} 天），但您申請了 ${workHours} 小時（${days} 天）`,
                    'error'
                );
                return;
            }
        }
    } catch (error) {
        console.error('❌ 檢查餘額失敗:', error);
        // 繼續提交（不阻擋）
    }
    
    const button = document.getElementById('submit-leave-btn');
    if (button) {
        button.disabled = true;
        button.textContent = '處理中...';
    }
    
    try {
        const response = await callApifetch(
            `submitLeave&leaveType=${encodeURIComponent(leaveType)}` +
            `&startDateTime=${encodeURIComponent(startTime)}` +
            `&endDateTime=${encodeURIComponent(endTime)}` +
            `&reason=${encodeURIComponent(reason)}`
        );
        
        console.log('📥 後端回應:', response);
        
        if (response.ok) {
            showNotification(`請假申請已提交！時數：${workHours} 小時`, 'success');
            
            // 清空表單
            document.getElementById('leave-type').value = '';
            document.getElementById('leave-start-datetime').value = '';
            document.getElementById('leave-end-datetime').value = '';
            document.getElementById('leave-reason').value = '';
            
            const previewEl = document.getElementById('work-hours-preview');
            if (previewEl) previewEl.classList.add('hidden');
            
            console.log('🔄 重新載入假期餘額...');
            await loadLeaveBalance();
            
            console.log('🔄 重新載入請假記錄...');
            await loadLeaveRecords();
            
            console.log('✅ 資料更新完成');
        } else {
            showNotification(response.msg || '提交失敗', 'error');
        }
    } catch (error) {
        console.error('❌ 提交請假申請失敗:', error);
        showNotification('網路錯誤，請稍後再試', 'error');
    } finally {
        if (button) {
            button.disabled = false;
            button.textContent = '提交請假申請';
        }
    }
}
/**
 * 驗證請假表單
 */
function validateLeaveForm() {
    const leaveType = document.getElementById('leave-type').value;
    const startTime = document.getElementById('leave-start-datetime').value;
    const endTime = document.getElementById('leave-end-datetime').value;
    const reason = document.getElementById('leave-reason').value;
    
    if (!leaveType) {
        showNotification('請選擇假別', 'error');
        return false;
    }
    
    if (!startTime) {
        showNotification('請選擇開始時間', 'error');
        return false;
    }
    
    if (!endTime) {
        showNotification('請選擇結束時間', 'error');
        return false;
    }
    
    // ⭐⭐⭐ 新增：檢查是否為整點時間
    const start = new Date(startTime);
    const end = new Date(endTime);
    
    if (start.getMinutes() !== 0 || start.getSeconds() !== 0) {
        showNotification('開始時間必須是整點（例如：09:00, 10:00）', 'error');
        return false;
    }
    
    if (end.getMinutes() !== 0 || end.getSeconds() !== 0) {
        showNotification('結束時間必須是整點（例如：09:00, 10:00）', 'error');
        return false;
    }
    
    if (!reason.trim() || reason.trim().length < 2) {
        showNotification('請填寫請假原因（至少2個字）', 'error');
        return false;
    }
    
    const workHours = calculateWorkHours(startTime, endTime);
    
    if (workHours <= 0) {
        showNotification('請假時數必須大於 0', 'error');
        return false;
    }
    
    if (!Number.isInteger(workHours)) {
        showNotification(`請假時數必須是整數小時，目前為 ${workHours} 小時`, 'error');
        return false;
    }
    
    // ⭐ 修正：使用已創建的 start 和 end 變數
    const startDate = new Date(start.getFullYear(), start.getMonth(), start.getDate());
    const endDate = new Date(end.getFullYear(), end.getMonth(), end.getDate());
    const isSameDay = startDate.getTime() === endDate.getTime();
    
    if (isSameDay && workHours > 8) {
        showNotification('單日請假不能超過 8 小時（已扣除午休）', 'error');
        return false;
    }
    
    return true;
}

/**
 * 載入假期餘額
 */
async function loadLeaveBalance() {
    const loadingEl = document.getElementById('leave-balance-loading');
    
    console.log('🔄 開始載入假期餘額...');
    
    if (loadingEl) loadingEl.style.display = 'block';
    
    try {
        const res = await callApifetch('getLeaveBalance');
        
        console.log('📥 後端返回的假期餘額:', res);
        
        if (res.ok && res.balance) {
            console.log('✅ 假期餘額數據:', res.balance);
            renderLeaveBalance(res.balance);
        } else {
            console.error('❌ 載入假期餘額失敗:', res);
        }
    } catch (err) {
        console.error('❌ 載入假期餘額錯誤:', err);
    } finally {
        if (loadingEl) loadingEl.style.display = 'none';
        console.log('✅ 假期餘額載入完成');
    }
}

/**
 * 渲染假期餘額（顯示小時數）
 */
function renderLeaveBalance(balance) {
    const listEl = document.getElementById('leave-balance-list');
    if (!listEl) return;
    
    console.log('📊 開始渲染假期餘額:', balance);
    
    listEl.innerHTML = '';
    
    const leaveOrder = [
        'ANNUAL_LEAVE', 'COMP_TIME_OFF', 'PERSONAL_LEAVE', 'SICK_LEAVE',
        'HOSPITALIZATION_LEAVE', 'BEREAVEMENT_LEAVE', 'MARRIAGE_LEAVE',
        'PATERNITY_LEAVE', 'MATERNITY_LEAVE', 'OFFICIAL_LEAVE',
        'WORK_INJURY_LEAVE', 'ABSENCE_WITHOUT_LEAVE', 'NATURAL_DISASTER_LEAVE',
        'FAMILY_CARE_LEAVE', 'MENSTRUAL_LEAVE'
    ];
    
    leaveOrder.forEach(leaveType => {
        if (balance[leaveType] !== undefined) {
            const item = document.createElement('div');
            item.className = 'flex justify-between items-center p-3 bg-gray-50 dark:bg-gray-700 rounded-lg';
            
            const typeSpan = document.createElement('span');
            typeSpan.className = 'font-medium text-gray-800 dark:text-white';
            typeSpan.textContent = t(leaveType);
            
            // ✅ 將天數轉換為小時數（1天 = 8小時）
            const days = balance[leaveType];
            const hours = days * 8;
            
            // 輸出主要假別的餘額（避免日誌過多）
            if (leaveType === 'ANNUAL_LEAVE' || leaveType === 'PERSONAL_LEAVE') {
                console.log(`  ${leaveType}: ${days} 天 = ${hours} 小時`);
            }
            
            const hoursSpan = document.createElement('span');
            hoursSpan.className = leaveType === 'ABSENCE_WITHOUT_LEAVE' 
                ? 'text-red-600 dark:text-red-400 font-bold'
                : 'text-indigo-600 dark:text-indigo-400 font-bold';
            hoursSpan.textContent = `${hours} 小時`;
            
            item.appendChild(typeSpan);
            item.appendChild(hoursSpan);
            listEl.appendChild(item);
        }
    });
    
    console.log('✅ 假期餘額渲染完成');
}

/**
 * 載入請假紀錄
 */
async function loadLeaveRecords() {
    const userId = localStorage.getItem('sessionUserId');
    const loadingEl = document.getElementById('leave-records-loading');
    const emptyEl = document.getElementById('leave-records-empty');
    const listEl = document.getElementById('leave-records-list');
    
    console.log('🔄 開始載入請假記錄...');
    
    try {
        if (loadingEl) loadingEl.style.display = 'block';
        if (emptyEl) emptyEl.style.display = 'none';
        if (listEl) listEl.innerHTML = '';
        
        const res = await callApifetch(`getEmployeeLeaveRecords&employeeId=${userId}`);
        
        console.log('📥 後端返回的請假記錄:', res);
        
        if (loadingEl) loadingEl.style.display = 'none';
        
        if (res.ok && res.records && res.records.length > 0) {
            console.log(`✅ 獲取到 ${res.records.length} 筆記錄`);
            
            // ✅ 去重處理：使用 rowNumber 作為唯一標識
            const uniqueRecords = [];
            const seenIds = new Set();
            
            res.records.forEach(record => {
                // 使用 rowNumber 或組合鍵作為唯一標識
                const uniqueKey = record.rowNumber || 
                    `${record.leaveType}_${record.startDateTime}_${record.endDateTime}_${record.reason}`;
                
                if (!seenIds.has(uniqueKey)) {
                    seenIds.add(uniqueKey);
                    uniqueRecords.push(record);
                } else {
                    console.warn('⚠️ 發現重複記錄:', record);
                }
            });
            
            console.log(`✅ 去重後剩餘 ${uniqueRecords.length} 筆記錄`);
            renderLeaveRecords(uniqueRecords);
        } else {
            console.log('ℹ️ 沒有請假記錄');
            if (emptyEl) emptyEl.style.display = 'block';
        }
    } catch (error) {
        console.error('❌ 載入請假紀錄失敗:', error);
        if (loadingEl) loadingEl.style.display = 'none';
        if (emptyEl) emptyEl.style.display = 'block';
    }
}

/**
 * 格式化日期時間顯示（ISO 8601 → 可讀格式）
 */
function formatDateTime(isoString) {
    if (!isoString) return '未設定';
    
    try {
        const date = new Date(isoString);
        
        // 檢查是否為有效日期
        if (isNaN(date.getTime())) return isoString;
        
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        const hours = String(date.getHours()).padStart(2, '0');
        const minutes = String(date.getMinutes()).padStart(2, '0');
        
        return `${year}-${month}-${day} ${hours}:${minutes}`;
    } catch (error) {
        console.error('❌ 日期格式化失敗:', error);
        return isoString;
    }
}

/**
 * 渲染請假記錄
 */
function renderLeaveRecords(records) {
    const listEl = document.getElementById('leave-records-list');
    if (!listEl) return;
    
    console.log(`📋 開始渲染 ${records.length} 筆請假記錄`);
    
    // ✅ 清空現有列表
    listEl.innerHTML = '';
    
    records.forEach((record, index) => {
        const card = document.createElement('div');
        card.className = 'card p-4 hover:shadow-lg transition-shadow';
        
        let statusClass = 'bg-yellow-100 text-yellow-800 dark:bg-yellow-900 dark:text-yellow-300';
        let statusText = '待審核';
        
        if (record.status === 'APPROVED') {
            statusClass = 'bg-green-100 text-green-800 dark:bg-green-900 dark:text-green-300';
            statusText = '已核准';
        } else if (record.status === 'REJECTED') {
            statusClass = 'bg-red-100 text-red-800 dark:bg-red-900 dark:text-red-300';
            statusText = '已拒絕';
        }
        
        const workHoursDisplay = record.workHours ? `${record.workHours} 小時` : '0 小時';
        
        // ✅ 修正：正確處理時間字段並格式化
        const startTime = formatDateTime(record.startDateTime || record.startTime);
        const endTime = formatDateTime(record.endDateTime || record.endTime);
        
        // 只在第一筆記錄時輸出範例日誌
        if (index === 0) {
            console.log('📋 記錄範例:', {
                leaveType: record.leaveType,
                startDateTime: record.startDateTime,
                endDateTime: record.endDateTime,
                workHours: record.workHours,
                status: record.status
            });
        }
        
        card.innerHTML = `
            <div class="flex justify-between items-start mb-3">
                <h3 class="font-bold text-lg text-gray-800 dark:text-white">
                    ${t(record.leaveType)}
                </h3>
                <span class="px-3 py-1 rounded-full text-sm font-medium ${statusClass}">
                    ${statusText}
                </span>
            </div>
            
            <div class="space-y-2 text-sm text-gray-600 dark:text-gray-400">
                <div class="flex items-center">
                    <svg class="w-4 h-4 mr-2" fill="currentColor" viewBox="0 0 20 20">
                        <path fill-rule="evenodd" d="M6 2a1 1 0 00-1 1v1H4a2 2 0 00-2 2v10a2 2 0 002 2h12a2 2 0 002-2V6a2 2 0 00-2-2h-1V3a1 1 0 10-2 0v1H7V3a1 1 0 00-1-1zm0 5a1 1 0 000 2h8a1 1 0 100-2H6z" clip-rule="evenodd"/>
                    </svg>
                    <span>${startTime} ~ ${endTime}</span>
                </div>
                
                <div class="flex items-center">
                    <svg class="w-4 h-4 mr-2" fill="currentColor" viewBox="0 0 20 20">
                        <path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm1-12a1 1 0 10-2 0v4a1 1 0 00.293.707l2.828 2.829a1 1 0 101.415-1.415L11 9.586V6z" clip-rule="evenodd"/>
                    </svg>
                    <span class="font-medium text-indigo-600 dark:text-indigo-400">
                        ${workHoursDisplay}
                    </span>
                </div>
                
                <div class="flex items-start">
                    <svg class="w-4 h-4 mr-2 mt-0.5" fill="currentColor" viewBox="0 0 20 20">
                        <path fill-rule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7-4a1 1 0 11-2 0 1 1 0 012 0zM9 9a1 1 0 000 2v3a1 1 0 001 1h1a1 1 0 100-2v-3a1 1 0 00-1-1H9z" clip-rule="evenodd"/>
                    </svg>
                    <span>請假原因:</span>
                    <span class="ml-1">${record.reason || '無'}</span>
                </div>
            </div>
        `;
        
        listEl.appendChild(card);
    });
    
    console.log(`✅ 渲染完成，共 ${records.length} 筆記錄`);
}

/**
 * 處理請假類型變更
 */
function handleLeaveTypeChange(e) {
    // 可以在這裡加入特殊處理邏輯
    console.log('請假類型變更:', e.target.value);
}

/**
 * 載入待審核的請假申請（管理員用）
 */
async function loadPendingLeaveRequests() {
    const loadingEl = document.getElementById('leave-requests-loading');
    const emptyEl = document.getElementById('leave-requests-empty');
    const listEl = document.getElementById('pending-leave-list');
    
    if (loadingEl) loadingEl.style.display = 'block';
    if (emptyEl) emptyEl.style.display = 'none';
    if (listEl) listEl.innerHTML = '';
    
    try {
        const res = await callApifetch('getPendingLeaveRequests');
        
        if (res.ok) {
            if (res.requests && res.requests.length > 0) {
                renderPendingLeaveRequests(res.requests);
            } else {
                if (emptyEl) emptyEl.style.display = 'block';
            }
        }
    } catch (err) {
        console.error('載入待審核請假失敗:', err);
    } finally {
        if (loadingEl) loadingEl.style.display = 'none';
    }
}

function renderPendingLeaveRequests(requests) {
    const listEl = document.getElementById('pending-leave-list');
    if (!listEl) return;
    
    listEl.innerHTML = '';
    
    requests.forEach(req => {
        const li = document.createElement('li');
        li.className = 'p-4 bg-gray-50 dark:bg-gray-700 rounded-lg';
        
        const timeDisplay = req.startDateTime && req.endDateTime
            ? `${formatDateTime(req.startDateTime)} ~ ${formatDateTime(req.endDateTime)}`
            : req.startDate && req.endDate
            ? `${req.startDate} ~ ${req.endDate}`
            : '時間未設定';
        
        const durationDisplay = req.workHours
            ? `${req.workHours} 小時`
            : req.days
            ? `${req.days} 天`
            : '時數未知';
        
        // ⭐⭐⭐ 新增：顯示餘額警告
        const balanceWarning = req.insufficientBalance 
            ? `<p class="text-xs text-red-600 dark:text-red-400 mt-2 font-semibold">
                   ⚠️ 該員工餘額不足（剩餘 ${req.remainingBalance} 天）
               </p>`
            : '';
        
        li.innerHTML = `
            <div class="flex flex-col space-y-2">
                <div class="flex items-center justify-between">
                    <div class="flex-1">
                        <p class="font-semibold text-gray-800 dark:text-white">
                            ${req.employeeName} - ${t(req.leaveType)}
                        </p>
                        <p class="text-sm text-gray-600 dark:text-gray-400">
                            ${timeDisplay}
                        </p>
                        <p class="text-xs text-gray-500 mt-1">
                            ${durationDisplay}
                        </p>
                        ${req.reason ? `
                            <p class="text-sm text-gray-600 dark:text-gray-400 mt-1">
                                原因：${req.reason}
                            </p>
                        ` : ''}
                        ${balanceWarning}
                    </div>
                </div>
                
                <div class="flex space-x-2 mt-2">
                    <button 
                        data-row="${req.rowNumber}" 
                        class="approve-leave-btn flex-1 px-3 py-2 rounded-md text-sm font-bold btn-primary">
                        核准
                    </button>
                    <button 
                        data-row="${req.rowNumber}" 
                        class="reject-leave-btn flex-1 px-3 py-2 rounded-md text-sm font-bold btn-warning">
                        拒絕
                    </button>
                </div>
            </div>
        `;
        
        listEl.appendChild(li);
    });
    
    // 綁定審核按鈕事件
    listEl.querySelectorAll('.approve-leave-btn').forEach(btn => {
        btn.addEventListener('click', (e) => handleReviewLeave(e.currentTarget, 'approve'));
    });
    
    listEl.querySelectorAll('.reject-leave-btn').forEach(btn => {
        btn.addEventListener('click', (e) => handleReviewLeave(e.currentTarget, 'reject'));
    });
}

/**
 * 處理請假審核
 */
async function handleReviewLeave(button, action) {
    const rowNumber = button.dataset.row;
    
    const comment = action === 'reject' 
        ? prompt('請輸入拒絕原因：') 
        : '';
    
    if (action === 'reject' && !comment) {
        showNotification('請輸入拒絕原因', 'warning');
        return;
    }
    
    button.disabled = true;
    button.textContent = '處理中...';
    
    try {
        const res = await callApifetch(
            `reviewLeave&rowNumber=${rowNumber}` +
            `&reviewAction=${action}` +
            `&comment=${encodeURIComponent(comment || '')}`
        );
        
        if (res.ok) {
            showNotification(action === 'approve' ? '已核准' : '已拒絕', 'success');
            await new Promise(resolve => setTimeout(resolve, 500));
            loadPendingLeaveRequests();
        } else {
            showNotification('審核失敗', 'error');
        }
    } catch (err) {
        console.error('審核請假失敗:', err);
        showNotification('網路錯誤', 'error');
    } finally {
        button.disabled = false;
        button.textContent = action === 'approve' ? '核准' : '拒絕';
    }
}