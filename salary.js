// salary.js - 薪資管理前端邏輯（完整版 v2.0 - 含所有津貼與扣款）
// ==================== 檢查依賴 ====================
if (typeof callApifetch !== 'function') {
    console.error('❌ callApifetch 函數未定義，請確認 script.js 已正確載入');
}

// ==================== 模組快取 ====================
let _sessionCache = null;
let _sessionCacheExpiry = 0;
const SESSION_CACHE_TTL = 30 * 60 * 1000; // 30 分鐘
let _currentNetSalary = 0;

async function getSession() {
    if (_sessionCache && Date.now() < _sessionCacheExpiry) return _sessionCache;
    _sessionCache = null;
    const session = await callApifetch('checkSession');
    if (session.ok && session.user) {
        _sessionCache = session;
        _sessionCacheExpiry = Date.now() + SESSION_CACHE_TTL;
    }
    return session;
}

// ==================== 初始化薪資頁面 ====================

/**
 * ✅ 初始化薪資頁面（完整版 + 多語言）
 */
async function initSalaryTab() {
    try {
        console.log('🎯 開始初始化薪資頁面（完整版 v2.0 + 多語言）');
        
        // 步驟 0：載入翻譯
        await loadTranslations(currentLang);
        
        // 步驟 1：驗證 Session
        const session = await getSession();

        if (!session.ok || !session.user) {
            showNotification(t('SALARY_LOGIN_REQUIRED'), 'error');
            return;
        }

        // 步驟 2：設定當前月份
        const now = new Date();
        const currentMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;

        const employeeSalaryMonth = document.getElementById('employee-salary-month');
        if (employeeSalaryMonth) {
            employeeSalaryMonth.value = currentMonth;
        }

        // 步驟 3：平行載入薪資資料與歷史
        await Promise.all([
            loadCurrentEmployeeSalary(),
            loadSalaryHistory()
        ]);

        // 步驟 4：綁定事件（管理員才需要）
        if (session.user.dept === "管理員") {
            bindSalaryEvents();
        }
        
    } catch (error) {
        console.error('❌ 初始化失敗:', error);
        console.error('錯誤堆疊:', error.stack);
        showNotification(t('SALARY_INIT_FAILED') + ': ' + error.message, 'error');
    }
}
// ==================== 員工薪資功能 ====================

/**
 * ✅ 載入當前員工的薪資（簡化版 - 後端自動計算）
 */
async function loadCurrentEmployeeSalary() {
    try {
        console.log(`💰 載入員工薪資`);
        
        const now = new Date();
        const currentMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
        
        const loadingEl = document.getElementById('current-salary-loading');
        const emptyEl = document.getElementById('current-salary-empty');
        const contentEl = document.getElementById('current-salary-content');
        
        if (loadingEl) loadingEl.style.display = 'block';
        if (emptyEl) emptyEl.style.display = 'none';
        if (contentEl) contentEl.style.display = 'none';
        
        // ⭐ 直接呼叫 getMySalary（後端會自動重新計算並儲存）
        const result = await callApifetch(`getMySalary&yearMonth=${currentMonth}`);
        
        console.log('📥 薪資資料回應:', result);
        
        if (loadingEl) loadingEl.style.display = 'none';
        
        if (result.ok && result.data) {
            displayEmployeeSalary(result.data);
            if (contentEl) contentEl.style.display = 'block';
            await Promise.all([
                loadAttendanceDetails(currentMonth, result.data),
                loadMonthlyLeaveStatus(currentMonth)
            ]);
        } else {
            if (emptyEl) {
                showNoSalaryMessage(currentMonth);
                emptyEl.style.display = 'block';
            }
        }
        
    } catch (error) {
        console.error('❌ 載入失敗:', error);
        const loadingEl = document.getElementById('current-salary-loading');
        const emptyEl = document.getElementById('current-salary-empty');
        if (loadingEl) loadingEl.style.display = 'none';
        if (emptyEl) emptyEl.style.display = 'block';
    }
}

/**
 * ✅ 按月份查詢薪資（修正版 - 先重新計算）
 */
async function loadEmployeeSalaryByMonth() {
    const monthInput = document.getElementById('employee-salary-month');
    const yearMonth = monthInput ? monthInput.value : '';
    
    if (!yearMonth) {
        showNotification(t('SALARY_SELECT_MONTH'), 'error');
        return;
    }
    
    const loadingEl = document.getElementById('current-salary-loading');
    const emptyEl = document.getElementById('current-salary-empty');
    const contentEl = document.getElementById('current-salary-content');
    
    if (!loadingEl || !emptyEl || !contentEl) {
        console.warn('薪資顯示元素未找到');
        return;
    }
    
    try {
        console.log(`🔍 查詢 ${yearMonth} 薪資（先重新計算）`);
        
        loadingEl.style.display = 'block';
        emptyEl.style.display = 'none';
        contentEl.style.display = 'none';
        
        const session = await getSession();
        if (!session.ok || !session.user) {
            throw new Error('Session 驗證失敗');
        }

        // getMySalary 後端已內建重新計算，不需要額外呼叫 calculateMonthlySalary
        const res = await callApifetch(`getMySalary&yearMonth=${yearMonth}`);
        
        console.log(`📥 查詢 ${yearMonth} 薪資回應:`, res);
        
        loadingEl.style.display = 'none';
        
        if (res.ok && res.data) {
            displayEmployeeSalary(res.data);
            contentEl.style.display = 'block';
            await Promise.all([
                loadAttendanceDetails(yearMonth, res.data),
                loadMonthlyLeaveStatus(yearMonth)
            ]);
        } else {
            console.log(`⚠️ 沒有 ${yearMonth} 的薪資記錄`);
            showNoSalaryMessage(yearMonth);
            emptyEl.style.display = 'block';
            const detailsSection = document.getElementById('attendance-details-section');
            if (detailsSection) detailsSection.style.display = 'none';
            const leaveSection = document.getElementById('leave-status-section');
            if (leaveSection) leaveSection.style.display = 'none';
        }
        
    } catch (error) {
        console.error(`❌ 載入 ${yearMonth} 薪資失敗:`, error);
        loadingEl.style.display = 'none';
        emptyEl.style.display = 'block';
    }
}
/**
 * ✅ 載入每日加班明細
 */
async function loadDailyOvertimeDetails(yearMonth) {
    const detailsContainer = document.getElementById('overtime-details');
    if (!detailsContainer) return;
    
    try {
        detailsContainer.innerHTML = '<p class="text-sm text-gray-400">載入中...</p>';
        
        // ⭐ 呼叫後端 API 取得加班記錄
        const res = await callApifetch(`getEmployeeMonthlyOvertime&yearMonth=${yearMonth}`);
        
        console.log('📥 加班記錄回應:', res);
        
        if (res.ok && res.records && res.records.length > 0) {
            detailsContainer.innerHTML = '';
            
            res.records.forEach(record => {
                const item = document.createElement('div');
                item.className = 'flex justify-between items-center p-2 bg-orange-800/10 rounded border border-orange-700/30';
                
                const hours = parseFloat(record.hours) || 0;
                
                item.innerHTML = `
                    <div>
                        <span class="font-semibold text-orange-200">${record.date}</span>
                        <span class="text-sm text-orange-400 ml-2">已核准</span>
                    </div>
                    <div class="text-right">
                        <span class="font-mono text-orange-300 font-bold">${hours.toFixed(1)}h</span>
                    </div>
                `;
                
                detailsContainer.appendChild(item);
            });
        } else {
            detailsContainer.innerHTML = '<p class="text-sm text-gray-400">本月無加班記錄</p>';
        }
        
    } catch (error) {
        console.error('❌ 載入加班明細失敗:', error);
        detailsContainer.innerHTML = '<p class="text-sm text-red-400">載入失敗</p>';
    }
}

async function loadOvertimeRecordsCard(yearMonth, salaryData) {
    console.log('📊 載入加班記錄卡片');
    
    // ⭐⭐⭐ 修正：讀取四種加班費
    const totalOvertimeHours = parseFloat(
        salaryData.totalOvertimeHours !== undefined 
            ? salaryData.totalOvertimeHours 
            : salaryData['總加班時數']
    ) || 0;
    
    const weekdayOvertimePay = parseFloat(
        salaryData.weekdayOvertimePay !== undefined 
            ? salaryData.weekdayOvertimePay 
            : salaryData['平日加班費']
    ) || 0;
    
    const restdayOvertimePay = parseFloat(
        salaryData.restdayOvertimePay !== undefined 
            ? salaryData.restdayOvertimePay 
            : salaryData['休息日加班費']
    ) || 0;
    
    const sundayOvertimePay = parseFloat(
        salaryData.sundayOvertimePay !== undefined 
            ? salaryData.sundayOvertimePay 
            : salaryData['例假日加班費']
    ) || 0;
    
    const holidayOvertimePay = parseFloat(
        salaryData.holidayOvertimePay !== undefined 
            ? salaryData.holidayOvertimePay 
            : salaryData['國定假日加班費']
    ) || 0;
    
    const holidayWorkPay = parseFloat(
        salaryData.holidayWorkPay !== undefined 
            ? salaryData.holidayWorkPay 
            : salaryData['國定假日出勤薪資']
    ) || 0;
    
    const totalOvertimePay = weekdayOvertimePay + restdayOvertimePay + 
                            sundayOvertimePay + holidayOvertimePay + holidayWorkPay;
    
    console.log(`⏰ 總加班: ${totalOvertimeHours}h`);
    console.log(`   平日: $${weekdayOvertimePay}`);
    console.log(`   休息日: $${restdayOvertimePay}`);
    console.log(`   例假日: $${sundayOvertimePay}`);
    console.log(`   國定假日加班費: $${holidayOvertimePay}`);
    console.log(`   國定假日出勤薪資: $${holidayWorkPay}`);
    
    // ... 顯示邏輯 ...
    
    if (totalOvertimeHours > 0) {
        overtimeCard.style.display = 'block';
        
        overtimeCard.innerHTML = `
            <h4 class="font-semibold mb-3 text-orange-400">⏰ 本月加班統計</h4>
            
            <div class="grid grid-cols-3 gap-4 mb-4">
                <div class="text-center p-3 bg-orange-800/20 rounded-lg">
                    <p class="text-sm text-orange-300 mb-1">總加班時數</p>
                    <p class="text-2xl font-bold text-orange-200">${totalOvertimeHours.toFixed(1)}h</p>
                </div>
                <div class="text-center p-3 bg-orange-800/20 rounded-lg">
                    <p class="text-sm text-orange-300 mb-1">平日加班費</p>
                    <p class="text-xl font-bold text-orange-200">${formatCurrency(weekdayOvertimePay)}</p>
                    <p class="text-xs text-orange-400 mt-1">(前2h ×1.34, 後2h ×1.67)</p>
                </div>
                <div class="text-center p-3 bg-orange-800/20 rounded-lg">
                    <p class="text-sm text-orange-300 mb-1">假日加班費</p>
                    <p class="text-xl font-bold text-orange-200">${formatCurrency(restdayOvertimePay + sundayOvertimePay + holidayOvertimePay + holidayWorkPay)}</p>
                    <p class="text-xs text-orange-400 mt-1">(週六/日/國定 ×1.34~2.67)</p>
                </div>
            </div>
            
            <!-- ⭐ 詳細分類 -->
            ${restdayOvertimePay > 0 || sundayOvertimePay > 0 || holidayOvertimePay > 0 || holidayWorkPay > 0 ? `
                <div class="p-3 bg-orange-800/10 rounded-lg mb-3">
                    <div class="text-sm space-y-1">
                        ${restdayOvertimePay > 0 ? `
                            <div class="flex justify-between">
                                <span class="text-orange-300">休息日（週六）</span>
                                <span class="font-mono text-orange-200">${formatCurrency(restdayOvertimePay)}</span>
                            </div>
                        ` : ''}
                        ${sundayOvertimePay > 0 ? `
                            <div class="flex justify-between">
                                <span class="text-orange-300">例假日（週日）×2.0</span>
                                <span class="font-mono text-orange-200">${formatCurrency(sundayOvertimePay)}</span>
                            </div>
                        ` : ''}
                        ${holidayWorkPay > 0 ? `
                            <div class="flex justify-between border-t border-orange-700/30 pt-2">
                                <span class="text-orange-300 font-semibold">國定假日出勤薪資</span>
                                <span class="font-mono text-orange-200 font-bold">${formatCurrency(holidayWorkPay)}</span>
                            </div>
                        ` : ''}
                        ${holidayOvertimePay > 0 ? `
                            <div class="flex justify-between">
                                <span class="text-orange-300 font-semibold">國定假日加班費 ×2.0</span>
                                <span class="font-mono text-orange-200 font-bold">${formatCurrency(holidayOvertimePay)}</span>
                            </div>
                        ` : ''}
                    </div>
                </div>
            ` : ''}
            
            <div class="p-3 bg-orange-800/20 rounded-lg">
                <div class="flex justify-between items-center">
                    <span class="font-semibold text-orange-200">加班費合計</span>
                    <span class="text-2xl font-bold text-orange-300">${formatCurrency(totalOvertimePay)}</span>
                </div>
            </div>
            
            <div id="overtime-details" class="mt-4 space-y-2">
                <!-- 每日加班明細將動態載入 -->
            </div>
        `;
        
        await loadDailyOvertimeDetails(yearMonth);
        
    } else {
        overtimeCard.style.display = 'none';
    }
}

/**
 * ✅ 載入每日工時明細
 */
async function loadDailyWorkHours(yearMonth) {
    const detailsContainer = document.getElementById('work-hours-details');
    if (!detailsContainer) return;
    
    try {
        detailsContainer.innerHTML = '<p class="text-sm text-gray-400">載入中...</p>';
        
        // ⭐ 呼叫後端 API 取得打卡記錄
        const res = await callApifetch(`getEmployeeMonthlyAttendance&yearMonth=${yearMonth}`);
        
        console.log('📥 打卡記錄回應:', res);
        
        if (res.ok && res.records && res.records.length > 0) {
            detailsContainer.innerHTML = '';
            
            res.records.forEach(record => {
                const item = document.createElement('div');
                item.className = 'flex justify-between items-center p-2 bg-purple-800/10 rounded border border-purple-700/30';
                
                const workHours = parseFloat(record.workHours) || 0;
                
                item.innerHTML = `
                    <div>
                        <span class="font-semibold text-purple-200">${record.date}</span>
                        <span class="text-sm text-purple-400 ml-2">
                            ${record.punchIn || '--'} ~ ${record.punchOut || '--'}
                        </span>
                    </div>
                    <div class="text-right">
                        <span class="font-mono text-purple-300 font-bold">${workHours.toFixed(1)}h</span>
                    </div>
                `;
                
                detailsContainer.appendChild(item);
            });
        } else {
            detailsContainer.innerHTML = '<p class="text-sm text-gray-400">本月無打卡記錄</p>';
        }
        
    } catch (error) {
        console.error('❌ 載入每日工時失敗:', error);
        detailsContainer.innerHTML = '<p class="text-sm text-red-400">載入失敗</p>';
    }
}

/**
 * ✅ 載入工作時數卡片（時薪專用）
 */
async function loadWorkHoursCard(yearMonth, salaryData) {
    console.log('📊 載入工作時數卡片');
    
    // 從薪資資料中取得工時資訊
    const totalWorkHours = parseFloat(salaryData['工作時數']) || 0;
    const hourlyRate = parseFloat(salaryData['時薪']) || 0;
    const baseSalary = parseFloat(salaryData['基本薪資']) || 0;
    const totalWorkHoursInt = Math.floor(totalWorkHours);
    console.log(`⏱️ 總工時: ${totalWorkHours}h, 時薪: $${hourlyRate}, 基本薪資: $${baseSalary}`);
    
    // 建立工時卡片
    let workHoursCard = document.getElementById('work-hours-card');
    
    if (!workHoursCard) {
        workHoursCard = document.createElement('div');
        workHoursCard.id = 'work-hours-card';
        workHoursCard.className = 'feature-box bg-purple-900/20 border-purple-700 mb-4';
        
        const detailsSection = document.getElementById('attendance-details-section');
        const firstChild = detailsSection.firstChild;
        detailsSection.insertBefore(workHoursCard, firstChild);
    }
    
    workHoursCard.innerHTML = `
        <h4 class="font-semibold mb-3 text-purple-400">⏰ 本月工作時數統計</h4>
        
        <div class="grid grid-cols-3 gap-4 mb-4">
            <div class="text-center p-3 bg-purple-800/20 rounded-lg">
                <p class="text-sm text-purple-300 mb-1">時薪</p>
                <p class="text-2xl font-bold text-purple-200">$${hourlyRate}</p>
            </div>
            <div class="text-center p-3 bg-purple-800/20 rounded-lg">
                <p class="text-sm text-purple-300 mb-1">總工作時數</p>
                <p class="text-2xl font-bold text-purple-200">${Math.floor(totalWorkHours)}h</p>
            </div>
            <div class="text-center p-3 bg-purple-800/20 rounded-lg">
                <p class="text-sm text-purple-300 mb-1">基本薪資</p>
                <p class="text-2xl font-bold text-purple-200">${formatCurrency(baseSalary)}</p>
                <p class="text-xs text-purple-400 mt-1">(時薪 × 工時)</p>
            </div>
        </div>
        
        <div id="work-hours-details" class="space-y-2">
            <!-- 每日工時明細將動態載入 -->
        </div>
    `;
    
    // 載入每日工時明細
    await loadDailyWorkHours(yearMonth);
}

function displayEmployeeSalary(data) {
    console.log('📊 顯示薪資明細（完整版）:', data);
    
    const safeSet = (id, value) => {
        const el = document.getElementById(id);
        if (el) {
            el.textContent = value;
        } else {
            console.warn(`⚠️ 元素 #${id} 未找到`);
        }
    };
    
    // ⭐⭐⭐ 關鍵修正：改用英文欄位
    const salaryType = data.salaryType || '月薪';
    const isHourly = salaryType === '時薪';
    
    // 應發總額與實發金額（直接使用後端計算值，保持前後端一致）
    const grossSalary = parseFloat(data.grossSalary) || 0;
    const netSalary   = parseFloat(data.netSalary)   || 0;
    safeSet('gross-salary',    formatCurrency(grossSalary));
    safeSet('net-salary',      formatCurrency(netSalary));
    // 扣款總額 = 應發 - 實發，完全跟後端一致，不在前端重算
    safeSet('total-deductions', formatCurrency(grossSalary - netSalary));
    
    // 應發項目（全部改成英文欄位）
    if (isHourly) {
        const hourlyRate = parseFloat(data.hourlyRate) || 0;
        const totalWorkHours = parseFloat(data.totalWorkHours) || 0;
        
        const baseSalaryLabel = document.querySelector('[for="detail-base-salary"]') || 
                                document.querySelector('#detail-base-salary')?.previousElementSibling;
        if (baseSalaryLabel) {
            baseSalaryLabel.textContent = '基本薪資 (時薪×工時)';
        }
        
        safeSet('detail-base-salary', formatCurrency(data.baseSalary));
        
        const baseSalaryEl = document.getElementById('detail-base-salary');
        if (baseSalaryEl && baseSalaryEl.parentElement) {
            let hourlyInfo = baseSalaryEl.parentElement.querySelector('.hourly-info');
            if (!hourlyInfo) {
                hourlyInfo = document.createElement('div');
                hourlyInfo.className = 'hourly-info text-xs text-purple-400 mt-1';
                baseSalaryEl.parentElement.appendChild(hourlyInfo);
            }
            hourlyInfo.textContent = `時薪 $${hourlyRate} × ${Math.floor(totalWorkHours)}h`;
        }
    } else {
        safeSet('detail-base-salary', formatCurrency(data.baseSalary));
        
        const baseSalaryEl = document.getElementById('detail-base-salary');
        if (baseSalaryEl && baseSalaryEl.parentElement) {
            const hourlyInfo = baseSalaryEl.parentElement.querySelector('.hourly-info');
            if (hourlyInfo) {
                hourlyInfo.remove();
            }
            
            const baseSalaryLabel = document.querySelector('[for="detail-base-salary"]') || 
                                    baseSalaryEl.previousElementSibling;
            if (baseSalaryLabel) {
                baseSalaryLabel.textContent = '基本薪資';
            }
        }
    }
    
    // ⭐ 工時統計資訊
    const totalWorkHours = parseFloat(data.totalWorkHours) || 0;
    const totalOvertimeHours = parseFloat(data.totalOvertimeHours) || 0;

    const weekdayOvertimeEl = document.getElementById('detail-weekday-overtime');
    if (weekdayOvertimeEl && weekdayOvertimeEl.parentElement) {
        // ⭐⭐⭐ 修正：先移除舊的工時統計區塊
        const oldWorkHoursInfo = weekdayOvertimeEl.parentElement.querySelector('.work-hours-summary');
        if (oldWorkHoursInfo) {
            oldWorkHoursInfo.remove();
        }
        
        // ⭐⭐⭐ 新增：檢查是否已存在工時統計區塊（在父容器層級）
        const container = weekdayOvertimeEl.closest('.space-y-2') || weekdayOvertimeEl.parentElement.parentElement;
        const existingSummaries = container.querySelectorAll('.work-hours-summary');
        existingSummaries.forEach(summary => summary.remove());
        
        if (totalWorkHours > 0 || totalOvertimeHours > 0) {
            const workHoursSummary = document.createElement('div');
            workHoursSummary.className = 'work-hours-summary mb-3 p-3 bg-blue-900/20 border border-blue-700/30 rounded-lg';
            
            let summaryHTML = '<div class="text-sm font-semibold text-blue-300 mb-2">本月工時統計</div>';
            
            if (isHourly && totalWorkHours > 0) {
                summaryHTML += `
                    <div class="flex justify-between text-sm mb-1">
                        <span class="text-blue-200">打卡工作時數：</span>
                        <span class="font-mono text-blue-100">${Math.floor(totalWorkHours)}h</span>
                    </div>
                `;
            }
            
            if (totalOvertimeHours > 0) {
                summaryHTML += `
                    <div class="flex justify-between text-sm">
                        <span class="text-orange-200">加班時數：</span>
                        <span class="font-mono text-orange-100">${totalOvertimeHours.toFixed(1)}h</span>
                    </div>
                `;
            }
            
            workHoursSummary.innerHTML = summaryHTML;
            
            // ⭐⭐⭐ 修正：只插入一次
            weekdayOvertimeEl.parentElement.parentElement.insertBefore(
                workHoursSummary,
                weekdayOvertimeEl.parentElement
            );
        }
    }
        
    // 其他津貼
    safeSet('detail-position-allowance', formatCurrency(data.positionAllowance || 0));
    safeSet('detail-meal-allowance', formatCurrency(data.mealAllowance || 0));
    safeSet('detail-transport-allowance', formatCurrency(data.transportAllowance || 0));
    safeSet('detail-attendance-bonus', formatCurrency(data.attendanceBonus || 0));
    safeSet('detail-performance-bonus', formatCurrency(data.performanceBonus || 0));
    
    // 加班費
    safeSet('detail-weekday-overtime', formatCurrency(data.weekdayOvertimePay || 0));
    safeSet('detail-restday-overtime', formatCurrency(data.restdayOvertimePay || 0));
    safeSet('detail-holiday-overtime', formatCurrency(data.holidayOvertimePay || 0));

    // 例假日（週日）加班費 + 國定假日出勤薪資：有值才顯示
    const sundayOT = parseFloat(data.sundayOvertimePay) || 0;
    const holidayWork = parseFloat(data.holidayWorkPay) || 0;
    safeSet('detail-sunday-overtime', formatCurrency(sundayOT));
    safeSet('detail-holiday-work-pay', formatCurrency(holidayWork));
    const rowSunday = document.getElementById('row-sunday-overtime');
    const rowHolidayWork = document.getElementById('row-holiday-work-pay');
    if (rowSunday) rowSunday.style.display = sundayOT > 0 ? 'flex' : 'none';
    if (rowHolidayWork) rowHolidayWork.style.display = holidayWork > 0 ? 'flex' : 'none';
    
    // 扣款項目
    safeSet('detail-labor-fee', formatCurrency(data.laborFee));
    safeSet('detail-health-fee', formatCurrency(data.healthFee));
    safeSet('detail-employment-fee', formatCurrency(data.employmentFee));
    
    const pensionRate = parseFloat(data.pensionSelfRate) || 0;
    safeSet('detail-pension-rate', `${pensionRate}%`);
    
    safeSet('detail-pension-self', formatCurrency(data.pensionSelf));
    safeSet('detail-income-tax', formatCurrency(data.incomeTax));
    safeSet('detail-leave-deduction', formatCurrency(data.leaveDeduction));
    
    // ⭐⭐⭐ 新增：早退扣款顯示
    const earlyLeaveDeduction = parseFloat(data.earlyLeaveDeduction || data['早退扣款']) || 0;
    safeSet('detail-early-leave-deduction', formatCurrency(earlyLeaveDeduction));

    const sickLeaveHours = parseFloat(data.sickLeaveHours) || 0;  // ⭐ 改名
    const sickLeaveDeduction = parseFloat(data.sickLeaveDeduction) || 0;
    const personalLeaveHours = parseFloat(data.personalLeaveHours) || 0;  // ⭐ 改名
    const personalLeaveDeduction = parseFloat(data.personalLeaveDeduction) || 0;
    
    console.log('🔍 請假資料檢查:');
    console.log('   病假時數:', sickLeaveHours);  // ⭐ 改名
    console.log('   病假扣款:', sickLeaveDeduction);
    console.log('   事假時數:', personalLeaveHours);  // ⭐ 改名
    console.log('   事假扣款:', personalLeaveDeduction);
    console.log('   早退扣款:', earlyLeaveDeduction); // ⭐ 新增

    const leaveDeductionEl = document.getElementById('detail-leave-deduction');

    if (leaveDeductionEl) {
        let container = leaveDeductionEl.closest('.space-y-2');
        
        if (!container) {
            container = leaveDeductionEl.parentElement;
        }
        
        if (container) {
            const oldLeaveDetails = container.querySelector('.leave-details');
            if (oldLeaveDetails) {
                oldLeaveDetails.remove();
            }
            
            if (sickLeaveHours > 0 || personalLeaveHours > 0) {  // ⭐ 改名
                const leaveDetails = document.createElement('div');
                leaveDetails.className = 'leave-details p-2 bg-yellow-900/20 rounded-lg mt-2 mb-2 border border-yellow-700/30';
                
                let detailsHTML = '<div class="text-xs space-y-1">';
                
                if (sickLeaveHours > 0) {  // ⭐ 改名
                    detailsHTML += `
                        <div class="flex justify-between">
                            <span class="text-yellow-300">病假 ${sickLeaveHours} 小時 (半薪)</span>
                            <span class="font-mono text-yellow-200 font-bold">${formatCurrency(sickLeaveDeduction)}</span>
                        </div>
                    `;
                }
                
                if (personalLeaveHours > 0) {  // ⭐ 改名
                    detailsHTML += `
                        <div class="flex justify-between">
                            <span class="text-yellow-300">事假 ${personalLeaveHours} 小時 (全薪)</span>
                            <span class="font-mono text-yellow-200 font-bold">${formatCurrency(personalLeaveDeduction)}</span>
                        </div>
                    `;
                }
                
                detailsHTML += '</div>';
                leaveDetails.innerHTML = detailsHTML;
                
                const leaveDeductionRow = leaveDeductionEl.closest('.flex');
                if (leaveDeductionRow && leaveDeductionRow.nextSibling) {
                    leaveDeductionRow.parentNode.insertBefore(leaveDetails, leaveDeductionRow.nextSibling);
                } else {
                    container.appendChild(leaveDetails);
                }
                
                console.log('請假明細已顯示');
            } else {
                console.log('本月無請假記錄');
            }
        }
    }
    
    const otherDeductions = 
        (parseFloat(data.welfareFee) || 0) +
        (parseFloat(data.dormitoryFee) || 0) +
        (parseFloat(data.groupInsurance) || 0) +
        (parseFloat(data.otherDeductions) || 0);
    safeSet('detail-other-deductions', formatCurrency(otherDeductions));

    // 預支扣款子項顯示
    const advanceDeduction = parseFloat(data.advanceDeduction) || 0;
    const advRow = document.getElementById('detail-advance-row');
    if (advRow) {
      advRow.style.display = advanceDeduction > 0 ? 'flex' : 'none';
    }
    safeSet('detail-advance-deduction', formatCurrency(advanceDeduction));

    // 銀行資訊
    let bankCode = data.bankCode;
    const bankAccount = data.bankAccount;
    
    if (bankCode) {
        bankCode = String(bankCode).padStart(3, '0');
    }
    
    safeSet('detail-bank-name', getBankName(bankCode));
    safeSet('detail-bank-account', bankAccount || '--');

    // 更新即時匯率換算面板
    _currentNetSalary = parseFloat(data.netSalary) || 0;
    const panel = document.getElementById('currency-converter-panel');
    if (panel) {
        panel.style.display = _currentNetSalary > 0 ? 'block' : 'none';
        refreshCurrencyDisplay();
    }
}

// ==================== 假勤狀況面板 ====================

const LEAVE_TYPE_LABELS = {
    SICK_LEAVE: '病假',
    PERSONAL_LEAVE: '事假',
    ANNUAL_LEAVE: '特休',
    COMPENSATORY_LEAVE: '補休',
    OFFICIAL_LEAVE: '公假',
    MATERNITY_LEAVE: '產假',
    PATERNITY_LEAVE: '陪產假',
    FUNERAL_LEAVE: '喪假',
};

/**
 * 載入並顯示本月假單狀況，讓員工了解哪些假已計入薪資
 */
async function loadMonthlyLeaveStatus(yearMonth) {
    const section = document.getElementById('leave-status-section');
    const loadingEl = document.getElementById('leave-status-loading');
    const contentEl = document.getElementById('leave-status-content');
    if (!section || !contentEl) return;

    section.style.display = 'block';
    loadingEl.style.display = 'block';
    contentEl.innerHTML = '';

    try {
        const res = await callApifetch('getEmployeeLeaveRecords');
        loadingEl.style.display = 'none';

        if (!res.ok || !res.records) {
            contentEl.innerHTML = '<p class="text-sm text-gray-400">無法載入假單資料</p>';
            return;
        }

        // 篩選本月假單
        const monthRecords = res.records.filter(r => {
            const dt = r.startDateTime instanceof Date
                ? r.startDateTime
                : new Date(r.startDateTime);
            const ym = `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}`;
            return ym === yearMonth;
        });

        if (monthRecords.length === 0) {
            contentEl.innerHTML = '<p class="text-sm text-gray-400 py-2">本月無請假記錄</p>';
            return;
        }

        // 分組
        const approved = monthRecords.filter(r => r.status === 'APPROVED' || r.status === '核准');
        const pending  = monthRecords.filter(r => r.status === 'PENDING'  || r.status === '待審核');
        const rejected = monthRecords.filter(r => r.status === 'REJECTED' || r.status === '拒絕');

        let html = '';

        // ⚠️ 待審核警示
        if (pending.length > 0) {
            html += `
            <div class="mb-3 p-3 rounded-lg border border-yellow-500/40" style="background:rgba(234,179,8,0.08);">
                <div class="flex items-center gap-2 font-semibold text-yellow-300 mb-2">
                    ⚠️ 待審核假單（尚未計入薪資，核准後將自動重新計算）
                </div>
                <div class="space-y-1">
                ${pending.map(r => leaveRecordRow(r, 'text-yellow-200', '⏳')).join('')}
                </div>
            </div>`;
        }

        // ✅ 已核准（已計入扣款）
        if (approved.length > 0) {
            html += `
            <div class="mb-3 p-3 rounded-lg border border-green-500/30" style="background:rgba(34,197,94,0.06);">
                <div class="font-semibold text-green-400 mb-2">✅ 已核准（已計入薪資扣款）</div>
                <div class="space-y-1">
                ${approved.map(r => leaveRecordRow(r, 'text-green-200', '✅')).join('')}
                </div>
            </div>`;
        }

        // ❌ 已拒絕（不計入）
        if (rejected.length > 0) {
            html += `
            <div class="p-3 rounded-lg border border-gray-600/30" style="background:rgba(100,116,139,0.06);">
                <div class="font-semibold text-gray-400 mb-2">❌ 已拒絕（不計入扣款）</div>
                <div class="space-y-1">
                ${rejected.map(r => leaveRecordRow(r, 'text-gray-400', '❌')).join('')}
                </div>
            </div>`;
        }

        contentEl.innerHTML = html;

    } catch (err) {
        loadingEl.style.display = 'none';
        contentEl.innerHTML = '<p class="text-sm text-red-400">載入失敗</p>';
    }
}

function leaveRecordRow(r, colorClass, icon) {
    const typeLabel = LEAVE_TYPE_LABELS[r.leaveType] || r.leaveType || '--';
    const days = parseFloat(r.days) || 0;
    const hours = parseFloat(r.workHours) || 0;
    const startDt = r.startDateTime ? new Date(r.startDateTime) : null;
    const dateStr = startDt ? startDt.toLocaleDateString('zh-TW', { timeZone: 'Asia/Taipei' }) : '--';
    return `
        <div class="flex justify-between items-center text-sm py-1 border-b border-white/5 last:border-0">
            <span class="${colorClass}">${icon} ${typeLabel}　${dateStr}</span>
            <span class="font-mono ${colorClass}">${days > 0 ? days + '天' : ''}${hours > 0 ? ' / ' + hours + 'h' : ''}</span>
        </div>`;
}

/**
 * 重新計算薪資按鈕的 handler
 */
async function recalculateSalary() {
    const btn = document.getElementById('recalc-btn');
    if (btn) {
        btn.disabled = true;
        btn.textContent = '⏳ 計算中...';
    }
    try {
        await loadEmployeeSalaryByMonth();
    } finally {
        if (btn) {
            btn.disabled = false;
            btn.textContent = '🔄 重新計算';
        }
    }
}

/**
 * ✅ 載入薪資歷史
 */
async function loadSalaryHistory() {
    const loadingEl = document.getElementById('salary-history-loading');
    const emptyEl = document.getElementById('salary-history-empty');
    const listEl = document.getElementById('salary-history-list');
    
    if (!loadingEl || !emptyEl || !listEl) {
        console.warn('薪資歷史元素未找到');
        return;
    }
    
    try {
        console.log('📋 載入薪資歷史');
        
        loadingEl.style.display = 'block';
        emptyEl.style.display = 'none';
        listEl.innerHTML = '';
        
        const res = await callApifetch('getMySalaryHistory&limit=12');
        
        console.log('📥 薪資歷史回應:', res);
        
        loadingEl.style.display = 'none';
        
        if (res.ok && res.data && res.data.length > 0) {
            console.log(`✅ 找到 ${res.data.length} 筆薪資歷史`);
            res.data.forEach(salary => {
                const item = createSalaryHistoryItem(salary);
                listEl.appendChild(item);
            });
        } else {
            console.log('⚠️ 沒有薪資歷史記錄');
            emptyEl.style.display = 'block';
        }
        
    } catch (error) {
        console.error('❌ 載入薪資歷史失敗:', error);
        loadingEl.style.display = 'none';
        emptyEl.style.display = 'block';
    }
}

/**
 * 建立薪資歷史項目
 */
function createSalaryHistoryItem(salary) {
    const div = document.createElement('div');
    div.className = 'feature-box flex justify-between items-center hover:bg-white/10 transition cursor-pointer';
    
    div.innerHTML = `
        <div>
            <div class="font-semibold text-lg">
                ${salary['年月'] || '--'}
            </div>
            <div class="text-sm text-gray-400 mt-1">
                ${salary['狀態'] || '已計算'}
            </div>
        </div>
        <div class="text-right">
            <div class="text-2xl font-bold text-purple-400">
                ${formatCurrency(salary['實發金額'])}
            </div>
            <div class="text-xs text-gray-400 mt-1">
                應發 ${formatCurrency(salary['應發總額'])}
            </div>
        </div>
    `;
    
    return div;
}

/**
 * 顯示無薪資訊息
 */
function showNoSalaryMessage(month) {
    const emptyEl = document.getElementById('current-salary-empty');
    if (emptyEl) {
        emptyEl.innerHTML = `
            <div class="empty-state-icon">📄</div>
            <div class="empty-state-title">尚無薪資記錄</div>
            <div class="empty-state-text">
                <p>${month} 還沒有薪資資料</p>
                <p style="margin-top: 0.5rem; font-size: 0.875rem;">
                    💡 提示：薪資需要由管理員先設定和計算<br>
                    請聯繫您的主管或人資部門
                </p>
            </div>
        `;
    }
}

// ==================== 管理員功能 ====================

function bindSalaryEvents() {
    console.log('🔗 綁定薪資表單事件（完整版）');
    
    const configForm = document.getElementById('salary-config-form');
    if (configForm) {
        // ⭐⭐⭐ 修正：移除舊的監聽器，避免重複綁定
        configForm.removeEventListener('submit', handleSalaryConfigSubmit);
        
        // ⭐⭐⭐ 修正：使用 addEventListener 而不是 onsubmit
        configForm.addEventListener('submit', async (e) => {
            e.preventDefault();  // ← 立即阻止預設行為
            e.stopPropagation(); // ← 阻止事件冒泡
            
            await handleSalaryConfigSubmit(e);
        });
        
        console.log('✅ 薪資設定表單已綁定');
    }
    
    const calculateBtn = document.getElementById('calculate-salary-btn');
    if (calculateBtn) {
        calculateBtn.addEventListener('click', handleSalaryCalculation);
        console.log('✅ 薪資計算按鈕已綁定');
    }
}

/**
 * ✅ 處理薪資設定表單提交（完整版 - 含所有津貼與扣款）
 */
async function handleSalaryConfigSubmit(e) {
    e.preventDefault();
    
    console.log('📝 開始提交薪資設定表單（完整版）');
    
    const safeGetValue = (id) => {
        const el = document.getElementById(id);
        return el ? el.value.trim() : '';
    };
    
    const toNumber = (value) => {
        if (value === '' || value === null || value === undefined) {
            return 0; // 空值視為 0
        }
        const num = parseFloat(value);
        return isNaN(num) ? 0 : Math.round(num);
    };
    // 基本資訊
    const employeeId = safeGetValue('config-employee-id');
    const employeeName = safeGetValue('config-employee-name');
    const idNumber = safeGetValue('config-id-number');
    const employeeType = safeGetValue('config-employee-type');
    const salaryType = safeGetValue('config-salary-type');
    const baseSalary = toNumber(safeGetValue('config-base-salary'));  // ⭐ 改這裡

    if (!employeeId || !employeeName || !salaryType) {
        showNotification(t('SALARY_FILL_REQUIRED'), 'error');
        return;
    }
    
    // ⭐⭐⭐ 只有月薪才需要檢查基本薪資 > 0
    if (salaryType === '月薪' && baseSalary <= 0) {
        showNotification('月薪的基本薪資必須大於 0', 'error');
        return;
    }
    
    // 時薪可以是 0（表示尚未設定）
    if (salaryType === '時薪' && isNaN(baseSalary)) {
        showNotification('請輸入有效的時薪', 'error');
        return;
    }

    // ⭐ 固定津貼（6項）
    const positionAllowance = toNumber(safeGetValue('config-position-allowance'));  // ⭐ 改這裡
    const mealAllowance = toNumber(safeGetValue('config-meal-allowance'));          // ⭐ 改這裡
    const transportAllowance = toNumber(safeGetValue('config-transport-allowance'));// ⭐ 改這裡
    const attendanceBonus = toNumber(safeGetValue('config-attendance-bonus'));      // ⭐ 改這裡
    const performanceBonus = toNumber(safeGetValue('config-performance-bonus'));    // ⭐ 改這裡
    const otherAllowances = toNumber(safeGetValue('config-other-allowances'));      // ⭐ 改這裡

    // 法定扣款
    const laborFee = toNumber(safeGetValue('config-labor-fee'));            // ⭐ 改這裡
    const healthFee = toNumber(safeGetValue('config-health-fee'));          // ⭐ 改這裡
    const employmentFee = toNumber(safeGetValue('config-employment-fee'));  // ⭐ 改這裡
    const pensionSelf = toNumber(safeGetValue('config-pension-self'));      // ⭐ 改這裡
    const incomeTax = toNumber(safeGetValue('config-income-tax'));          // ⭐ 改這裡
    const pensionSelfRate = toNumber(safeGetValue('config-pension-rate'));  // ⭐ 改這裡

    // ⭐ 其他扣款（4項）
    const welfareFee = toNumber(safeGetValue('config-welfare-fee'));        // ⭐ 改這裡
    const dormitoryFee = toNumber(safeGetValue('config-dormitory-fee'));    // ⭐ 改這裡
    const groupInsurance = toNumber(safeGetValue('config-group-insurance'));// ⭐ 改這裡
    const otherDeductions = toNumber(safeGetValue('config-other-deductions'));// ⭐ 改這裡

    // 其他資訊
    const bankCodeRaw = document.getElementById('config-bank-code').value;
    const bankCode = bankCodeRaw ? String(bankCodeRaw).padStart(3, '0') : '';
    const bankAccount = safeGetValue('config-bank-account');
    const hireDate = safeGetValue('config-hire-date');
    const paymentDay = safeGetValue('config-payment-day') || '5';
    const note = safeGetValue('config-note');
    
    // 驗證
    if (!employeeId || !employeeName || !baseSalary || parseFloat(baseSalary) <= 0) {
        showNotification(t('SALARY_FILL_REQUIRED'), 'error');
        return;
    }
    
    if (!employeeType || !salaryType) {
        showNotification(t('SALARY_SELECT_TYPE'), 'error');
        return;
    }
    
    try {
        showNotification(t('SALARY_SAVING'), 'info');
        
        // ⭐ 重新排序參數，與後端 Sheet 欄位順序一致
        const queryString = 
            // 基本資訊 (6個參數)
            `employeeId=${encodeURIComponent(employeeId)}` +
            `&employeeName=${encodeURIComponent(employeeName)}` +
            `&idNumber=${encodeURIComponent(idNumber)}` +                    // ⭐ 新增
            `&employeeType=${encodeURIComponent(employeeType)}` +            // ⭐ 新增
            `&salaryType=${encodeURIComponent(salaryType)}` +                // ⭐ 新增
            `&baseSalary=${encodeURIComponent(baseSalary)}` +
            
            // 固定津貼 (6個參數)
            `&positionAllowance=${encodeURIComponent(positionAllowance)}` +
            `&mealAllowance=${encodeURIComponent(mealAllowance)}` +
            `&transportAllowance=${encodeURIComponent(transportAllowance)}` +
            `&attendanceBonus=${encodeURIComponent(attendanceBonus)}` +
            `&performanceBonus=${encodeURIComponent(performanceBonus)}` +
            `&otherAllowances=${encodeURIComponent(otherAllowances)}` +
            
            // 銀行資訊 (4個參數)
            `&bankCode=${encodeURIComponent(bankCode)}` +
            `&bankAccount=${encodeURIComponent(bankAccount)}` +
            `&hireDate=${encodeURIComponent(hireDate)}` +
            `&paymentDay=${encodeURIComponent(paymentDay)}` +
            
            // 法定扣款 (6個參數)
            `&pensionSelfRate=${encodeURIComponent(pensionSelfRate)}` +
            `&laborFee=${encodeURIComponent(laborFee)}` +
            `&healthFee=${encodeURIComponent(healthFee)}` +
            `&employmentFee=${encodeURIComponent(employmentFee)}` +
            `&pensionSelf=${encodeURIComponent(pensionSelf)}` +
            `&incomeTax=${encodeURIComponent(incomeTax)}` +
            
            // 其他扣款 (4個參數)
            `&welfareFee=${encodeURIComponent(welfareFee)}` +
            `&dormitoryFee=${encodeURIComponent(dormitoryFee)}` +
            `&groupInsurance=${encodeURIComponent(groupInsurance)}` +
            `&otherDeductions=${encodeURIComponent(otherDeductions)}` +
            
            // 備註
            `&note=${encodeURIComponent(note)}`;
        
        console.log('📤 送出參數:', queryString);
        
        const res = await callApifetch(`setEmployeeSalaryTW&${queryString}`);
        
        if (res.ok) {
            showNotification(t('SALARY_SAVE_SUCCESS'), 'success');
            e.target.reset();
            
            // 重置所有輸入欄位為 0
            const resetFields = [
                'config-position-allowance',
                'config-meal-allowance',
                'config-transport-allowance',
                'config-attendance-bonus',
                'config-performance-bonus',
                'config-other-allowances',
                'config-welfare-fee',
                'config-dormitory-fee',
                'config-group-insurance',
                'config-other-deductions',
                'config-labor-fee',
                'config-health-fee',
                'config-employment-fee',
                'config-pension-self',
                'config-income-tax',
                'config-pension-rate'
            ];
            
            resetFields.forEach(id => {
                const el = document.getElementById(id);
                if (el) el.value = '0';
            });
            
            // 重置試算預覽
            if (typeof setCalculatedValues === 'function') {
                setCalculatedValues(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
            }
        } else {
            showNotification(t('SALARY_SAVE_FAILED') + ': ' + (res.msg || res.message || t('UNKNOWN_ERROR')), 'error');

        }
        
    } catch (error) {
        console.error('❌ 設定薪資失敗:', error);
        showNotification(t('SALARY_SAVE_ERROR'), 'error');
    }
}
/**
 * ✅ 處理薪資計算
 */
async function handleSalaryCalculation() {
    const employeeIdEl = document.getElementById('calc-employee-id');
    const yearMonthEl = document.getElementById('calc-year-month');
    const resultEl = document.getElementById('salary-calculation-result');
    
    if (!employeeIdEl || !yearMonthEl || !resultEl) return;
    
    const employeeId = employeeIdEl.value.trim();
    const yearMonth = yearMonthEl.value;
    
    if (!employeeId || !yearMonth) {
        showNotification(t('SALARY_INPUT_EMPLOYEE_MONTH'), 'error');
        return;
    }
    
    try {
        showNotification(t('SALARY_CALCULATING'), 'info');
        
        const res = await callApifetch(`calculateMonthlySalary&employeeId=${encodeURIComponent(employeeId)}&yearMonth=${encodeURIComponent(yearMonth)}`);
        
        if (res.ok && res.data) {
            displaySalaryCalculation(res.data, resultEl);
            resultEl.style.display = 'block';
            showNotification(t('SALARY_CALC_SUCCESS'), 'success');
            
            if (confirm('是否儲存此薪資單？')) {
                await saveSalaryRecord(res.data);
            }
        } else {
            showNotification(t('SALARY_CALC_FAILED') + ': ' + (res.msg || t('UNKNOWN_ERROR')), 'error');

        }
        
    } catch (error) {
        console.error('❌ 計算薪資失敗:', error);
        showNotification(t('SALARY_CALC_ERROR'), 'error');
    }
}

/**
 * ✅ 顯示薪資計算結果（支援月薪/時薪區分 + 國定假日完整版）
 */
function displaySalaryCalculation(data, container) {
    if (!container) return;
    
    // 計算扣款總額
    const totalDeductions = 
        (parseFloat(data.laborFee) || 0) + 
        (parseFloat(data.healthFee) || 0) + 
        (parseFloat(data.employmentFee) || 0) + 
        (parseFloat(data.pensionSelf) || 0) + 
        (parseFloat(data.incomeTax) || 0) + 
        (parseFloat(data.leaveDeduction) || 0) +
        (parseFloat(data.earlyLeaveDeduction || data['早退扣款']) || 0) +
        (parseFloat(data.welfareFee) || 0) +
        (parseFloat(data.dormitoryFee) || 0) +
        (parseFloat(data.groupInsurance) || 0) +
        (parseFloat(data.otherDeductions) || 0);
    
    const isHourly = data.salaryType === '時薪';
    
    // ⭐⭐⭐ 修正：讀取四種加班費 + 國定假日出勤薪資
    const weekdayOvertimePay = parseFloat(data.weekdayOvertimePay) || 0;
    const restdayOvertimePay = parseFloat(data.restdayOvertimePay) || 0;
    const holidayWorkPay = parseFloat(data.holidayWorkPay) || 0;          // ⭐ 新增
    const holidayOvertimePay = parseFloat(data.holidayOvertimePay) || 0;
    const totalOvertimeHours = parseFloat(data.totalOvertimeHours) || 0;
    
    // ⭐⭐⭐ 修正：病假/事假明細（改用時數）
    const sickLeaveHours = parseFloat(data.sickLeaveHours) || 0;  // ⭐ 改名
    const sickLeaveDeduction = parseFloat(data.sickLeaveDeduction) || 0;
    const personalLeaveHours = parseFloat(data.personalLeaveHours) || 0;  // ⭐ 改名
    const personalLeaveDeduction = parseFloat(data.personalLeaveDeduction) || 0;
    // ⭐⭐⭐ 新增：早退扣款變數
    const earlyLeaveDeduction = parseFloat(data.earlyLeaveDeduction || data['早退扣款']) || 0;
    container.innerHTML = `
        <div class="calculation-card">
            <h3 class="text-xl font-bold mb-4">
                ${data.employeeName || '--'} - ${data.yearMonth || '--'} 薪資計算結果
                <span class="ml-2 px-3 py-1 text-sm rounded-full ${isHourly ? 'bg-purple-100 text-purple-800' : 'bg-blue-100 text-blue-800'}">
                    ${data.salaryType || '月薪'}
                </span>
            </h3>
            
            <!-- 三大金額卡片 -->
            <div class="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
                <div class="info-card" style="background: rgba(34, 197, 94, 0.1);">
                    <div class="info-label">應發總額</div>
                    <div class="info-value" style="color: #22c55e;">${formatCurrency(data.grossSalary)}</div>
                </div>
                <div class="info-card" style="background: rgba(239, 68, 68, 0.1);">
                    <div class="info-label">扣款總額</div>
                    <div class="info-value" style="color: #ef4444;">${formatCurrency(totalDeductions)}</div>
                </div>
                <div class="info-card" style="background: rgba(168, 85, 247, 0.1);">
                    <div class="info-label">實發金額</div>
                    <div class="info-value" style="color: #a855f7;">${formatCurrency(data.netSalary)}</div>
                </div>
            </div>
            
            <!-- ⭐ 時薪工時統計區塊 -->
            ${isHourly ? `
                <div class="bg-purple-50 dark:bg-purple-900/20 border-2 border-purple-200 dark:border-purple-700 rounded-lg p-4 mb-6">
                    <h4 class="font-semibold text-purple-800 dark:text-purple-300 mb-3">⏰ 時薪工時統計</h4>
                    <div class="grid grid-cols-3 gap-4 text-center">
                        <div>
                            <p class="text-sm text-purple-600 dark:text-purple-400">時薪</p>
                            <p class="text-2xl font-bold text-purple-800 dark:text-purple-200">$${data.hourlyRate || 0}</p>
                        </div>
                        <div>
                            <p class="text-sm text-purple-600 dark:text-purple-400">工作時數</p>
                            <p class="text-2xl font-bold text-purple-800 dark:text-purple-200">${Math.floor(data.totalWorkHours || 0)}h</p>
                        </div>
                        <div>
                            <p class="text-sm text-purple-600 dark:text-purple-400">基本薪資</p>
                            <p class="text-xl font-bold text-purple-800 dark:text-purple-200">${formatCurrency(data.baseSalary)}</p>
                            <p class="text-xs text-purple-500">(時薪 × 工時)</p>
                        </div>
                    </div>
                </div>
            ` : ''}
            
            <!-- ⭐⭐⭐ 加班統計區塊（完整版：含國定假日） -->
            ${totalOvertimeHours > 0 ? `
                <div class="bg-orange-50 dark:bg-orange-900/20 border-2 border-orange-200 dark:border-orange-700 rounded-lg p-4 mb-6">
                    <h4 class="font-semibold text-orange-800 dark:text-orange-300 mb-3">⏰ 本月加班統計</h4>
                    
                    <!-- 總時數 -->
                    <div class="text-center p-3 bg-orange-100 dark:bg-orange-800/30 rounded-lg mb-3">
                        <p class="text-sm text-orange-600 dark:text-orange-400">總加班時數</p>
                        <p class="text-3xl font-bold text-orange-800 dark:text-orange-200">${totalOvertimeHours.toFixed(1)}h</p>
                    </div>
                    
                    <!-- 分類明細 -->
                    <div class="space-y-2">
                        ${weekdayOvertimePay > 0 ? `
                            <div class="p-3 bg-blue-100 dark:bg-blue-900/30 rounded-lg border border-blue-300 dark:border-blue-700">
                                <div class="flex justify-between items-center">
                                    <div>
                                        <span class="font-semibold text-blue-800 dark:text-blue-300">平日加班</span>
                                        <span class="text-xs text-blue-600 dark:text-blue-400 ml-2">（週一～五）</span>
                                    </div>
                                    <span class="text-lg font-bold text-blue-800 dark:text-blue-200">${formatCurrency(weekdayOvertimePay)}</span>
                                </div>
                                <p class="text-xs text-blue-600 dark:text-blue-400 mt-1">前2h ×1.34 | 第3h起 ×1.67</p>
                            </div>
                        ` : ''}
                        
                        ${restdayOvertimePay > 0 ? `
                            <div class="p-3 bg-purple-100 dark:bg-purple-900/30 rounded-lg border border-purple-300 dark:border-purple-700">
                                <div class="flex justify-between items-center">
                                    <div>
                                        <span class="font-semibold text-purple-800 dark:text-purple-300">休息日加班</span>
                                        <span class="text-xs text-purple-600 dark:text-purple-400 ml-2">（週六）</span>
                                    </div>
                                    <span class="text-lg font-bold text-purple-800 dark:text-purple-200">${formatCurrency(restdayOvertimePay)}</span>
                                </div>
                                <p class="text-xs text-purple-600 dark:text-purple-400 mt-1">前2h ×1.34 | 3-8h ×1.67 | 9h起 ×2.67</p>
                            </div>
                        ` : ''}
                        
                        ${holidayWorkPay > 0 || holidayOvertimePay > 0 ? `
                            <div class="p-3 bg-red-100 dark:bg-red-900/30 rounded-lg border border-red-300 dark:border-red-700">
                                <div class="flex justify-between items-center mb-2">
                                    <div>
                                        <span class="font-semibold text-red-800 dark:text-red-300">🎊 國定假日出勤</span>
                                    </div>
                                    <span class="text-lg font-bold text-red-800 dark:text-red-200">${formatCurrency(holidayWorkPay + holidayOvertimePay)}</span>
                                </div>
                                
                                <!-- ⭐⭐⭐ 分開顯示正常薪資與加班費 -->
                                <div class="text-xs space-y-1 mt-2 border-t border-red-700/30 pt-2">
                                    ${holidayWorkPay > 0 ? `
                                        <div class="flex justify-between">
                                            <span class="text-red-600 dark:text-red-400">正常出勤薪資 ×1.0</span>
                                            <span class="font-mono text-red-700 dark:text-red-300">${formatCurrency(holidayWorkPay)}</span>
                                        </div>
                                    ` : ''}
                                    ${holidayOvertimePay > 0 ? `
                                        <div class="flex justify-between">
                                            <span class="text-red-600 dark:text-red-400">加班費 ×2.0</span>
                                            <span class="font-mono text-red-700 dark:text-red-300">${formatCurrency(holidayOvertimePay)}</span>
                                        </div>
                                    ` : ''}
                                </div>
                            </div>
                        ` : ''}
                    </div>
                </div>
            ` : ''}
            
            <!-- 應發項目 vs 扣款項目 -->
            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                <!-- 👈 應發項目 -->
                <div class="calculation-detail">
                    <h4 class="font-semibold mb-3 text-green-400">💰 應發項目</h4>
                    
                    ${isHourly ? `
                        <div class="calculation-row">
                            <span>時薪</span>
                            <span class="font-mono">$${data.hourlyRate || 0}</span>
                        </div>
                        <div class="calculation-row">
                            <span>工作時數</span>
                            <span class="font-mono">${(data.totalWorkHours || 0).toFixed(1)}h</span>
                        </div>
                        <div class="calculation-row">
                            <span>基本薪資 (時薪×工時)</span>
                            <span class="font-mono">${formatCurrency(data.baseSalary)}</span>
                        </div>
                    ` : `
                        <div class="calculation-row">
                            <span>基本薪資</span>
                            <span class="font-mono">${formatCurrency(data.baseSalary)}</span>
                        </div>
                    `}
                    
                    <div class="calculation-row">
                        <span>職務加給</span>
                        <span class="font-mono">${formatCurrency(data.positionAllowance || 0)}</span>
                    </div>
                    <div class="calculation-row">
                        <span>伙食費</span>
                        <span class="font-mono">${formatCurrency(data.mealAllowance || 0)}</span>
                    </div>
                    <div class="calculation-row">
                        <span>交通補助</span>
                        <span class="font-mono">${formatCurrency(data.transportAllowance || 0)}</span>
                    </div>
                    <div class="calculation-row">
                        <span>全勤獎金</span>
                        <span class="font-mono">${formatCurrency(data.attendanceBonus || 0)}</span>
                    </div>
                    <div class="calculation-row">
                        <span>業績獎金</span>
                        <span class="font-mono">${formatCurrency(data.performanceBonus || 0)}</span>
                    </div>
                    
                    ${weekdayOvertimePay > 0 ? `
                        <div class="calculation-row">
                            <span>平日加班費</span>
                            <span class="font-mono">${formatCurrency(weekdayOvertimePay)}</span>
                        </div>
                    ` : ''}
                    
                    ${restdayOvertimePay > 0 ? `
                        <div class="calculation-row">
                            <span>休息日加班費</span>
                            <span class="font-mono">${formatCurrency(restdayOvertimePay)}</span>
                        </div>
                    ` : ''}
                    
                    ${holidayWorkPay > 0 ? `
                        <div class="calculation-row">
                            <span>國定假日出勤薪資</span>
                            <span class="font-mono">${formatCurrency(holidayWorkPay)}</span>
                        </div>
                    ` : ''}
                    
                    ${holidayOvertimePay > 0 ? `
                        <div class="calculation-row">
                            <span>國定假日加班費</span>
                            <span class="font-mono">${formatCurrency(holidayOvertimePay)}</span>
                        </div>
                    ` : ''}
                    
                    <div class="calculation-row total">
                        <span>應發總額</span>
                        <span>${formatCurrency(data.grossSalary)}</span>
                    </div>
                </div>
                
                <!-- 👈 扣款項目 -->
                <div class="calculation-detail">
                    <h4 class="font-semibold mb-3 text-red-400">⚠️ 扣款項目</h4>
                    
                    <div class="calculation-row">
                        <span>勞保費</span>
                        <span class="font-mono">${formatCurrency(data.laborFee)}</span>
                    </div>
                    <div class="calculation-row">
                        <span>健保費</span>
                        <span class="font-mono">${formatCurrency(data.healthFee)}</span>
                    </div>
                    <div class="calculation-row">
                        <span>就業保險費</span>
                        <span class="font-mono">${formatCurrency(data.employmentFee)}</span>
                    </div>
                    <div class="calculation-row">
                        <span>勞退自提 (${data.pensionSelfRate || 0}%)</span>
                        <span class="font-mono">${formatCurrency(data.pensionSelf)}</span>
                    </div>
                    <div class="calculation-row">
                        <span>所得稅</span>
                        <span class="font-mono">${formatCurrency(data.incomeTax)}</span>
                    </div>
                    
                    ${!isHourly && data.leaveDeduction > 0 ? `
                        <div class="calculation-row">
                            <span>請假扣款</span>
                            <span class="font-mono">${formatCurrency(data.leaveDeduction)}</span>
                        </div>

                        <!-- ⭐⭐⭐ 病假/事假明細 -->
                        ${sickLeaveHours > 0 || personalLeaveHours > 0 ? `
                            <div class="p-2 bg-yellow-900/20 rounded-lg mt-2 mb-2 border border-yellow-700/30">
                                <div class="text-xs space-y-1">
                                    ${sickLeaveHours > 0 ? `
                                        <div class="flex justify-between">
                                            <span class="text-yellow-300">病假 ${sickLeaveHours} 小時 (半薪)</span>
                                            <span class="font-mono text-yellow-200">${formatCurrency(sickLeaveDeduction)}</span>
                                        </div>
                                    ` : ''}
                                    ${personalLeaveHours > 0 ? `
                                        <div class="flex justify-between">
                                            <span class="text-yellow-300">事假 ${personalLeaveHours} 小時 (全薪)</span>
                                            <span class="font-mono text-yellow-200">${formatCurrency(personalLeaveDeduction)}</span>
                                        </div>
                                    ` : ''}
                                </div>
                            </div>
                        ` : ''}
                    ` : ''}
                    
                    <!-- ⭐⭐⭐ 新增：早退扣款 -->
                    ${!isHourly && earlyLeaveDeduction > 0 ? `
                        <div class="calculation-row">
                            <span>早退扣款</span>
                            <span class="font-mono">${formatCurrency(earlyLeaveDeduction)}</span>
                        </div>
                    ` : ''}
                    <div class="calculation-row">
                        <span>福利金</span>
                        <span class="font-mono">${formatCurrency(data.welfareFee || 0)}</span>
                    </div>
                    <div class="calculation-row">
                        <span>宿舍費用</span>
                        <span class="font-mono">${formatCurrency(data.dormitoryFee || 0)}</span>
                    </div>
                    <div class="calculation-row">
                        <span>團保費用</span>
                        <span class="font-mono">${formatCurrency(data.groupInsurance || 0)}</span>
                    </div>
                    <div class="calculation-row">
                        <span>其他扣款</span>
                        <span class="font-mono">${formatCurrency(data.otherDeductions || 0)}</span>
                    </div>
                    
                    <div class="calculation-row total">
                        <span>實發金額</span>
                        <span>${formatCurrency(data.netSalary)}</span>
                    </div>
                </div>
            </div>
        </div>
    `;
}
/**
 * ✅ 儲存薪資記錄（修正版 - 包含所有必要欄位）
 */
async function saveSalaryRecord(data) {
    try {
        showNotification(t('SALARY_SAVING_RECORD'), 'info');
        
        // ⭐⭐⭐ 修正：加入完整的欄位（特別是 salaryType, hourlyRate, totalWorkHours）
        const queryString = 
            `employeeId=${encodeURIComponent(data.employeeId)}` +
            `&employeeName=${encodeURIComponent(data.employeeName)}` +
            `&yearMonth=${encodeURIComponent(data.yearMonth)}` +
            
            // ⭐ 新增：薪資類型相關欄位
            `&salaryType=${encodeURIComponent(data.salaryType || '月薪')}` +
            `&hourlyRate=${encodeURIComponent(data.hourlyRate || 0)}` +
            `&totalWorkHours=${encodeURIComponent(data.totalWorkHours || 0)}` +
            `&totalOvertimeHours=${encodeURIComponent(data.totalOvertimeHours || 0)}` +
            
            // 應發項目
            `&baseSalary=${encodeURIComponent(data.baseSalary)}` +
            `&positionAllowance=${encodeURIComponent(data.positionAllowance || 0)}` +
            `&mealAllowance=${encodeURIComponent(data.mealAllowance || 0)}` +
            `&transportAllowance=${encodeURIComponent(data.transportAllowance || 0)}` +
            `&attendanceBonus=${encodeURIComponent(data.attendanceBonus || 0)}` +
            `&performanceBonus=${encodeURIComponent(data.performanceBonus || 0)}` +
            `&otherAllowances=${encodeURIComponent(data.otherAllowances || 0)}` +
            
            // ⭐ 修正：加班費（三種）
            `&weekdayOvertimePay=${encodeURIComponent(data.weekdayOvertimePay || 0)}` +
            `&restdayOvertimePay=${encodeURIComponent(data.restdayOvertimePay || 0)}` +
            `&holidayOvertimePay=${encodeURIComponent(data.holidayOvertimePay || 0)}` +
            
            // 法定扣款
            `&laborFee=${encodeURIComponent(data.laborFee || 0)}` +
            `&healthFee=${encodeURIComponent(data.healthFee || 0)}` +
            `&employmentFee=${encodeURIComponent(data.employmentFee || 0)}` +
            `&pensionSelf=${encodeURIComponent(data.pensionSelf || 0)}` +
            `&pensionSelfRate=${encodeURIComponent(data.pensionSelfRate || 0)}` +
            `&incomeTax=${encodeURIComponent(data.incomeTax || 0)}` +
            
            // 其他扣款
            `&leaveDeduction=${encodeURIComponent(data.leaveDeduction || 0)}` +
            `&welfareFee=${encodeURIComponent(data.welfareFee || 0)}` +
            `&dormitoryFee=${encodeURIComponent(data.dormitoryFee || 0)}` +
            `&groupInsurance=${encodeURIComponent(data.groupInsurance || 0)}` +
            `&otherDeductions=${encodeURIComponent(data.otherDeductions || 0)}` +
            
            // 總計
            `&grossSalary=${encodeURIComponent(data.grossSalary)}` +
            `&netSalary=${encodeURIComponent(data.netSalary)}` +
            
            // 銀行資訊
            `&bankCode=${encodeURIComponent(data.bankCode || '')}` +
            `&bankAccount=${encodeURIComponent(data.bankAccount || '')}` +
            
            // 狀態
            `&status=${encodeURIComponent(data.status || '已計算')}` +
            `&note=${encodeURIComponent(data.note || '')}`;
        
        console.log('📤 儲存薪資記錄，包含參數:', {
            employeeId: data.employeeId,
            yearMonth: data.yearMonth,
            salaryType: data.salaryType,  // ⭐ 確認有傳遞
            hourlyRate: data.hourlyRate,
            totalWorkHours: data.totalWorkHours
        });
        
        const res = await callApifetch(`saveMonthlySalary&${queryString}`);
        
        if (res.ok) {
            showNotification(t('SALARY_RECORD_SAVED'), 'success');
        } else {
            showNotification(t('SALARY_SAVE_FAILED') + ': ' + (res.msg || t('UNKNOWN_ERROR')), 'error');
        }
        
    } catch (error) {
        console.error('❌ 儲存薪資單失敗:', error);
        showNotification(t('SALARY_SAVE_ERROR'), 'error');
    }
}

/**
 * 載入所有員工薪資列表
 */
async function loadAllEmployeeSalaryFromList() {
    const yearMonthEl = document.getElementById('filter-year-month-list');
    const loadingEl = document.getElementById('all-salary-loading-list');
    const listEl = document.getElementById('all-salary-list-content');
    
    if (!yearMonthEl || !loadingEl || !listEl) return;
    
    const yearMonth = yearMonthEl.value;
    
    if (!yearMonth) {
        showNotification(t('SALARY_SELECT_MONTH'), 'error');
        return;
    }
    
    try {
        loadingEl.style.display = 'block';
        listEl.innerHTML = '';
        
        const res = await callApifetch(`getAllMonthlySalary&yearMonth=${encodeURIComponent(yearMonth)}`);
        
        loadingEl.style.display = 'none';
        
        if (res.ok && res.data && res.data.length > 0) {
            res.data.forEach(salary => {
                const item = createAllSalaryItem(salary);
                listEl.appendChild(item);
            });
        } else {
            listEl.innerHTML = '<p class="text-center text-gray-400 py-8">尚無薪資記錄</p>';
        }
        
    } catch (error) {
        console.error('❌ 載入薪資列表失敗:', error);
        loadingEl.style.display = 'none';
        listEl.innerHTML = '<p class="text-center text-red-400 py-8">載入失敗</p>';
    }
}

/**
 * 建立所有員工薪資項目
 */
function createAllSalaryItem(salary) {
    const div = document.createElement('div');
    div.className = 'feature-box flex justify-between items-center hover:bg-white/10 transition cursor-pointer';
    
    div.innerHTML = `
        <div>
            <div class="font-semibold text-lg">
                ${salary['員工姓名'] || '--'} <span class="text-gray-400 text-sm">(${salary['員工ID'] || '--'})</span>
            </div>
            <div class="text-sm text-gray-400 mt-1">
                ${salary['年月'] || '--'} | ${salary['狀態'] || '--'}
            </div>
        </div>
        <div class="text-right">
            <div class="text-2xl font-bold text-green-400">
                ${formatCurrency(salary['實發金額'])}
            </div>
            <div class="text-xs text-gray-400 mt-1">
                ${getBankName(salary['銀行代碼'])} ${salary['銀行帳號'] || '--'}
            </div>
        </div>
    `;
    
    return div;
}

// ==================== 即時匯率換算 ====================

const EXCHANGE_RATE_TTL = 3600000; // 1 小時快取

async function fetchExchangeRates() {
    const now = Date.now();
    const cached = localStorage.getItem('_ex_rates');
    const expiry = parseInt(localStorage.getItem('_ex_rates_exp') || '0');
    if (cached && now < expiry) return JSON.parse(cached);

    try {
        const res = await fetch('https://open.er-api.com/v6/latest/TWD');
        const data = await res.json();
        if (data.result === 'success') {
            localStorage.setItem('_ex_rates', JSON.stringify(data.rates));
            localStorage.setItem('_ex_rates_exp', String(now + EXCHANGE_RATE_TTL));
            localStorage.setItem('_ex_rates_time', new Date().toLocaleTimeString('zh-TW', { timeZone: 'Asia/Taipei' }));
            return data.rates;
        }
    } catch (e) {
        console.warn('匯率 API 無法連線，使用備用匯率', e);
    }
    // 備用匯率（近似值）
    return { USD: 0.031, JPY: 4.65, EUR: 0.029, KRW: 41.5, HKD: 0.242, VND: 779, THB: 1.08 };
}

const CURRENCY_SYMBOLS = { USD: '$', JPY: '¥', EUR: '€', KRW: '₩', HKD: 'HK$', VND: '₫', THB: '฿' };
const ZERO_DECIMAL_CURRENCIES = new Set(['JPY', 'KRW', 'VND']);

function refreshCurrencyDisplay() {
    const select = document.getElementById('currency-select');
    if (!select || !_currentNetSalary) return;

    const targetCurrency = select.value;
    const rates = JSON.parse(localStorage.getItem('_ex_rates') || 'null');

    if (rates) {
        applyCurrencyConversion(_currentNetSalary, targetCurrency, rates);
    } else {
        fetchExchangeRates().then(r => applyCurrencyConversion(_currentNetSalary, targetCurrency, r));
    }
}

function applyCurrencyConversion(twdAmount, targetCurrency, rates) {
    const rate = rates[targetCurrency];
    if (!rate) return;

    const converted = twdAmount * rate;
    const decimals = ZERO_DECIMAL_CURRENCIES.has(targetCurrency) ? 0 : 2;
    const formatted = new Intl.NumberFormat('zh-TW', {
        minimumFractionDigits: 0,
        maximumFractionDigits: decimals
    }).format(converted);

    const symbol = CURRENCY_SYMBOLS[targetCurrency] || targetCurrency + ' ';
    const resultEl = document.getElementById('currency-result');
    if (resultEl) resultEl.textContent = symbol + formatted;

    const rateEl = document.getElementById('exchange-rate-display');
    if (rateEl) {
        const rateStr = rate < 0.01 ? rate.toFixed(4) : rate < 1 ? rate.toFixed(4) : rate.toFixed(2);
        rateEl.textContent = `1 TWD = ${rateStr} ${targetCurrency}`;
    }

    const timeEl = document.getElementById('exchange-rate-time');
    if (timeEl) {
        const t = localStorage.getItem('_ex_rates_time');
        timeEl.textContent = t ? `更新時間 ${t}` : '';
    }
}

// ==================== 工具函數 ====================

/**
 * 格式化貨幣
 */
function formatCurrency(amount) {
    if (amount === null || amount === undefined || isNaN(amount)) return '$0';
    const num = parseFloat(amount);
    if (isNaN(num)) return '$0';
    return '$' + num.toLocaleString('zh-TW', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

/**
 * 取得銀行名稱
 */
function getBankName(code) {
    const banks = {
        // 公股銀行
        "004": "臺灣銀行",
        "005": "臺灣土地銀行",
        "006": "合作金庫商業銀行",
        "007": "第一商業銀行",
        "008": "華南商業銀行",
        "009": "彰化商業銀行",
        "011": "上海商業儲蓄銀行",
        "012": "台北富邦商業銀行",
        "013": "國泰世華商業銀行",
        "016": "高雄銀行",
        "017": "兆豐國際商業銀行",
        "050": "臺灣中小企業銀行",
        
        // 民營銀行
        "103": "臺灣新光商業銀行",
        "108": "陽信商業銀行",
        "118": "板信商業銀行",
        "147": "三信商業銀行",
        "803": "聯邦商業銀行",
        "805": "遠東國際商業銀行",
        "806": "元大商業銀行",
        "807": "永豐商業銀行",
        "808": "玉山商業銀行",
        "809": "凱基商業銀行",
        "810": "星展（台灣）商業銀行",
        "812": "台新國際商業銀行",
        "816": "安泰商業銀行",
        "822": "中國信託商業銀行",
        "826": "樂天國際商業銀行",
        
        // 外商銀行
        "052": "渣打國際商業銀行",
        "081": "匯豐（台灣）商業銀行",
        "101": "瑞興商業銀行",
        "102": "華泰商業銀行",
        "815": "日盛國際商業銀行",
        "824": "連線商業銀行",
        
        // 郵局
        "700": "中華郵政"
    };
    
    return banks[code] || "未知銀行";
}

// ✅ 新方法：直接用 calculateMonthlySalary（跟時薪計算一樣）
// salaryData：由呼叫端傳入已取得的薪資資料，避免第三次呼叫 calculateMonthlySalary
async function loadAttendanceDetails(yearMonth, salaryData = null) {
    try {
        const detailsSection = document.getElementById('attendance-details-section');
        if (!detailsSection) return;

        let data = salaryData;

        if (!data) {
            // 僅在沒有傳入資料時才呼叫 API（相容舊呼叫點）
            const session = await getSession();
            if (!session.ok || !session.user) {
                detailsSection.style.display = 'none';
                return;
            }
            const employeeId = session.user.userId;
            const res = await callApifetch(`calculateMonthlySalary&employeeId=${encodeURIComponent(employeeId)}&yearMonth=${encodeURIComponent(yearMonth)}`);
            if (!res.ok || !res.data) {
                detailsSection.style.display = 'none';
                return;
            }
            data = res.data;
        }

        const isHourly = (data.salaryType || '月薪') === '時薪';
        detailsSection.style.display = 'block';

        if (isHourly) displayWorkHoursFromCalculation(data);
        if (data.totalOvertimeHours > 0) displayOvertimeFromCalculation(data);

    } catch (error) {
        console.error('❌ 載入出勤明細失敗:', error);
    }
}

function displayWorkHoursFromCalculation(data) {
    const detailsSection = document.getElementById('attendance-details-section');
    if (!detailsSection) return;
    
    const oldCard = document.getElementById('work-hours-card');
    if (oldCard) oldCard.remove();
    
    const workHoursCard = document.createElement('div');
    workHoursCard.id = 'work-hours-card';
    workHoursCard.className = 'feature-box bg-purple-900/20 border-purple-700 mb-4';
    
    // ⭐⭐⭐ 修正：保留小數位數
    const totalWorkHours = parseFloat(data.totalWorkHours || 0).toFixed(1);
    const hourlyRate = data.hourlyRate || 0;
    const baseSalary = data.baseSalary || 0;
    
    workHoursCard.innerHTML = `
      <h4 class="font-semibold mb-3 text-purple-400">本月工作時數統計</h4>
      
      <div class="grid grid-cols-3 gap-4 mb-4">
        <div class="text-center p-3 bg-purple-800/20 rounded-lg">
          <p class="text-sm text-purple-300 mb-1">時薪</p>
          <p class="text-2xl font-bold text-purple-200">$${hourlyRate}</p>
        </div>
        <div class="text-center p-3 bg-purple-800/20 rounded-lg">
          <p class="text-sm text-purple-300 mb-1">總工作時數</p>
          <p class="text-2xl font-bold text-purple-200">${totalWorkHours}h</p>
        </div>
        <div class="text-center p-3 bg-purple-800/20 rounded-lg">
          <p class="text-sm text-purple-300 mb-1">基本薪資</p>
          <p class="text-2xl font-bold text-purple-200">${formatCurrency(baseSalary)}</p>
          <p class="text-xs text-purple-400 mt-1">(時薪 × 工時)</p>
        </div>
      </div>
      
      <div class="p-3 bg-purple-800/10 rounded-lg text-sm text-purple-300">
        💡 工作時數已包含在薪資計算中
      </div>
    `;
    
    detailsSection.insertBefore(workHoursCard, detailsSection.firstChild);
  }


  function displayOvertimeFromCalculation(data) {
    const detailsSection = document.getElementById('attendance-details-section');
    if (!detailsSection) return;
    
    const oldCard = document.getElementById('overtime-card');
    if (oldCard) oldCard.remove();
    
    const overtimeCard = document.createElement('div');
    overtimeCard.id = 'overtime-card';
    overtimeCard.className = 'feature-box bg-orange-900/20 border-orange-700 mt-4';
    
    // ⭐⭐⭐ 修正：正確讀取四種加班費 + 國定假日出勤薪資
    const totalOvertimeHours = Math.floor(data.totalOvertimeHours || 0);
    const weekdayOvertimePay = parseFloat(data.weekdayOvertimePay) || 0;
    const restdayOvertimePay = parseFloat(data.restdayOvertimePay) || 0;
    const holidayWorkPay = parseFloat(data.holidayWorkPay) || 0;          // ⭐ 新增這行
    const holidayOvertimePay = parseFloat(data.holidayOvertimePay) || 0;
    
    console.log('🔍 displayOvertimeFromCalculation 讀取的加班費:');
    console.log('   平日:', weekdayOvertimePay);
    console.log('   休息日:', restdayOvertimePay);
    console.log('   國定假日出勤薪資:', holidayWorkPay);               // ⭐ 新增這行
    console.log('   國定假日加班費:', holidayOvertimePay);
    
    overtimeCard.innerHTML = `
        <h4 class="font-semibold mb-3 text-orange-400">⏰ 本月加班統計</h4>
        
        <!-- 總時數 -->
        <div class="text-center p-3 bg-orange-800/20 rounded-lg mb-3">
            <p class="text-sm text-orange-300 mb-1">總加班時數</p>
            <p class="text-3xl font-bold text-orange-200">${totalOvertimeHours}h</p>
        </div>
        
        <!-- ⭐⭐⭐ 關鍵修正：使用 space-y-2 垂直排列 -->
        <div class="space-y-2 mb-3">
            ${weekdayOvertimePay > 0 ? `
                <div class="p-3 bg-blue-100 dark:bg-blue-900/30 rounded-lg border border-blue-300 dark:border-blue-700">
                    <div class="flex justify-between items-center">
                        <div>
                            <span class="font-semibold text-blue-800 dark:text-blue-300">平日加班</span>
                            <span class="text-xs text-blue-600 dark:text-blue-400 ml-2">（週一～五）</span>
                        </div>
                        <span class="text-lg font-bold text-blue-800 dark:text-blue-200">${formatCurrency(weekdayOvertimePay)}</span>
                    </div>
                    <p class="text-xs text-blue-600 dark:text-blue-400 mt-1">前2h ×1.34 | 第3h起 ×1.67</p>
                </div>
            ` : ''}
            
            ${restdayOvertimePay > 0 ? `
                <div class="p-3 bg-purple-100 dark:bg-purple-900/30 rounded-lg border border-purple-300 dark:border-purple-700">
                    <div class="flex justify-between items-center">
                        <div>
                            <span class="font-semibold text-purple-800 dark:text-purple-300">休息日加班</span>
                            <span class="text-xs text-purple-600 dark:text-purple-400 ml-2">（週六）</span>
                        </div>
                        <span class="text-lg font-bold text-purple-800 dark:text-purple-200">${formatCurrency(restdayOvertimePay)}</span>
                    </div>
                    <p class="text-xs text-purple-600 dark:text-purple-400 mt-1">前2h ×1.34 | 3-8h ×1.67 | 9h起 ×2.67</p>
                </div>
            ` : ''}
            
            ${holidayWorkPay > 0 || holidayOvertimePay > 0 ? `
                <div class="p-3 bg-red-100 dark:bg-red-900/30 rounded-lg border border-red-300 dark:border-red-700">
                    <div class="flex justify-between items-center mb-2">
                        <div>
                            <span class="font-semibold text-red-800 dark:text-red-300">國定假日出勤</span>
                        </div>
                        <span class="text-lg font-bold text-red-800 dark:text-red-200">${formatCurrency(holidayWorkPay + holidayOvertimePay)}</span>
                    </div>
                    
                    <!-- ⭐⭐⭐ 分開顯示正常薪資與加班費 -->
                    <div class="text-xs space-y-1 mt-2 border-t border-red-700/30 pt-2">
                        ${holidayWorkPay > 0 ? `
                            <div class="flex justify-between">
                                <span class="text-red-600 dark:text-red-400">正常出勤薪資 ×1.0</span>
                                <span class="font-mono text-red-700 dark:text-red-300">${formatCurrency(holidayWorkPay)}</span>
                            </div>
                        ` : ''}
                        ${holidayOvertimePay > 0 ? `
                            <div class="flex justify-between">
                                <span class="text-red-600 dark:text-red-400">加班費 ×2.0</span>
                                <span class="font-mono text-red-700 dark:text-red-300">${formatCurrency(holidayOvertimePay)}</span>
                            </div>
                        ` : ''}
                    </div>
                </div>
            ` : ''}
        </div>
    `;
    
    detailsSection.appendChild(overtimeCard);
}
/**
 * ✅ 載入打卡記錄
 */
async function loadPunchRecords(yearMonth) {
    const loadingEl = document.getElementById('punch-records-loading');
    const emptyEl = document.getElementById('punch-records-empty');
    const listEl = document.getElementById('punch-records-list');
    const totalEl = document.getElementById('total-work-hours');
    
    if (!loadingEl || !emptyEl || !listEl) return;
    
    try {
        loadingEl.style.display = 'block';
        emptyEl.style.display = 'none';
        listEl.innerHTML = '';
        
        // 呼叫後端 API 取得打卡記錄
        const res = await callApifetch(`getEmployeeMonthlyAttendance&yearMonth=${yearMonth}`);
        
        loadingEl.style.display = 'none';
        
        if (res.ok && res.records && res.records.length > 0) {
            let totalHours = 0;
            
            res.records.forEach(record => {
                const item = document.createElement('div');
                item.className = 'flex justify-between items-center p-2 bg-white/5 rounded';
                
                const workHours = record.workHours || 0;
                totalHours += workHours;
                
                item.innerHTML = `
                    <div>
                        <span class="font-semibold">${record.date}</span>
                        <span class="text-sm text-gray-400 ml-2">
                            ${record.punchIn || '--'} ~ ${record.punchOut || '--'}
                        </span>
                    </div>
                    <div class="text-right">
                        <span class="font-mono text-blue-400">${workHours.toFixed(1)}h</span>
                    </div>
                `;
                
                listEl.appendChild(item);
            });
            
            if (totalEl) {
                totalEl.textContent = `${totalHours.toFixed(1)} 小時`;
            }
            
        } else {
            emptyEl.style.display = 'block';
            if (totalEl) totalEl.textContent = '0.0 小時';
        }
        
    } catch (error) {
        console.error('❌ 載入打卡記錄失敗:', error);
        loadingEl.style.display = 'none';
        emptyEl.style.display = 'block';
    }
}

/**
 * ✅ 載入加班記錄
 */
async function loadOvertimeRecords(yearMonth) {
    const loadingEl = document.getElementById('overtime-records-loading');
    const emptyEl = document.getElementById('overtime-records-empty');
    const listEl = document.getElementById('overtime-records-list');
    const totalEl = document.getElementById('total-overtime-hours');
    
    if (!loadingEl || !emptyEl || !listEl) return;
    
    try {
        loadingEl.style.display = 'block';
        emptyEl.style.display = 'none';
        listEl.innerHTML = '';
        
        // 呼叫後端 API 取得加班記錄
        const res = await callApifetch(`getEmployeeMonthlyOvertime&yearMonth=${yearMonth}`);
        
        loadingEl.style.display = 'none';
        
        if (res.ok && res.records && res.records.length > 0) {
            let totalHours = 0;
            
            res.records.forEach(record => {
                const item = document.createElement('div');
                item.className = 'flex justify-between items-center p-2 bg-white/5 rounded';
                
                const hours = record.hours || 0;
                totalHours += hours;
                
                item.innerHTML = `
                    <div>
                        <span class="font-semibold">${record.date}</span>
                        <span class="text-sm text-gray-400 ml-2">已核准</span>
                    </div>
                    <div class="text-right">
                        <span class="font-mono text-orange-400">${hours.toFixed(1)}h</span>
                    </div>
                `;
                
                listEl.appendChild(item);
            });
            
            if (totalEl) {
                totalEl.textContent = `${totalHours.toFixed(1)} 小時`;
            }
            
        } else {
            emptyEl.style.display = 'block';
            if (totalEl) totalEl.textContent = '0.0 小時';
        }
        
    } catch (error) {
        console.error('❌ 載入加班記錄失敗:', error);
        loadingEl.style.display = 'none';
        emptyEl.style.display = 'block';
    }
}

// ==================== 薪資匯出功能（管理員專用） ====================

/**
 * ✅ 匯出所有員工薪資總表為 Excel（管理員專用）
 */
async function exportAllSalaryExcel() {
    try {
        console.log('🔍 開始匯出薪資總表');
        
        // 取得月份
        const yearMonthEl = document.getElementById('filter-year-month-list');
        const yearMonth = yearMonthEl ? yearMonthEl.value : '';
        
        if (!yearMonth) {
            showNotification('請先選擇要匯出的月份', 'error');
            return;
        }
        
        const token = localStorage.getItem('sessionToken');
        if (!token) {
            showNotification('請先登入', 'error');
            return;
        }
        
        console.log('📤 準備匯出:', { yearMonth, token: token ? '存在' : '不存在' });
        
        // 顯示進度
        showExportProgress('正在生成薪資總表 Excel...');
        
        // ⭐⭐⭐ 修正：使用正確的 API URL 格式
        const apiUrl = `${API_CONFIG.apiUrl}?action=exportAllSalaryExcel&token=${encodeURIComponent(token)}&yearMonth=${encodeURIComponent(yearMonth)}`;
        
        console.log('🌐 API URL:', apiUrl);
        
        // 呼叫 API
        const response = await fetch(apiUrl, {
            method: 'GET'
        });
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const result = await response.json();
        
        console.log('📥 收到回應:', result);
        
        hideExportProgress();
        
        // ⭐⭐⭐ 修正：正確判斷成功
        if (result.ok && result.fileUrl) {
            // 成功：開啟下載連結
            window.open(result.fileUrl, '_blank');
            
            showNotification(
                `✅ 匯出成功！\n檔案：${result.fileName || '薪資總表'}\n記錄數：${result.recordCount || 0}`,
                'success'
            );
            
            // 顯示結果區塊（備用）
            displayExportResult({
                fileName: result.fileName,
                fileUrl: result.fileUrl,
                recordCount: result.recordCount
            });
            
        } else {
            throw new Error(result.msg || result.message || '匯出失敗');
        }
        
    } catch (error) {
        hideExportProgress();
        console.error('❌ 匯出失敗:', error);
        showNotification('匯出失敗: ' + error.message, 'error');
    }
}

/**
 * ✅ 顯示匯出結果（備用方案）
 */
function displayExportResult(data) {
    // 建立結果提示區塊
    let resultDiv = document.getElementById('export-result-box');
    
    if (!resultDiv) {
        resultDiv = document.createElement('div');
        resultDiv.id = 'export-result-box';
        resultDiv.className = 'mt-4 p-4 bg-green-50 dark:bg-green-900/20 border-2 border-green-500 rounded-lg';
        
        const exportSection = document.querySelector('#admin-view .feature-box');
        if (exportSection) {
            exportSection.appendChild(resultDiv);
        }
    }
    
    resultDiv.innerHTML = `
        <div class="flex items-center justify-between">
            <div>
                <p class="font-semibold text-green-800 dark:text-green-300">
                    ✅ 薪資總表已生成！
                </p>
                <p class="text-sm text-green-700 dark:text-green-400">
                    檔案名稱：${data.fileName}.xlsx<br>
                    共 ${data.recordCount} 筆記錄
                </p>
            </div>
            <a href="${data.fileUrl}" 
               download="${data.fileName}.xlsx"
               class="px-6 py-3 bg-green-600 hover:bg-green-700 text-white font-bold rounded-lg transition-colors">
                📥 重新下載
            </a>
        </div>
    `;
    
    resultDiv.style.display = 'block';
    resultDiv.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

/**
 * ✅ 顯示匯出進度
 */
function showExportProgress(message) {
    // 移除舊的進度提示（如果存在）
    const oldProgress = document.getElementById('export-progress-overlay');
    if (oldProgress) {
        oldProgress.remove();
    }
    
    // 建立新的進度提示
    const overlay = document.createElement('div');
    overlay.id = 'export-progress-overlay';
    overlay.className = 'export-progress-overlay';
    
    overlay.innerHTML = `
        <div class="export-progress">
            <div class="export-progress-spinner"></div>
            <div class="export-progress-text">${message}</div>
            <p style="color: #94a3b8; font-size: 0.875rem; margin-top: 1rem;">
                請稍候，這可能需要幾秒鐘...
            </p>
        </div>
    `;
    
    document.body.appendChild(overlay);
}

/**
 * ✅ 隱藏匯出進度
 */
function hideExportProgress() {
    const overlay = document.getElementById('export-progress-overlay');
    if (overlay) {
        overlay.remove();
    }
}

console.log('✅ 薪資匯出功能已載入（管理員專用）');

console.log('✅ 薪資管理系統（完整版 v2.0）JS 已載入');
console.log('📋 包含：基本薪資 + 6項津貼 + 10項扣款');

/**
 * ✅ 呼叫 API：取得員工總工作時數
 * 
 * @param {string} yearMonth - 年月 (YYYY-MM)
 * @returns {Promise<Object>} { ok, totalWorkHours, workDays, records }
 */
async function getEmployeeWorkHours(yearMonth) {
    try {
      console.log(`📡 呼叫 API: getEmployeeWorkHours, 年月: ${yearMonth}`);
      
      const res = await callApifetch(`getEmployeeWorkHours&yearMonth=${encodeURIComponent(yearMonth)}`);
      
      if (res.ok && res.data) {
        console.log(`✅ 總工作時數: ${res.data.totalWorkHours}h`);
        console.log(`📊 工作天數: ${res.data.workDays} 天`);
        return res;
      } else {
        console.error('❌ 取得工作時數失敗:', res.msg);
        return { ok: false, msg: res.msg };
      }
      
    } catch (error) {
      console.error('❌ 呼叫 API 失敗:', error);
      return { ok: false, msg: error.toString() };
    }
  }


// ==================== 系統設定（管理員） ====================

let _systemConfig = null; // module-level cache
let _holidayList = [];    // current working list of holidays

const SYSTEM_CONFIG_FIELD_KEYS = [
  'OVERTIME_WEEKDAY_1', 'OVERTIME_WEEKDAY_2',
  'OVERTIME_RESTDAY_1', 'OVERTIME_RESTDAY_2', 'OVERTIME_RESTDAY_3',
  'OVERTIME_SUNDAY', 'OVERTIME_HOLIDAY',
  'MAX_OT_WEEKDAY', 'MAX_OT_RESTDAY', 'MAX_OT_HOLIDAY',
  'DAILY_WORK_HOURS', 'MONTHLY_WORK_DAYS', 'DEFAULT_PAYMENT_DAY',
  'SICK_LEAVE_RATE', 'CANCEL_BONUS_PERSONAL', 'CANCEL_BONUS_SICK',
  'HOLIDAY_WORK_AS_NORMAL',
];

/**
 * 載入系統設定並填入表單
 */
async function loadSystemConfig() {
  const loading = document.getElementById('system-config-loading');
  const content = document.getElementById('system-config-content');
  if (loading) loading.style.display = 'block';
  if (content) content.style.display = 'none';

  try {
    const res = await callApifetch('getSystemConfig');
    if (!res.ok) {
      showNotification('載入系統設定失敗：' + (res.msg || '未知錯誤'), 'error');
      return;
    }
    _systemConfig = res.data || {};

    // Fill numeric/text fields
    SYSTEM_CONFIG_FIELD_KEYS.forEach(key => {
      const el = document.getElementById('cfg-' + key);
      if (!el) return;
      if (el.type === 'checkbox') {
        el.checked = parseInt(_systemConfig[key]) === 1;
      } else {
        el.value = _systemConfig[key] !== undefined ? _systemConfig[key] : '';
      }
    });

    // Init holiday year selector
    const yearSel = document.getElementById('holiday-year-select');
    if (yearSel) {
      const currentYear = new Date().getFullYear();
      yearSel.innerHTML = '';
      for (let y = currentYear - 1; y <= currentYear + 2; y++) {
        const opt = document.createElement('option');
        opt.value = y;
        opt.textContent = y + ' 年';
        if (y === currentYear) opt.selected = true;
        yearSel.appendChild(opt);
      }
    }

    if (loading) loading.style.display = 'none';
    if (content) content.style.display = 'block';

    // Load holidays for current year
    await loadHolidayList();

    // Init batch-calc month selector to current month
    const batchMonthEl = document.getElementById('batch-calc-month');
    if (batchMonthEl && !batchMonthEl.value) {
      batchMonthEl.value = toTWDateString().slice(0, 7);
    }

    // Init advance date to today
    const advDateEl = document.getElementById('advance-date');
    if (advDateEl && !advDateEl.value) {
      advDateEl.value = toTWDateString();
    }

    // Load employee list for advance dropdown
    await loadEmployeeListForAdvance();
    await loadSalaryAdvances();

  } catch (err) {
    console.error('loadSystemConfig error:', err);
    showNotification('載入系統設定失敗', 'error');
  }
}

/**
 * 儲存指定區塊的系統設定
 * section: 'overtime' | 'hours' | 'leave'
 */
async function saveSystemConfigSection(section) {
  const sectionKeys = {
    overtime: ['OVERTIME_WEEKDAY_1','OVERTIME_WEEKDAY_2','OVERTIME_RESTDAY_1','OVERTIME_RESTDAY_2','OVERTIME_RESTDAY_3','OVERTIME_SUNDAY','OVERTIME_HOLIDAY'],
    hours:    ['DAILY_WORK_HOURS','MONTHLY_WORK_DAYS','DEFAULT_PAYMENT_DAY','MAX_OT_WEEKDAY','MAX_OT_RESTDAY','MAX_OT_HOLIDAY'],
    leave:    ['SICK_LEAVE_RATE','CANCEL_BONUS_PERSONAL','CANCEL_BONUS_SICK','HOLIDAY_WORK_AS_NORMAL'],
  };

  const keys = sectionKeys[section];
  if (!keys) return;

  const updates = {};
  keys.forEach(key => {
    const el = document.getElementById('cfg-' + key);
    if (!el) return;
    if (el.type === 'checkbox') {
      updates[key] = el.checked ? 1 : 0;
    } else {
      updates[key] = parseFloat(el.value) || el.value;
    }
  });

  try {
    const res = await callApifetch('saveSystemConfig&configJson=' + encodeURIComponent(JSON.stringify(updates)));
    if (res.ok) {
      showNotification('✅ 設定已儲存', 'success');
    } else {
      showNotification('儲存失敗：' + (res.msg || '未知錯誤'), 'error');
    }
  } catch (err) {
    console.error('saveSystemConfig error:', err);
    showNotification('儲存失敗', 'error');
  }
}

/**
 * 載入並顯示指定年份的國定假日
 */
async function loadHolidayList() {
  const yearSel = document.getElementById('holiday-year-select');
  const year = yearSel ? parseInt(yearSel.value) : new Date().getFullYear();

  try {
    const res = await callApifetch('getHolidays&year=' + year);
    if (res.ok) {
      _holidayList = res.holidays || [];
    } else {
      _holidayList = [];
    }
    renderHolidayList();
  } catch (err) {
    console.error('loadHolidayList error:', err);
    _holidayList = [];
    renderHolidayList();
  }
}

/**
 * 渲染假日標籤列表
 */
function renderHolidayList() {
  const container = document.getElementById('holiday-list-items');
  if (!container) return;

  if (_holidayList.length === 0) {
    container.innerHTML = '<span style="color:#64748b;font-size:0.875rem;">（尚無假日）</span>';
    return;
  }

  const sorted = [..._holidayList].sort();
  container.innerHTML = sorted.map(date => `
    <span style="display:inline-flex;align-items:center;gap:0.4rem;padding:0.3rem 0.7rem;background:rgba(59,130,246,0.15);border:1px solid rgba(59,130,246,0.3);border-radius:6px;color:#93c5fd;font-size:0.875rem;">
      ${date}
      <button onclick="removeHolidayFromList('${date}')"
              style="background:none;border:none;color:#f87171;cursor:pointer;font-size:1rem;line-height:1;padding:0;"
              title="移除">×</button>
    </span>
  `).join('');
}

/**
 * 新增假日到暫存清單（尚未存到後端）
 */
function addHolidayToList() {
  const dateInput = document.getElementById('holiday-add-date');
  if (!dateInput || !dateInput.value) {
    showNotification('請選擇日期', 'error');
    return;
  }
  const date = dateInput.value; // YYYY-MM-DD
  if (_holidayList.includes(date)) {
    showNotification('該日期已存在清單中', 'error');
    return;
  }
  _holidayList.push(date);
  renderHolidayList();
  dateInput.value = '';
  const labelInput = document.getElementById('holiday-add-label');
  if (labelInput) labelInput.value = '';

  const status = document.getElementById('holiday-save-status');
  if (status) status.textContent = '已新增，請按「儲存假日清單」以存入系統';
}

/**
 * 從暫存清單移除假日
 */
function removeHolidayFromList(date) {
  _holidayList = _holidayList.filter(d => d !== date);
  renderHolidayList();
  const status = document.getElementById('holiday-save-status');
  if (status) status.textContent = '已移除，請按「儲存假日清單」以存入系統';
}

/**
 * 儲存假日清單到後端
 */
async function saveHolidayList() {
  const yearSel = document.getElementById('holiday-year-select');
  const year = yearSel ? parseInt(yearSel.value) : new Date().getFullYear();
  const status = document.getElementById('holiday-save-status');

  try {
    const res = await callApifetch(
      'saveHolidays&year=' + year +
      '&holidaysJson=' + encodeURIComponent(JSON.stringify(_holidayList))
    );
    if (res.ok) {
      if (status) status.textContent = '✅ ' + year + ' 年假日清單已儲存（共 ' + _holidayList.length + ' 天）';
      showNotification('✅ 假日清單已儲存', 'success');
    } else {
      if (status) status.textContent = '❌ 儲存失敗：' + (res.msg || '未知錯誤');
      showNotification('儲存失敗：' + (res.msg || ''), 'error');
    }
  } catch (err) {
    console.error('saveHolidayList error:', err);
    showNotification('儲存失敗', 'error');
  }
}

// ==================== 批次計算薪資 ====================

async function batchCalculateSalary() {
  const monthEl = document.getElementById('batch-calc-month');
  const yearMonth = monthEl ? monthEl.value : '';
  if (!yearMonth) {
    showNotification('請選擇年月', 'error');
    return;
  }
  if (!confirm(`確定要計算 ${yearMonth} 所有員工的薪資嗎？\n此操作會覆蓋已存在的薪資記錄並扣除待扣預支。`)) return;

  const resultEl = document.getElementById('batch-calc-result');
  const detailsEl = document.getElementById('batch-calc-details');
  const tbodyEl = document.getElementById('batch-calc-tbody');
  if (resultEl) resultEl.innerHTML = '<span style="color:#94a3b8;">計算中，請稍候...</span>';
  if (detailsEl) detailsEl.style.display = 'none';

  try {
    const res = await callApifetch(`batchCalculateSalary&yearMonth=${encodeURIComponent(yearMonth)}`);
    if (!res.ok) {
      if (resultEl) resultEl.innerHTML = `<span style="color:#f87171;">失敗：${res.msg || '未知錯誤'}</span>`;
      return;
    }
    if (resultEl) {
      resultEl.innerHTML = `<span style="color:#10b981;">✅ 計算完成：成功 ${res.successCount} 人，失敗 ${res.failCount} 人</span>`;
    }
    if (tbodyEl && res.details && res.details.length > 0) {
      tbodyEl.innerHTML = res.details.map(d => `
        <tr style="border-bottom:1px solid rgba(255,255,255,0.05);">
          <td style="padding:0.4rem 0.5rem;color:#cbd5e1;">${d.employeeName || d.employeeId}</td>
          <td style="padding:0.4rem 0.5rem;text-align:right;color:#e2e8f0;">${d.ok ? formatCurrency(d.netSalary) : '-'}</td>
          <td style="padding:0.4rem 0.5rem;text-align:center;">
            ${d.ok ? '<span style="color:#10b981;">✅</span>' : `<span style="color:#f87171;" title="${d.msg || ''}">❌</span>`}
          </td>
        </tr>`).join('');
      if (detailsEl) detailsEl.style.display = 'block';
    }
  } catch (err) {
    console.error('batchCalculateSalary error:', err);
    if (resultEl) resultEl.innerHTML = '<span style="color:#f87171;">執行失敗，請查看 console</span>';
  }
}

// ==================== 薪資預支管理 ====================

async function loadEmployeeListForAdvance() {
  const sel = document.getElementById('advance-employee-id');
  if (!sel) return;
  try {
    const res = await callApifetch('getAllEmployeeBasicInfo');
    if (res.ok && res.data) {
      const current = sel.value;
      sel.innerHTML = '<option value="">— 選擇員工 —</option>';
      res.data.forEach(emp => {
        const id = emp['員工ID'] || emp.userId || '';
        const name = emp['員工姓名'] || emp.name || id;
        const opt = document.createElement('option');
        opt.value = id;
        opt.textContent = name;
        if (id === current) opt.selected = true;
        sel.appendChild(opt);
      });
    }
  } catch (err) {
    console.error('loadEmployeeListForAdvance error:', err);
  }
}

async function recordSalaryAdvance() {
  const employeeId = document.getElementById('advance-employee-id')?.value;
  const amount = parseFloat(document.getElementById('advance-amount')?.value || '0');
  const date = document.getElementById('advance-date')?.value;
  const note = document.getElementById('advance-note')?.value || '';
  const msgEl = document.getElementById('advance-msg');

  if (!employeeId) { if (msgEl) msgEl.textContent = '請選擇員工'; return; }
  if (amount <= 0) { if (msgEl) msgEl.textContent = '請輸入正確金額'; return; }

  try {
    const res = await callApifetch(
      `recordSalaryAdvance&employeeId=${encodeURIComponent(employeeId)}&amount=${amount}&date=${encodeURIComponent(date)}&note=${encodeURIComponent(note)}`
    );
    if (res.ok) {
      if (msgEl) msgEl.innerHTML = `<span style="color:#10b981;">✅ ${res.msg}</span>`;
      document.getElementById('advance-amount').value = '';
      document.getElementById('advance-note').value = '';
      await loadSalaryAdvances();
    } else {
      if (msgEl) msgEl.innerHTML = `<span style="color:#f87171;">❌ ${res.msg || '新增失敗'}</span>`;
    }
  } catch (err) {
    console.error('recordSalaryAdvance error:', err);
    if (msgEl) msgEl.textContent = '執行失敗';
  }
}

async function loadSalaryAdvances() {
  const listEl = document.getElementById('advance-list');
  if (!listEl) return;
  try {
    const res = await callApifetch('getSalaryAdvances');
    if (!res.ok || !res.data || res.data.length === 0) {
      listEl.innerHTML = '<span style="color:#64748b;">目前無預支記錄</span>';
      return;
    }
    listEl.innerHTML = `
      <table style="width:100%;border-collapse:collapse;">
        <thead>
          <tr style="color:#94a3b8;border-bottom:1px solid rgba(255,255,255,0.1);">
            <th style="text-align:left;padding:0.3rem 0.5rem;">員工</th>
            <th style="text-align:left;padding:0.3rem 0.5rem;">日期</th>
            <th style="text-align:right;padding:0.3rem 0.5rem;">金額</th>
            <th style="text-align:center;padding:0.3rem 0.5rem;">狀態</th>
            <th style="text-align:left;padding:0.3rem 0.5rem;">備註</th>
          </tr>
        </thead>
        <tbody>
          ${res.data.map(r => `
            <tr style="border-bottom:1px solid rgba(255,255,255,0.05);">
              <td style="padding:0.3rem 0.5rem;color:#cbd5e1;">${r.employeeName || r.employeeId}</td>
              <td style="padding:0.3rem 0.5rem;color:#94a3b8;">${r.date}</td>
              <td style="padding:0.3rem 0.5rem;text-align:right;color:#fcd34d;">${formatCurrency(r.amount)}</td>
              <td style="padding:0.3rem 0.5rem;text-align:center;">
                <span style="color:${r.status === '待扣' ? '#f59e0b' : '#10b981'};">${r.status}</span>
              </td>
              <td style="padding:0.3rem 0.5rem;color:#64748b;">${r.note || ''}</td>
            </tr>`).join('')}
        </tbody>
      </table>`;
  } catch (err) {
    console.error('loadSalaryAdvances error:', err);
    listEl.innerHTML = '<span style="color:#f87171;">載入失敗</span>';
  }
}
