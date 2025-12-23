// salary.js - 薪資管理前端邏輯（完整版 v2.0 - 含所有津貼與扣款）
// ==================== 檢查依賴 ====================
if (typeof callApifetch !== 'function') {
    console.error('❌ callApifetch 函數未定義，請確認 script.js 已正確載入');
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
        console.log('📡 正在驗證 Session...');
        const session = await callApifetch("checkSession");
        
        if (!session.ok || !session.user) {
            console.error('❌ Session 驗證失敗:', session);
            showNotification(t('SALARY_LOGIN_REQUIRED'), 'error');
            return;
        }
        
        console.log('✅ Session 驗證成功');
        console.log('👤 使用者:', session.user.name);
        console.log('🔐 權限:', session.user.dept);
        console.log('📌 員工ID:', session.user.userId);
        
        // 步驟 2：設定當前月份
        const now = new Date();
        const currentMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
        console.log('📅 當前月份:', currentMonth);
        
        const employeeSalaryMonth = document.getElementById('employee-salary-month');
        if (employeeSalaryMonth) {
            employeeSalaryMonth.value = currentMonth;
        }
        
        // 步驟 3：載入薪資資料
        console.log('💰 開始載入薪資資料...');
        await loadCurrentEmployeeSalary();
        
        console.log('📋 開始載入薪資歷史...');
        await loadSalaryHistory();
        
        // 步驟 4：綁定事件（管理員才需要）
        if (session.user.dept === "管理員") {
            console.log('🔧 綁定管理員功能...');
            bindSalaryEvents();
        }
        
        console.log('✅ 薪資頁面初始化完成（完整版 v2.0 + 多語言）！');
        
    } catch (error) {
        console.error('❌ 初始化失敗:', error);
        console.error('錯誤堆疊:', error.stack);
        showNotification(t('SALARY_INIT_FAILED') + ': ' + error.message, 'error');
    }
}
// ==================== 員工薪資功能 ====================

/**
 * ✅ 載入當前員工的薪資
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
        
        const result = await callApifetch(`getMySalary&yearMonth=${currentMonth}`);
        
        console.log('📥 薪資資料回應:', result);
        
        if (loadingEl) loadingEl.style.display = 'none';
        
        if (result.ok && result.data) {
            console.log('✅ 成功載入薪資資料');
            displayEmployeeSalary(result.data);
            if (contentEl) contentEl.style.display = 'block';
            await loadAttendanceDetails(currentMonth);
        } else {
            console.log(`⚠️ 沒有 ${currentMonth} 的薪資記錄`);
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
 * ✅ 按月份查詢薪資
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
        console.log(`🔍 查詢 ${yearMonth} 薪資`);
        
        loadingEl.style.display = 'block';
        emptyEl.style.display = 'none';
        contentEl.style.display = 'none';
        
        const res = await callApifetch(`getMySalary&yearMonth=${yearMonth}`);
        
        console.log(`📥 查詢 ${yearMonth} 薪資回應:`, res);
        
        loadingEl.style.display = 'none';
        
        if (res.ok && res.data) {
            console.log(`✅ 找到 ${yearMonth} 的薪資記錄`);
            displayEmployeeSalary(res.data);
            contentEl.style.display = 'block';
            await loadAttendanceDetails(yearMonth);
        } else {
            console.log(`⚠️ 沒有 ${yearMonth} 的薪資記錄`);
            showNoSalaryMessage(yearMonth);
            emptyEl.style.display = 'block';
            const detailsSection = document.getElementById('attendance-details-section');
            if (detailsSection) detailsSection.style.display = 'none';
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

/**
 * ✅ 載入加班記錄卡片（月薪/時薪通用）
 */
async function loadOvertimeRecordsCard(yearMonth, salaryData) {
    console.log('📊 載入加班記錄卡片');
    
    // 從薪資資料中取得加班資訊
    const totalOvertimeHours = parseFloat(salaryData['總加班時數']) || 0;
    const weekdayOvertimePay = parseFloat(salaryData['平日加班費']) || 0;
    const extendedOvertimePay = parseFloat(salaryData['國定假日加班費']) || 0;
    const totalOvertimePay = weekdayOvertimePay + extendedOvertimePay;
    
    console.log(`⏰ 總加班: ${totalOvertimeHours}h, 前2h: $${weekdayOvertimePay}, 後2h: $${extendedOvertimePay}`);
    
    // 建立加班卡片
    let overtimeCard = document.getElementById('overtime-records-card');
    
    if (!overtimeCard) {
        overtimeCard = document.createElement('div');
        overtimeCard.id = 'overtime-records-card';
        overtimeCard.className = 'feature-box bg-orange-900/20 border-orange-700 mt-4';
        
        const detailsSection = document.getElementById('attendance-details-section');
        detailsSection.appendChild(overtimeCard);
    }
    
    // 只有當有加班記錄時才顯示
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
                    <p class="text-sm text-orange-300 mb-1">前2小時加班費</p>
                    <p class="text-xl font-bold text-orange-200">${formatCurrency(weekdayOvertimePay)}</p>
                    <p class="text-xs text-orange-400 mt-1">(× 1.34)</p>
                </div>
                <div class="text-center p-3 bg-orange-800/20 rounded-lg">
                    <p class="text-sm text-orange-300 mb-1">後2小時加班費</p>
                    <p class="text-xl font-bold text-orange-200">${formatCurrency(extendedOvertimePay)}</p>
                    <p class="text-xs text-orange-400 mt-1">(× 1.67)</p>
                </div>
            </div>
            
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
        
        // 載入每日加班明細
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

/**
 * ✅ 顯示薪資明細（完整版 - 支援時薪顯示 + 工時統計）
 */
function displayEmployeeSalary(data) {
    console.log('顯示薪資明細（完整版）:', data);
    
    const safeSet = (id, value) => {
        const el = document.getElementById(id);
        if (el) {
            el.textContent = value;
        } else {
            console.warn(`⚠️ 元素 #${id} 未找到`);
        }
    };
    
    // ⭐ 判斷是否為時薪
    const salaryType = data['薪資類型'] || '月薪';
    const isHourly = salaryType === '時薪';
    
    // 應發總額與實發金額
    safeSet('gross-salary', formatCurrency(data['應發總額']));
    safeSet('net-salary', formatCurrency(data['實發金額']));
    
    // 計算扣款總額
    const deductions = 
        (parseFloat(data['勞保費']) || 0) + 
        (parseFloat(data['健保費']) || 0) + 
        (parseFloat(data['就業保險費']) || 0) + 
        (parseFloat(data['勞退自提']) || 0) + 
        (parseFloat(data['所得稅']) || 0) +
        (parseFloat(data['請假扣款']) || 0) +
        (parseFloat(data['福利金扣款']) || 0) +
        (parseFloat(data['宿舍費用']) || 0) +
        (parseFloat(data['團保費用']) || 0) +
        (parseFloat(data['其他扣款']) || 0);
    
    safeSet('total-deductions', formatCurrency(deductions));
    
    // ⭐⭐⭐ 應發項目（時薪 vs 月薪顯示不同）
    if (isHourly) {
        // 時薪顯示方式
        const hourlyRate = parseFloat(data['時薪']) || 0;
        const totalWorkHours = parseFloat(data['工作時數']) || 0;
        const totalWorkHoursInt = Math.floor(totalWorkHours);
        // 修改基本薪資的顯示文字
        const baseSalaryLabel = document.querySelector('[for="detail-base-salary"]') || 
                                document.querySelector('#detail-base-salary')?.previousElementSibling;
        if (baseSalaryLabel) {
            baseSalaryLabel.textContent = '基本薪資 (時薪×工時)';
        }
        
        // 在基本薪資下方顯示時薪資訊
        safeSet('detail-base-salary', formatCurrency(data['基本薪資']));
        
        // 可以考慮添加一個額外的顯示區域
        const baseSalaryEl = document.getElementById('detail-base-salary');
        if (baseSalaryEl && baseSalaryEl.parentElement) {
            // 檢查是否已存在時薪資訊
            let hourlyInfo = baseSalaryEl.parentElement.querySelector('.hourly-info');
            if (!hourlyInfo) {
                hourlyInfo = document.createElement('div');
                hourlyInfo.className = 'hourly-info text-xs text-purple-400 mt-1';
                baseSalaryEl.parentElement.appendChild(hourlyInfo);
            }
            hourlyInfo.textContent = `時薪 $${hourlyRate} × ${Math.floor(totalWorkHours)}h`;
        }
    } else {
        // 月薪顯示方式（原本的邏輯）
        safeSet('detail-base-salary', formatCurrency(data['基本薪資']));
        
        // 移除時薪資訊（如果存在）
        const baseSalaryEl = document.getElementById('detail-base-salary');
        if (baseSalaryEl && baseSalaryEl.parentElement) {
            const hourlyInfo = baseSalaryEl.parentElement.querySelector('.hourly-info');
            if (hourlyInfo) {
                hourlyInfo.remove();
            }
            
            // 恢復原本的標籤文字
            const baseSalaryLabel = document.querySelector('[for="detail-base-salary"]') || 
                                    baseSalaryEl.previousElementSibling;
            if (baseSalaryLabel) {
                baseSalaryLabel.textContent = '基本薪資';
            }
        }
    }
    
    // ⭐⭐⭐ 在加班費上方加入工時統計資訊
    const totalWorkHours = parseFloat(data['工作時數']) || 0;
    const totalOvertimeHours = parseFloat(data['總加班時數']) || 0;
    // ✅ 強制取整數
    const totalWorkHoursInt = Math.floor(totalWorkHours);
    const totalOvertimeHoursInt = Math.floor(totalOvertimeHours);
    // 找到平日加班費的元素
    const weekdayOvertimeEl = document.getElementById('detail-weekday-overtime');
    if (weekdayOvertimeEl && weekdayOvertimeEl.parentElement) {
        // 移除舊的工時資訊（如果存在）
        const oldWorkHoursInfo = weekdayOvertimeEl.parentElement.querySelector('.work-hours-summary');
        if (oldWorkHoursInfo) {
            oldWorkHoursInfo.remove();
        }
        
        // 只有在有工時或加班時數時才顯示
        if (totalWorkHours > 0 || totalOvertimeHours > 0) {
            const workHoursSummary = document.createElement('div');
            workHoursSummary.className = 'work-hours-summary mb-3 p-3 bg-blue-900/20 border border-blue-700/30 rounded-lg';
            
            let summaryHTML = '<div class="text-sm font-semibold text-blue-300 mb-2">📊 本月工時統計</div>';
            
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
            
            // 插入到平日加班費之前
            weekdayOvertimeEl.parentElement.parentElement.insertBefore(
                workHoursSummary,
                weekdayOvertimeEl.parentElement
            );
        }
    }
    
    // 其他津貼（時薪和月薪都顯示）
    safeSet('detail-position-allowance', formatCurrency(data['職務加給'] || 0));
    safeSet('detail-meal-allowance', formatCurrency(data['伙食費'] || 0));
    safeSet('detail-transport-allowance', formatCurrency(data['交通補助'] || 0));
    safeSet('detail-attendance-bonus', formatCurrency(data['全勤獎金'] || 0));
    safeSet('detail-performance-bonus', formatCurrency(data['績效獎金'] || 0));
    safeSet('detail-weekday-overtime', formatCurrency(data['平日加班費']));
    safeSet('detail-restday-overtime', formatCurrency(data['休息日加班費']));
    safeSet('detail-holiday-overtime', formatCurrency(data['國定假日加班費']));
    
    // 扣款項目
    safeSet('detail-labor-fee', formatCurrency(data['勞保費']));
    safeSet('detail-health-fee', formatCurrency(data['健保費']));
    safeSet('detail-employment-fee', formatCurrency(data['就業保險費']));
    
    const pensionRate = parseFloat(data['勞退自提率']) || 0;
    safeSet('detail-pension-rate', `${pensionRate}%`);
    
    safeSet('detail-pension-self', formatCurrency(data['勞退自提']));
    safeSet('detail-income-tax', formatCurrency(data['所得稅']));
    safeSet('detail-leave-deduction', formatCurrency(data['請假扣款']));
    
    const otherDeductions = 
        (parseFloat(data['福利金扣款']) || 0) +
        (parseFloat(data['宿舍費用']) || 0) +
        (parseFloat(data['團保費用']) || 0) +
        (parseFloat(data['其他扣款']) || 0);
    safeSet('detail-other-deductions', formatCurrency(otherDeductions));
    
    // 銀行資訊
    let bankCode = data['銀行代碼'];
    const bankAccount = data['銀行帳號'];
    
    if (bankCode) {
        bankCode = String(bankCode).padStart(3, '0');
    }
    
    safeSet('detail-bank-name', getBankName(bankCode));
    safeSet('detail-bank-account', bankAccount || '--');
    
    console.log('✅ 薪資明細顯示完成（完整版 - 支援時薪 + 工時統計）');
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

/**
 * 綁定表單事件
 */
function bindSalaryEvents() {
    console.log('🔗 綁定薪資表單事件（完整版）');
    
    const configForm = document.getElementById('salary-config-form');
    if (configForm) {
        configForm.addEventListener('submit', handleSalaryConfigSubmit);
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
    
    // 基本資訊
    const employeeId = safeGetValue('config-employee-id');
    const employeeName = safeGetValue('config-employee-name');
    const idNumber = safeGetValue('config-id-number');           // ⭐ 加入
    const employeeType = safeGetValue('config-employee-type');   // ⭐ 加入
    const salaryType = safeGetValue('config-salary-type');       // ⭐ 加入
    const baseSalary = safeGetValue('config-base-salary');
    
    // ⭐ 固定津貼（6項）
    const positionAllowance = safeGetValue('config-position-allowance') || '0';
    const mealAllowance = safeGetValue('config-meal-allowance') || '0';
    const transportAllowance = safeGetValue('config-transport-allowance') || '0';
    const attendanceBonus = safeGetValue('config-attendance-bonus') || '0';
    const performanceBonus = safeGetValue('config-performance-bonus') || '0';
    const otherAllowances = safeGetValue('config-other-allowances') || '0';
    
    // 法定扣款
    const laborFee = safeGetValue('config-labor-fee') || '0';
    const healthFee = safeGetValue('config-health-fee') || '0';
    const employmentFee = safeGetValue('config-employment-fee') || '0';
    const pensionSelf = safeGetValue('config-pension-self') || '0';
    const incomeTax = safeGetValue('config-income-tax') || '0';
    const pensionSelfRate = safeGetValue('config-pension-rate') || '0';
    
    // ⭐ 其他扣款（4項）
    const welfareFee = safeGetValue('config-welfare-fee') || '0';
    const dormitoryFee = safeGetValue('config-dormitory-fee') || '0';
    const groupInsurance = safeGetValue('config-group-insurance') || '0';
    const otherDeductions = safeGetValue('config-other-deductions') || '0';
    
    // 其他資訊
    const bankCodeRaw = document.getElementById('config-bank-code').value;
    const bankCode = bankCodeRaw ? String(bankCodeRaw).padStart(3, '0') : '';
    // const bankCode = safeGetValue('config-bank-code');
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
 * ✅ 顯示薪資計算結果（支援月薪/時薪區分）
 */
function displaySalaryCalculation(data, container) {
    if (!container) return;
    
    const totalDeductions = 
        (parseFloat(data.laborFee) || 0) + 
        (parseFloat(data.healthFee) || 0) + 
        (parseFloat(data.employmentFee) || 0) + 
        (parseFloat(data.pensionSelf) || 0) + 
        (parseFloat(data.incomeTax) || 0) + 
        (parseFloat(data.leaveDeduction) || 0) +
        (parseFloat(data.welfareFee) || 0) +
        (parseFloat(data.dormitoryFee) || 0) +
        (parseFloat(data.groupInsurance) || 0) +
        (parseFloat(data.otherDeductions) || 0);
    
    // ⭐ 判斷是月薪還是時薪
    const isHourly = data.salaryType === '時薪';
    
    container.innerHTML = `
        <div class="calculation-card">
            <h3 class="text-xl font-bold mb-4">
                ${data.employeeName || '--'} - ${data.yearMonth || '--'} 薪資計算結果
                <span class="ml-2 px-3 py-1 text-sm rounded-full ${isHourly ? 'bg-purple-100 text-purple-800' : 'bg-blue-100 text-blue-800'}">
                    ${data.salaryType || '月薪'}
                </span>
            </h3>
            
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
            
            ${isHourly ? `
                <!-- ⭐ 時薪統計卡片 -->
                <div class="bg-purple-50 dark:bg-purple-900/20 border-2 border-purple-200 dark:border-purple-700 rounded-lg p-4 mb-6">
                    <h4 class="font-semibold text-purple-800 dark:text-purple-300 mb-3">
                        時薪工時統計
                    </h4>
                    <div class="grid grid-cols-3 gap-4 text-center">
                        <div>
                            <p class="text-sm text-purple-600 dark:text-purple-400">時薪</p>
                            <p class="text-2xl font-bold text-purple-800 dark:text-purple-200">
                                $${data.hourlyRate || 0}
                            </p>
                        </div>
                        <div>
                            <p class="text-sm text-purple-600 dark:text-purple-400">工作時數</p>
                            <p class="text-2xl font-bold text-purple-800 dark:text-purple-200">
                                ${Math.floor(data.totalWorkHours || 0)}h
                            </p>
                        </div>
                        <div>
                            <p class="text-sm text-purple-600 dark:text-purple-400">基本薪資</p>
                            <p class="text-xl font-bold text-purple-800 dark:text-purple-200">
                                ${formatCurrency(data.baseSalary)}
                            </p>
                            <p class="text-xs text-purple-500">(時薪 × 工時)</p>
                        </div>
                    </div>
                </div>
            ` : ''}
            
            <!-- ⭐ 加班統計卡片（月薪/時薪都顯示） -->
            ${data.totalOvertimeHours > 0 ? `
                <div class="bg-orange-50 dark:bg-orange-900/20 border-2 border-orange-200 dark:border-orange-700 rounded-lg p-4 mb-6">
                    <h4 class="font-semibold text-orange-800 dark:text-orange-300 mb-3">
                        本月加班統計
                    </h4>
                    <div class="grid grid-cols-3 gap-4 text-center">
                        <div>
                            <p class="text-sm text-orange-600 dark:text-orange-400">總加班時數</p>
                            <p class="text-2xl font-bold text-orange-800 dark:text-orange-200">
                                ${data.totalOvertimeHours.toFixed(1)}h
                            </p>
                        </div>
                        <div>
                            <p class="text-sm text-orange-600 dark:text-orange-400">前2小時加班費</p>
                            <p class="text-xl font-bold text-orange-800 dark:text-orange-200">
                                ${formatCurrency(data.weekdayOvertimePay)}
                            </p>
                            <p class="text-xs text-orange-500">(× 1.34)</p>
                        </div>
                        <div>
                            <p class="text-sm text-orange-600 dark:text-orange-400">後2小時加班費</p>
                            <p class="text-xl font-bold text-orange-800 dark:text-orange-200">
                                ${formatCurrency(data.holidayOvertimePay)}
                            </p>
                            <p class="text-xs text-orange-500">(× 1.67)</p>
                        </div>
                    </div>
                </div>
            ` : ''}
            
            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div class="calculation-detail">
                    <h4 class="font-semibold mb-3 text-green-400">應發項目</h4>
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
                        <span>績效獎金</span>
                        <span class="font-mono">${formatCurrency(data.performanceBonus || 0)}</span>
                    </div>
                    <div class="calculation-row">
                        <span>前2小時加班費 (×1.34)</span>
                        <span class="font-mono">${formatCurrency(data.weekdayOvertimePay || 0)}</span>
                    </div>
                    <div class="calculation-row">
                        <span>後2小時加班費 (×1.67)</span>
                        <span class="font-mono">${formatCurrency(data.holidayOvertimePay || 0)}</span>
                    </div>
                    <div class="calculation-row total">
                        <span>應發總額</span>
                        <span>${formatCurrency(data.grossSalary)}</span>
                    </div>
                </div>
                
                <div class="calculation-detail">
                    <h4 class="font-semibold mb-3 text-red-400">扣款項目</h4>
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
 * ✅ 儲存薪資記錄
 */
async function saveSalaryRecord(data) {
    try {
        showNotification(t('SALARY_SAVING_RECORD'), 'info');
        
        const queryString = 
            `employeeId=${encodeURIComponent(data.employeeId)}` +
            `&employeeName=${encodeURIComponent(data.employeeName)}` +
            `&yearMonth=${encodeURIComponent(data.yearMonth)}` +
            `&baseSalary=${encodeURIComponent(data.baseSalary)}` +
            `&positionAllowance=${encodeURIComponent(data.positionAllowance || 0)}` +
            `&mealAllowance=${encodeURIComponent(data.mealAllowance || 0)}` +
            `&transportAllowance=${encodeURIComponent(data.transportAllowance || 0)}` +
            `&attendanceBonus=${encodeURIComponent(data.attendanceBonus || 0)}` +
            `&performanceBonus=${encodeURIComponent(data.performanceBonus || 0)}` +
            `&weekdayOvertimePay=${encodeURIComponent(data.weekdayOvertimePay || 0)}` +
            `&restdayOvertimePay=${encodeURIComponent(data.restdayOvertimePay || 0)}` +
            `&holidayOvertimePay=${encodeURIComponent(data.holidayOvertimePay || 0)}` +
            `&laborFee=${encodeURIComponent(data.laborFee || 0)}` +
            `&healthFee=${encodeURIComponent(data.healthFee || 0)}` +
            `&employmentFee=${encodeURIComponent(data.employmentFee || 0)}` +
            `&pensionSelf=${encodeURIComponent(data.pensionSelf || 0)}` +
            `&incomeTax=${encodeURIComponent(data.incomeTax || 0)}` +
            `&leaveDeduction=${encodeURIComponent(data.leaveDeduction || 0)}` +
            `&welfareFee=${encodeURIComponent(data.welfareFee || 0)}` +
            `&dormitoryFee=${encodeURIComponent(data.dormitoryFee || 0)}` +
            `&groupInsurance=${encodeURIComponent(data.groupInsurance || 0)}` +
            `&otherDeductions=${encodeURIComponent(data.otherDeductions || 0)}` +
            `&grossSalary=${encodeURIComponent(data.grossSalary)}` +
            `&netSalary=${encodeURIComponent(data.netSalary)}` +
            `&bankCode=${encodeURIComponent(data.bankCode || '')}` +
            `&bankAccount=${encodeURIComponent(data.bankAccount || '')}`;
        
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
async function loadAttendanceDetails(yearMonth) {
    try {
        console.log(`📋 載入 ${yearMonth} 出勤明細`);
        
        const detailsSection = document.getElementById('attendance-details-section');
        if (!detailsSection) return;
        
        // ⭐⭐⭐ 改用跟時薪計算一樣的 API
        // 先取得當前使用者的 session
        const session = await callApifetch('checkSession');
        if (!session.ok || !session.user) {
            detailsSection.style.display = 'none';
            return;
        }
        
        const employeeId = session.user.userId;
        
        // ⭐ 呼叫 calculateMonthlySalary（跟時薪計算完全一樣）
        const res = await callApifetch(`calculateMonthlySalary&employeeId=${encodeURIComponent(employeeId)}&yearMonth=${encodeURIComponent(yearMonth)}`);
        
        if (!res.ok || !res.data) {
            detailsSection.style.display = 'none';
            return;
        }
        
        const data = res.data;
        const salaryType = data.salaryType || '月薪';
        const isHourly = salaryType === '時薪';
        
        console.log(`💼 薪資類型: ${salaryType}, 是否為時薪: ${isHourly}`);
        
        // 顯示出勤明細區塊
        detailsSection.style.display = 'block';
        
        // ⭐ 如果是時薪，顯示工作時數卡片（直接用 API 回傳的資料）
        if (isHourly) {
            displayWorkHoursFromCalculation(data);
        }
        
        // 顯示加班記錄（直接用 API 回傳的資料）
        if (data.totalOvertimeHours > 0) {
            displayOvertimeFromCalculation(data);
        }
        
    } catch (error) {
        console.error('❌ 載入出勤明細失敗:', error);
    }
}

/**
 * ✅ 從薪資計算結果顯示工作時數（只顯示統計，不顯示明細）
 */
function displayWorkHoursFromCalculation(data) {
    const detailsSection = document.getElementById('attendance-details-section');
    if (!detailsSection) return;
    
    // 移除舊的工時卡片
    const oldCard = document.getElementById('work-hours-card');
    if (oldCard) oldCard.remove();
    
    // 建立新的工時卡片
    const workHoursCard = document.createElement('div');
    workHoursCard.id = 'work-hours-card';
    workHoursCard.className = 'feature-box bg-purple-900/20 border-purple-700 mb-4';
    
    const totalWorkHours = Math.floor(data.totalWorkHours || 0);
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


/**
 * ✅ 從薪資計算結果顯示加班統計（只顯示統計，不顯示明細）
 */
function displayOvertimeFromCalculation(data) {
    const detailsSection = document.getElementById('attendance-details-section');
    if (!detailsSection) return;
    
    // 移除舊的加班卡片
    const oldCard = document.getElementById('overtime-card');
    if (oldCard) oldCard.remove();
    
    // 建立新的加班卡片
    const overtimeCard = document.createElement('div');
    overtimeCard.id = 'overtime-card';
    overtimeCard.className = 'feature-box bg-orange-900/20 border-orange-700 mt-4';
    
    const totalOvertimeHours = Math.floor(data.totalOvertimeHours || 0);
    const weekdayOvertimePay = data.weekdayOvertimePay || 0;
    const extendedOvertimePay = data.holidayOvertimePay || 0;
    const totalOvertimePay = weekdayOvertimePay + extendedOvertimePay;
    
    overtimeCard.innerHTML = `
        <h4 class="font-semibold mb-3 text-orange-400">本月加班統計</h4>
        
        <div class="grid grid-cols-3 gap-4 mb-4">
            <div class="text-center p-3 bg-orange-800/20 rounded-lg">
                <p class="text-sm text-orange-300 mb-1">總加班時數</p>
                <p class="text-2xl font-bold text-orange-200">${totalOvertimeHours}h</p>
            </div>
            <div class="text-center p-3 bg-orange-800/20 rounded-lg">
                <p class="text-sm text-orange-300 mb-1">前2小時加班費</p>
                <p class="text-xl font-bold text-orange-200">${formatCurrency(weekdayOvertimePay)}</p>
                <p class="text-xs text-orange-400 mt-1">(× 1.34)</p>
            </div>
            <div class="text-center p-3 bg-orange-800/20 rounded-lg">
                <p class="text-sm text-orange-300 mb-1">後2小時加班費</p>
                <p class="text-xl font-bold text-orange-200">${formatCurrency(extendedOvertimePay)}</p>
                <p class="text-xs text-orange-400 mt-1">(× 1.67)</p>
            </div>
        </div>
        
        <div class="p-3 bg-orange-800/20 rounded-lg">
            <div class="flex justify-between items-center">
                <span class="font-semibold text-orange-200">加班費合計</span>
                <span class="text-2xl font-bold text-orange-300">${formatCurrency(totalOvertimePay)}</span>
            </div>
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
 * ✅ 匯出所有員工的薪資總表（Excel格式）- 管理員專用
 */
async function exportAllSalaryExcel() {
    try {
        // 1. 檢查管理員權限
        if (currentUserRole !== 'admin') {
            showNotification('❌ 此功能僅限管理員使用', 'error');
            return;
        }
        
        // 2. 取得選擇的月份
        const yearMonthEl = document.getElementById('filter-year-month-list');
        const yearMonth = yearMonthEl ? yearMonthEl.value : '';
        
        if (!yearMonth) {
            showNotification('❌ 請先選擇要匯出的月份', 'error');
            return;
        }
        
        // 3. 確認是否要匯出
        if (!confirm(`確定要匯出 ${yearMonth} 的薪資總表嗎？\n\n將包含所有員工的薪資明細。`)) {
            return;
        }
        
        // 4. 顯示進度提示
        showExportProgress('正在生成薪資總表 Excel...');
        
        // 5. 呼叫後端 API
        const res = await callApifetch(`action=exportAllSalaryExcel&yearMonth=${encodeURIComponent(yearMonth)}`);
        
        // 6. 隱藏進度提示
        hideExportProgress();
        
        // 7. 處理結果
        if (res.ok && res.fileUrl) {
            // 開啟新視窗下載 Excel
            window.open(res.fileUrl, '_blank');
            
            showNotification(`✅ 薪資總表已生成！\n共 ${res.recordCount} 筆記錄\n\n請檢查下載或前往 Google Drive 查看`, 'success');
            
            console.log('✅ 匯出成功:', {
                檔案名稱: res.fileName,
                記錄數: res.recordCount,
                檔案連結: res.fileUrl
            });
        } else {
            showNotification('❌ 匯出失敗: ' + (res.msg || res.message || '未知錯誤'), 'error');
        }
        
    } catch (error) {
        hideExportProgress();
        console.error('❌ 匯出 Excel 失敗:', error);
        showNotification('❌ 匯出失敗：' + error.message, 'error');
    }
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