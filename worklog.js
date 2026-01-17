// ==================== 工作日誌功能 ====================

/**
 * 初始化工作日誌分頁
 */
async function initWorklogTab() {
    await loadWorklogRecords();
    setupWorklogForm();
}

/**
 * 設定工作日誌表單
 */
function setupWorklogForm() {
    const dateInput = document.getElementById('worklog-date');
    const hoursInput = document.getElementById('worklog-hours');
    const contentInput = document.getElementById('worklog-content');
    const submitBtn = document.getElementById('submit-worklog-btn');
    
    // 設定預設日期為今天
    if (dateInput) {
        const today = new Date().toISOString().split('T')[0];
        dateInput.value = today;
        dateInput.max = today; // 不允許選擇未來日期
    }
    
    // 工時輸入驗證
    if (hoursInput) {
        hoursInput.addEventListener('input', (e) => {
            const value = parseFloat(e.target.value);
            if (value < 0) e.target.value = 0;
            if (value > 24) e.target.value = 24;
        });
    }
    
    // 提交按鈕
    if (submitBtn) {
        submitBtn.addEventListener('click', submitWorklog);
    }
}

/**
 * 提交工作日誌
 */
async function submitWorklog() {
    const dateInput = document.getElementById('worklog-date');
    const hoursInput = document.getElementById('worklog-hours');
    const contentInput = document.getElementById('worklog-content');
    const submitBtn = document.getElementById('submit-worklog-btn');
    
    const date = dateInput?.value;
    const hours = parseFloat(hoursInput?.value);
    const content = contentInput?.value.trim();
    
    // 驗證
    if (!date) {
        showNotification(t('WORKLOG_DATE_REQUIRED') || '請選擇日期', 'error');
        return;
    }
    
    if (!hours || hours <= 0) {
        showNotification(t('WORKLOG_HOURS_REQUIRED') || '請輸入工作時數', 'error');
        return;
    }
    
    if (!content || content.length < 10) {
        showNotification(t('WORKLOG_CONTENT_REQUIRED') || '請填寫工作內容（至少 10 個字）', 'error');
        return;
    }
    
    const loadingText = t('LOADING') || '提交中...';
    generalButtonState(submitBtn, 'processing', loadingText);
    
    try {
        const userId = localStorage.getItem('sessionUserId');
        
        const params = new URLSearchParams({
            date: date,
            hours: hours,
            content: content,
            userId: userId
        });
        
        const res = await callApifetch(`submitWorklog&${params.toString()}`);
        
        if (res.ok) {
            showNotification(t('WORKLOG_SUBMIT_SUCCESS') || '工作日誌提交成功！', 'success');
            
            // 清空表單
            if (hoursInput) hoursInput.value = '';
            if (contentInput) contentInput.value = '';
            
            // 重新載入記錄
            await loadWorklogRecords();
        } else {
            showNotification(res.msg || t('WORKLOG_SUBMIT_FAILED') || '提交失敗', 'error');
        }
        
    } catch (error) {
        console.error('提交工作日誌失敗:', error);
        showNotification(t('NETWORK_ERROR') || '網路錯誤', 'error');
        
    } finally {
        generalButtonState(submitBtn, 'idle');
    }
}

/**
 * 載入工作日誌記錄
 */
async function loadWorklogRecords() {
    const loadingEl = document.getElementById('worklog-records-loading');
    const emptyEl = document.getElementById('worklog-records-empty');
    const listEl = document.getElementById('worklog-records-list');
    
    if (!listEl) return;
    
    try {
        if (loadingEl) loadingEl.style.display = 'block';
        if (emptyEl) emptyEl.style.display = 'none';
        listEl.innerHTML = '';
        
        const userId = localStorage.getItem('sessionUserId');
        const res = await callApifetch(`getWorklogs&userId=${userId}`);
        
        if (loadingEl) loadingEl.style.display = 'none';
        
        if (res.ok && res.worklogs && res.worklogs.length > 0) {
            renderWorklogRecords(res.worklogs);
        } else {
            if (emptyEl) emptyEl.style.display = 'block';
        }
        
    } catch (error) {
        console.error('載入工作日誌失敗:', error);
        if (loadingEl) loadingEl.style.display = 'none';
        if (emptyEl) emptyEl.style.display = 'block';
    }
}

/**
 * 渲染工作日誌記錄（修正日期顯示）
 */
function renderWorklogRecords(worklogs) {
    const listEl = document.getElementById('worklog-records-list');
    if (!listEl) return;
    
    listEl.innerHTML = '';
    
    // 按日期排序（新到舊）
    const sortedWorklogs = worklogs.sort((a, b) => {
        return new Date(b.date) - new Date(a.date);
    });
    
    sortedWorklogs.forEach((log, index) => {
        const li = document.createElement('li');
        li.className = 'bg-white dark:bg-gray-800 rounded-lg p-4 border border-gray-200 dark:border-gray-700';
        
        // ⭐ 修正：格式化工作日期（只顯示日期部分）
        let workDateStr = log.date;
        if (log.date) {
            try {
                // 如果是 ISO 格式，只取日期部分
                if (log.date.includes('T')) {
                    workDateStr = log.date.split('T')[0];
                }
            } catch (e) {
                workDateStr = log.date;
            }
        }
        
        // ⭐ 格式化提交時間
        let submittedTimeStr = '';
        if (log.submittedAt) {
            try {
                const submittedDate = new Date(log.submittedAt);
                submittedTimeStr = submittedDate.toLocaleString('zh-TW', {
                    year: 'numeric',
                    month: '2-digit',
                    day: '2-digit',
                    hour: '2-digit',
                    minute: '2-digit',
                    hour12: false
                });
            } catch (e) {
                submittedTimeStr = log.submittedAt;
            }
        }
        
        // 狀態樣式
        let statusClass = '';
        let statusText = '';
        let statusIcon = '';
        let actionButtons = '';
        
        switch(log.status) {
            case 'PENDING':
                statusClass = 'bg-yellow-100 text-yellow-800 dark:bg-yellow-900/30 dark:text-yellow-300';
                statusText = t('STATUS_PENDING') || '待審核';
                statusIcon = '⏳';
                actionButtons = `
                    <button onclick="editWorklog('${log.id}')" 
                            class="px-3 py-1.5 text-sm bg-blue-500 hover:bg-blue-600 text-white rounded-md transition-colors">
                        ✏️ ${t('BTN_EDIT') || '編輯'}
                    </button>
                    <button onclick="deleteWorklog('${log.id}')" 
                            class="px-3 py-1.5 text-sm bg-red-500 hover:bg-red-600 text-white rounded-md transition-colors">
                        🗑️ ${t('BTN_DELETE') || '刪除'}
                    </button>
                `;
                break;
            case 'APPROVED':
                statusClass = 'bg-green-100 text-green-800 dark:bg-green-900/30 dark:text-green-300';
                statusText = t('STATUS_APPROVED') || '已核准';
                statusIcon = '✅';
                break;
            case 'REJECTED':
                statusClass = 'bg-red-100 text-red-800 dark:bg-red-900/30 dark:text-red-300';
                statusText = t('STATUS_REJECTED') || '已拒絕';
                statusIcon = '❌';
                actionButtons = `
                    <button onclick="editWorklog('${log.id}')" 
                            class="px-3 py-1.5 text-sm bg-orange-500 hover:bg-orange-600 text-white rounded-md transition-colors">
                        🔄 ${t('BTN_RESUBMIT') || '重新提交'}
                    </button>
                `;
                break;
        }
        
        li.innerHTML = `
            <div class="flex justify-between items-start mb-3">
                <div class="flex-1">
                    <div class="flex items-center space-x-2 mb-2">
                        <span class="font-bold text-gray-800 dark:text-white">${workDateStr}</span>
                        <span class="px-2 py-1 text-xs font-semibold rounded-full ${statusClass}">
                            ${statusIcon} ${statusText}
                        </span>
                    </div>
                    <div class="flex items-center space-x-4 text-sm text-gray-600 dark:text-gray-400">
                        <span>⏱️ ${log.hours} ${t('UNIT_HOURS') || '小時'}</span>
                        ${submittedTimeStr ? `<span>📅 ${submittedTimeStr}</span>` : ''}
                    </div>
                </div>
            </div>
            
            <div class="bg-gray-50 dark:bg-gray-700/50 rounded-lg p-3 mb-3">
                <p class="text-sm text-gray-700 dark:text-gray-300 whitespace-pre-wrap">${log.content}</p>
            </div>
            
            ${log.reviewComment ? `
                <div class="bg-blue-50 dark:bg-blue-900/20 border-l-4 border-blue-400 dark:border-blue-600 rounded p-3 mb-3">
                    <p class="text-xs font-semibold text-blue-800 dark:text-blue-300 mb-1">
                        💬 ${t('REVIEW_COMMENT') || '審核意見'}：
                    </p>
                    <p class="text-sm text-blue-700 dark:text-blue-400">${log.reviewComment}</p>
                </div>
            ` : ''}
            
            ${actionButtons ? `
                <div class="flex space-x-2 pt-3 border-t border-gray-200 dark:border-gray-600">
                    ${actionButtons}
                </div>
            ` : ''}
        `;
        
        listEl.appendChild(li);
    });
}
/**
 * 編輯工作日誌
 */
async function editWorklog(logId) {
    try {
        const res = await callApifetch(`getWorklogDetail&id=${logId}`);
        
        if (res.ok && res.worklog) {
            const log = res.worklog;
            
            // 填充表單
            const dateInput = document.getElementById('worklog-date');
            const hoursInput = document.getElementById('worklog-hours');
            const contentInput = document.getElementById('worklog-content');
            
            if (dateInput) dateInput.value = log.date;
            if (hoursInput) hoursInput.value = log.hours;
            if (contentInput) contentInput.value = log.content;
            
            // 滾動到表單
            document.getElementById('worklog-form-container')?.scrollIntoView({ 
                behavior: 'smooth',
                block: 'start'
            });
            
            // 變更提交按鈕為更新
            const submitBtn = document.getElementById('submit-worklog-btn');
            if (submitBtn) {
                submitBtn.textContent = t('BTN_UPDATE') || '更新';
                submitBtn.onclick = () => updateWorklog(logId);
            }
            
            showNotification(t('WORKLOG_EDIT_MODE') || '進入編輯模式', 'info');
        }
        
    } catch (error) {
        console.error('載入工作日誌失敗:', error);
        showNotification(t('LOAD_FAILED') || '載入失敗', 'error');
    }
}

/**
 * 更新工作日誌
 */
async function updateWorklog(logId) {
    const dateInput = document.getElementById('worklog-date');
    const hoursInput = document.getElementById('worklog-hours');
    const contentInput = document.getElementById('worklog-content');
    const submitBtn = document.getElementById('submit-worklog-btn');
    
    const date = dateInput?.value;
    const hours = parseFloat(hoursInput?.value);
    const content = contentInput?.value.trim();
    
    // 驗證
    if (!date || !hours || !content || content.length < 10) {
        showNotification(t('FORM_INCOMPLETE') || '請完整填寫表單', 'error');
        return;
    }
    
    const loadingText = t('UPDATING') || '更新中...';
    generalButtonState(submitBtn, 'processing', loadingText);
    
    try {
        const params = new URLSearchParams({
            id: logId,
            date: date,
            hours: hours,
            content: content
        });
        
        const res = await callApifetch(`updateWorklog&${params.toString()}`);
        
        if (res.ok) {
            showNotification(t('WORKLOG_UPDATE_SUCCESS') || '工作日誌更新成功！', 'success');
            
            // 清空表單並恢復提交按鈕
            if (hoursInput) hoursInput.value = '';
            if (contentInput) contentInput.value = '';
            submitBtn.textContent = t('BTN_SUBMIT_WORKLOG') || '提交工作日誌';
            submitBtn.onclick = submitWorklog;
            
            // 重新載入記錄
            await loadWorklogRecords();
        } else {
            showNotification(res.msg || t('UPDATE_FAILED') || '更新失敗', 'error');
        }
        
    } catch (error) {
        console.error('更新工作日誌失敗:', error);
        showNotification(t('NETWORK_ERROR') || '網路錯誤', 'error');
        
    } finally {
        generalButtonState(submitBtn, 'idle');
    }
}

/**
 * 刪除工作日誌
 */
async function deleteWorklog(logId) {
    if (!confirm(t('WORKLOG_DELETE_CONFIRM') || '確定要刪除此工作日誌嗎？')) {
        return;
    }
    
    try {
        const res = await callApifetch(`deleteWorklog&id=${logId}`);
        
        if (res.ok) {
            showNotification(t('WORKLOG_DELETE_SUCCESS') || '工作日誌已刪除', 'success');
            await loadWorklogRecords();
        } else {
            showNotification(res.msg || t('DELETE_FAILED') || '刪除失敗', 'error');
        }
        
    } catch (error) {
        console.error('刪除工作日誌失敗:', error);
        showNotification(t('NETWORK_ERROR') || '網路錯誤', 'error');
    }
}

// ==================== 管理員功能 ====================

/**
 * 載入待審核的工作日誌
 */
async function loadPendingWorklogs() {
    const loadingEl = document.getElementById('worklog-requests-loading');
    const emptyEl = document.getElementById('worklog-requests-empty');
    const listEl = document.getElementById('pending-worklog-list');
    
    if (!listEl) return;
    
    try {
        if (loadingEl) loadingEl.style.display = 'block';
        if (emptyEl) emptyEl.style.display = 'none';
        listEl.innerHTML = '';
        
        const res = await callApifetch('getPendingWorklogs');
        
        if (loadingEl) loadingEl.style.display = 'none';
        
        if (res.ok && res.worklogs && res.worklogs.length > 0) {
            renderPendingWorklogs(res.worklogs);
        } else {
            if (emptyEl) emptyEl.style.display = 'block';
        }
        
    } catch (error) {
        console.error('載入待審核工作日誌失敗:', error);
        if (loadingEl) loadingEl.style.display = 'none';
        if (emptyEl) emptyEl.style.display = 'block';
    }
}

/**
 * 渲染待審核的工作日誌（完全修正版）
 */
function renderPendingWorklogs(worklogs) {
    const listEl = document.getElementById('pending-worklog-list');
    if (!listEl) return;
    
    listEl.innerHTML = '';
    
    worklogs.forEach((log) => {
        const li = document.createElement('li');
        li.className = 'bg-white dark:bg-gray-800 rounded-lg p-4 border border-gray-200 dark:border-gray-700';
        
        // ⭐ 格式化工作日期
        let workDateStr = log.date;
        if (log.date && log.date.includes('T')) {
            workDateStr = log.date.split('T')[0];
        }
        
        // ⭐ 格式化提交時間
        let submittedTimeStr = '';
        if (log.submittedAt) {
            try {
                const submittedDate = new Date(log.submittedAt);
                submittedTimeStr = submittedDate.toLocaleString('zh-TW', {
                    year: 'numeric',
                    month: '2-digit',
                    day: '2-digit',
                    hour: '2-digit',
                    minute: '2-digit',
                    hour12: false
                });
            } catch (e) {
                submittedTimeStr = log.submittedAt;
            }
        }
        
        li.innerHTML = `
            <div class="flex justify-between items-start mb-3">
                <div class="flex-1">
                    <div class="flex items-center space-x-2 mb-2">
                        <span class="font-bold text-gray-800 dark:text-white">${log.userName || '未知員工'}</span>
                        <span class="text-xs px-2 py-1 rounded-full bg-indigo-100 dark:bg-indigo-900 text-indigo-700 dark:text-indigo-300">
                            ${log.department || '未分類'}
                        </span>
                    </div>
                    <div class="flex items-center space-x-4 text-sm text-gray-600 dark:text-gray-400">
                        <span>📅 ${workDateStr}</span>
                        <span>⏱️ ${log.hours} ${t('UNIT_HOURS') || '小時'}</span>
                        ${submittedTimeStr ? `<span>🕐 提交於 ${submittedTimeStr}</span>` : ''}
                    </div>
                </div>
            </div>
            
            <div class="bg-gray-50 dark:bg-gray-700/50 rounded-lg p-3 mb-3">
                <p class="text-sm font-semibold text-gray-700 dark:text-gray-300 mb-1">
                    📝 工作內容：
                </p>
                <p class="text-sm text-gray-600 dark:text-gray-400 whitespace-pre-wrap">${log.content}</p>
            </div>
            
            <div class="mb-3">
                <label for="review-comment-${log.id}" class="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1">
                    💬 審核意見 <span class="text-xs text-gray-500">(選填)</span>
                </label>
                <textarea id="review-comment-${log.id}" 
                          rows="2" 
                          placeholder="填寫審核意見（選填）..."
                          class="w-full p-2 border border-gray-300 dark:border-gray-600 rounded-md dark:bg-gray-700 dark:text-white text-sm"></textarea>
            </div>
            
            <div class="flex space-x-2">
                <button onclick="approveWorklog('${log.id}')" 
                        class="flex-1 px-4 py-2 bg-green-500 hover:bg-green-600 text-white rounded-md font-semibold transition-colors">
                    ✅ 核准
                </button>
                <button onclick="rejectWorklog('${log.id}')" 
                        class="flex-1 px-4 py-2 bg-red-500 hover:bg-red-600 text-white rounded-md font-semibold transition-colors">
                    ❌ 拒絕
                </button>
            </div>
        `;
        
        listEl.appendChild(li);
    });
}

/**
 * 核准工作日誌（除錯版本）
 */
async function approveWorklog(logId) {
    const commentInput = document.getElementById(`review-comment-${logId}`);
    const comment = commentInput?.value.trim() || '';
    
    console.log('🔍 核准工作日誌 - 開始');
    console.log('   日誌ID:', logId);
    console.log('   審核意見:', comment);
    
    try {
        // ⭐ 建立 URL（除錯版）
        const url = `reviewWorklog&id=${encodeURIComponent(logId)}&action=approve&comment=${encodeURIComponent(comment)}`;
        console.log('📤 發送 URL:', url);
        
        const res = await callApifetch(url);
        
        console.log('📥 收到回應:', res);
        
        if (res.ok) {
            showNotification('工作日誌已核准', 'success');
            await loadPendingWorklogs();
        } else {
            showNotification(res.msg || '核准失敗', 'error');
            console.error('❌ 錯誤訊息:', res.msg);
        }
        
    } catch (error) {
        console.error('❌ 核准失敗:', error);
        showNotification('網路錯誤', 'error');
    }
}
/**
 * 拒絕工作日誌（除錯版本）
 */
async function rejectWorklog(logId) {
    const commentInput = document.getElementById(`review-comment-${logId}`);
    const comment = commentInput?.value.trim();
    
    if (!comment) {
        showNotification('請填寫拒絕原因', 'error');
        commentInput?.focus();
        return;
    }
    
    console.log('🔍 拒絕工作日誌 - 開始');
    console.log('   日誌ID:', logId);
    console.log('   拒絕原因:', comment);
    
    try {
        const url = `reviewWorklog&id=${encodeURIComponent(logId)}&action=reject&comment=${encodeURIComponent(comment)}`;
        console.log('📤 發送 URL:', url);
        
        const res = await callApifetch(url);
        
        console.log('📥 收到回應:', res);
        
        if (res.ok) {
            showNotification('工作日誌已拒絕', 'success');
            await loadPendingWorklogs();
        } else {
            showNotification(res.msg || '拒絕失敗', 'error');
            console.error('❌ 錯誤訊息:', res.msg);
        }
        
    } catch (error) {
        console.error('❌ 拒絕失敗:', error);
        showNotification('網路錯誤', 'error');
    }
}

/**
 * 匯出工作日誌報表（PDF）
 */
async function exportWorklogReport() {
    const employeeSelect = document.getElementById('worklog-export-employee');
    const monthInput = document.getElementById('worklog-export-month');
    const exportBtn = document.getElementById('export-worklog-btn');
    
    if (!employeeSelect || !monthInput) return;
    
    const employeeId = employeeSelect.value;
    const yearMonth = monthInput.value;
    
    if (!employeeId) {
        showNotification(t('SELECT_EMPLOYEE_FIRST') || '請先選擇員工', 'error');
        return;
    }
    
    if (!yearMonth) {
        showNotification(t('SELECT_MONTH_FIRST') || '請先選擇月份', 'error');
        return;
    }
    
    const loadingText = t('EXPORT_LOADING') || '正在準備報表...';
    showNotification(loadingText, 'warning');
    
    if (exportBtn) {
        generalButtonState(exportBtn, 'processing', loadingText);
    }
    
    try {
        const employeeName = employeeSelect.options[employeeSelect.selectedIndex].text.split(' (')[0];
        
        const res = await callApifetch(`getWorklogReport&employeeId=${employeeId}&yearMonth=${yearMonth}`);
        
        if (!res.ok || !res.worklogs || res.worklogs.length === 0) {
            showNotification(t('EXPORT_NO_DATA') || '本月沒有工作日誌', 'warning');
            return;
        }
        
        // 使用 jsPDF 生成 PDF
        await generateWorklogPDF(employeeName, yearMonth, res.worklogs);
        
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

/**
 * 生成工作日誌 PDF
 */
async function generateWorklogPDF(employeeName, yearMonth, worklogs) {
    // 這裡需要引入 jsPDF 庫
    // 由於系統限制，這裡提供基本框架
    // 實際使用時需要在 HTML 中引入 jsPDF
    
    const [year, month] = yearMonth.split('-');
    
    // 如果有 jsPDF，可以這樣使用：
    // const { jsPDF } = window.jspdf;
    // const doc = new jsPDF();
    
    // 或者先使用 Excel 格式作為替代
    const exportData = worklogs.map(log => ({
        '日期': log.date,
        '工作時數': log.hours,
        '工作內容': log.content,
        '狀態': getStatusText(log.status),
        '審核意見': log.reviewComment || '-',
        '提交時間': log.submittedAt ? new Date(log.submittedAt).toLocaleString() : '-'
    }));
    
    const ws = XLSX.utils.json_to_sheet(exportData);
    
    const wscols = [
        { wch: 12 },  // 日期
        { wch: 10 },  // 工作時數
        { wch: 50 },  // 工作內容
        { wch: 12 },  // 狀態
        { wch: 30 },  // 審核意見
        { wch: 20 }   // 提交時間
    ];
    ws['!cols'] = wscols;
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `${month}月工作日誌`);
    
    const fileName = `${employeeName}_${year}年${month}月_工作日誌.xlsx`;
    XLSX.writeFile(wb, fileName);
}

/**
 * 獲取狀態文字
 */
function getStatusText(status) {
    const statusMap = {
        'PENDING': t('STATUS_PENDING') || '待審核',
        'APPROVED': t('STATUS_APPROVED') || '已核准',
        'REJECTED': t('STATUS_REJECTED') || '已拒絕'
    };
    return statusMap[status] || status;
}