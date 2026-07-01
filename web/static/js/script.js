/**
 * 智览金融 财数贯通 - 前端交互逻辑
 * 新布局：标签导航 + 左侧表单 + 右侧统一结果面板
 */
const POLL_INTERVAL = 1500;
const DEFAULT_COMPANIES = `平安银行股份有限公司
招商银行股份有限公司
国银金融租赁股份有限公司
招商证券股份有限公司
国信证券股份有限公司
长城证券股份有限公司
第一创业证券股份有限公司
中信证券股份有限公司
中国平安保险（集团）股份有限公司
阳光保险集团股份有限公司`;

// ==================== 统一结果面板 ====================

const resultsPanel = {
    list: document.getElementById('resultsList'),
    empty: document.getElementById('resultsEmpty'),
    count: document.getElementById('resultsCount'),
    tasks: new Map(),  // taskId -> card element

    showEmpty(show) {
        this.empty.style.display = show ? '' : 'none';
        this.list.style.display = show ? 'none' : '';
    },

    createCard(taskId, label) {
        this.showEmpty(false);
        const card = document.createElement('div');
        card.className = 'task-card';
        card.id = 'task-' + taskId;
        card.innerHTML = `
            <div class="task-card-header">
                <div class="task-card-title">
                    <span class="dot running"></span>
                    <span>${escapeHtml(label)}</span>
                </div>
                <span class="task-card-time">${new Date().toLocaleTimeString()}</span>
            </div>
            <div class="task-progress">
                <div class="progress-track">
                    <div class="progress-fill" style="width:0%"></div>
                </div>
                <div class="progress-text">正在启动...</div>
            </div>
        `;
        this.list.prepend(card);
        this.tasks.set(taskId, card);
        this.updateCount();
        return card;
    },

    updateProgress(taskId, task) {
        const card = this.tasks.get(taskId);
        if (!card) return;
        const current = task.progress.current;
        const total = task.progress.total;
        const pct = total > 0 ? Math.round((current / total) * 100) : 0;
        const fill = card.querySelector('.progress-fill');
        const text = card.querySelector('.progress-text');
        if (fill) fill.style.width = pct + '%';
        if (text) text.textContent = task.progress.message || '';
    },

    setDone(taskId, task) {
        const card = this.tasks.get(taskId);
        if (!card) return;
        const dot = card.querySelector('.dot');
        if (dot) { dot.className = 'dot done'; }
        const fill = card.querySelector('.progress-fill');
        if (fill) fill.style.width = '100%';
        const text = card.querySelector('.progress-text');
        if (text) text.textContent = '完成';

        // 移除进度条，添加文件列表
        const progress = card.querySelector('.task-progress');
        if (progress) progress.remove();

        if (task.files && task.files.length > 0) {
            const filesDiv = document.createElement('div');
            filesDiv.className = 'task-files';
            task.files.forEach(f => {
                const size = f.size ? (f.size >= 1048576 ? (f.size / 1048576).toFixed(2) + ' MB' : (f.size / 1024).toFixed(0) + ' KB') : '';
                filesDiv.innerHTML += `
                    <div class="file-row">
                        <div class="file-row-info">
                            <div class="file-row-name">${escapeHtml(f.display_name || f.name)}</div>
                            <div class="file-row-size">${size}</div>
                        </div>
                        <a class="file-row-dl" href="/download/${encodeURIComponent(f.name)}" download>
                            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                                <polyline points="7 10 12 15 17 10"/>
                                <line x1="12" y1="15" x2="12" y2="3"/>
                            </svg>
                            下载
                        </a>
                    </div>`;
            });
            card.appendChild(filesDiv);
        }
    },

    setError(taskId, message) {
        const card = this.tasks.get(taskId);
        if (!card) return;
        const dot = card.querySelector('.dot');
        if (dot) { dot.className = 'dot error'; }
        const progress = card.querySelector('.task-progress');
        if (progress) progress.remove();
        const err = document.createElement('div');
        err.className = 'task-error';
        err.textContent = '错误: ' + message;
        card.appendChild(err);
    },

    updateCount() {
        this.count.textContent = this.tasks.size + ' 个任务';
    }
};

// ==================== 标签切换 ====================

document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
        document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
        tab.classList.add('active');
        const name = tab.dataset.tab;
        document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
        document.getElementById('panel-' + name).classList.add('active');
    });
});

// ==================== 通用任务提交 ====================

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function setBtnLoading(btn, loading) {
    btn.disabled = loading;
    if (loading) {
        btn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="spinning"><circle cx="12" cy="12" r="10" stroke-opacity="0.3"/><path d="M12 2a10 10 0 0 1 10 10" stroke-linecap="round"/></svg>处理中...`;
    } else {
        btn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="5 3 19 12 5 21 5 3"/></svg>开始获取`;
    }
}

function pollTask(taskId, label, btn) {
    const timer = setInterval(async () => {
        try {
            const resp = await fetch('/api/progress/' + taskId);
            const data = await resp.json();
            if (!data.success) {
                clearInterval(timer);
                resultsPanel.setError(taskId, data.error || '查询失败');
                setBtnLoading(btn, false);
                return;
            }
            const task = data.task;
            resultsPanel.updateProgress(taskId, task);
            if (task.status === 'done') {
                clearInterval(timer);
                resultsPanel.setDone(taskId, task);
                setBtnLoading(btn, false);
            } else if (task.status === 'error') {
                clearInterval(timer);
                resultsPanel.setError(taskId, task.error || '处理失败');
                setBtnLoading(btn, false);
            }
        } catch (err) {
            clearInterval(timer);
            resultsPanel.setError(taskId, '网络错误: ' + err.message);
            setBtnLoading(btn, false);
        }
    }, POLL_INTERVAL);
}

async function submitFormData(apiUrl, formData, label, btn) {
    setBtnLoading(btn, true);
    try {
        const resp = await fetch(apiUrl, { method: 'POST', body: formData });
        const data = await resp.json();
        if (!data.success) {
            resultsPanel.createCard('err-' + Date.now(), label);
            resultsPanel.setError('err-' + Date.now(), data.error || '请求失败');
            setBtnLoading(btn, false);
            return;
        }
        resultsPanel.createCard(data.task_id, label);
        pollTask(data.task_id, label, btn);
    } catch (err) {
        resultsPanel.createCard('err-' + Date.now(), label);
        resultsPanel.setError('err-' + Date.now(), '网络错误: ' + err.message);
        setBtnLoading(btn, false);
    }
}

async function submitJson(apiUrl, body, label, btn) {
    setBtnLoading(btn, true);
    try {
        const resp = await fetch(apiUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body),
        });
        const data = await resp.json();
        if (!data.success) {
            resultsPanel.createCard('err-' + Date.now(), label);
            resultsPanel.setError('err-' + Date.now(), data.error || '请求失败');
            setBtnLoading(btn, false);
            return;
        }
        resultsPanel.createCard(data.task_id, label);
        pollTask(data.task_id, label, btn);
    } catch (err) {
        resultsPanel.createCard('err-' + Date.now(), label);
        resultsPanel.setError('err-' + Date.now(), '网络错误: ' + err.message);
        setBtnLoading(btn, false);
    }
}

// ==================== 模块1: 年报数据 ====================

const startYearEl = document.getElementById('startYear');
const endYearEl = document.getElementById('endYear');
const textInput = document.getElementById('textInput');
const fileInput = document.getElementById('fileInput');
const uploadArea = document.getElementById('uploadArea');
const uploadPlaceholder = document.getElementById('uploadPlaceholder');
const fileInfo = document.getElementById('fileInfo');
const fileNameSpan = document.getElementById('fileName');
const fileRemoveBtn = document.getElementById('fileRemove');
const industryAvgToggle = document.getElementById('industryAvgToggle');
const mergeReportsToggle = document.getElementById('mergeReportsToggle');

function initYearSelectors() {
    const cy = new Date().getFullYear();
    for (let y = cy; y >= 1990; y--) {
        const o1 = document.createElement('option'); o1.value = y; o1.textContent = y + '年'; startYearEl.appendChild(o1);
        const o2 = document.createElement('option'); o2.value = y; o2.textContent = y + '年'; endYearEl.appendChild(o2);
    }
    startYearEl.value = 2021; endYearEl.value = 2025;
}

function validateYearRange() {
    const s = parseInt(startYearEl.value), e = parseInt(endYearEl.value);
    if (s && e && s > e) { [startYearEl.value, endYearEl.value] = [endYearEl.value, startYearEl.value]; }
}

function showAnnualFile(file) {
    if (!['txt','csv'].includes(file.name.split('.').pop().toLowerCase())) {
        alert('仅支持 .txt 和 .csv 文件'); fileInput.value = ''; return;
    }
    uploadPlaceholder.style.display = 'none';
    fileInfo.style.display = 'flex';
    fileNameSpan.textContent = file.name;
    textInput.value = '';
}

function removeAnnualFile() {
    fileInput.value = '';
    uploadPlaceholder.style.display = '';
    fileInfo.style.display = 'none';
    if (!textInput.value.trim()) textInput.value = DEFAULT_COMPANIES;
}

uploadArea.addEventListener('dragover', e => { e.preventDefault(); uploadArea.style.borderColor = 'var(--accent)'; });
uploadArea.addEventListener('dragleave', () => { uploadArea.style.borderColor = ''; });
uploadArea.addEventListener('drop', e => {
    e.preventDefault(); uploadArea.style.borderColor = '';
    const f = e.dataTransfer.files[0];
    if (f) { fileInput.files = e.dataTransfer.files; showAnnualFile(f); }
});
fileInput.addEventListener('change', e => { if (e.target.files[0]) showAnnualFile(e.target.files[0]); });
fileRemoveBtn.addEventListener('click', removeAnnualFile);
startYearEl.addEventListener('change', validateYearRange);
endYearEl.addEventListener('change', validateYearRange);

document.getElementById('annualForm').addEventListener('submit', async e => {
    e.preventDefault();
    const s = parseInt(startYearEl.value), e2 = parseInt(endYearEl.value);
    const text = textInput.value.trim(), file = fileInput.files[0];
    if (!s || !e2) return alert('请选择年份范围');
    if (!text && !file) return alert('请输入企业名单或上传文件');

    const fd = new FormData();
    fd.append('start_year', s); fd.append('end_year', e2);
    fd.append('text_input', text);
    if (file) fd.append('file', file);
    fd.append('industry_avg', industryAvgToggle.checked ? '1' : '0');
    fd.append('merge_reports', mergeReportsToggle.checked ? '1' : '0');
    await submitFormData('/api/fetch', fd, '年报数据', document.getElementById('annualSubmitBtn'));
});

document.getElementById('annualResetBtn').addEventListener('click', () => {
    document.getElementById('annualForm').reset();
    textInput.value = DEFAULT_COMPANIES;
    removeAnnualFile();
    industryAvgToggle.checked = false;
    mergeReportsToggle.checked = false;
    startYearEl.value = 2021; endYearEl.value = 2025;
});

// ==================== 模块2: 机构名录 ====================

const instCheckboxes = document.querySelectorAll('.inst-checkbox');
const instBtn = document.getElementById('institutionSubmitBtn');

function updateInstBtn() {
    const any = Array.from(instCheckboxes).some(cb => cb.checked);
    instBtn.disabled = !any;
    instBtn.style.opacity = any ? '1' : '0.5';
}
instCheckboxes.forEach(cb => cb.addEventListener('change', updateInstBtn));

document.getElementById('institutionForm').addEventListener('submit', async e => {
    e.preventDefault();
    const types = [];
    if (document.getElementById('instAmac').checked) types.push('amac');
    if (document.getElementById('instBankIns').checked) types.push('bank_insurance');
    if (document.getElementById('instSecFund').checked) types.push('securities_fund');
    if (!types.length) return alert('请至少勾选一种名录');
    await submitJson('/api/institutions', { types }, '机构名录', instBtn);
});
updateInstBtn();

// ==================== 模块3: 深交所概况 ====================

const szseYearEl = document.getElementById('szseYear');
function initSzseYear() {
    const cy = new Date().getFullYear();
    for (let y = cy; y >= 2000; y--) {
        const o = document.createElement('option'); o.value = y; o.textContent = y + '年'; szseYearEl.appendChild(o);
    }
    szseYearEl.value = cy;
}
document.getElementById('szseForm').addEventListener('submit', async e => {
    e.preventDefault();
    const fd = new FormData(); fd.append('year', szseYearEl.value);
    await submitFormData('/api/szse', fd, '深交所概况', document.getElementById('szseSubmitBtn'));
});

// ==================== 模块4: 分红查询 ====================

const dividendStartYearEl = document.getElementById('dividendStartYear');
const dividendEndYearEl = document.getElementById('dividendEndYear');
const dividendTextInput = document.getElementById('dividendTextInput');
const dividendFileInput = document.getElementById('dividendFileInput');
const dividendUploadArea = document.getElementById('dividendUploadArea');
const dividendUploadPlaceholder = document.getElementById('dividendUploadPlaceholder');
const dividendFileInfo = document.getElementById('dividendFileInfo');
const dividendFileName = document.getElementById('dividendFileName');
const dividendFileRemove = document.getElementById('dividendFileRemove');

function initDividendYears() {
    const cy = new Date().getFullYear();
    for (let y = cy; y >= 2000; y--) {
        const o1 = document.createElement('option'); o1.value = y; o1.textContent = y + '年'; dividendStartYearEl.appendChild(o1);
        const o2 = document.createElement('option'); o2.value = y; o2.textContent = y + '年'; dividendEndYearEl.appendChild(o2);
    }
    dividendStartYearEl.value = 2024; dividendEndYearEl.value = 2025;
}

function validateDividendYears() {
    const s = parseInt(dividendStartYearEl.value), e = parseInt(dividendEndYearEl.value);
    if (s && e && s > e) { [dividendStartYearEl.value, dividendEndYearEl.value] = [dividendEndYearEl.value, dividendStartYearEl.value]; }
}

function showDividendFile(file) {
    if (!['txt','csv'].includes(file.name.split('.').pop().toLowerCase())) {
        alert('仅支持 .txt 和 .csv 文件'); dividendFileInput.value = ''; return;
    }
    dividendUploadPlaceholder.style.display = 'none';
    dividendFileInfo.style.display = 'flex';
    dividendFileName.textContent = file.name;
}

function removeDividendFile() {
    dividendFileInput.value = '';
    dividendUploadPlaceholder.style.display = '';
    dividendFileInfo.style.display = 'none';
}

dividendUploadArea.addEventListener('dragover', e => { e.preventDefault(); dividendUploadArea.style.borderColor = 'var(--accent)'; });
dividendUploadArea.addEventListener('dragleave', () => { dividendUploadArea.style.borderColor = ''; });
dividendUploadArea.addEventListener('drop', e => {
    e.preventDefault(); dividendUploadArea.style.borderColor = '';
    const f = e.dataTransfer.files[0];
    if (f) { dividendFileInput.files = e.dataTransfer.files; showDividendFile(f); }
});
dividendFileInput.addEventListener('change', e => { if (e.target.files[0]) showDividendFile(e.target.files[0]); });
dividendFileRemove.addEventListener('click', removeDividendFile);
dividendStartYearEl.addEventListener('change', validateDividendYears);
dividendEndYearEl.addEventListener('change', validateDividendYears);

document.getElementById('dividendForm').addEventListener('submit', async e => {
    e.preventDefault();
    const s = parseInt(dividendStartYearEl.value), e2 = parseInt(dividendEndYearEl.value);
    const text = dividendTextInput.value.trim(), file = dividendFileInput.files[0];
    if (!s || !e2) return alert('请选择年度范围');
    if (!text && !file) return alert('请输入企业名单或上传文件');

    const fd = new FormData();
    fd.append('start_year', s); fd.append('end_year', e2);
    fd.append('text_input', text);
    if (file) fd.append('file', file);
    await submitFormData('/api/dividend', fd, '分红查询', document.getElementById('dividendSubmitBtn'));
});

// ==================== 模块5: 处罚信息 ====================

document.getElementById('penaltyForm').addEventListener('submit', async e => {
    e.preventDefault();
    await submitJson('/api/penalty', {}, '处罚信息', document.getElementById('penaltySubmitBtn'));
});

// ==================== 启动 ====================
initYearSelectors();
initSzseYear();
initDividendYears();