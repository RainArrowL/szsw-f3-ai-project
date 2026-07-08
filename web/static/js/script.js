/**
 * 智览金融 财数贯通 - 前端交互逻辑
 * 布局：左侧模块导航 + 右侧表单区 + 结果区
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

// ==================== 模块配置 ====================
const MODULES = {
    annual: {
        label: '年报数据爬取',
        apiUrl: '/api/fetch',
        isFormData: true,
        buildFormData: () => {
            const startYear = parseInt(document.getElementById('startYear').value);
            const endYear = parseInt(document.getElementById('endYear').value);
            const textVal = document.getElementById('textInput').value.trim();
            const file = document.getElementById('fileInput').files[0];
            if (!startYear || !endYear) { alert('请选择年度范围'); return null; }
            if (!textVal && !file) { alert('请输入企业名单或上传文件'); return null; }
            const fd = new FormData();
            fd.append('start_year', startYear);
            fd.append('end_year', endYear);
            fd.append('text_input', textVal);
            if (file) fd.append('file', file);
            fd.append('industry_avg', document.getElementById('industryAvgToggle').checked ? '1' : '0');
            fd.append('merge_reports', document.getElementById('mergeReportsToggle').checked ? '1' : '0');
            return fd;
        }
    },
    institutions: {
        label: '机构法人名录',
        apiUrl: '/api/institutions',
        isFormData: false,
        buildFormData: () => {
            const types = [];
            if (document.getElementById('bank_insurance').checked) types.push('bank_insurance');
            if (document.getElementById('securities_fund').checked) types.push('securities_fund');
            if (document.getElementById('amac').checked) types.push('amac');
            if (!types.length) { alert('请至少选择一种名录类型'); return null; }
            return JSON.stringify({ types });
        }
    },
    szse: {
        label: '深交所日度概况',
        apiUrl: '/api/szse',
        isFormData: true,
        buildFormData: () => {
            const year = document.getElementById('szseYear').value;
            if (!year) { alert('请选择年度'); return null; }
            const fd = new FormData();
            fd.append('year', year);
            return fd;
        }
    },
    dividend: {
        label: '分红公告查询',
        apiUrl: '/api/dividend',
        isFormData: true,
        buildFormData: () => {
            const startYear = parseInt(document.getElementById('dividendStartYear').value);
            const endYear = parseInt(document.getElementById('dividendEndYear').value);
            const textVal = document.getElementById('dividendTextInput').value.trim();
            const file = document.getElementById('dividendFileInput').files[0];
            if (!startYear || !endYear) { alert('请选择年度范围'); return null; }
            if (!textVal && !file) { alert('请输入企业名单或上传文件'); return null; }
            const fd = new FormData();
            fd.append('start_year', startYear);
            fd.append('end_year', endYear);
            fd.append('text_input', textVal);
            if (file) fd.append('file', file);
            return fd;
        }
    },
    external: {
        label: '外部信息获取',
        apiUrl: '/api/external',
        isFormData: false,
        buildFormData: () => {
            const penalty = document.getElementById('extPenalty').checked;
            const stats = document.getElementById('extStats').checked;
            if (!penalty && !stats) { alert('请至少勾选一个信息类型'); return null; }
            return JSON.stringify({ penalty, stats });
        }
    }
};

// ==================== 模块切换 ====================
let currentModule = 'annual';

function switchModule(name) {
    currentModule = name;
    document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
    document.querySelector(`.nav-item[data-module="${name}"]`).classList.add('active');
    document.querySelectorAll('.form-panel').forEach(p => p.classList.remove('active'));
    document.getElementById(`form-${name}`).classList.add('active');
}

document.querySelectorAll('.nav-item').forEach(btn => {
    btn.addEventListener('click', () => switchModule(btn.dataset.module));
});

// ==================== 结果卡片管理 ====================
function createResultCard(taskLabel, cardId) {
    const placeholder = document.querySelector('.results-placeholder');
    if (placeholder) placeholder.style.display = 'none';

    const card = document.createElement('div');
    card.className = 'result-card';
    card.id = cardId;
    card.innerHTML = `
        <div class="result-card-header">
            <span class="result-card-title">${escapeHtml(taskLabel)}</span>
            <span class="result-card-badge processing">处理中...</span>
        </div>
        <div class="progress-section">
            <div class="progress-header">
                <span class="progress-label">等待开始...</span>
                <span class="progress-percent">0%</span>
            </div>
            <div class="progress-track">
                <div class="progress-fill" style="width:0%"></div>
            </div>
            <p class="progress-message"></p>
        </div>
        <div class="file-list" style="display:none;"></div>
    `;
    document.getElementById('resultsPanel').prepend(card);
    return card;
}

function updateCardProgress(card, task) {
    const current = task.progress.current;
    const total = task.progress.total;
    const pct = total > 0 ? Math.round((current / total) * 100) : 0;

    card.querySelector('.progress-label').textContent = task.status === 'done' ? '完成' : '处理中';
    card.querySelector('.progress-percent').textContent = pct + '%';
    card.querySelector('.progress-fill').style.width = pct + '%';
    card.querySelector('.progress-message').textContent = task.progress.message || '';
}

function setCardDone(card, task) {
    const badge = card.querySelector('.result-card-badge');
    badge.className = 'result-card-badge success';
    badge.textContent = '完成';

    card.querySelector('.progress-section').style.display = 'none';

    const fileList = card.querySelector('.file-list');
    fileList.style.display = 'flex';
    fileList.innerHTML = '';

    task.files.forEach(f => {
        const item = document.createElement('div');
        item.className = 'file-item';
        item.innerHTML = `
            <div class="file-info">
                <div class="file-name">${escapeHtml(f.display_name || f.name)}</div>
                <div class="file-size">${formatFileSize(f.size)}</div>
            </div>
            <a class="download-btn" href="/download/${encodeURIComponent(f.name)}" download>
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
                下载
            </a>
        `;
        fileList.appendChild(item);
    });

    // 2+文件时加一键下载按钮，放在文件列表上方
    if (task.files.length >= 2) {
        const btn = document.createElement('button');
        btn.className = 'btn-download-all';
        btn.innerHTML = `
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            一键下载全部 (${task.files.length}个文件)
        `;
        btn.addEventListener('click', async () => {
            btn.disabled = true;
            btn.textContent = '打包中...';
            try {
                const label = MODULES[currentModule] ? MODULES[currentModule].label : '批量下载';
                const filenames = task.files.map(f => f.name);
                const resp = await fetch('/api/download_all', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ files: filenames, label })
                });
                if (resp.ok) {
                    const blob = await resp.blob();
                    const url = URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = '';
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    URL.revokeObjectURL(url);
                } else {
                    const err = await resp.json();
                    alert('下载失败: ' + (err.error || '未知错误'));
                }
            } catch (err) {
                alert('下载失败: ' + err.message);
            }
            btn.disabled = false;
            btn.innerHTML = `
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
                一键下载全部 (${task.files.length}个文件)
            `;
        });
        card.insertBefore(btn, fileList);
    }
}

function setCardError(card, message) {
    const badge = card.querySelector('.result-card-badge');
    badge.className = 'result-card-badge error';
    badge.textContent = '失败';

    card.querySelector('.progress-section').style.display = 'none';
    card.querySelector('.progress-message').textContent = '错误: ' + message;
    card.querySelector('.progress-message').style.display = 'block';
}

// ==================== 任务提交 & 轮询 ====================
function setBtnLoading(btn, loading) {
    if (loading) {
        btn.disabled = true;
        btn.dataset.origHtml = btn.innerHTML;
        btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="spinning"><circle cx="12" cy="12" r="10" stroke-opacity="0.3"/><path d="M12 2a10 10 0 0 1 10 10" stroke-linecap="round"/></svg> 处理中...`;
    } else {
        btn.disabled = false;
        if (btn.dataset.origHtml) btn.innerHTML = btn.dataset.origHtml;
    }
}

async function submitModule(name) {
    const cfg = MODULES[name];
    const body = cfg.buildFormData();
    if (!body) return;

    const btn = document.getElementById(name + 'SubmitBtn');
    setBtnLoading(btn, true);

    const cardId = 'card-' + Date.now();
    const card = createResultCard(cfg.label, cardId);

    try {
        const opts = { method: 'POST' };
        if (cfg.isFormData) {
            opts.body = body;
        } else {
            opts.headers = { 'Content-Type': 'application/json' };
            opts.body = body;
        }

        const resp = await fetch(cfg.apiUrl, opts);
        const data = await resp.json();

        if (!data.success) {
            setCardError(card, data.error || '请求失败');
            setBtnLoading(btn, false);
            return;
        }

        pollProgress(data.task_id, card, btn);
    } catch (err) {
        setCardError(card, '网络错误: ' + err.message);
        setBtnLoading(btn, false);
    }
}

function pollProgress(taskId, card, btn) {
    const timer = setInterval(async () => {
        try {
            const resp = await fetch('/api/progress/' + taskId);
            const data = await resp.json();
            if (!data.success) {
                clearInterval(timer);
                setCardError(card, data.error || '查询进度失败');
                setBtnLoading(btn, false);
                return;
            }

            const task = data.task;
            updateCardProgress(card, task);

            if (task.status === 'done') {
                clearInterval(timer);
                setCardDone(card, task);
                setBtnLoading(btn, false);
            } else if (task.status === 'error') {
                clearInterval(timer);
                setCardError(card, task.error || '处理失败');
                setBtnLoading(btn, false);
            }
        } catch (err) {
            clearInterval(timer);
            setCardError(card, '网络错误: ' + err.message);
            setBtnLoading(btn, false);
        }
    }, POLL_INTERVAL);
}

// ==================== 表单绑定 ====================
Object.keys(MODULES).forEach(name => {
    const form = document.getElementById(name + 'Form');
    if (form) {
        form.addEventListener('submit', e => {
            e.preventDefault();
            submitModule(name);
        });
    }
});

// ==================== 年份选择器 ====================
function initYearSelectors() {
    const cy = new Date().getFullYear();
    [document.getElementById('startYear'), document.getElementById('endYear')].forEach(sel => {
        for (let y = cy; y >= 1990; y--) {
            const opt = document.createElement('option');
            opt.value = y;
            opt.textContent = y + '年';
            sel.appendChild(opt);
        }
    });
    document.getElementById('startYear').value = 2021;
    document.getElementById('endYear').value = 2025;
}

function initSzseYearSelector() {
    const sel = document.getElementById('szseYear');
    const cy = new Date().getFullYear();
    for (let y = cy; y >= 2000; y--) {
        const opt = document.createElement('option');
        opt.value = y;
        opt.textContent = y + '年';
        sel.appendChild(opt);
    }
    sel.value = cy;
}

function initDividendYearSelectors() {
    const cy = new Date().getFullYear();
    [document.getElementById('dividendStartYear'), document.getElementById('dividendEndYear')].forEach(sel => {
        for (let y = cy; y >= 2000; y--) {
            const opt = document.createElement('option');
            opt.value = y;
            opt.textContent = y + '年';
            sel.appendChild(opt);
        }
    });
    document.getElementById('dividendStartYear').value = 2024;
    document.getElementById('dividendEndYear').value = 2025;
}

// 年份范围自动交换
function validateYearRange(startEl, endEl) {
    const s = parseInt(startEl.value), e = parseInt(endEl.value);
    if (s && e && s > e) {
        [startEl.value, endEl.value] = [endEl.value, startEl.value];
    }
}
document.getElementById('startYear').addEventListener('change', function() {
    validateYearRange(document.getElementById('startYear'), document.getElementById('endYear'));
});
document.getElementById('endYear').addEventListener('change', function() {
    validateYearRange(document.getElementById('startYear'), document.getElementById('endYear'));
});
document.getElementById('dividendStartYear').addEventListener('change', function() {
    validateYearRange(document.getElementById('dividendStartYear'), document.getElementById('dividendEndYear'));
});
document.getElementById('dividendEndYear').addEventListener('change', function() {
    validateYearRange(document.getElementById('dividendStartYear'), document.getElementById('dividendEndYear'));
});

// ==================== 文件上传 ====================
function setupFileUpload(uploadAreaId, fileInputId, fileInfoId, fileNameId, fileRemoveId, textInputId) {
    const uploadArea = document.getElementById(uploadAreaId);
    const fileInput = document.getElementById(fileInputId);
    const fileInfo = document.getElementById(fileInfoId);
    const fileName = document.getElementById(fileNameId);
    const fileRemove = document.getElementById(fileRemoveId);
    const textInput = document.getElementById(textInputId);
    const uploadContent = uploadArea.querySelector('.upload-content');

    function showFile(file) {
        if (!['txt', 'csv'].includes(file.name.split('.').pop().toLowerCase())) {
            alert('仅支持 .txt 和 .csv 文件');
            fileInput.value = '';
            return;
        }
        uploadContent.style.display = 'none';
        fileInfo.style.display = 'flex';
        fileName.textContent = file.name;
        if (textInput) textInput.value = '';
    }

    function removeFile() {
        fileInput.value = '';
        uploadContent.style.display = '';
        fileInfo.style.display = 'none';
        if (textInput && !textInput.value.trim()) textInput.value = DEFAULT_COMPANIES;
    }

    uploadArea.addEventListener('dragover', e => { e.preventDefault(); uploadArea.style.borderColor = 'var(--accent)'; });
    uploadArea.addEventListener('dragleave', () => { uploadArea.style.borderColor = ''; });
    uploadArea.addEventListener('drop', e => {
        e.preventDefault();
        uploadArea.style.borderColor = '';
        const file = e.dataTransfer.files[0];
        if (file) { fileInput.files = e.dataTransfer.files; showFile(file); }
    });
    fileInput.addEventListener('change', e => { const f = e.target.files[0]; if (f) showFile(f); });
    fileRemove.addEventListener('click', removeFile);
}

setupFileUpload('uploadArea', 'fileInput', 'fileInfo', 'fileName', 'fileRemove', 'textInput');
setupFileUpload('dividendUploadArea', 'dividendFileInput', 'dividendFileInfo', 'dividendFileName', 'dividendFileRemove', 'dividendTextInput');

// ==================== 年报重置 ====================
document.getElementById('annualResetBtn').addEventListener('click', () => {
    document.getElementById('annualForm').reset();
    document.getElementById('textInput').value = DEFAULT_COMPANIES;
    document.getElementById('fileInput').value = '';
    document.getElementById('fileInfo').style.display = 'none';
    document.querySelector('#uploadArea .upload-content').style.display = '';
    document.getElementById('startYear').value = 2021;
    document.getElementById('endYear').value = 2025;
});

// ==================== 外部信息按钮状态 ====================
function updateExternalBtn() {
    const btn = document.getElementById('externalSubmitBtn');
    const checked = document.getElementById('extPenalty').checked || document.getElementById('extStats').checked;
    btn.disabled = !checked;
}
document.getElementById('extPenalty').addEventListener('change', updateExternalBtn);
document.getElementById('extStats').addEventListener('change', updateExternalBtn);

// ==================== 工具函数 ====================
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function formatFileSize(bytes) {
    if (!bytes) return '';
    if (bytes > 1024 * 1024) return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
    return (bytes / 1024).toFixed(0) + ' KB';
}

// ==================== 启动 ====================
initYearSelectors();
initSzseYearSelector();
initDividendYearSelectors();

document.head.insertAdjacentHTML('beforeend', '<style>.spinning{animation:spin 1s linear infinite}@keyframes spin{from{transform:rotate(0)}to{transform:rotate(360deg)}}</style>');