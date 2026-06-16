/**
 * 年报财务数据获取工具 - 双模块前端交互逻辑
 * 模块1: 年报数据爬取 (含行业均值)
 * 模块2: 公募基金管理人名录
 */

// ==================== 常量 ====================
const POLL_INTERVAL = 1500;

// 默认企业名单
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

// ==================== 模块抽象 ====================

/**
 * 创建一个独立的任务模块
 * @param {Object} cfg - 配置
 * @param {string} cfg.formId - 表单元素ID
 * @param {string} cfg.submitBtnId - 提交按钮ID
 * @param {string} cfg.apiUrl - 提交API地址
 * @param {Function} cfg.buildFormData - 构建FormData的函数，返回FormData或null（校验失败）
 * @param {string} cfg.resultId - 结果容器ID
 * @param {string} cfg.progressId - 进度区域ID
 * @param {string} cfg.progressLabelId - 进度标签ID
 * @param {string} cfg.progressPercentId - 进度百分比ID
 * @param {string} cfg.progressFillId - 进度条填充ID
 * @param {string} cfg.progressMessageId - 进度消息ID
 * @param {string} cfg.resultSectionId - 结果区域ID
 * @param {string} cfg.resultStatusId - 结果状态ID
 * @param {string} cfg.fileListId - 文件列表ID
 */
function createTaskModule(cfg) {
    const els = {
        form: document.getElementById(cfg.formId),
        submitBtn: document.getElementById(cfg.submitBtnId),
        result: document.getElementById(cfg.resultId),
        progress: document.getElementById(cfg.progressId),
        progressLabel: document.getElementById(cfg.progressLabelId),
        progressPercent: document.getElementById(cfg.progressPercentId),
        progressFill: document.getElementById(cfg.progressFillId),
        progressMessage: document.getElementById(cfg.progressMessageId),
        resultSection: document.getElementById(cfg.resultSectionId),
        resultStatus: document.getElementById(cfg.resultStatusId),
        fileList: document.getElementById(cfg.fileListId),
    };

    let pollTimer = null;

    function setDisabled(disabled) {
        els.submitBtn.disabled = disabled;
        if (disabled) {
            els.submitBtn.innerHTML = `
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="spinning">
                    <circle cx="12" cy="12" r="10" stroke-opacity="0.3"/>
                    <path d="M12 2a10 10 0 0 1 10 10" stroke-linecap="round"/>
                </svg>
                处理中...
            `;
        } else {
            els.submitBtn.innerHTML = `
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <polygon points="5 3 19 12 5 21 5 3"/>
                </svg>
                开始获取
            `;
        }
    }

    function showResult() {
        els.result.style.display = 'block';
        els.progress.style.display = 'block';
        els.resultSection.style.display = 'none';
        els.progressFill.style.width = '0%';
        els.progressPercent.textContent = '0%';
        els.progressMessage.textContent = '正在启动...';
    }

    function showSuccess(task) {
        els.progress.style.display = 'none';
        els.resultSection.style.display = 'block';
        els.resultStatus.className = 'result-status success';
        els.resultStatus.textContent = '任务完成！共生成 ' + task.files.length + ' 个文件';

        els.fileList.innerHTML = '';
        task.files.forEach((f) => {
            const sizeMB = (f.size / (1024 * 1024)).toFixed(2);
            const item = document.createElement('div');
            item.className = 'file-item';
            item.innerHTML = `
                <div class="file-info">
                    <div class="file-name">${escapeHtml(f.display_name || f.name)}</div>
                    <div class="file-size">${sizeMB > 0 ? sizeMB + ' MB' : (f.size ? (f.size / 1024).toFixed(0) + ' KB' : '')}</div>
                </div>
                <a class="download-btn" href="/download/${encodeURIComponent(f.name)}" download>
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                        <polyline points="7 10 12 15 17 10"/>
                        <line x1="12" y1="15" x2="12" y2="3"/>
                    </svg>
                    下载
                </a>
            `;
            els.fileList.appendChild(item);
        });
    }

    function showError(message) {
        els.progress.style.display = 'none';
        els.resultSection.style.display = 'block';
        els.resultStatus.className = 'result-status error';
        els.resultStatus.textContent = '错误: ' + message;
        els.fileList.innerHTML = '';
    }

    function updateProgress(task) {
        const current = task.progress.current;
        const total = task.progress.total;
        const pct = total > 0 ? Math.round((current / total) * 100) : 0;

        els.progressPercent.textContent = pct + '%';
        els.progressFill.style.width = pct + '%';
        els.progressMessage.textContent = task.progress.message;
        els.progressLabel.textContent = task.status === 'done' ? '完成' : '正在处理...';
    }

    function pollProgress(taskId) {

        pollTimer = setInterval(async () => {
            try {
                const resp = await fetch(`/api/progress/${taskId}`);
                const data = await resp.json();

                if (!data.success) {
                    clearInterval(pollTimer);
                    showError(data.error || '查询进度失败');
                    setDisabled(false);
                    return;
                }

                const task = data.task;
                updateProgress(task);

                if (task.status === 'done') {
                    clearInterval(pollTimer);
                    showSuccess(task);
                    setDisabled(false);
                } else if (task.status === 'error') {
                    clearInterval(pollTimer);
                    showError(task.error || '处理失败');
                    setDisabled(false);
                }
            } catch (err) {
                clearInterval(pollTimer);
                showError('网络错误: ' + err.message);
                setDisabled(false);
            }
        }, POLL_INTERVAL);
    }

    // 表单提交
    els.form.addEventListener('submit', async (e) => {
        e.preventDefault();

        const formData = cfg.buildFormData();
        if (!formData) return;

        setDisabled(true);
        showResult();

        try {
            const resp = await fetch(cfg.apiUrl, {
                method: 'POST',
                body: formData,
            });
            const data = await resp.json();

            if (!data.success) {
                showError(data.error || '请求失败');
                setDisabled(false);
                return;
            }

            pollProgress(data.task_id);
        } catch (err) {
            showError('网络错误: ' + err.message);
            setDisabled(false);
        }
    });

    return {
        setDisabled,
        showResult,
        showError,
        showSuccess,
        pollProgress,
    };
}

// ==================== 模块1: 年报数据爬取 ====================

const startYearEl = document.getElementById('startYear');
const endYearEl = document.getElementById('endYear');
const textInput = document.getElementById('textInput');
const fileInput = document.getElementById('fileInput');
const uploadArea = document.getElementById('uploadArea');
const fileInfoDiv = document.getElementById('fileInfo');
const fileNameSpan = document.getElementById('fileName');
const fileRemoveBtn = document.getElementById('fileRemove');
const uploadContent = document.querySelector('.upload-area .upload-content');
const industryAvgToggle = document.getElementById('industryAvgToggle');
const mergeReportsToggle = document.getElementById('mergeReportsToggle');

function initYearSelectors() {
    const currentYear = new Date().getFullYear();
    const startYear = 1990;

    for (let y = currentYear; y >= startYear; y--) {
        const opt1 = document.createElement('option');
        opt1.value = y;
        opt1.textContent = y + '年';
        startYearEl.appendChild(opt1);

        const opt2 = document.createElement('option');
        opt2.value = y;
        opt2.textContent = y + '年';
        endYearEl.appendChild(opt2);
    }

    startYearEl.value = 2021;
    endYearEl.value = 2025;
}

function validateYearRange() {
    const start = parseInt(startYearEl.value);
    const end = parseInt(endYearEl.value);
    if (start && end && start > end) {
        [startYearEl.value, endYearEl.value] = [endYearEl.value, startYearEl.value];
    }
}

// 文件上传处理
function showFileInfo(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['txt', 'csv'].includes(ext)) {
        alert('仅支持 .txt 和 .csv 文件');
        fileInput.value = '';
        return;
    }
    uploadContent.style.display = 'none';
    fileInfoDiv.style.display = 'flex';
    fileNameSpan.textContent = file.name;
    // 上传文件时清空文本输入框，避免默认名单干扰
    textInput.value = '';
}

function removeFile() {
    fileInput.value = '';
    uploadContent.style.display = '';
    fileInfoDiv.style.display = 'none';
    // 恢复默认名单
    if (!textInput.value.trim()) {
        textInput.value = DEFAULT_COMPANIES;
    }
}

uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.style.borderColor = 'var(--accent)';
    uploadArea.style.background = 'rgba(79, 195, 247, 0.08)';
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.style.borderColor = '';
    uploadArea.style.background = '';
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.style.borderColor = '';
    uploadArea.style.background = '';
    const file = e.dataTransfer.files[0];
    if (file) {
        fileInput.files = e.dataTransfer.files;
        showFileInfo(file);
    }
});

fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) showFileInfo(file);
});
fileRemoveBtn.addEventListener('click', removeFile);
startYearEl.addEventListener('change', validateYearRange);
endYearEl.addEventListener('change', validateYearRange);

const annualModule = createTaskModule({
    formId: 'annualForm',
    submitBtnId: 'annualSubmitBtn',
    apiUrl: '/api/fetch',
    buildFormData: () => {
        const startYear = parseInt(startYearEl.value);
        const endYear = parseInt(endYearEl.value);
        const textVal = textInput.value.trim();
        const file = fileInput.files[0];

        if (!startYear || !endYear) {
            alert('请选择年份范围');
            return null;
        }
        if (!textVal && !file) {
            alert('请输入企业名单或上传文件');
            return null;
        }

        const formData = new FormData();
        formData.append('start_year', startYear);
        formData.append('end_year', endYear);
        formData.append('text_input', textVal);
        if (file) formData.append('file', file);
        formData.append('industry_avg', industryAvgToggle.checked ? '1' : '0');
        formData.append('merge_reports', mergeReportsToggle.checked ? '1' : '0');
        // 年报模块不再传 amac_list
        return formData;
    },
    resultId: 'annualResult',
    progressId: 'annualProgress',
    progressLabelId: 'annualProgressLabel',
    progressPercentId: 'annualProgressPercent',
    progressFillId: 'annualProgressFill',
    progressMessageId: 'annualProgressMessage',
    resultSectionId: 'annualResultSection',
    resultStatusId: 'annualResultStatus',
    fileListId: 'annualFileList',
});

// 年报重置按钮
document.getElementById('annualResetBtn').addEventListener('click', () => {
    document.getElementById('annualForm').reset();
    textInput.value = DEFAULT_COMPANIES;
    removeFile();
    industryAvgToggle.checked = false;
    mergeReportsToggle.checked = false;
    document.getElementById('annualResult').style.display = 'none';
    startYearEl.value = 2021;
    endYearEl.value = 2025;
});

// ==================== 模块2: 公募基金管理人名录 ====================

const amacModule = createTaskModule({
    formId: 'amacForm',
    submitBtnId: 'amacSubmitBtn',
    apiUrl: '/api/amac',
    buildFormData: () => new FormData(),
    resultId: 'amacResult',
    progressId: 'amacProgress',
    progressLabelId: 'amacProgressLabel',
    progressPercentId: 'amacProgressPercent',
    progressFillId: 'amacProgressFill',
    progressMessageId: 'amacProgressMessage',
    resultSectionId: 'amacResultSection',
    resultStatusId: 'amacResultStatus',
    fileListId: 'amacFileList',
});

// ==================== 模块3: 深交所日度概况 ====================

const szseYearEl = document.getElementById('szseYear');

function initSzseYearSelector() {
    const currentYear = new Date().getFullYear();
    for (let y = currentYear; y >= 2000; y--) {
        const opt = document.createElement('option');
        opt.value = y;
        opt.textContent = y + '年';
        szseYearEl.appendChild(opt);
    }
    szseYearEl.value = currentYear;
}

const szseModule = createTaskModule({
    formId: 'szseForm',
    submitBtnId: 'szseSubmitBtn',
    apiUrl: '/api/szse',
    buildFormData: () => {
        const formData = new FormData();
        formData.append('year', szseYearEl.value);
        return formData;
    },
    resultId: 'szseResult',
    progressId: 'szseProgress',
    progressLabelId: 'szseProgressLabel',
    progressPercentId: 'szseProgressPercent',
    progressFillId: 'szseProgressFill',
    progressMessageId: 'szseProgressMessage',
    resultSectionId: 'szseResultSection',
    resultStatusId: 'szseResultStatus',
    fileListId: 'szseFileList',
});

// ==================== 模块4: 分红公告查询 ====================

const dividendStartYearEl = document.getElementById('dividendStartYear');
const dividendEndYearEl = document.getElementById('dividendEndYear');
const dividendTextInput = document.getElementById('dividendTextInput');
const dividendFileInput = document.getElementById('dividendFileInput');
const dividendUploadArea = document.getElementById('dividendUploadArea');
const dividendFileInfo = document.getElementById('dividendFileInfo');
const dividendFileName = document.getElementById('dividendFileName');
const dividendFileRemove = document.getElementById('dividendFileRemove');
const dividendUploadContent = document.querySelector('#dividendUploadArea .upload-content');

function initDividendYearSelectors() {
    const currentYear = new Date().getFullYear();
    const startYear = 2000;

    for (let y = currentYear; y >= startYear; y--) {
        const opt1 = document.createElement('option');
        opt1.value = y;
        opt1.textContent = y + '年';
        dividendStartYearEl.appendChild(opt1);

        const opt2 = document.createElement('option');
        opt2.value = y;
        opt2.textContent = y + '年';
        dividendEndYearEl.appendChild(opt2);
    }

    dividendStartYearEl.value = 2024;
    dividendEndYearEl.value = 2025;
}

function validateDividendYearRange() {
    const start = parseInt(dividendStartYearEl.value);
    const end = parseInt(dividendEndYearEl.value);
    if (start && end && start > end) {
        [dividendStartYearEl.value, dividendEndYearEl.value] = [dividendEndYearEl.value, dividendStartYearEl.value];
    }
}

// 分红文件上传处理
function showDividendFileInfo(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['txt', 'csv'].includes(ext)) {
        alert('仅支持 .txt 和 .csv 文件');
        dividendFileInput.value = '';
        return;
    }
    dividendUploadContent.style.display = 'none';
    dividendFileInfo.style.display = 'flex';
    dividendFileName.textContent = file.name;
}

function removeDividendFile() {
    dividendFileInput.value = '';
    dividendUploadContent.style.display = '';
    dividendFileInfo.style.display = 'none';
}

dividendUploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    dividendUploadArea.style.borderColor = 'var(--dividend-accent, #AB47BC)';
    dividendUploadArea.style.background = 'rgba(171, 71, 188, 0.08)';
});

dividendUploadArea.addEventListener('dragleave', () => {
    dividendUploadArea.style.borderColor = '';
    dividendUploadArea.style.background = '';
});

dividendUploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    dividendUploadArea.style.borderColor = '';
    dividendUploadArea.style.background = '';
    const file = e.dataTransfer.files[0];
    if (file) {
        dividendFileInput.files = e.dataTransfer.files;
        showDividendFileInfo(file);
    }
});

dividendFileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) showDividendFileInfo(file);
});
dividendFileRemove.addEventListener('click', removeDividendFile);
dividendStartYearEl.addEventListener('change', validateDividendYearRange);
dividendEndYearEl.addEventListener('change', validateDividendYearRange);

const dividendModule = createTaskModule({
    formId: 'dividendForm',
    submitBtnId: 'dividendSubmitBtn',
    apiUrl: '/api/dividend',
    buildFormData: () => {
        const startYear = parseInt(dividendStartYearEl.value);
        const endYear = parseInt(dividendEndYearEl.value);
        const textVal = dividendTextInput.value.trim();
        const file = dividendFileInput.files[0];

        if (!startYear || !endYear) {
            alert('请选择年度范围');
            return null;
        }
        if (!textVal && !file) {
            alert('请输入企业名单或上传文件');
            return null;
        }

        const formData = new FormData();
        formData.append('start_year', startYear);
        formData.append('end_year', endYear);
        formData.append('text_input', textVal);
        if (file) formData.append('file', file);
        return formData;
    },
    resultId: 'dividendResult',
    progressId: 'dividendProgress',
    progressLabelId: 'dividendProgressLabel',
    progressPercentId: 'dividendProgressPercent',
    progressFillId: 'dividendProgressFill',
    progressMessageId: 'dividendProgressMessage',
    resultSectionId: 'dividendResultSection',
    resultStatusId: 'dividendResultStatus',
    fileListId: 'dividendFileList',
});

// ==================== 工具函数 ====================

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// ==================== 启动 ====================
initYearSelectors();
initSzseYearSelector();
initDividendYearSelectors();

// 旋转动画样式
const spinStyle = document.createElement('style');
spinStyle.textContent = `
    .spinning { animation: spin 1s linear infinite; }
    @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
`;
document.head.appendChild(spinStyle);