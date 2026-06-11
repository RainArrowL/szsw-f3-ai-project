/**
 * 年报财务数据获取工具 - 前端交互逻辑
 * 深色科技金融风格
 */

// ==================== 常量 ====================
const POLL_INTERVAL = 1500; // 轮询间隔（毫秒）
const MAX_POLL_TIME = 10 * 60 * 1000; // 最大轮询时间（10分钟）

// ==================== DOM 元素 ====================
const startYearEl = document.getElementById('startYear');
const endYearEl = document.getElementById('endYear');
const fetchForm = document.getElementById('fetchForm');
const textInput = document.getElementById('textInput');
const fileInput = document.getElementById('fileInput');
const submitBtn = document.getElementById('submitBtn');
const resetBtn = document.getElementById('resetBtn');
const uploadArea = document.getElementById('uploadArea');
const fileInfoDiv = document.getElementById('fileInfo');
const fileNameSpan = document.getElementById('fileName');
const fileRemoveBtn = document.getElementById('fileRemove');
const resultCard = document.getElementById('resultCard');
const progressSection = document.getElementById('progressSection');
const progressLabelEl = document.getElementById('progressLabel');
const progressPercentEl = document.getElementById('progressPercent');
const progressMessageEl = document.getElementById('progressMessage');
const progressFill = document.getElementById('progressFill');
const resultSection = document.getElementById('resultSection');
const resultStatus = document.getElementById('resultStatus');
const fileListEl = document.getElementById('fileList');
const uploadContent = document.querySelector('.upload-content');
const uploadHint = document.querySelector('.upload-hint');

const industryAvgToggle = document.getElementById('industryAvgToggle');

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

// ==================== 初始化 ====================

/**
 * 初始化年份选择器
 */
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

    // 默认值：最近5年
    startYearEl.value = currentYear - 4;
    endYearEl.value = currentYear;
}

/**
 * 验证年份范围
 */
function validateYearRange() {
    const start = parseInt(startYearEl.value);
    const end = parseInt(endYearEl.value);
    if (start && end && start > end) {
        // 自动互换
        [startYearEl.value, endYearEl.value] = [endYearEl.value, startYearEl.value];
    }
}

// ==================== 文件上传处理 ====================

function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;
    showFileInfo(file);
}

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
}

function removeFile() {
    fileInput.value = '';
    uploadContent.style.display = 'block';
    fileInfoDiv.style.display = 'none';
}

// 拖拽支持
uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('drag-over');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('drag-over');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) {
        fileInput.files = e.dataTransfer.files;
        showFileInfo(file);
    }
});

fileInput.addEventListener('change', handleFileSelect);
fileRemoveBtn.addEventListener('click', removeFile);
startYearEl.addEventListener('change', validateYearRange);
endYearEl.addEventListener('change', validateYearRange);

// ==================== 表单提交 ====================

fetchForm.addEventListener('submit', async (e) => {
    e.preventDefault();

    const startYear = parseInt(startYearEl.value);
    const endYear = parseInt(endYearEl.value);
    const textVal = textInput.value.trim();
    const file = fileInput.files[0];

    // 校验
    if (!startYear || !endYear) {
        alert('请选择年份范围');
        return;
    }
    if (!textVal && !file) {
        alert('请输入企业名单或上传文件');
        return;
    }

    // 禁用表单
    setFormDisabled(true);

    // 构建 FormData
    const formData = new FormData();
    formData.append('start_year', startYear);
    formData.append('end_year', endYear);
    formData.append('text_input', textVal);
    if (file) {
        formData.append('file', file);
    }
    formData.append('industry_avg', industryAvgToggle.checked ? '1' : '0');

    // 显示结果区域
    showResultCard('processing');

    try {
        const resp = await fetch('/api/fetch', {
            method: 'POST',
            body: formData,
        });
        const data = await resp.json();

        if (!data.success) {
            showError(data.error || '请求失败');
            setFormDisabled(false);
            return;
        }

        // 开始轮询进度
        pollProgress(data.task_id);

    } catch (err) {
        showError('网络错误: ' + err.message);
        setFormDisabled(false);
    }
});

// ==================== 进度轮询 ====================

let pollTimer = null;
let pollStartedAt = 0;

function pollProgress(taskId) {
    pollStartedAt = Date.now();

    pollTimer = setInterval(async () => {
        try {
            const resp = await fetch(`/api/progress/${taskId}`);
            const data = await resp.json();

            if (!data.success) {
                clearInterval(pollTimer);
                showError(data.error || '查询进度失败');
                setFormDisabled(false);
                return;
            }

            const task = data.task;
            updateProgress(task);

            if (task.status === 'done') {
                clearInterval(pollTimer);
                showSuccess(task);
                setFormDisabled(false);
            } else if (task.status === 'error') {
                clearInterval(pollTimer);
                showError(task.error || '处理失败');
                setFormDisabled(false);
            }

            // 超时保护
            if (Date.now() - pollStartedAt > MAX_POLL_TIME) {
                clearInterval(pollTimer);
                showError('处理超时，请检查服务器状态');
                setFormDisabled(false);
            }

        } catch (err) {
            clearInterval(pollTimer);
            showError('网络错误: ' + err.message);
            setFormDisabled(false);
        }
    }, POLL_INTERVAL);
}

function updateProgress(task) {
    const current = task.progress.current;
    const total = task.progress.total;
    const pct = total > 0 ? Math.round((current / total) * 100) : 0;

    progressPercentEl.textContent = pct + '%';
    progressFill.style.width = pct + '%';
    progressMessageEl.textContent = task.progress.message;

    if (task.status === 'done') {
        document.querySelector('.progress-fill').style.background =
            'linear-gradient(90deg, var(--success), var(--accent))';
    }
}

function showResultCard(status) {
    resultCard.style.display = 'block';
    progressSection.style.display = 'block';
    resultSection.style.display = 'none';
    progressFill.style.width = '0%';
    progressPercentEl.textContent = '0%';
    progressMessageEl.textContent = '正在启动...';
    // 滚动到结果区域
    resultCard.scrollIntoView({ behavior: 'smooth' });
}

// ==================== 结果展示 ====================

function showSuccess(task) {
    progressSection.style.display = 'none';
    resultSection.style.display = 'block';
    resultStatus.className = 'result-status success';
    resultStatus.textContent = '任务完成！共生成 ' + task.files.length + ' 个文件';

    fileListEl.innerHTML = '';
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
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                    <polyline points="7 10 12 15 17 10"/>
                    <line x1="12" y1="15" x2="12" y2="3"/>
                </svg>
                下载
            </a>
        `;
        fileListEl.appendChild(item);
    });

    resultCard.scrollIntoView({ behavior: 'smooth' });
}

function showError(message) {
    progressSection.style.display = 'none';
    resultSection.style.display = 'block';
    resultStatus.className = 'result-status error';
    resultStatus.textContent = '错误: ' + message;
    fileListEl.innerHTML = '';
    resultCard.scrollIntoView({ behavior: 'smooth' });
}

// ==================== 工具函数 ====================

function setFormDisabled(disabled) {
    submitBtn.disabled = disabled;
    if (disabled) {
        submitBtn.innerHTML = `
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="spinning">
                <circle cx="12" cy="12" r="10" stroke-opacity="0.3"/>
                <path d="M12 2a10 10 0 0 1 10 10" stroke-linecap="round"/>
            </svg>
            处理中...
        `;
    } else {
        submitBtn.innerHTML = `
            <span class="btn-icon">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <polygon points="5 3 19 12 5 21 5 3"/>
                </svg>
            </span>
            开始获取数据
        `;
    }
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// ==================== 重置按钮 ====================

resetBtn.addEventListener('click', () => {
    fetchForm.reset();
    textInput.value = DEFAULT_COMPANIES;
    removeFile();
    industryAvgToggle.checked = false;
    resultCard.style.display = 'none';
    // 重置年份默认值
    const currentYear = new Date().getFullYear();
    startYearEl.value = currentYear - 4;
    endYearEl.value = currentYear;
});

// ==================== 启动 ====================
initYearSelectors();

// 添加旋转动画样式
const spinStyle = document.createElement('style');
spinStyle.textContent = `
    .spinning { animation: spin 1s linear infinite; }
    @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
`;
document.head.appendChild(spinStyle);