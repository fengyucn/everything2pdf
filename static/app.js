/**
 * Everything to PDF - 前端交互逻辑
 */

// 文件列表
let files = [];

// DOM元素
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const fileListSection = document.getElementById('file-list-section');
const fileList = document.getElementById('file-list');
const fileCount = document.getElementById('file-count');
const clearBtn = document.getElementById('clear-btn');
const convertSection = document.getElementById('convert-section');
const convertBtn = document.getElementById('convert-btn');
const resultSection = document.getElementById('result-section');
const progressBar = document.getElementById('progress-bar');
const progressFill = document.getElementById('progress-fill');
const resultMessage = document.getElementById('result-message');
const downloadLink = document.getElementById('download-link');

// 文件类型图标
const fileIcons = {
    'image': '&#128247;',  // 相机
    'office': '&#128196;', // 文档
    'pdf': '&#128213;',    // 书本
    'unknown': '&#128196;' // 默认
};

// 格式化文件大小
function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

// 更新界面显示
function updateUI() {
    const count = files.length;
    fileCount.textContent = `(${count})`;
    
    if (count > 0) {
        fileListSection.style.display = 'block';
        convertSection.style.display = 'block';
    } else {
        fileListSection.style.display = 'none';
        convertSection.style.display = 'none';
    }
    
    renderFileList();
}

// 渲染文件列表
function renderFileList() {
    fileList.innerHTML = '';
    
    files.forEach((file, index) => {
        const li = document.createElement('li');
        li.className = 'file-item';
        li.draggable = true;
        li.dataset.index = index;
        
        li.innerHTML = `
            <span class="drag-handle">&#9776;</span>
            <span class="file-icon">${fileIcons[file.type] || fileIcons.unknown}</span>
            <div class="file-info">
                <div class="file-name">${escapeHtml(file.name)}</div>
                <div class="file-meta">${file.type} - ${formatSize(file.size)}</div>
            </div>
            <button class="remove-btn" data-id="${file.id}" title="移除">&times;</button>
        `;
        
        // 拖拽事件
        li.addEventListener('dragstart', handleDragStart);
        li.addEventListener('dragend', handleDragEnd);
        li.addEventListener('dragover', handleDragOver);
        li.addEventListener('drop', handleDrop);
        
        // 删除按钮
        li.querySelector('.remove-btn').addEventListener('click', (e) => {
            e.stopPropagation();
            removeFile(file.id);
        });
        
        fileList.appendChild(li);
    });
}

// HTML转义
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// 拖拽排序相关
let draggedItem = null;

function handleDragStart(e) {
    draggedItem = this;
    this.classList.add('dragging');
    e.dataTransfer.effectAllowed = 'move';
}

function handleDragEnd(e) {
    this.classList.remove('dragging');
    document.querySelectorAll('.file-item').forEach(item => {
        item.classList.remove('drag-over');
    });
}

function handleDragOver(e) {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    
    if (this !== draggedItem) {
        this.classList.add('drag-over');
    }
}

function handleDrop(e) {
    e.preventDefault();
    
    if (this !== draggedItem) {
        const fromIndex = parseInt(draggedItem.dataset.index);
        const toIndex = parseInt(this.dataset.index);
        
        // 重新排序
        const [moved] = files.splice(fromIndex, 1);
        files.splice(toIndex, 0, moved);
        
        renderFileList();
    }
    
    this.classList.remove('drag-over');
}

// 上传文件
async function uploadFiles(fileObjects) {
    const formData = new FormData();
    
    for (const file of fileObjects) {
        formData.append('files', file);
    }
    
    try {
        const response = await fetch('/api/upload', {
            method: 'POST',
            body: formData
        });
        
        if (response.ok) {
            const data = await response.json();
            files.push(...data.files);
            updateUI();
        } else {
            const error = await response.json();
            alert('上传失败: ' + (error.error || '未知错误'));
        }
    } catch (err) {
        alert('上传失败: ' + err.message);
    }
}

// 删除文件
async function removeFile(fileId) {
    try {
        await fetch(`/api/remove/${fileId}`, { method: 'DELETE' });
        files = files.filter(f => f.id !== fileId);
        updateUI();
    } catch (err) {
        console.error('删除失败:', err);
    }
}

// 清空所有文件
async function clearFiles() {
    try {
        await fetch('/api/clear', { method: 'POST' });
        files = [];
        updateUI();
        hideResult();
    } catch (err) {
        console.error('清空失败:', err);
    }
}

// 显示结果
function showResult(type, message) {
    resultSection.style.display = 'block';
    resultMessage.className = 'result-message ' + type;
    resultMessage.textContent = message;
}

// 隐藏结果
function hideResult() {
    resultSection.style.display = 'none';
    progressBar.style.display = 'none';
    downloadLink.style.display = 'none';
}

// 转换文件
async function convertFiles() {
    if (files.length === 0) {
        alert('请先添加文件');
        return;
    }
    
    // 显示进度
    resultSection.style.display = 'block';
    progressBar.style.display = 'block';
    progressFill.style.width = '0%';
    downloadLink.style.display = 'none';
    showResult('processing', '正在转换中，请稍候...');
    
    convertBtn.disabled = true;
    
    // 模拟进度
    let progress = 0;
    const progressInterval = setInterval(() => {
        progress += Math.random() * 15;
        if (progress > 90) progress = 90;
        progressFill.style.width = progress + '%';
    }, 200);
    
    try {
        const response = await fetch('/api/convert', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                file_ids: files.map(f => f.id)
            })
        });
        
        clearInterval(progressInterval);
        progressFill.style.width = '100%';
        
        if (response.ok) {
            const data = await response.json();
            showResult('success', data.message);
            downloadLink.href = data.download_url;
            downloadLink.style.display = 'inline-block';
        } else {
            const error = await response.json();
            showResult('error', '转换失败: ' + (error.error || '未知错误'));
        }
    } catch (err) {
        clearInterval(progressInterval);
        showResult('error', '转换失败: ' + err.message);
    } finally {
        convertBtn.disabled = false;
        setTimeout(() => {
            progressBar.style.display = 'none';
        }, 500);
    }
}

// 事件绑定

// 拖拽上传
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
});

dropZone.addEventListener('dragleave', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    
    const droppedFiles = e.dataTransfer.files;
    if (droppedFiles.length > 0) {
        uploadFiles(droppedFiles);
    }
});

// 点击上传区域也触发文件选择
dropZone.addEventListener('click', (e) => {
    if (e.target === dropZone || e.target.closest('.drop-zone-content')) {
        if (!e.target.closest('.upload-btn')) {
            fileInput.click();
        }
    }
});

// 文件选择
fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        uploadFiles(e.target.files);
        e.target.value = ''; // 清空以便再次选择相同文件
    }
});

// 清空按钮
clearBtn.addEventListener('click', clearFiles);

// 转换按钮
convertBtn.addEventListener('click', convertFiles);

// 初始化
updateUI();
