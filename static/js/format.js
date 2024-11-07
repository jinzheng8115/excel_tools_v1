// 等待DOM加载完成后初始化事件监听
document.addEventListener('DOMContentLoaded', function() {
    // 文件处理相关的状态管理
    window.selectedFiles = window.selectedFiles || [];
    window.currentSheets = window.currentSheets || [];

    // 初始化事件监听
    const splitFileInput = document.getElementById('split-file');
    if (splitFileInput) {
        splitFileInput.addEventListener('change', handleSplitFileSelect);
    }

    const mergeFilesInput = document.getElementById('merge-files');
    if (mergeFilesInput) {
        mergeFilesInput.addEventListener('change', handleMergeFilesSelect);
    }

    // 初始化拖放区域
    initializeDragAndDrop();
});

// 初始化拖放功能
function initializeDragAndDrop() {
    const mergeUploadArea = document.getElementById('merge-upload-area');
    if (!mergeUploadArea) return;

    mergeUploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        mergeUploadArea.classList.add('drag-over');
    });

    mergeUploadArea.addEventListener('dragleave', (e) => {
        e.preventDefault();
        mergeUploadArea.classList.remove('drag-over');
    });

    mergeUploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        mergeUploadArea.classList.remove('drag-over');
        
        const files = Array.from(e.dataTransfer.files).filter(file => 
            file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
        );
        
        if (files.length > 0) {
            handleMergeFilesSelect({ target: { files } });
        }
    });
}

// 文件拆分相关函数
async function handleSplitFileSelect() {
    resetError('formatStatus');
    const file = document.getElementById('split-file').files[0];
    if (!file) return;

    const formData = new FormData();
    formData.append('file', file);

    try {
        const response = await fetch('/api/get-sheets', {
            method: 'POST',
            body: formData
        });
        const data = await response.json();

        if (data.success) {
            // 清空并重新填充工作表列表
            const sheetList = document.querySelector('.sheet-list');
            sheetList.innerHTML = '';
            window.currentSheets = data.sheets;

            data.sheets.forEach(sheet => {
                const sheetItem = document.createElement('div');
                sheetItem.className = 'sheet-item';
                sheetItem.innerHTML = `
                    <input type="checkbox" id="sheet-${sheet}" value="${sheet}" checked>
                    <label for="sheet-${sheet}">${sheet}</label>
                `;
                sheetList.appendChild(sheetItem);
            });

            // 显示工作表选择区域
            document.getElementById('sheet-selection').style.display = 'block';
        } else {
            showError(data.error || '获取工作表失败');
        }
    } catch (error) {
        showError('处理文件时发生错误: ' + error.message);
    }
}

// 全选工作表
function selectAllSheets() {
    document.querySelectorAll('.sheet-list input[type="checkbox"]')
        .forEach(checkbox => checkbox.checked = true);
}

// 取消全选工作表
function deselectAllSheets() {
    document.querySelectorAll('.sheet-list input[type="checkbox"]')
        .forEach(checkbox => checkbox.checked = false);
}

// 文件合并相关函数
function handleMergeFilesSelect(event) {
    const files = Array.from(event.target.files || []);
    if (files.length === 0) return;

    // 更新已选文件列表
    window.selectedFiles = [...window.selectedFiles, ...files];
    updateMergeUI();
}

// 更新合并UI
function updateMergeUI() {
    const filesList = document.getElementById('selected-files-list');
    if (!filesList) return;
    
    filesList.innerHTML = '';
    
    window.selectedFiles.forEach((file, index) => {
        const fileItem = document.createElement('div');
        fileItem.className = 'file-item';
        
        const fileName = document.createElement('div');
        fileName.className = 'file-name';
        fileName.textContent = file.name;
        
        const removeBtn = document.createElement('button');
        removeBtn.className = 'remove-file';
        removeBtn.textContent = '删除';
        removeBtn.onclick = () => removeFile(index);
        
        fileItem.appendChild(fileName);
        fileItem.appendChild(removeBtn);
        filesList.appendChild(fileItem);
    });
    
    // 更新上传提示和合并按钮
    const uploadTip = document.querySelector('.upload-tip');
    if (uploadTip) {
        uploadTip.textContent = window.selectedFiles.length > 0 ? 
            '继续添加文件或开始合并' : '点击或拖放文件到此处';
    }
}

// 删除选中的文件
function removeFile(index) {
    window.selectedFiles = window.selectedFiles.filter((_, i) => i !== index);
    updateMergeUI();
}

// 文件拆分处理
async function splitFile() {
    const file = document.getElementById('split-file').files[0];
    if (!file) {
        showError('请选择要拆分的文件');
        return;
    }

    const checkedSheets = Array.from(
        document.querySelectorAll('.sheet-list input[type="checkbox"]:checked')
    ).map(checkbox => checkbox.value);

    if (checkedSheets.length === 0) {
        showError('请至少选择一个工作表');
        return;
    }

    const status = document.getElementById('formatStatus');
    const processStats = document.getElementById('process-stats');
    
    status.textContent = '正在处理...';
    status.className = 'status-message status-processing';
    processStats.textContent = '';

    const formData = new FormData();
    formData.append('file', file);
    formData.append('splitAll', 'false');
    formData.append('sheets', JSON.stringify(checkedSheets));

    try {
        const response = await fetch('/api/split-excel', {
            method: 'POST',
            body: formData
        });
        const data = await response.json();

        if (data.success) {
            status.textContent = '处理完成！';
            status.className = 'status-message status-success';
            
            if (data.files && data.files.length > 0) {
                processStats.textContent = `成功拆分 ${data.files.length} 个工作表`;
                // 依次下载所有文件
                data.files.forEach((file, index) => {
                    setTimeout(() => {
                        window.location.href = file.downloadUrl;
                    }, index * 1000); // 每个文件间隔1秒下载
                });
            }
        } else {
            showError(data.error || '拆分失败');
        }
    } catch (error) {
        showError('处理文件时发生错误: ' + error.message);
    }
}

// 文件合并处理
async function mergeFiles() {
    if (window.selectedFiles.length === 0) {
        showError('请选择要合并的文件');
        return;
    }

    const status = document.getElementById('formatStatus');
    const processStats = document.getElementById('process-stats');
    
    status.textContent = '正在处理...';
    status.className = 'status-message status-processing';
    processStats.textContent = '';

    const formData = new FormData();
    window.selectedFiles.forEach(file => {
        formData.append('files[]', file);
    });

    const mergeMode = document.querySelector('input[name="merge-mode"]:checked').value;
    const mergeAll = document.getElementById('merge-all-sheets').checked;
    const addSource = document.getElementById('add-source-column').checked;
    const filename = document.getElementById('merged-filename').value.trim() || '合并文件';

    formData.append('mergeMode', mergeMode);
    formData.append('mergeAll', mergeAll.toString());
    formData.append('addSource', addSource.toString());
    formData.append('filename', filename);

    try {
        const response = await fetch('/api/merge-excel', {
            method: 'POST',
            body: formData
        });
        const data = await response.json();

        if (data.success) {
            status.textContent = '处理完成！';
            status.className = 'status-message status-success';
            processStats.textContent = `成功合并 ${window.selectedFiles.length} 个文件`;
            
            // 下载结果文件
            window.location.href = data.downloadUrl;
            
            // 清空已选文件列表
            window.selectedFiles = [];
            updateMergeUI();
        } else {
            showError(data.error || '合并失败');
        }
    } catch (error) {
        showError('处理文件时发生错误: ' + error.message);
    }
}

// 错误提示函数
function showError(message) {
    const status = document.getElementById('formatStatus');
    const processStats = document.getElementById('process-stats');
    
    status.textContent = message;
    status.className = 'status-message status-error';
    processStats.textContent = '';
}

// 重置错误提示
function resetError(statusId) {
    const status = document.getElementById(statusId);
    const processStats = document.getElementById('process-stats');
    
    status.textContent = '等待上传文件...';
    status.className = 'status-message';
    processStats.textContent = '';
} 