// 使用window对象来存储状态，避免重复声明
window.importModule = {
    currentFile: null,
    previewData: null,
    selectedFiles: []
};

// 等待DOM加载完成后初始化事件监听
document.addEventListener('DOMContentLoaded', function() {
    console.log('初始化导入模块事件监听器');
    
    // 监听文件选择变化
    const importFileInput = document.getElementById('import-file');
    if (importFileInput) {
        console.log('找到文件输入框');
        importFileInput.addEventListener('change', function(event) {
            console.log('文件选择事件触发');
            handleFileSelect(event);
        });
    }

    // 监听导入类型变化
    document.querySelectorAll('input[name="import-type"]').forEach(radio => {
        radio.addEventListener('change', function(event) {
            const fileInput = document.getElementById('import-file');
            const folderInput = document.getElementById('import-folder');
            const uploadTip = document.querySelector('.upload-tip');
            
            if (event.target.value === 'files') {
                fileInput.style.display = 'block';
                folderInput.style.display = 'none';
                uploadTip.textContent = '点击或拖放 TXT/CSV 文件到此处';
            } else {
                fileInput.style.display = 'none';
                folderInput.style.display = 'block';
                uploadTip.textContent = '点击选择文件夹';
            }
            
            // 重置文件选择
            fileInput.value = '';
            folderInput.value = '';
            window.importModule.selectedFiles = [];
            hidePreview();
            
            // 隐藏文件信息
            const fileInfo = document.getElementById('file-info');
            if (fileInfo) {
                fileInfo.style.display = 'none';
            }
        });
    });

    // 添加拖放处理
    const uploadArea = document.getElementById('import-upload-area');
    if (uploadArea) {
        uploadArea.addEventListener('dragover', function(e) {
            e.preventDefault();
            e.stopPropagation();
            this.classList.add('drag-over');
        });

        uploadArea.addEventListener('dragleave', function(e) {
            e.preventDefault();
            e.stopPropagation();
            this.classList.remove('drag-over');
        });

        uploadArea.addEventListener('drop', function(e) {
            e.preventDefault();
            e.stopPropagation();
            this.classList.remove('drag-over');
            
            const files = Array.from(e.dataTransfer.files).filter(file => 
                file.name.endsWith('.txt') || file.name.endsWith('.csv')
            );
            
            if (files.length > 0) {
                handleFileSelect({ target: { files } });
            }
        });

        // 添加点击处理
        uploadArea.addEventListener('click', function() {
            const importType = document.querySelector('input[name="import-type"]:checked').value;
            const input = document.getElementById(importType === 'files' ? 'import-file' : 'import-folder');
            if (input) {
                input.click();
            }
        });
    }

    // 监听文件和文件夹选择
    const fileInput = document.getElementById('import-file');
    const folderInput = document.getElementById('import-folder');
    
    if (fileInput) {
        fileInput.addEventListener('change', handleFileSelect);
    }
    
    if (folderInput) {
        folderInput.addEventListener('change', handleFileSelect);
    }

    // 监听自动分列选项变化
    const autoSplitCheckbox = document.getElementById('auto-split');
    if (autoSplitCheckbox) {
        autoSplitCheckbox.addEventListener('change', function() {
            const manualDelimiter = document.getElementById('manual-delimiter');
            if (manualDelimiter) {
                manualDelimiter.style.display = this.checked ? 'none' : 'block';
                if (!this.checked) {
                    // 取消自动分列时，聚焦到分隔符输入框
                    document.getElementById('custom-delimiter')?.focus();
                }
            }
        });
    }

    // 监听分隔符选择变化
    const delimiterSelect = document.getElementById('delimiter');
    if (delimiterSelect) {
        delimiterSelect.addEventListener('change', function() {
            const customDelimiter = document.getElementById('custom-delimiter');
            if (customDelimiter) {
                customDelimiter.style.display = this.value === 'custom' ? 'block' : 'none';
                if (this.value === 'custom') {
                    customDelimiter.focus();
                }
            }
        });
    }
});

// 添加 formatFileSize 函数
function formatFileSize(bytes) {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// 处理文件选择
function handleFileSelect(event) {
    console.log('handleFileSelect 函数被调用');
    
    // 获取所有txt和csv文件
    const files = Array.from(event.target.files).filter(file => {
        // 如果是文件夹选择，需要检查文件路径
        if (event.target.webkitdirectory) {
            return file.name.endsWith('.txt') || file.name.endsWith('.csv');
        }
        return file.name.endsWith('.txt') || file.name.endsWith('.csv');
    });
    
    if (!files.length) {
        showError('没有找到支持的文件格式');
        return;
    }

    console.log('选择的文件数量:', files.length);
    window.importModule.selectedFiles = files;

    // 更新文件信息显示
    const fileInfo = document.getElementById('file-info');
    const fileList = document.getElementById('file-list');
    
    if (fileInfo && fileList) {
        // 更新文件列表
        fileList.innerHTML = files.map((file, index) => `
            <div class="file-item">
                <div class="file-info-content">
                    <span class="file-name">${file.webkitRelativePath || file.name}</span>
                    <span class="file-size">${formatFileSize(file.size)}</span>
                </div>
                <button class="delete-btn" onclick="removeFile(${index})" title="删除此文件">
                    <span class="delete-icon">×</span>
                </button>
            </div>
        `).join('');
        
        // 显示文件信息区域
        fileInfo.style.display = 'block';
    }

    // 更新上传提示
    const uploadTip = document.querySelector('.upload-tip');
    if (uploadTip) {
        uploadTip.textContent = `已选择 ${files.length} 个文件`;
    }

    // 显示导入选项
    const importOptions = document.getElementById('import-options');
    if (importOptions) {
        importOptions.style.display = 'block';
    }

    // 启用预览按钮
    const previewBtn = document.getElementById('preview-btn');
    if (previewBtn) {
        previewBtn.disabled = false;
    }
}

// 处理分隔符选择变化
function handleDelimiterChange(value) {
    const customDelimiterContainer = document.getElementById('custom-delimiter-container');
    const customDelimiter = document.getElementById('custom-delimiter');
    
    if (customDelimiterContainer && customDelimiter) {
        customDelimiterContainer.style.display = value === 'custom' ? 'block' : 'none';
        if (value === 'custom') {
            customDelimiter.focus();
        }
    }
}

// 修改 setupUploadAreaClick 函数
function setupUploadAreaClick(uploadArea, importType) {
    // 移除之前的点击事件监听器
    uploadArea.onclick = null;
    
    // 添加新的点击事件监听器
    uploadArea.addEventListener('click', function(e) {
        e.preventDefault();
        e.stopPropagation();
        
        console.log('上传区域被点击，当前模式:', importType);
        
        // 根据当前模式获取正确的输入元素
        const input = document.getElementById(importType === 'files' ? 'import-file' : 'import-folder');
            
        if (input) {
            console.log('准备触发文件选择，input元素:', input.id);
            
            // 确保输入框可见
            input.style.display = 'block';
            input.style.position = 'absolute';
            input.style.left = '-9999px';  // 移出可视区域但保持可点击
            
            // 设置正确的属性
            if (importType === 'folder') {
                input.setAttribute('webkitdirectory', '');
                input.setAttribute('directory', '');
                input.removeAttribute('multiple');
            } else {
                input.removeAttribute('webkitdirectory');
                input.removeAttribute('directory');
                input.setAttribute('multiple', '');
                input.accept = '.txt,.csv';
            }
            
            // 触发点击
            input.click();
            
            // 点击后恢复原始状态
            setTimeout(() => {
                input.style.display = importType === 'files' ? '' : 'none';
                input.style.position = '';
                input.style.left = '';
            }, 100);
        } else {
            console.error('未找到对应的输入元素');
        }
    });
    
    // 重新添加拖放事件处理
    setupDragAndDrop(uploadArea);
}

// 修改 handleImportTypeChange 函数
function handleImportTypeChange(event) {
    const fileInput = document.getElementById('import-file');
    const folderInput = document.getElementById('import-folder');
    const uploadArea = document.getElementById('import-upload-area');
    
    console.log('切换到:', event.target.value);
    
    if (event.target.value === 'files') {
        console.log('切换到文件选择模式');
        fileInput.style.display = '';
        folderInput.style.display = 'none';
        fileInput.removeAttribute('webkitdirectory');
        fileInput.removeAttribute('directory');
        fileInput.setAttribute('multiple', '');
        fileInput.accept = '.txt,.csv';
        uploadArea.querySelector('.upload-tip').textContent = '点击或拖放 TXT/CSV 文件到此处';
    } else {
        console.log('切换到文件夹选择模式');
        fileInput.style.display = 'none';
        folderInput.style.display = '';
        folderInput.setAttribute('webkitdirectory', '');
        folderInput.setAttribute('directory', '');
        folderInput.removeAttribute('multiple');
        uploadArea.querySelector('.upload-tip').textContent = '点击选择文件夹';
    }
    
    // 重置文件选择
    fileInput.value = '';
    folderInput.value = '';
    window.importModule.selectedFiles = [];
    hidePreview();
    
    // 隐藏文件信息
    const fileInfo = document.getElementById('file-info');
    if (fileInfo) {
        fileInfo.style.display = 'none';
    }
}

// 添加预览数据处理函数
window.handlePreviewData = async function() {
    if (!window.importModule.selectedFiles?.length) {
        showError('请先选择文件');
        return;
    }

    try {
        showLoading('正在加载预览...');
        
        const firstFile = window.importModule.selectedFiles[0];
        console.log(`预览第一个文件: ${firstFile.name}`);

        const formData = new FormData();
        formData.append('file', firstFile);
        formData.append('sourceType', 'text');
        formData.append('startRow', document.getElementById('start-row').value);
        formData.append('headerRow', document.getElementById('header-row').value);

        // 添加分隔符设置
        const autoSplit = document.getElementById('auto-split').checked;
        formData.append('autoSplit', autoSplit.toString());
        
        if (!autoSplit) {
            // 直接获取自定义分隔符的值
            const customDelimiter = document.getElementById('custom-delimiter');
            if (customDelimiter && customDelimiter.value) {
                formData.append('delimiter', customDelimiter.value);
            }
        }

        console.log('发送预览请求...');
        const response = await fetch('/api/preview-data', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();
        console.log('预览响应:', result);

        if (!response.ok) {
            throw new Error(result.error || '预览失败');
        }
        
        if (!result.data) {
            throw new Error('预览数据格式错误');
        }

        window.importModule.previewData = result.data;
        updatePreviewTable(result.data);
        document.getElementById('data-preview').style.display = 'block';
        document.getElementById('import-btn').disabled = false;

    } catch (error) {
        console.error('预览错误:', error);
        showError(error.message || '预览失败，请检查文件格式和选项设置');
    } finally {
        hideLoading();
    }
};

// 修改导入数据处理函数
window.handleImportData = async function() {
    if (!window.importModule.selectedFiles?.length) {
        showError('请先选择文件');
        return;
    }

    try {
        showLoading('正在导入数据...');
        
        const formData = new FormData();
        window.importModule.selectedFiles.forEach(file => {
            formData.append('files[]', file);
        });

        formData.append('sourceType', 'text');
        formData.append('startRow', document.getElementById('start-row').value);
        formData.append('headerRow', document.getElementById('header-row').value);

        // 添加分隔符设置
        const autoSplit = document.getElementById('auto-split').checked;
        formData.append('autoSplit', autoSplit.toString());
        
        if (!autoSplit) {
            // 直接获取自定义分隔符的值
            const customDelimiter = document.getElementById('custom-delimiter');
            if (customDelimiter && customDelimiter.value) {
                formData.append('delimiter', customDelimiter.value);
            }
        }

        console.log('发送导入请求...');
        const response = await fetch('/api/import-data', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();
        console.log('导入响应:', result);

        if (!response.ok) {
            throw new Error(result.error || '导入失败');
        }

        showSuccess('导入成功');
        
        // 显示下载按钮，并存储下载链接
        if (result.downloadUrl) {
            window.importModule.downloadUrl = result.downloadUrl;
            const downloadBtn = document.getElementById('download-btn');
            if (downloadBtn) {
                downloadBtn.style.display = 'inline-block';
            }
        }

    } catch (error) {
        console.error('导入错误:', error);
        showError(error.message || '导入失败，请检查文件格式和选项设置');
    } finally {
        hideLoading();
    }
};

// 添加下载结果处理函数
window.handleDownloadResult = function() {
    if (window.importModule.downloadUrl) {
        window.location.href = window.importModule.downloadUrl;
    } else {
        showError('下载链接不可用');
    }
};

// 添加辅助函数
function showLoading(message) {
    const loadingDiv = document.createElement('div');
    loadingDiv.id = 'loading-overlay';
    loadingDiv.innerHTML = `
        <div class="loading-content">
            <div class="loading-spinner"></div>
            <div class="loading-message">${message}</div>
        </div>
    `;
    document.body.appendChild(loadingDiv);
}

function hideLoading() {
    const loadingDiv = document.getElementById('loading-overlay');
    if (loadingDiv) {
        document.body.removeChild(loadingDiv);
    }
}

function showError(message) {
    const errorDiv = document.createElement('div');
    errorDiv.className = 'alert error';
    errorDiv.innerHTML = message;
    document.body.appendChild(errorDiv);
    
    setTimeout(() => {
        errorDiv.classList.add('fade-out');
        setTimeout(() => {
            document.body.removeChild(errorDiv);
        }, 300);
    }, 3000);
}

function showSuccess(message) {
    const successDiv = document.createElement('div');
    successDiv.className = 'alert success';
    successDiv.innerHTML = message;
    document.body.appendChild(successDiv);
    
    setTimeout(() => {
        successDiv.classList.add('fade-out');
        setTimeout(() => {
            document.body.removeChild(successDiv);
        }, 300);
    }, 3000);
}

// 添加预览表格更新函数
function updatePreviewTable(data) {
    if (!data || !data.columns || !data.rows) {
        console.error('预览数据格式不正确:', data);
        return;
    }

    const table = document.querySelector('.preview-table');
    if (!table) {
        console.error('未找到预览表格元素');
        return;
    }
    
    // 更新表头
    const thead = table.querySelector('thead');
    thead.innerHTML = `
        <tr>
            ${data.columns.map(col => `<th>${col}</th>`).join('')}
        </tr>
    `;

    // 更新数据行
    const tbody = table.querySelector('tbody');
    tbody.innerHTML = data.rows.map(row => `
        <tr>
            ${row.map(cell => `<td>${cell || ''}</td>`).join('')}
        </tr>
    `).join('');

    // 更新列设置
    const columnSettings = document.querySelector('.columns-container');
    if (columnSettings) {
        columnSettings.innerHTML = data.columns.map((col, index) => `
            <div class="column-setting">
                <label>
                    <input type="checkbox" checked data-column="${index}">
                    ${col}
                </label>
                <select data-column="${index}">
                    <option value="auto">自动</option>
                    <option value="text">文本</option>
                    <option value="number">数字</option>
                    <option value="date">日期</option>
                </select>
                <input type="text" class="column-name" value="${col}" placeholder="自定义列名">
            </div>
        `).join('');
    }
}

// 添加隐藏预览函数
function hidePreview() {
    const importOptions = document.getElementById('import-options');
    const dataPreview = document.getElementById('data-preview');
    const previewBtn = document.getElementById('preview-btn');
    const importBtn = document.getElementById('import-btn');
    const downloadBtn = document.getElementById('download-btn');

    if (importOptions) importOptions.style.display = 'none';
    if (dataPreview) dataPreview.style.display = 'none';
    if (previewBtn) previewBtn.disabled = true;
    if (importBtn) importBtn.disabled = true;
    if (downloadBtn) downloadBtn.style.display = 'none';
    
    window.importModule.previewData = null;
    window.importModule.downloadUrl = null;
}

// 添加删除文件函数
window.removeFile = function(index) {
    console.log('删除文件:', index);
    
    // 从选中文件列表中移除
    window.importModule.selectedFiles = Array.from(window.importModule.selectedFiles)
        .filter((_, i) => i !== index);
    
    // 更新文件列表显示
    const fileInfo = document.getElementById('file-info');
    const fileList = document.getElementById('file-list');
    
    if (fileInfo && fileList) {
        if (window.importModule.selectedFiles.length > 0) {
            // 更新文件列表
            fileList.innerHTML = window.importModule.selectedFiles.map((file, idx) => `
                <div class="file-item">
                    <div class="file-info-content">
                        <span class="file-name">${file.webkitRelativePath || file.name}</span>
                        <span class="file-size">${formatFileSize(file.size)}</span>
                    </div>
                    <button class="delete-btn" onclick="removeFile(${idx})" title="删除此文件">
                        <span class="delete-icon">×</span>
                    </button>
                </div>
            `).join('');
        } else {
            // 如果没有文件了，隐藏文件信息区域
            fileInfo.style.display = 'none';
        }
    }

    // 更新上传提示
    const uploadTip = document.querySelector('.upload-tip');
    if (uploadTip) {
        uploadTip.textContent = window.importModule.selectedFiles.length > 0 ? 
            `已选择 ${window.importModule.selectedFiles.length} 个文件` : 
            '点击或拖放 TXT/CSV 文件到此处';
    }

    // 如果没有文件了，重置界面
    if (window.importModule.selectedFiles.length === 0) {
        // 重置文件输入框
        const fileInput = document.getElementById('import-file');
        const folderInput = document.getElementById('import-folder');
        if (fileInput) fileInput.value = '';
        if (folderInput) folderInput.value = '';
        
        // 隐藏预览和导入选项
        hidePreview();
        
        const importOptions = document.getElementById('import-options');
        if (importOptions) {
            importOptions.style.display = 'none';
        }
    }
};