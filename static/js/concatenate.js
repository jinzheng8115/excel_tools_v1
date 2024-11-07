async function handleConcatenateFileUpload() {
    resetError('concatenateStatus');
    const file = document.getElementById('concatenateFile').files[0];
    if (file) {
        const formData = new FormData();
        formData.append('file', file);
        
        try {
            const response = await fetch('/api/get-sheets', {
                method: 'POST',
                body: formData
            });
            const data = await response.json();
            
            if (data.success) {
                const select = document.getElementById('concatenateSheet');
                select.innerHTML = '<option value="">请选择工作表</option>' + 
                    data.sheets.map(sheet => `<option value="${sheet}">${sheet}</option>`).join('');
                document.getElementById('concatenateFileConfig').style.display = 'block';
            } else {
                showError(data.error || '获取工作表失败', 'concatenateStatus');
            }
        } catch (error) {
            showError('处理文件时发生错误: ' + error.message, 'concatenateStatus');
        }
    }
}

// 修改变量名避免冲突
let concatenateColumns = null;

async function handleConcatenateSheetChange() {
    resetError('concatenateStatus');
    const file = document.getElementById('concatenateFile').files[0];
    const sheet = document.getElementById('concatenateSheet').value;
    
    if (file && sheet) {
        const formData = new FormData();
        formData.append('file', file);
        formData.append('sheet', sheet);
        formData.append('concatenateFile', 'true');
        
        try {
            const response = await fetch('/api/get-columns', {
                method: 'POST',
                body: formData
            });
            const data = await response.json();
            
            if (data.success) {
                // 使用新的变量名
                concatenateColumns = data.columns;
                document.querySelector('#mergeColumns .column-list').innerHTML = '';
                addMergeColumn();
            } else {
                showError(data.error || '获取列信息失败', 'concatenateStatus');
            }
        } catch (error) {
            showError('处理文件时发生错误: ' + error.message, 'concatenateStatus');
        }
    }
}

function addMergeColumn() {
    const columnList = document.querySelector('#mergeColumns .column-list');
    const file = document.getElementById('concatenateFile').files[0];
    const sheet = document.getElementById('concatenateSheet').value;
    
    if (!file || !sheet) {
        showError('请先选择文件和工作表', 'concatenateStatus');
        return;
    }
    
    if (!concatenateColumns) {
        showError('请先选择工作表', 'concatenateStatus');
        return;
    }
    
    const columnItem = document.createElement('div');
    columnItem.className = 'column-item';
    
    // 列选择
    const select = document.createElement('select');
    select.className = 'merge-column-select';
    select.innerHTML = '<option value="">请选择列</option>';
    concatenateColumns.columns.forEach((col, index) => {
        const header = concatenateColumns.headers[index] || '';
        const displayText = header ? `${col} (${header})` : col;
        select.innerHTML += `<option value="${col}">${displayText}</option>`;
    });
    
    // 分隔符配置
    const separatorConfig = document.createElement('div');
    separatorConfig.className = 'separator-config';
    
    const separatorType = document.createElement('div');
    separatorType.className = 'separator-type';
    separatorType.innerHTML = `
        <label>
            <input type="checkbox" class="use-space" onchange="handleSeparatorTypeChange(this)"> 使用空格
        </label>
        <input type="text" class="separator-input" placeholder="自定义分隔符">
    `;
    
    const separatorPosition = document.createElement('label');
    separatorPosition.className = 'separator-position';
    separatorPosition.innerHTML = `
        <input type="checkbox" class="separator-before"> 分隔符在前
    `;
    
    separatorConfig.appendChild(separatorType);
    separatorConfig.appendChild(separatorPosition);
    
    // 删除按钮
    const removeBtn = document.createElement('button');
    removeBtn.className = 'remove-column-btn';
    removeBtn.textContent = '删除';
    removeBtn.onclick = function() {
        columnList.removeChild(columnItem);
    };
    
    columnItem.appendChild(select);
    columnItem.appendChild(separatorConfig);
    columnItem.appendChild(removeBtn);
    columnList.appendChild(columnItem);
}

function handleSeparatorChange(type) {
    const customInput = document.getElementById('separator');
    switch(type) {
        case 'none':
            customInput.style.display = 'none';
            customInput.value = '';
            break;
        case 'space':
            customInput.style.display = 'none';
            customInput.value = ' ';
            break;
        case 'custom':
            customInput.style.display = 'block';
            customInput.value = '';
            customInput.focus();
            break;
    }
}

async function processConcatenate() {
    resetError('concatenateStatus');
    const file = document.getElementById('concatenateFile').files[0];
    const sheet = document.getElementById('concatenateSheet').value;
    
    // 获取所有列和分隔符配置
    const columnItems = document.querySelectorAll('.column-item');
    const columnsData = Array.from(columnItems).map(item => {
        const useSpace = item.querySelector('.use-space').checked;
        const customSeparator = item.querySelector('.separator-input').value;
        
        return {
            column: item.querySelector('.merge-column-select').value,
            separator: useSpace ? ' ' : customSeparator,
            separator_before: item.querySelector('.separator-before').checked
        };
    }).filter(data => data.column);
    
    if (!file || !sheet) {
        showError('请选择文件和工作表', 'concatenateStatus');
        return;
    }

    if (columnsData.length === 0) {
        showError('请选择要合并的列', 'concatenateStatus');
        return;
    }

    const formData = new FormData();
    formData.append('file', file);
    formData.append('sheet', sheet);
    formData.append('columnsData', JSON.stringify(columnsData));

    const status = document.getElementById('concatenateStatus');
    status.textContent = '正在处理...';

    try {
        const response = await fetch('/api/concatenate', {
            method: 'POST',
            body: formData
        });
        const data = await response.json();
        
        if (data.success) {
            status.textContent = '处理完成！';
            document.getElementById('concatenateDownloadBtn').disabled = false;
        } else {
            showError(data.error || '处理失败', 'concatenateStatus');
        }
    } catch (error) {
        showError('处理文件时发生错误: ' + error.message, 'concatenateStatus');
    }
}

function downloadConcatenateResult() {
    window.location.href = '/api/download-result';
}

function showError(message, statusId) {
    const status = document.getElementById(statusId);
    status.textContent = message;
    status.style.color = '#ff3b30';
}

function resetError(statusId) {
    const status = document.getElementById(statusId);
    status.textContent = '等待操作...';
    status.style.color = '#1d1d1f';
} 