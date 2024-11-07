async function handleMainFileUpload() {
    resetError('status');
    const file = document.getElementById('mainFile').files[0];
    
    // 无论是否选择了文件，都重置主表相关的选择
    resetMainTableSelections();
    
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
                const select = document.getElementById('mainSheet');
                select.innerHTML = '<option value="">请选择工作表</option>' + 
                    data.sheets.map(sheet => `<option value="${sheet}">${sheet}</option>`).join('');
                document.getElementById('mainFileConfig').style.display = 'block';
            } else {
                showError(data.error || '获取工作表失败', 'status');
            }
        } catch (error) {
            showError('处理文件时发生错误: ' + error.message, 'status');
        }
    } else {
        // 如果没有选择文件，隐藏配置区域
        document.getElementById('mainFileConfig').style.display = 'none';
    }
}

async function handleMainSheetChange() {
    resetError('status');
    const file = document.getElementById('mainFile').files[0];
    const sheet = document.getElementById('mainSheet').value;
    
    if (file && sheet) {
        const formData = new FormData();
        formData.append('file', file);
        formData.append('sheet', sheet);
        
        try {
            const response = await fetch('/api/get-columns', {
                method: 'POST',
                body: formData
            });
            const data = await response.json();
            
            if (data.success) {
                updateSelectOptions('mainLookupValue', data);
                document.querySelector('#mainLookupColumns .column-list').innerHTML = '';
            } else {
                showError(data.error || '获取列信息失败', 'status');
            }
        } catch (error) {
            showError('处理文件时发生错误: ' + error.message, 'status');
        }
    }
}

async function handleLookupFileUpload() {
    resetError('status');
    const file = document.getElementById('lookupFile').files[0];
    
    // 无论是否选择了文件，都重置查找表相关的选择
    resetLookupTableSelections();
    
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
                const select = document.getElementById('lookupSheet');
                select.innerHTML = '<option value="">请选择工作表</option>' + 
                    data.sheets.map(sheet => `<option value="${sheet}">${sheet}</option>`).join('');
                document.getElementById('lookupFileConfig').style.display = 'block';
            } else {
                showError(data.error || '获取工作表失败', 'status');
            }
        } catch (error) {
            showError('处理文件时发生错误: ' + error.message, 'status');
        }
    } else {
        // 如果没有选择文件，隐藏配置区域
        document.getElementById('lookupFileConfig').style.display = 'none';
    }
}

async function handleLookupSheetChange() {
    resetError('status');
    const file = document.getElementById('lookupFile').files[0];
    const sheet = document.getElementById('lookupSheet').value;
    
    if (file && sheet) {
        const formData = new FormData();
        formData.append('file', file);
        formData.append('sheet', sheet);
        
        try {
            const response = await fetch('/api/get-columns', {
                method: 'POST',
                body: formData
            });
            const data = await response.json();
            
            if (data.success) {
                // 更新匹配列选择器
                updateSelectOptions('lookupMatchValue', data);
                document.querySelector('#lookupMatchColumns .column-list').innerHTML = '';
                
                // 更新返回列选择器
                updateSelectOptions('lookupReturnValue', data);
                document.querySelector('#returnColumns .column-list').innerHTML = '';
                
                // 重置返回列选择为单列模式
                document.querySelector('input[name="returnType"][value="single"]').checked = true;
                switchReturnType('single');
                
                // 如果是多列匹配模式，同步匹配列数
                const mainMatchType = document.querySelector('input[name="mainMatchType"]:checked').value;
                if (mainMatchType === 'multiple') {
                    syncMatchColumns();
                }
            } else {
                showError(data.error || '获取列信息失败', 'status');
            }
        } catch (error) {
            showError('处理文件时发生错误: ' + error.message, 'status');
        }
    }
}

function switchMainMatchType(type) {
    const singleMatch = document.getElementById('mainSingleMatch');
    const multiMatch = document.getElementById('mainMultiMatch');
    const lookupMatchTypeContainer = document.getElementById('lookupMatchTypeContainer');
    
    if (type === 'single') {
        singleMatch.style.display = 'block';
        multiMatch.style.display = 'none';
        document.querySelector('#mainLookupColumns .column-list').innerHTML = '';
        
        document.querySelector('input[name="lookupMatchType"][value="single"]').checked = true;
        document.querySelector('input[name="lookupMatchType"][value="single"]').disabled = false;
        document.querySelector('input[name="lookupMatchType"][value="multiple"]').disabled = true;
        switchLookupMatchType('single');
    } else {
        singleMatch.style.display = 'none';
        multiMatch.style.display = 'block';
        
        document.querySelector('input[name="lookupMatchType"][value="multiple"]').checked = true;
        document.querySelector('input[name="lookupMatchType"][value="multiple"]').disabled = false;
        document.querySelector('input[name="lookupMatchType"][value="single"]').disabled = true;
        switchLookupMatchType('multiple');
        
        if (document.querySelector('#mainLookupColumns .column-list').children.length === 0) {
            addMainColumn();
        }
    }
    
    updateMatchTypeHint(type);
}

function updateMatchTypeHint(type) {
    const mainTypeHint = document.createElement('div');
    mainTypeHint.className = 'match-type-hint';
    mainTypeHint.style.color = '#666';
    mainTypeHint.style.fontSize = '12px';
    mainTypeHint.style.marginTop = '5px';
    
    if (type === 'single') {
        mainTypeHint.textContent = '已选择单列匹配，查找表也将使用单列匹配';
    } else {
        mainTypeHint.textContent = '已选择多列匹配，查找表也将使用多列匹配';
    }
    
    const oldHint = document.querySelector('.match-type-hint');
    if (oldHint) {
        oldHint.remove();
    }
    
    const matchTypeSelector = document.querySelector('.match-type-selector');
    matchTypeSelector.appendChild(mainTypeHint);
}

function switchLookupMatchType(type) {
    const singleMatch = document.getElementById('lookupSingleMatch');
    const multiMatch = document.getElementById('lookupMultiMatch');
    
    if (type === 'single') {
        singleMatch.style.display = 'block';
        multiMatch.style.display = 'none';
        document.querySelector('#lookupMatchColumns .column-list').innerHTML = '';
    } else {
        singleMatch.style.display = 'none';
        multiMatch.style.display = 'block';
        syncMatchColumns();
    }
}

function syncMatchColumns() {
    const mainColumns = document.querySelectorAll('#mainLookupColumns .column-list .column-item');
    const lookupColumns = document.querySelector('#lookupMatchColumns .column-list');
    
    lookupColumns.innerHTML = '';
    
    mainColumns.forEach(() => {
        addLookupMatchColumn();
    });
}

function addMainColumn() {
    const columnList = document.querySelector('#mainLookupColumns .column-list');
    const mainSelect = document.getElementById('mainLookupValue');
    
    if (!mainSelect || !mainSelect.options.length) {
        showError('请先选择工作表', 'status');
        return;
    }
    
    const columnItem = document.createElement('div');
    columnItem.className = 'column-item';
    
    const select = document.createElement('select');
    select.className = 'mainLookupValue-select';
    select.innerHTML = mainSelect.innerHTML;
    
    const removeBtn = document.createElement('button');
    removeBtn.className = 'remove-column-btn';
    removeBtn.textContent = '删除';
    removeBtn.onclick = function() {
        columnList.removeChild(columnItem);
        syncMatchColumns();
    };
    
    columnItem.appendChild(select);
    columnItem.appendChild(removeBtn);
    columnList.appendChild(columnItem);
    
    if (document.querySelector('input[name="lookupMatchType"]:checked').value === 'multiple') {
        addLookupMatchColumn();
    }
}

function addLookupMatchColumn() {
    const columnList = document.querySelector('#lookupMatchColumns .column-list');
    const lookupSelect = document.getElementById('lookupMatchValue');
    
    if (!lookupSelect || !lookupSelect.options.length) {
        showError('请先选择工作表', 'status');
        return;
    }
    
    const columnItem = document.createElement('div');
    columnItem.className = 'column-item';
    
    const select = document.createElement('select');
    select.className = 'lookupMatchValue-select';
    select.innerHTML = lookupSelect.innerHTML;
    
    const removeBtn = document.createElement('button');
    removeBtn.className = 'remove-column-btn';
    removeBtn.textContent = '删除';
    removeBtn.onclick = function() {
        columnList.removeChild(columnItem);
    };
    
    columnItem.appendChild(select);
    columnItem.appendChild(removeBtn);
    columnList.appendChild(columnItem);
}

function addReturnColumn() {
    const columnList = document.querySelector('#returnColumns .column-list');
    const lookupSelect = document.getElementById('lookupReturnValue');
    
    if (!lookupSelect || !lookupSelect.options.length) {
        showError('请先选择工作表', 'status');
        return;
    }
    
    const columnItem = document.createElement('div');
    columnItem.className = 'column-item';
    
    const select = document.createElement('select');
    select.className = 'lookupReturnValue-select';
    select.innerHTML = lookupSelect.innerHTML;
    
    const removeBtn = document.createElement('button');
    removeBtn.className = 'remove-column-btn';
    removeBtn.textContent = '删除';
    removeBtn.onclick = function() {
        if (document.querySelectorAll('.lookupReturnValue-select').length > 1) {
            columnList.removeChild(columnItem);
        } else {
            showError('至少需要一个返回列', 'status');
        }
    };
    
    columnItem.appendChild(select);
    columnItem.appendChild(removeBtn);
    columnList.appendChild(columnItem);
}

function processFiles() {
    // 验证文件和工作表选择
    const mainFile = document.getElementById('mainFile').files[0];
    const mainSheet = document.getElementById('mainSheet').value;
    const lookupFile = document.getElementById('lookupFile').files[0];
    const lookupSheet = document.getElementById('lookupSheet').value;
    
    if (!mainFile || !mainSheet) {
        showError('请选择主数据表文件并选择工作表');
        return;
    }
    
    if (!lookupFile || !lookupSheet) {
        showError('请选择查找数据表文件并选择工作表');
        return;
    }
    
    const mainMatchType = document.querySelector('input[name="mainMatchType"]:checked').value;
    const lookupMatchType = document.querySelector('input[name="lookupMatchType"]:checked').value;
    const returnType = document.querySelector('input[name="returnType"]:checked').value;
    
    // 获取主表匹配列
    let mainColumns = [];
    if (mainMatchType === 'single') {
        const mainColumn = document.getElementById('mainLookupValue').value;
        if (!mainColumn) {
            showError('请选择主表的匹配列');
            return;
        }
        mainColumns.push(mainColumn);
    } else {
        mainColumns = Array.from(document.querySelectorAll('#mainLookupColumns select'))
            .map(select => select.value)
            .filter(value => value);
        if (mainColumns.length === 0) {
            showError('请至少选择一个主表匹配列');
            return;
        }
    }
    
    // 获取查找表匹配列
    let lookupMatchColumns = [];
    if (lookupMatchType === 'single') {
        const lookupColumn = document.getElementById('lookupMatchValue').value;
        if (!lookupColumn) {
            showError('请选择查找表的匹配列');
            return;
        }
        lookupMatchColumns.push(lookupColumn);
    } else {
        lookupMatchColumns = Array.from(document.querySelectorAll('#lookupMatchColumns select'))
            .map(select => select.value)
            .filter(value => value);
        if (lookupMatchColumns.length === 0) {
            showError('请至少选择一个查找表匹配列');
            return;
        }
    }
    
    // 获取返回列
    let returnColumns = [];
    if (returnType === 'single') {
        const returnColumn = document.getElementById('lookupReturnValue').value;
        if (!returnColumn) {
            showError('请选择返回列');
            return;
        }
        returnColumns.push(returnColumn);
    } else {
        returnColumns = Array.from(document.querySelectorAll('#returnColumns select'))
            .map(select => select.value)
            .filter(value => value);
        if (returnColumns.length === 0) {
            showError('请至少选择一个返回列');
            return;
        }
    }
    
    // 验证匹配列数
    if (mainMatchType === 'multiple' && mainColumns.length !== lookupMatchColumns.length) {
        showError('主表和查找表的匹配列数必须相同');
        return;
    }
    
    // 验证重复列
    if (new Set(mainColumns).size !== mainColumns.length) {
        showError('主表存在重复的匹配列，请检查');
        return;
    }
    
    if (new Set(lookupMatchColumns).size !== lookupMatchColumns.length) {
        showError('查找表存在重复的匹配列，请检查');
        return;
    }
    
    if (new Set(returnColumns).size !== returnColumns.length) {
        showError('返回列存在重复选择，请检查');
        return;
    }
    
    // 所有验证通过，开始处理
    const formData = new FormData();
    formData.append('mainFile', mainFile);
    formData.append('mainSheet', mainSheet);
    formData.append('mainMatchType', mainMatchType);
    formData.append('mainColumns', JSON.stringify(mainColumns));
    
    formData.append('lookupFile', lookupFile);
    formData.append('lookupSheet', lookupSheet);
    formData.append('lookupMatchType', lookupMatchType);
    formData.append('lookupMatchColumns', JSON.stringify(lookupMatchColumns));
    formData.append('returnType', returnType);
    formData.append('returnColumns', JSON.stringify(returnColumns));
    
    const status = document.getElementById('status');
    const matchStats = document.getElementById('match-stats');
    
    status.textContent = '正在处理...';
    status.className = 'status-message status-processing';
    matchStats.textContent = '';
    
    fetch('/api/vlookup', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            status.textContent = '处理完成！';
            status.className = 'status-message status-success';
            matchStats.textContent = data.error || '';  // 显示匹配统计信息
            document.getElementById('downloadBtn').disabled = false;
        } else {
            showError(data.error || '处理失败，请检查数据是否正确');
            document.getElementById('downloadBtn').disabled = true;
        }
    })
    .catch(error => {
        showError('处理文件时发生错误，请重试');
        document.getElementById('downloadBtn').disabled = true;
    });
}

function downloadResult() {
    window.location.href = '/api/download-result';
}

function updateSelectOptions(selectId, data) {
    const select = document.getElementById(selectId);
    
    select.innerHTML = '<option value="">请选择</option>';
    data.columns.columns.forEach((col, index) => {
        const header = data.columns.headers[index] || '';
        const displayText = header ? `${col} (${header})` : col;
        select.innerHTML += `<option value="${col}">${displayText}</option>`;
    });
}

// 添加返回类型切换函数
function switchReturnType(type) {
    const singleReturn = document.getElementById('returnSingleColumn');
    const multiReturn = document.getElementById('returnMultiColumns');
    
    if (type === 'single') {
        singleReturn.style.display = 'block';
        multiReturn.style.display = 'none';
        // 清空多列返回的列表
        document.querySelector('#returnColumns .column-list').innerHTML = '';
    } else {
        singleReturn.style.display = 'none';
        multiReturn.style.display = 'block';
        // 如果没有返回列，添加一个
        if (document.querySelector('#returnColumns .column-list').children.length === 0) {
            addReturnColumn();
        }
    }
}

// 新增：重置主表相关的选择
function resetMainTableSelections() {
    // 重置工作表选择
    document.getElementById('mainSheet').value = '';
    
    // 重置单列匹配选择
    document.getElementById('mainLookupValue').innerHTML = '<option value="">请先选择工作表</option>';
    
    // 重置多列匹配选择
    document.querySelector('#mainLookupColumns .column-list').innerHTML = '';
    
    // 重置匹配方式为单列
    document.querySelector('input[name="mainMatchType"][value="single"]').checked = true;
    switchMainMatchType('single');
    
    // 重置状态
    document.getElementById('status').textContent = '等待上传文件...';
    document.getElementById('status').style.color = '#1d1d1f';
    document.getElementById('downloadBtn').disabled = true;
}

// 新增：重置查找表相关的选择
function resetLookupTableSelections() {
    // 重置工作表选择
    document.getElementById('lookupSheet').value = '';
    
    // 重置单列匹配选择
    document.getElementById('lookupMatchValue').innerHTML = '<option value="">请选择列</option>';
    
    // 重置多列匹配选择
    document.querySelector('#lookupMatchColumns .column-list').innerHTML = '';
    
    // 重置返回列选择
    document.getElementById('lookupReturnValue').innerHTML = '<option value="">请选择返回列</option>';
    document.querySelector('#returnColumns .column-list').innerHTML = '';
    
    // 重置返回类型为单列
    document.querySelector('input[name="returnType"][value="single"]').checked = true;
    switchReturnType('single');
    
    // 重置状态
    document.getElementById('status').textContent = '等待上传文件...';
    document.getElementById('status').style.color = '#1d1d1f';
    document.getElementById('downloadBtn').disabled = true;
}

// 修改错误提示函数
function showError(message) {
    const status = document.getElementById('status');
    const matchStats = document.getElementById('match-stats');
    
    status.textContent = message;
    status.className = 'status-message status-error';
    matchStats.textContent = '';  // 清空匹配统计
}