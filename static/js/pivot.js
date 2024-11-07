function validateColumnSelection() {
    // 获取所有选择的列
    const rowLabels = Array.from(document.querySelectorAll('.row-label-select'))
        .map(select => select.value)
        .filter(value => value);
    
    const valueFields = Array.from(document.querySelectorAll('.value-field-select'))
        .map(select => select.value)
        .filter(value => value);

    // 只检查必选项
    if (rowLabels.length === 0) {
        showError('请至少选择一个行标签', 'pivotStatus');
        return false;
    }

    if (valueFields.length === 0) {
        showError('请至少选择一个值字段', 'pivotStatus');
        return false;
    }

    return true;
} 