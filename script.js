/**
 * Excel模拟器主类
 * 实现Excel的基本功能：单元格编辑、行列操作、右键菜单等
 */
class ExcelSimulator {
    constructor(container, options = {}) {
        this.container = container;
        this.options = {
            rows: options.rows || 100,
            columns: options.columns || 100,
            defaultRowHeight: 30,
            defaultColumnWidth: 100,
            ...options
        };
        
        // 数据存储
        this.data = new Map(); // 存储单元格数据
        this.selectedCells = new Set(); // 当前选中的单元格
        this.clipboard = null; // 剪贴板数据
        this.isSelecting = false; // 是否正在选择
        this.selectionStart = null; // 选择起始位置
        
        // DOM元素引用
        this.elements = {};
    
        // 编辑状态
        this.editingCell = null;
        this.cellEditor = null;
        
        // 调整大小相关
        this.isResizing = false;
        this.resizeType = null; // 'row' or 'column'
        this.resizeIndex = null;
        this.resizeAnimationFrame = null; // 用于节流的动画帧ID
        this.cachedResizeElements = null; // 缓存resize时需要操作的DOM元素
        
        // 行列尺寸
        this.rowHeights = new Array(this.options.rows).fill(this.options.defaultRowHeight);
        this.columnWidths = new Array(this.options.columns).fill(this.options.defaultColumnWidth);
        
        this.init();
    }
    
    /**
     * 初始化Excel模拟器
     */
    init() {
        this.createStructure();
        this.bindEvents();
        this.generateTable();
    }
    
    /**
     * 创建基本DOM结构
     */
    createStructure() {
        this.elements.columnHeaders = document.getElementById('columnHeaders');
        this.elements.rowHeaders = document.getElementById('rowHeaders');
        this.elements.tableContainer = document.getElementById('tableContainer');
        this.elements.tableBody = document.getElementById('tableBody');
        this.elements.contextMenu = document.getElementById('contextMenu');
        this.cellEditor = document.getElementById('cellEditor');
    }
    
    /**
     * 绑定事件监听器
     */
    bindEvents() {
        // 工具栏事件
        document.getElementById('addRow').addEventListener('click', () => this.addRow());
        document.getElementById('addColumn').addEventListener('click', () => this.addColumn());
        document.getElementById('regenerate').addEventListener('click', () => this.regenerateTable());
        
        // 表格容器事件 - 改为监听excel-wrapper的滚动
        const excelWrapper = document.querySelector('.excel-wrapper');
        excelWrapper.addEventListener('scroll', (e) => this.handleScroll(e));
        
        // 全局事件
        document.addEventListener('mousedown', (e) => this.handleMouseDown(e));
        document.addEventListener('mousemove', (e) => this.handleMouseMove(e));
        document.addEventListener('mouseup', (e) => this.handleMouseUp(e));
        document.addEventListener('keydown', (e) => this.handleKeyDown(e));
        document.addEventListener('contextmenu', (e) => this.handleContextMenu(e));
        document.addEventListener('click', (e) => this.handleDocumentClick(e));
        
        // 单元格编辑器事件
        this.cellEditor.addEventListener('blur', () => this.finishEditing());
        this.cellEditor.addEventListener('keydown', (e) => this.handleEditorKeyDown(e));
        
        // 右键菜单事件
        this.elements.contextMenu.addEventListener('click', (e) => this.handleMenuClick(e));
    }
    
    /**
     * 生成表格
     */
    generateTable() {
        this.generateColumnHeaders();
        this.generateRowHeaders();
        this.generateTableBody();
        // 初始化后确保所有尺寸正确
        this.updateTableWidth();
    }
    
    /**
     * 生成列头
     */
    generateColumnHeaders() {
        this.elements.columnHeaders.innerHTML = '';
        
        for (let i = 0; i < this.options.columns; i++) {
            const header = document.createElement('div');
            header.className = 'column-header';
            header.textContent = this.getColumnName(i);
            header.style.width = this.columnWidths[i] + 'px';
            header.dataset.column = i;
            
            // 添加调整手柄
            const resizeHandle = document.createElement('div');
            resizeHandle.className = 'column-resize-handle';
            resizeHandle.dataset.column = i;
            header.appendChild(resizeHandle);
            
            this.elements.columnHeaders.appendChild(header);
        }
    }
    
    /**
     * 生成行头
     */
    generateRowHeaders() {
        this.elements.rowHeaders.innerHTML = '';
        
        for (let i = 0; i < this.options.rows; i++) {
            const header = document.createElement('div');
            header.className = 'row-header';
            header.textContent = i + 1;
            header.style.height = this.rowHeights[i] + 'px';
            header.dataset.row = i;
            
            // 添加调整手柄
            const resizeHandle = document.createElement('div');
            resizeHandle.className = 'row-resize-handle';
            resizeHandle.dataset.row = i;
            header.appendChild(resizeHandle);
            
            this.elements.rowHeaders.appendChild(header);
        }
    }
    
    /**
     * 生成表格主体
     */
    generateTableBody() {
        this.elements.tableBody.innerHTML = '';
        
        for (let row = 0; row < this.options.rows; row++) {
            const tr = document.createElement('tr');
            tr.dataset.row = row;
            
            for (let col = 0; col < this.options.columns; col++) {
                const td = document.createElement('td');
                td.dataset.row = row;
                td.dataset.column = col;
                td.style.width = this.columnWidths[col] + 'px';
                td.style.height = this.rowHeights[row] + 'px';
                
                const cellKey = `${row}-${col}`;
                const cellValue = this.data.get(cellKey) || '';
                td.textContent = cellValue;
                
                tr.appendChild(td);
            }
            
            this.elements.tableBody.appendChild(tr);
        }
        
        // 设置表格总宽度以确保列对齐
        this.updateTableWidth();
    }
    
    /**
     * 更新表格宽度
     */
    updateTableWidth() {
        const totalWidth = this.columnWidths.reduce((sum, width) => sum + width, 0);
        const totalHeight = this.rowHeights.reduce((sum, height) => sum + height, 0);
        
        // 更新列头容器宽度
        this.elements.columnHeaders.style.width = totalWidth + 'px';
        
        // 更新行头容器高度
        this.elements.rowHeaders.style.height = totalHeight + 'px';
        
        // 更新表格宽度和高度
        const table = this.elements.tableBody.parentElement;
        if (table) {
            table.style.width = totalWidth + 'px';
            table.style.height = totalHeight + 'px';
        }
        
        // 确保所有行宽度一致
        const rows = this.elements.tableBody.querySelectorAll('tr');
        rows.forEach(row => {
            row.style.width = totalWidth + 'px';
        });
    }

    /**
     * 获取列名（A, B, C, ..., Z, AA, AB, ...）
     */
    getColumnName(index) {
        let result = '';
        while (index >= 0) {
            result = String.fromCharCode(65 + (index % 26)) + result;
            index = Math.floor(index / 26) - 1;
        }
        return result;
    }
    
    /**
     * 处理鼠标按下事件
     */
    handleMouseDown(e) {
        const target = e.target;
        
        // 如果点击的不是编辑器，且正在编辑，先完成编辑
        if (this.editingCell && target !== this.cellEditor) {
            this.finishEditing();
        }
        
        // 处理列调整
        if (target.classList.contains('column-resize-handle')) {
            this.startResize('column', parseInt(target.dataset.column));
            e.preventDefault();
            return;
        }
        
        // 处理行调整
        if (target.classList.contains('row-resize-handle')) {
            this.startResize('row', parseInt(target.dataset.row));
            e.preventDefault();
            return;
        }
        
        // 处理列头选择
        if (target.classList.contains('column-header')) {
            this.selectColumn(parseInt(target.dataset.column), e.ctrlKey);
            e.preventDefault();
            return;
        }
        
        // 处理行头选择
        if (target.classList.contains('row-header')) {
            this.selectRow(parseInt(target.dataset.row), e.ctrlKey);
            e.preventDefault();
            return;
        }
        
        // 处理单元格选择
        if (target.tagName === 'TD') {
            const row = parseInt(target.dataset.row);
            const col = parseInt(target.dataset.column);
            
            if (e.detail === 2) { // 双击
                this.startEditing(row, col);
            } else {
                this.selectCell(row, col, e.ctrlKey, e.shiftKey);
                this.isSelecting = true;
                this.selectionStart = { row, col };
            }
            e.preventDefault();
        }
    }
    
    /**
     * 处理鼠标移动事件
     */
    handleMouseMove(e) {
        if (this.isResizing) {
            // 使用requestAnimationFrame节流，避免过度频繁的DOM操作
            if (this.resizeAnimationFrame) {
                cancelAnimationFrame(this.resizeAnimationFrame);
            }
            this.resizeAnimationFrame = requestAnimationFrame(() => {
                this.performResize(e);
            });
            return;
        }
        
        if (this.isSelecting && this.selectionStart) {
            const target = e.target;
            if (target.tagName === 'TD') {
                const row = parseInt(target.dataset.row);
                const col = parseInt(target.dataset.column);
                this.extendSelection(this.selectionStart.row, this.selectionStart.col, row, col);
            }
        }
    }
    
    /**
     * 处理鼠标释放事件
     */
    handleMouseUp(e) {
        if (this.isResizing) {
            this.stopResize();
        }
        
        this.isSelecting = false;
        this.selectionStart = null;
    }
    
    /**
     * 开始调整大小
     */
    startResize(type, index) {
        this.isResizing = true;
        this.resizeType = type;
        this.resizeIndex = index;
        document.body.style.cursor = type === 'column' ? 'col-resize' : 'row-resize';
        
        // 阻止页面滚动和其他交互
        document.body.style.userSelect = 'none';
        
        // 缓存需要操作的DOM元素，避免重复查询
        this.cacheResizeElements();
        
        // 记录开始resize时的滚动位置，用于调试
        if (type === 'column') {
            this.resizeStartScrollLeft = this.elements.tableContainer.scrollLeft;
        } else {
            this.resizeStartScrollTop = this.elements.tableContainer.scrollTop;
        }
    }
    
    /**
     * 缓存resize时需要操作的DOM元素
     */
    cacheResizeElements() {
        if (this.resizeType === 'column') {
            this.cachedResizeElements = {
                header: this.elements.columnHeaders.querySelector(`[data-column="${this.resizeIndex}"]`),
                cells: this.elements.tableBody.querySelectorAll(`td[data-column="${this.resizeIndex}"]`)
            };
        } else if (this.resizeType === 'row') {
            this.cachedResizeElements = {
                header: this.elements.rowHeaders.querySelector(`[data-row="${this.resizeIndex}"]`),
                row: this.elements.tableBody.querySelector(`tr[data-row="${this.resizeIndex}"]`)
            };
        }
    }
    
    /**
     * 执行调整大小
     */
    performResize(e) {
        if (!this.isResizing || !this.cachedResizeElements) return;
        
        if (this.resizeType === 'column') {
            // 获取excel-wrapper的滚动位置
            const excelWrapper = document.querySelector('.excel-wrapper');
            const scrollLeft = excelWrapper.scrollLeft;
            
            // 使用列头容器作为基准，因为resize手柄在列头中
            const headerRect = this.elements.columnHeaders.getBoundingClientRect();
            
            // 计算鼠标在列头容器中的位置（考虑滚动偏移）
            const mouseXInHeader = e.clientX - headerRect.left + scrollLeft;
            
            // 计算当前要调整的列的起始位置
            let columnStartX = 0;
            for (let i = 0; i < this.resizeIndex; i++) {
                columnStartX += this.columnWidths[i];
            }
            
            // 计算新的列宽（确保最小宽度）
            const newWidth = Math.max(30, mouseXInHeader - columnStartX);
            this.columnWidths[this.resizeIndex] = newWidth;
            
            // 直接更新缓存的DOM元素，不调用updateColumnWidths避免重复查询
            this.updateCachedColumnElements(newWidth);
            
        } else if (this.resizeType === 'row') {
            // 获取当前正在调整的行头元素的位置
            const headerRect = this.cachedResizeElements.header.getBoundingClientRect();
            
            // 计算新的行高：鼠标Y位置减去当前行头的顶部位置
            const newHeight = Math.max(20, e.clientY - headerRect.top);
            this.rowHeights[this.resizeIndex] = newHeight;
            
            // 直接更新缓存的DOM元素，不调用updateRowHeights避免重复查询
            this.updateCachedRowElements(newHeight);
        }
    }
    
    /**
     * 更新缓存的列元素（性能优化版本）
     */
    updateCachedColumnElements(width) {
        const widthPx = width + 'px';
        
        // 更新列头
        if (this.cachedResizeElements.header) {
            this.cachedResizeElements.header.style.width = widthPx;
        }
        
        // 更新该列的所有单元格
        this.cachedResizeElements.cells.forEach(cell => {
            cell.style.width = widthPx;
        });
    }
    
    /**
     * 更新缓存的行元素（性能优化版本）
     */
    updateCachedRowElements(height) {
        const heightPx = height + 'px';
        
        // 更新行头
        if (this.cachedResizeElements.header) {
            this.cachedResizeElements.header.style.height = heightPx;
        }
        
        // 更新该行及其所有单元格
        if (this.cachedResizeElements.row) {
            this.cachedResizeElements.row.style.height = heightPx;
            const cells = this.cachedResizeElements.row.querySelectorAll('td');
            cells.forEach(cell => {
                cell.style.height = heightPx;
            });
        }
    }
    
    /**
     * 停止调整大小
     */
    stopResize() {
        if (!this.isResizing) return;
        
        // 取消动画帧
        if (this.resizeAnimationFrame) {
            cancelAnimationFrame(this.resizeAnimationFrame);
            this.resizeAnimationFrame = null;
        }
        
        // 在resize结束时才更新总尺寸，避免频繁重计算
        if (this.resizeType === 'column') {
            this.updateTableWidth();
        } else if (this.resizeType === 'row') {
            // 只在resize结束时更新总高度
            const totalHeight = this.rowHeights.reduce((sum, height) => sum + height, 0);
            this.elements.rowHeaders.style.height = totalHeight + 'px';
            
            const table = this.elements.tableBody.parentElement;
            if (table) {
                table.style.height = totalHeight + 'px';
            }
        }
        
        this.isResizing = false;
        this.resizeType = null;
        this.resizeIndex = null;
        this.cachedResizeElements = null; // 清除缓存
        document.body.style.cursor = 'default';
        document.body.style.userSelect = '';
        
        // 清除调试信息
        this.resizeStartScrollLeft = null;
        this.resizeStartScrollTop = null;
    }
    
    /**
     * 更新列宽
     */
    updateColumnWidths() {
        // 更新特定列的列头宽度
        const columnHeader = this.elements.columnHeaders.querySelector(`[data-column="${this.resizeIndex}"]`);
        if (columnHeader) {
            columnHeader.style.width = this.columnWidths[this.resizeIndex] + 'px';
        }
        
        // 更新特定列的所有单元格宽度
        const cells = this.elements.tableBody.querySelectorAll(`td[data-column="${this.resizeIndex}"]`);
        cells.forEach(cell => {
            cell.style.width = this.columnWidths[this.resizeIndex] + 'px';
        });
        
        // 更新表格总宽度以保持对齐
        this.updateTableWidth();
    }
    
    /**
     * 更新行高
     */
    updateRowHeights() {
        // 更新特定行的行头高度
        const rowHeader = this.elements.rowHeaders.querySelector(`[data-row="${this.resizeIndex}"]`);
        if (rowHeader) {
            rowHeader.style.height = this.rowHeights[this.resizeIndex] + 'px';
        }
        
        // 更新特定行的所有单元格高度
        const row = this.elements.tableBody.querySelector(`tr[data-row="${this.resizeIndex}"]`);
        if (row) {
            const cells = row.querySelectorAll('td');
            cells.forEach(cell => {
                cell.style.height = this.rowHeights[this.resizeIndex] + 'px';
            });
            // 同时更新行的高度
            row.style.height = this.rowHeights[this.resizeIndex] + 'px';
        }
        
        // 重新计算并更新行头容器总高度
        const totalHeight = this.rowHeights.reduce((sum, height) => sum + height, 0);
        this.elements.rowHeaders.style.height = totalHeight + 'px';
        
        // 更新表格总高度
        const table = this.elements.tableBody.parentElement;
        if (table) {
            table.style.height = totalHeight + 'px';
        }
    }
    
    /**
     * 更新所有列宽（用于表格重新生成）
     */
    updateAllColumnWidths() {
        // 更新所有列头宽度
        const columnHeaders = this.elements.columnHeaders.querySelectorAll('.column-header');
        columnHeaders.forEach((header, index) => {
            header.style.width = this.columnWidths[index] + 'px';
        });
        
        // 更新所有表格单元格宽度
        const rows = this.elements.tableBody.querySelectorAll('tr');
        rows.forEach(row => {
            const cells = row.querySelectorAll('td');
            cells.forEach((cell, index) => {
                cell.style.width = this.columnWidths[index] + 'px';
            });
        });
    }
    
    /**
     * 更新所有行高（用于表格重新生成）
     */
    updateAllRowHeights() {
        // 更新所有行头高度
        const rowHeaders = this.elements.rowHeaders.querySelectorAll('.row-header');
        rowHeaders.forEach((header, index) => {
            header.style.height = this.rowHeights[index] + 'px';
        });
        
        // 更新所有表格行高度
        const rows = this.elements.tableBody.querySelectorAll('tr');
        rows.forEach((row, index) => {
            const cells = row.querySelectorAll('td');
            cells.forEach(cell => {
                cell.style.height = this.rowHeights[index] + 'px';
            });
        });
    }
    
    /**
     * 选择单元格
     */
    selectCell(row, col, ctrlKey = false, shiftKey = false) {
        if (!ctrlKey && !shiftKey) {
            this.clearSelection();
        }
        
        const cellKey = `${row}-${col}`;
        
        if (shiftKey && this.selectedCells.size > 0) {
            // Shift选择：选择范围
            const firstSelected = Array.from(this.selectedCells)[0];
            const [startRow, startCol] = firstSelected.split('-').map(Number);
            this.selectRange(startRow, startCol, row, col);
        } else if (ctrlKey) {
            // Ctrl选择：切换选择状态
            if (this.selectedCells.has(cellKey)) {
                this.selectedCells.delete(cellKey);
            } else {
                this.selectedCells.add(cellKey);
            }
        } else {
            // 普通选择
            this.selectedCells.add(cellKey);
        }
        
        this.updateSelection();
    }
    
    /**
     * 选择范围
     */
    selectRange(startRow, startCol, endRow, endCol) {
        this.clearSelection();
        
        const minRow = Math.min(startRow, endRow);
        const maxRow = Math.max(startRow, endRow);
        const minCol = Math.min(startCol, endCol);
        const maxCol = Math.max(startCol, endCol);
        
        for (let row = minRow; row <= maxRow; row++) {
            for (let col = minCol; col <= maxCol; col++) {
                this.selectedCells.add(`${row}-${col}`);
            }
        }
    }
    
    /**
     * 扩展选择
     */
    extendSelection(startRow, startCol, endRow, endCol) {
        this.selectRange(startRow, startCol, endRow, endCol);
        this.updateSelection();
    }
    
    /**
     * 选择整列
     */
    selectColumn(colIndex, ctrlKey = false) {
        if (!ctrlKey) {
            this.clearSelection();
        }
        
        for (let row = 0; row < this.options.rows; row++) {
            this.selectedCells.add(`${row}-${colIndex}`);
        }
        
        this.updateSelection();
    }
    
    /**
     * 选择整行
     */
    selectRow(rowIndex, ctrlKey = false) {
        if (!ctrlKey) {
            this.clearSelection();
        }
        
        for (let col = 0; col < this.options.columns; col++) {
            this.selectedCells.add(`${rowIndex}-${col}`);
        }
        
        this.updateSelection();
    }
    
    /**
     * 清除选择
     */
    clearSelection() {
        this.selectedCells.clear();
    }
    
    /**
     * 更新选择显示
     */
    updateSelection() {
        // 清除所有选择样式
        const allCells = this.elements.tableBody.querySelectorAll('td');
        allCells.forEach(cell => {
            cell.classList.remove('selected', 'multi-selected');
        });
        
        // 应用新的选择样式
        let isFirst = true;
        this.selectedCells.forEach(cellKey => {
            const [row, col] = cellKey.split('-').map(Number);
            const cell = this.elements.tableBody.querySelector(`td[data-row="${row}"][data-column="${col}"]`);
            if (cell) {
                if (isFirst && this.selectedCells.size === 1) {
                    cell.classList.add('selected');
                    isFirst = false;
                } else {
                    cell.classList.add('multi-selected');
                }
            }
        });
        
        // 自动更新头部选择状态
        this.updateHeaderSelection();
    }
    
    /**
     * 更新行列头选择状态
     */
    updateHeaderSelection() {
        // 清除所有头部选择和高亮
        const columnHeaders = this.elements.columnHeaders.querySelectorAll('.column-header');
        const rowHeaders = this.elements.rowHeaders.querySelectorAll('.row-header');
        
        columnHeaders.forEach(header => {
            header.classList.remove('selected', 'cell-highlighted');
        });
        rowHeaders.forEach(header => {
            header.classList.remove('selected', 'cell-highlighted');
        });
        
        // 检查是否选择了整列或整行
        const selectedColumns = new Set();
        const selectedRows = new Set();
        
        this.selectedCells.forEach(cellKey => {
            const [row, col] = cellKey.split('-').map(Number);
            selectedColumns.add(col);
            selectedRows.add(row);
        });
        
        // 高亮完全选中的列
        selectedColumns.forEach(col => {
            let isFullColumn = true;
            for (let row = 0; row < this.options.rows; row++) {
                if (!this.selectedCells.has(`${row}-${col}`)) {
                    isFullColumn = false;
                    break;
                }
            }
            if (isFullColumn) {
                const header = this.elements.columnHeaders.querySelector(`[data-column="${col}"]`);
                if (header) header.classList.add('selected');
            } else {
                // 如果不是完全选中，但有单元格被选中，添加高亮
                const header = this.elements.columnHeaders.querySelector(`[data-column="${col}"]`);
                if (header) header.classList.add('cell-highlighted');
            }
        });
        
        // 高亮完全选中的行
        selectedRows.forEach(row => {
            let isFullRow = true;
            for (let col = 0; col < this.options.columns; col++) {
                if (!this.selectedCells.has(`${row}-${col}`)) {
                    isFullRow = false;
                    break;
                }
            }
            if (isFullRow) {
                const header = this.elements.rowHeaders.querySelector(`[data-row="${row}"]`);
                if (header) header.classList.add('selected');
            } else {
                // 如果不是完全选中，但有单元格被选中，添加高亮
                const header = this.elements.rowHeaders.querySelector(`[data-row="${row}"]`);
                if (header) header.classList.add('cell-highlighted');
            }
        });
    }
    
    /**
     * 更新单元格对应的行首列首高亮（已由updateHeaderSelection统一处理）
     */
    updateCellHighlight(row, col) {
        // 这个方法现在主要由updateHeaderSelection处理
        // 保留此方法以便向后兼容
    }
    
    /**
     * 清除单元格对应的行首列首高亮
     */
    clearCellHighlight() {
        const columnHeaders = this.elements.columnHeaders.querySelectorAll('.column-header');
        const rowHeaders = this.elements.rowHeaders.querySelectorAll('.row-header');
        
        columnHeaders.forEach(header => header.classList.remove('cell-highlighted'));
        rowHeaders.forEach(header => header.classList.remove('cell-highlighted'));
    }
    
    /**
     * 开始编辑单元格
     */
    startEditing(row, col) {
        if (this.editingCell) {
            this.finishEditing();
        }
        
        const cell = this.elements.tableBody.querySelector(`td[data-row="${row}"][data-column="${col}"]`);
        if (!cell) return;
        
        this.editingCell = { row, col, cell };
        
        // 获取单元格在页面中的位置和尺寸
        const rect = cell.getBoundingClientRect();
        
        // 定位编辑器 - 稍微内缩以不遮挡边框
        this.cellEditor.style.position = 'fixed';
        this.cellEditor.style.left = (rect.left + 1) + 'px';
        this.cellEditor.style.top = (rect.top + 1) + 'px';
        this.cellEditor.style.width = (rect.width - 2) + 'px';
        this.cellEditor.style.height = (rect.height - 2) + 'px';
        this.cellEditor.style.display = 'block';
        
        // 设置编辑器内容
        const cellKey = `${row}-${col}`;
        this.cellEditor.value = this.data.get(cellKey) || '';
        
        // 标记单元格为编辑状态
        cell.classList.add('editing');
        
        // 聚焦并选中文本
        this.cellEditor.focus();
        this.cellEditor.select();
    }
    
    /**
     * 完成编辑
     */
    finishEditing() {
        if (!this.editingCell) return;
        
        const { row, col, cell } = this.editingCell;
        const cellKey = `${row}-${col}`;
        const value = this.cellEditor.value;
        
        // 保存数据
        if (value) {
            this.data.set(cellKey, value);
        } else {
            this.data.delete(cellKey);
        }
        
        // 更新单元格显示
        cell.textContent = value;
        cell.classList.remove('editing');
        
        // 隐藏编辑器
        this.cellEditor.style.display = 'none';
        this.editingCell = null;
    }
    
    /**
     * 处理键盘事件
     */
    handleKeyDown(e) {
        if (this.editingCell) return; // 编辑状态下不处理
        
        switch (e.key) {
            case 'Delete':
            case 'Backspace':
                this.clearSelectedCells();
                e.preventDefault();
                break;
            case 'Enter':
                if (this.selectedCells.size === 1) {
                    const cellKey = Array.from(this.selectedCells)[0];
                    const [row, col] = cellKey.split('-').map(Number);
                    this.startEditing(row, col);
                }
                e.preventDefault();
                break;
            case 'F2':
                if (this.selectedCells.size === 1) {
                    const cellKey = Array.from(this.selectedCells)[0];
                    const [row, col] = cellKey.split('-').map(Number);
                    this.startEditing(row, col);
                }
                e.preventDefault();
                break;
            case 'c':
                if (e.ctrlKey) {
                    this.copySelectedCells();
                    e.preventDefault();
                }
                break;
            case 'v':
                if (e.ctrlKey) {
                    this.pasteClipboard();
                    e.preventDefault();
                }
                break;
            case 'a':
                if (e.ctrlKey) {
                    this.selectAll();
                    e.preventDefault();
                }
                break;
        }
    }
    
    /**
     * 处理编辑器键盘事件
     */
    handleEditorKeyDown(e) {
        switch (e.key) {
            case 'Enter':
                this.finishEditing();
                e.preventDefault();
                break;
            case 'Escape':
                this.cancelEditing();
                e.preventDefault();
                break;
        }
    }
    
    /**
     * 取消编辑
     */
    cancelEditing() {
        if (!this.editingCell) return;
        
        const { cell } = this.editingCell;
        cell.classList.remove('editing');
        this.cellEditor.style.display = 'none';
        this.editingCell = null;
    }
    
    /**
     * 清除选中单元格的内容
     */
    clearSelectedCells() {
        this.selectedCells.forEach(cellKey => {
            const [row, col] = cellKey.split('-').map(Number);
            this.data.delete(cellKey);
            
            const cell = this.elements.tableBody.querySelector(`td[data-row="${row}"][data-column="${col}"]`);
            if (cell) {
                cell.textContent = '';
            }
        });
    }
    
    /**
     * 复制选中单元格
     */
    copySelectedCells() {
        const copyData = [];
        this.selectedCells.forEach(cellKey => {
            const [row, col] = cellKey.split('-').map(Number);
            copyData.push({
                row, col,
                value: this.data.get(cellKey) || ''
            });
        });
        
        this.clipboard = copyData;
        
        // 显示复制状态
        this.selectedCells.forEach(cellKey => {
            const [row, col] = cellKey.split('-').map(Number);
            const cell = this.elements.tableBody.querySelector(`td[data-row="${row}"][data-column="${col}"]`);
            if (cell) {
                cell.classList.add('copied');
                setTimeout(() => cell.classList.remove('copied'), 1000);
            }
        });
    }
    
    /**
     * 粘贴剪贴板内容
     */
    pasteClipboard() {
        if (!this.clipboard || this.selectedCells.size === 0) return;
        
        const firstSelected = Array.from(this.selectedCells)[0];
        const [startRow, startCol] = firstSelected.split('-').map(Number);
        
        this.clipboard.forEach(item => {
            const newRow = startRow + (item.row - this.clipboard[0].row);
            const newCol = startCol + (item.col - this.clipboard[0].col);
            
            if (newRow < this.options.rows && newCol < this.options.columns) {
                const cellKey = `${newRow}-${newCol}`;
                this.data.set(cellKey, item.value);
                
                const cell = this.elements.tableBody.querySelector(`td[data-row="${newRow}"][data-column="${newCol}"]`);
                if (cell) {
                    cell.textContent = item.value;
                }
            }
        });
    }
    
    /**
     * 选择全部
     */
    selectAll() {
        this.clearSelection();
        for (let row = 0; row < this.options.rows; row++) {
            for (let col = 0; col < this.options.columns; col++) {
                this.selectedCells.add(`${row}-${col}`);
            }
        }
        this.updateSelection();
    }
    
    /**
     * 处理右键菜单
     */
    handleContextMenu(e) {
        e.preventDefault();
        
        const contextMenu = this.elements.contextMenu;
        contextMenu.style.left = e.pageX + 'px';
        contextMenu.style.top = e.pageY + 'px';
        contextMenu.style.display = 'block';
    }
    
    /**
     * 处理文档点击（隐藏右键菜单）
     */
    handleDocumentClick(e) {
        if (!this.elements.contextMenu.contains(e.target)) {
            this.elements.contextMenu.style.display = 'none';
        }
    }
    
    /**
     * 处理菜单点击
     */
    handleMenuClick(e) {
        const action = e.target.dataset.action;
        if (!action) return;
        
        this.elements.contextMenu.style.display = 'none';
        
        switch (action) {
            case 'insertRowAbove':
                this.insertRow(this.getSelectedRow(), 'above');
                break;
            case 'insertRowBelow':
                this.insertRow(this.getSelectedRow(), 'below');
                break;
            case 'deleteRow':
                this.deleteRow(this.getSelectedRow());
                break;
            case 'insertColumnLeft':
                this.insertColumn(this.getSelectedColumn(), 'left');
                break;
            case 'insertColumnRight':
                this.insertColumn(this.getSelectedColumn(), 'right');
                break;
            case 'deleteColumn':
                this.deleteColumn(this.getSelectedColumn());
                break;
            case 'copy':
                this.copySelectedCells();
                break;
            case 'paste':
                this.pasteClipboard();
                break;
            case 'clear':
                this.clearSelectedCells();
                break;
        }
    }
    
    /**
     * 获取选中的行
     */
    getSelectedRow() {
        if (this.selectedCells.size === 0) return 0;
        const firstSelected = Array.from(this.selectedCells)[0];
        return parseInt(firstSelected.split('-')[0]);
    }
    
    /**
     * 获取选中的列
     */
    getSelectedColumn() {
        if (this.selectedCells.size === 0) return 0;
        const firstSelected = Array.from(this.selectedCells)[0];
        return parseInt(firstSelected.split('-')[1]);
    }
    
    /**
     * 插入行
     */
    insertRow(rowIndex, position = 'below') {
        const insertIndex = position === 'above' ? rowIndex : rowIndex + 1;
        
        // 更新数据
        const newData = new Map();
        this.data.forEach((value, key) => {
            const [row, col] = key.split('-').map(Number);
            if (row >= insertIndex) {
                newData.set(`${row + 1}-${col}`, value);
            } else {
                newData.set(key, value);
            }
        });
        this.data = newData;
        
        // 更新尺寸数组
        this.rowHeights.splice(insertIndex, 0, this.options.defaultRowHeight);
        this.options.rows++;
        
        // 重新生成表格
        this.generateTable();
    }
    
    /**
     * 删除行
     */
    deleteRow(rowIndex) {
        if (this.options.rows <= 1) return;
        
        // 更新数据
        const newData = new Map();
        this.data.forEach((value, key) => {
            const [row, col] = key.split('-').map(Number);
            if (row < rowIndex) {
                newData.set(key, value);
            } else if (row > rowIndex) {
                newData.set(`${row - 1}-${col}`, value);
            }
        });
        this.data = newData;
        
        // 更新尺寸数组
        this.rowHeights.splice(rowIndex, 1);
        this.options.rows--;
        
        // 清除选择
        this.clearSelection();
        
        // 重新生成表格
        this.generateTable();
    }
    
    /**
     * 插入列
     */
    insertColumn(colIndex, position = 'right') {
        const insertIndex = position === 'left' ? colIndex : colIndex + 1;
        
        // 更新数据
        const newData = new Map();
        this.data.forEach((value, key) => {
            const [row, col] = key.split('-').map(Number);
            if (col >= insertIndex) {
                newData.set(`${row}-${col + 1}`, value);
            } else {
                newData.set(key, value);
            }
        });
        this.data = newData;
        
        // 更新尺寸数组
        this.columnWidths.splice(insertIndex, 0, this.options.defaultColumnWidth);
        this.options.columns++;
        
        // 重新生成表格
        this.generateTable();
    }
    
    /**
     * 删除列
     */
    deleteColumn(colIndex) {
        if (this.options.columns <= 1) return;
        
        // 更新数据
        const newData = new Map();
        this.data.forEach((value, key) => {
            const [row, col] = key.split('-').map(Number);
            if (col < colIndex) {
                newData.set(key, value);
            } else if (col > colIndex) {
                newData.set(`${row}-${col - 1}`, value);
            }
        });
        this.data = newData;
        
        // 更新尺寸数组
        this.columnWidths.splice(colIndex, 1);
        this.options.columns--;
        
        // 清除选择
        this.clearSelection();
        
        // 重新生成表格
        this.generateTable();
    }
    
    /**
     * 添加行
     */
    addRow() {
        this.rowHeights.push(this.options.defaultRowHeight);
        this.options.rows++;
        this.generateTable();
    }
    
    /**
     * 添加列
     */
    addColumn() {
        this.columnWidths.push(this.options.defaultColumnWidth);
        this.options.columns++;
        this.generateTable();
    }
    
    /**
     * 重新生成表格
     */
    regenerateTable() {
        const newRows = parseInt(document.getElementById('rowCount').value);
        const newCols = parseInt(document.getElementById('colCount').value);
        
        if (newRows > 0 && newCols > 0 && newRows <= 1000 && newCols <= 100) {
            this.options.rows = newRows;
            this.options.columns = newCols;
            
            // 调整尺寸数组
            this.rowHeights = new Array(newRows).fill(this.options.defaultRowHeight);
            this.columnWidths = new Array(newCols).fill(this.options.defaultColumnWidth);
            
            // 清除数据和选择
            this.data.clear();
            this.clearSelection();
            
            // 重新生成表格
            this.generateTable();
        }
    }
    
    /**
     * 处理滚动同步
     */
    handleScroll(e) {
        // 如果正在编辑，隐藏编辑器或重新定位
        if (this.editingCell) {
            this.finishEditing();
        }
    }
}

// 初始化Excel模拟器
document.addEventListener('DOMContentLoaded', () => {
    const container = document.querySelector('.excel-container');
    window.excelSimulator = new ExcelSimulator(container, {
        rows: 100,
        columns: 100
    });
});
