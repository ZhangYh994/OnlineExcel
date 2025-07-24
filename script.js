/**
 * 虚拟Excel模拟器
 * 使用虚拟滚动技术，只渲染可见区域的单元格
 * 动态生成行列，提升性能
 */
class VirtualExcelSimulator {
    constructor(container, options = {}) {
        this.container = container;
        this.options = {
            defaultRowHeight: 40,
            defaultColumnWidth: 120,
            visibleRowBuffer: 5,  // 可见区域外缓冲的行数
            visibleColumnBuffer: 3, // 可见区域外缓冲的列数
            minRows: 100,
            minColumns: 26,
            maxRows: 10000,
            maxColumns: 1000,
            ...options
        };
        
        // 数据存储
        this.data = new Map(); // 存储单元格数据
        this.modifiedCells = new Set(); // 记录被修改过的单元格
        this.selectedCells = new Set(); // 当前选中的单元格
        this.clipboard = null; // 剪贴板数据
        
        // 虚拟滚动相关
        this.scrollTop = 0;
        this.scrollLeft = 0;
        this.containerWidth = 0;
        this.containerHeight = 0;
        this.visibleStartRow = 0;
        this.visibleEndRow = 0;
        this.visibleStartColumn = 0;
        this.visibleEndColumn = 0;
        
        // 行列尺寸
        this.rowHeights = new Array(this.options.minRows).fill(this.options.defaultRowHeight);
        this.columnWidths = new Array(this.options.minColumns).fill(this.options.defaultColumnWidth);
        this.rowOffsets = [0]; // 行的累积偏移量
        this.columnOffsets = [0]; // 列的累积偏移量
        
        // DOM元素引用
        this.elements = {};
        
        // 编辑和选择状态
        this.editingCell = null;
        this.cellEditor = null;
        this.isSelecting = false;
        this.selectionStart = null;
        
        // 调整大小相关
        this.isResizing = false;
        this.resizeType = null;
        this.resizeIndex = null;
        this.resizeAnimationFrame = null;
        this.scrollAnimationFrame = null;
        this.selectionAnimationFrame = null; // 添加选择操作的动画帧
        
        this.init();
    }
    
    /**
     * 初始化
     */
    init() {
        this.createStructure();
        this.bindEvents();
        this.calculateOffsets();
        this.updateContainerSize();
        // 延迟渲染以确保DOM准备就绪
        setTimeout(() => {
            this.calculateVisibleRange();
            this.renderVisibleCells();
            this.updateHeadersPosition();
        }, 0);
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
        
        // 创建虚拟滚动容器
        this.elements.virtualContainer = document.createElement('div');
        this.elements.virtualContainer.className = 'virtual-container';
        this.elements.virtualContainer.style.position = 'relative';
        this.elements.virtualContainer.style.overflow = 'hidden';
        
        // 创建表格元素
        this.elements.table = document.createElement('table');
        this.elements.table.className = 'excel-table';
        this.elements.table.style.position = 'absolute';
        this.elements.table.style.top = '0';
        this.elements.table.style.left = '0';
        
        this.elements.virtualContainer.appendChild(this.elements.table);
        this.elements.tableContainer.appendChild(this.elements.virtualContainer);
    }
    
    /**
     * 绑定事件监听器
     */
    bindEvents() {
        // 工具栏事件
        document.getElementById('addRow').addEventListener('click', () => this.addRows(10));
        document.getElementById('addColumn').addEventListener('click', () => this.addColumns(5));
        document.getElementById('regenerate').addEventListener('click', () => this.regenerateTable());
        
        // 滚动事件
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
        
        // 窗口大小变化
        window.addEventListener('resize', () => this.updateContainerSize());
    }
    
    /**
     * 计算行列偏移量
     */
    calculateOffsets() {
        // 计算行偏移量
        this.rowOffsets = [0];
        for (let i = 0; i < this.rowHeights.length; i++) {
            this.rowOffsets.push(this.rowOffsets[i] + this.rowHeights[i]);
        }
        
        // 计算列偏移量
        this.columnOffsets = [0];
        for (let i = 0; i < this.columnWidths.length; i++) {
            this.columnOffsets.push(this.columnOffsets[i] + this.columnWidths[i]);
        }
    }
    
    /**
     * 更新容器尺寸
     */
    updateContainerSize() {
        const wrapper = document.querySelector('.excel-wrapper');
        this.containerWidth = wrapper.clientWidth - 60; // 减去行头宽度
        this.containerHeight = wrapper.clientHeight - 30; // 减去列头高度
        
        // 更新虚拟容器尺寸
        const totalHeight = this.rowOffsets[this.rowOffsets.length - 1];
        const totalWidth = this.columnOffsets[this.columnOffsets.length - 1];
        
        this.elements.virtualContainer.style.width = totalWidth + 'px';
        this.elements.virtualContainer.style.height = totalHeight + 'px';
        
        this.calculateVisibleRange();
    }
    
    /**
     * 计算可见范围
     */
    calculateVisibleRange() {
        const wrapper = document.querySelector('.excel-wrapper');
        this.scrollTop = wrapper.scrollTop;
        this.scrollLeft = wrapper.scrollLeft;
        
        // 计算可见行范围
        this.visibleStartRow = Math.max(0, this.findRowByOffset(this.scrollTop) - this.options.visibleRowBuffer);
        this.visibleEndRow = Math.min(
            this.rowHeights.length - 1,
            this.findRowByOffset(this.scrollTop + this.containerHeight) + this.options.visibleRowBuffer
        );
        
        // 计算可见列范围
        this.visibleStartColumn = Math.max(0, this.findColumnByOffset(this.scrollLeft) - this.options.visibleColumnBuffer);
        this.visibleEndColumn = Math.min(
            this.columnWidths.length - 1,
            this.findColumnByOffset(this.scrollLeft + this.containerWidth) + this.options.visibleColumnBuffer
        );
        
        // 检查是否需要扩展行列
        this.checkAndExpandGrid();
    }
    
    /**
     * 根据偏移量查找行索引
     */
    findRowByOffset(offset) {
        let left = 0, right = this.rowOffsets.length - 1;
        while (left < right) {
            const mid = Math.floor((left + right) / 2);
            if (this.rowOffsets[mid] <= offset) {
                left = mid + 1;
            } else {
                right = mid;
            }
        }
        return Math.max(0, left - 1);
    }
    
    /**
     * 根据偏移量查找列索引
     */
    findColumnByOffset(offset) {
        let left = 0, right = this.columnOffsets.length - 1;
        while (left < right) {
            const mid = Math.floor((left + right) / 2);
            if (this.columnOffsets[mid] <= offset) {
                left = mid + 1;
            } else {
                right = mid;
            }
        }
        return Math.max(0, left - 1);
    }
    
    /**
     * 检查并扩展网格
     */
    checkAndExpandGrid() {
        let needsUpdate = false;
        
        // 检查是否需要添加更多行
        if (this.visibleEndRow >= this.rowHeights.length - 10 && this.rowHeights.length < this.options.maxRows) {
            this.addRows(50);
            needsUpdate = true;
        }
        
        // 检查是否需要添加更多列
        if (this.visibleEndColumn >= this.columnWidths.length - 5 && this.columnWidths.length < this.options.maxColumns) {
            this.addColumns(20);
            needsUpdate = true;
        }
        
        if (needsUpdate) {
            this.calculateOffsets();
            this.updateContainerSize();
        }
    }
    
    /**
     * 添加行
     */
    addRows(count) {
        const startIndex = this.rowHeights.length;
        for (let i = 0; i < count; i++) {
            this.rowHeights.push(this.options.defaultRowHeight);
        }
        this.calculateOffsets();
        this.updateRowHeaders();
    }
    
    /**
     * 添加列
     */
    addColumns(count) {
        const startIndex = this.columnWidths.length;
        for (let i = 0; i < count; i++) {
            this.columnWidths.push(this.options.defaultColumnWidth);
        }
        this.calculateOffsets();
        this.updateColumnHeaders();
    }
    
    /**
     * 处理滚动事件
     */
    handleScroll(e) {
        // 使用 requestAnimationFrame 来节流滚动事件
        if (this.scrollAnimationFrame) {
            cancelAnimationFrame(this.scrollAnimationFrame);
        }
        
        this.scrollAnimationFrame = requestAnimationFrame(() => {
            this.calculateVisibleRange();
            this.renderVisibleCells();
            this.updateHeadersPosition();
            
            // 如果正在编辑，隐藏编辑器
            if (this.editingCell) {
                this.cellEditor.style.display = 'none';
            }
        });
    }
    
    /**
     * 渲染可见的单元格
     */
    renderVisibleCells() {
        // 清空现有内容
        this.elements.table.innerHTML = '';
        
        for (let row = this.visibleStartRow; row <= this.visibleEndRow; row++) {
            for (let col = this.visibleStartColumn; col <= this.visibleEndColumn; col++) {
                const td = document.createElement('td');
                td.dataset.row = row;
                td.dataset.column = col;
                td.style.position = 'absolute';
                td.style.left = this.columnOffsets[col] + 'px';
                td.style.top = this.rowOffsets[row] + 'px';
                td.style.width = this.columnWidths[col] + 'px';
                td.style.height = this.rowHeights[row] + 'px';
                
                // 设置单元格内容
                const cellKey = `${row}-${col}`;
                const cellValue = this.data.get(cellKey) || '';
                td.textContent = cellValue;
                
                // 应用选择样式
                if (this.selectedCells.has(cellKey)) {
                    td.classList.add(this.selectedCells.size === 1 ? 'selected' : 'multi-selected');
                    // 如果是多选，应用智能边框
                    if (this.selectedCells.size > 1) {
                        this.applySmartBorders(td, row, col);
                    }
                }
                
                this.elements.table.appendChild(td);
            }
        }
        
        // 更新行首列首高亮
        this.updateHeaderHighlights();
    }
    
    /**
     * 轻量级选择更新（不重建DOM）- 优化版
     */
    updateSelection() {
        // 批量清除所有现有的选择样式
        const allCells = this.elements.table.querySelectorAll('td');
        
        // 使用一次性样式重置以提高性能
        allCells.forEach(cell => {
            cell.className = cell.className.replace(/\b(selected|multi-selected)\b/g, '').trim();
            // 重置边框样式
            if (cell.style.borderTop) {
                cell.style.borderTop = '';
                cell.style.borderBottom = '';
                cell.style.borderLeft = '';
                cell.style.borderRight = '';
            }
        });
        
        // 如果没有选择，直接返回
        if (this.selectedCells.size === 0) {
            this.updateHeaderHighlights();
            return;
        }
        
        // 批量应用新的选择样式
        const isMultiSelect = this.selectedCells.size > 1;
        this.selectedCells.forEach(cellKey => {
            const [row, col] = cellKey.split('-').map(Number);
            const cell = this.elements.table.querySelector(`td[data-row="${row}"][data-column="${col}"]`);
            if (cell) {
                if (isMultiSelect) {
                    cell.classList.add('multi-selected');
                    this.applySmartBorders(cell, row, col);
                } else {
                    cell.classList.add('selected');
                }
            }
        });
        
        // 更新行首列首高亮
        this.updateHeaderHighlights();
    }
    
    /**
     * 为多选单元格应用智能边框（优化版）
     */
    applySmartBorders(td, row, col) {
        // 只有在多选时才应用智能边框
        if (this.selectedCells.size <= 1) return;
        
        // 检查相邻单元格是否也被选中
        const topSelected = this.selectedCells.has(`${row-1}-${col}`);
        const bottomSelected = this.selectedCells.has(`${row+1}-${col}`);
        const leftSelected = this.selectedCells.has(`${row}-${col-1}`);
        const rightSelected = this.selectedCells.has(`${row}-${col+1}`);
        
        // 设置边框样式 - 使用CSS类而不是内联样式以提高性能
        const thickBorder = '2px solid #0078d4';  // 选区外边框
        const thinBorder = '1px solid #dee2e6';   // 选区内边框（正常边框）
        
        // 批量设置边框以减少重排
        const styles = {
            borderTop: topSelected ? thinBorder : thickBorder,
            borderBottom: bottomSelected ? thinBorder : thickBorder,
            borderLeft: leftSelected ? thinBorder : thickBorder,
            borderRight: rightSelected ? thinBorder : thickBorder
        };
        
        Object.assign(td.style, styles);
    }
    
    /**
     * 更新行首列首高亮显示
     */
    updateHeaderHighlights() {
        // 如果元素不存在，直接返回
        if (!this.elements.rowHeaders || !this.elements.columnHeaders) {
            return;
        }
        
        // 获取选中单元格的行列索引
        const selectedRows = new Set();
        const selectedCols = new Set();
        
        this.selectedCells.forEach(cellKey => {
            const [row, col] = cellKey.split('-').map(Number);
            selectedRows.add(row);
            selectedCols.add(col);
        });
        
        // 清除所有行头高亮和选中样式
        const allRowHeaders = this.elements.rowHeaders.querySelectorAll('.row-header');
        allRowHeaders.forEach(header => {
            header.classList.remove('cell-highlighted');
            header.classList.remove('selected');
        });
        
        // 清除所有列头高亮和选中样式
        const allColHeaders = this.elements.columnHeaders.querySelectorAll('.column-header');
        allColHeaders.forEach(header => {
            header.classList.remove('cell-highlighted');
            header.classList.remove('selected');
        });
        
        // 高亮选中的行头
        selectedRows.forEach(row => {
            const rowHeader = this.elements.rowHeaders.querySelector(`[data-row="${row}"]`);
            if (rowHeader) {
                // 检查是否选中了整行
                let isFullRowSelected = true;
                for (let col = 0; col < this.columnWidths.length; col++) {
                    if (!this.selectedCells.has(`${row}-${col}`)) {
                        isFullRowSelected = false;
                        break;
                    }
                }
                // 如果整行都被选中，则添加selected类，否则添加cell-highlighted类
                if (isFullRowSelected) {
                    rowHeader.classList.add('selected');
                } else {
                    rowHeader.classList.add('cell-highlighted');
                }
            }
        });
        
        // 高亮选中的列头
        selectedCols.forEach(col => {
            const colHeader = this.elements.columnHeaders.querySelector(`[data-column="${col}"]`);
            if (colHeader) {
                // 检查是否选中了整列
                let isFullColSelected = true;
                for (let row = 0; row < this.rowHeights.length; row++) {
                    if (!this.selectedCells.has(`${row}-${col}`)) {
                        isFullColSelected = false;
                        break;
                    }
                }
                // 如果整列都被选中，则添加selected类，否则添加cell-highlighted类
                if (isFullColSelected) {
                    colHeader.classList.add('selected');
                } else {
                    colHeader.classList.add('cell-highlighted');
                }
            }
        });
    }
    
    /**
     * 更新列头
     */
    updateColumnHeaders() {
        this.elements.columnHeaders.innerHTML = '';
        
        for (let i = this.visibleStartColumn; i <= this.visibleEndColumn; i++) {
            const header = document.createElement('div');
            header.className = 'column-header';
            header.textContent = this.getColumnName(i);
            header.style.position = 'absolute';
            header.style.left = this.columnOffsets[i] + 'px';
            header.style.width = this.columnWidths[i] + 'px';
            header.style.height = '30px';
            header.dataset.column = i;
            
            // 添加调整手柄
            const resizeHandle = document.createElement('div');
            resizeHandle.className = 'column-resize-handle';
            resizeHandle.dataset.column = i;
            header.appendChild(resizeHandle);
            
            this.elements.columnHeaders.appendChild(header);
        }
        
        // 设置列头容器总宽度
        this.elements.columnHeaders.style.width = this.columnOffsets[this.columnOffsets.length - 1] + 'px';
    }
    
    /**
     * 更新行头
     */
    updateRowHeaders() {
        this.elements.rowHeaders.innerHTML = '';
        
        for (let i = this.visibleStartRow; i <= this.visibleEndRow; i++) {
            const header = document.createElement('div');
            header.className = 'row-header';
            header.textContent = i + 1;
            header.style.position = 'absolute';
            header.style.top = this.rowOffsets[i] + 'px';
            header.style.width = '60px';
            header.style.height = this.rowHeights[i] + 'px';
            header.dataset.row = i;
            
            // 添加调整手柄
            const resizeHandle = document.createElement('div');
            resizeHandle.className = 'row-resize-handle';
            resizeHandle.dataset.row = i;
            header.appendChild(resizeHandle);
            
            this.elements.rowHeaders.appendChild(header);
        }
        
        // 设置行头容器总高度
        this.elements.rowHeaders.style.height = this.rowOffsets[this.rowOffsets.length - 1] + 'px';
    }
    
    /**
     * 更新头部位置（滚动时同步）
     */
    updateHeadersPosition() {
        // 更新列头位置
        this.updateColumnHeaders();
        
        // 更新行头位置
        this.updateRowHeaders();
    }
    
    /**
     * 获取列名
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
                if (!e.shiftKey) {
                    this.isSelecting = true;
                    this.selectionStart = { row, col };
                }
            }
            e.preventDefault();
        }
    }
    
    /**
     * 处理鼠标移动事件
     */
    handleMouseMove(e) {
        if (this.isResizing) {
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
                
                // 使用节流机制优化拖拽选择性能
                if (this.selectionAnimationFrame) {
                    cancelAnimationFrame(this.selectionAnimationFrame);
                }
                this.selectionAnimationFrame = requestAnimationFrame(() => {
                    this.extendSelection(this.selectionStart.row, this.selectionStart.col, row, col);
                    this.selectionAnimationFrame = null;
                });
            }
        }
    }
    
    /**
     * 处理鼠标释放事件
     */
    handleMouseUp(e) {
        if (this.isResizing) {
            this.stopResize();
            return;
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
        document.body.style.userSelect = 'none';
    }
    
    /**
     * 执行调整大小
     */
    performResize(e) {
        if (!this.isResizing) return;
        
        if (this.resizeType === 'column') {
            // 获取当前正在调整的列头元素
            const currentColumnHeader = this.elements.columnHeaders.querySelector(`[data-column="${this.resizeIndex}"]`);
            if (!currentColumnHeader) return;
            
            // 获取列头的当前位置和鼠标位置
            const headerRect = currentColumnHeader.getBoundingClientRect();
            
            // 计算新的列宽：鼠标X位置减去当前列头的左侧位置
            const newWidth = Math.max(30, e.clientX - headerRect.left);
            this.columnWidths[this.resizeIndex] = newWidth;
            
            this.calculateOffsets();
            this.updateContainerSize();
            this.renderVisibleCells();
            this.updateColumnHeaders();
            
        } else if (this.resizeType === 'row') {
            // 获取当前正在调整的行头元素
            const currentRowHeader = this.elements.rowHeaders.querySelector(`[data-row="${this.resizeIndex}"]`);
            if (!currentRowHeader) return;
            
            // 获取行头的当前位置和鼠标位置
            const headerRect = currentRowHeader.getBoundingClientRect();
            
            // 计算新的行高：鼠标Y位置减去当前行头的顶部位置
            const newHeight = Math.max(20, e.clientY - headerRect.top);
            this.rowHeights[this.resizeIndex] = newHeight;
            
            this.calculateOffsets();
            this.updateContainerSize();
            this.renderVisibleCells();
            this.updateRowHeaders();
        }
    }
    
    /**
     * 停止调整大小
     */
    stopResize() {
        if (this.resizeAnimationFrame) {
            cancelAnimationFrame(this.resizeAnimationFrame);
            this.resizeAnimationFrame = null;
        }
        
        this.isResizing = false;
        this.resizeType = null;
        this.resizeIndex = null;
        document.body.style.cursor = 'default';
        document.body.style.userSelect = '';
    }
    
    /**
     * 选择单元格
     */
    selectCell(row, col, ctrlKey = false, shiftKey = false) {
        if (!ctrlKey && !shiftKey) {
            this.selectedCells.clear();
            // 清除行首和列首的selected样式
            const allRowHeaders = this.elements.rowHeaders.querySelectorAll('.row-header');
            allRowHeaders.forEach(header => {
                header.classList.remove('selected');
            });
            
            const allColHeaders = this.elements.columnHeaders.querySelectorAll('.column-header');
            allColHeaders.forEach(header => {
                header.classList.remove('selected');
            });
        }
        
        const cellKey = `${row}-${col}`;
        
        if (shiftKey && this.selectedCells.size > 0) {
            // 扩展选择
            const firstSelected = Array.from(this.selectedCells)[0];
            const [startRow, startCol] = firstSelected.split('-').map(Number);
            this.selectRange(startRow, startCol, row, col);
        } else {
            this.selectedCells.add(cellKey);
        }
        
        this.updateSelection();
    }
    
    /**
     * 选择范围
     */
    selectRange(startRow, startCol, endRow, endCol) {
        this.selectedCells.clear();
        
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
            this.selectedCells.clear();
            // 清除之前选中的行头样式
            const allRowHeaders = this.elements.rowHeaders.querySelectorAll('.row-header');
            allRowHeaders.forEach(header => {
                header.classList.remove('selected');
            });
        }
        
        // 为选中的列头添加selected样式
        const colHeader = this.elements.columnHeaders.querySelector(`[data-column="${colIndex}"]`);
        if (colHeader) {
            colHeader.classList.add('selected');
        }
        
        for (let row = 0; row < this.rowHeights.length; row++) {
            this.selectedCells.add(`${row}-${colIndex}`);
        }
        
        this.updateSelection();
    }
    
    /**
     * 选择整行
     */
    selectRow(rowIndex, ctrlKey = false) {
        if (!ctrlKey) {
            this.selectedCells.clear();
            // 清除之前选中的列头样式
            const allColHeaders = this.elements.columnHeaders.querySelectorAll('.column-header');
            allColHeaders.forEach(header => {
                header.classList.remove('selected');
            });
        }
        
        // 为选中的行头添加selected样式
        const rowHeader = this.elements.rowHeaders.querySelector(`[data-row="${rowIndex}"]`);
        if (rowHeader) {
            rowHeader.classList.add('selected');
        }
        
        for (let col = 0; col < this.columnWidths.length; col++) {
            this.selectedCells.add(`${rowIndex}-${col}`);
        }
        
        this.updateSelection();
    }
    
    /**
     * 开始编辑单元格
     */
    startEditing(row, col) {
        if (this.editingCell) {
            this.finishEditing();
        }
        
        const cell = this.elements.table.querySelector(`td[data-row="${row}"][data-column="${col}"]`);
        if (!cell) return;
        
        this.editingCell = { row, col, cell };
        
        // 获取单元格在页面中的位置和尺寸
        const rect = cell.getBoundingClientRect();
        
        // 定位编辑器
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
            this.modifiedCells.add(cellKey);
        } else {
            this.data.delete(cellKey);
            this.modifiedCells.delete(cellKey);
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
            case 'Enter':
                if (this.selectedCells.size === 1) {
                    const cellKey = Array.from(this.selectedCells)[0];
                    const [row, col] = cellKey.split('-').map(Number);
                    this.startEditing(row, col);
                }
                e.preventDefault();
                break;
            case 'Delete':
                this.clearSelectedCells();
                break;
            case 'Escape':
                this.selectedCells.clear();
                // 清除行首和列首的selected样式
                const allRowHeaders = this.elements.rowHeaders.querySelectorAll('.row-header');
                allRowHeaders.forEach(header => {
                    header.classList.remove('selected');
                });
                
                const allColHeaders = this.elements.columnHeaders.querySelectorAll('.column-header');
                allColHeaders.forEach(header => {
                    header.classList.remove('selected');
                });
                this.renderVisibleCells();
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
            this.data.delete(cellKey);
            this.modifiedCells.delete(cellKey);
        });
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
     * 复制选中单元格
     */
    copySelectedCells() {
        const copyData = [];
        this.selectedCells.forEach(cellKey => {
            const value = this.data.get(cellKey) || '';
            copyData.push({ cellKey, value });
        });
        
        this.clipboard = copyData;
    }
    
    /**
     * 粘贴剪贴板内容
     */
    pasteClipboard() {
        if (!this.clipboard || this.selectedCells.size === 0) return;
        
        const firstSelected = Array.from(this.selectedCells)[0];
        const [startRow, startCol] = firstSelected.split('-').map(Number);
        
        this.clipboard.forEach(item => {
            const [origRow, origCol] = item.cellKey.split('-').map(Number);
            const newRow = startRow;
            const newCol = startCol;
            const newCellKey = `${newRow}-${newCol}`;
            
            if (item.value) {
                this.data.set(newCellKey, item.value);
                this.modifiedCells.add(newCellKey);
            }
        });
        
        this.renderVisibleCells();
    }
    
    /**
     * 重新生成表格
     */
    regenerateTable() {
        const newRows = parseInt(document.getElementById('rowCount').value);
        const newCols = parseInt(document.getElementById('colCount').value);
        
        if (newRows > 0 && newCols > 0 && newRows <= 10000 && newCols <= 1000) {
            // 清除所有数据和状态
            this.data.clear();
            this.modifiedCells.clear();
            this.selectedCells.clear();
            this.clipboard = null;
            
            // 重置行列数量
            this.rowHeights = new Array(Math.max(newRows, this.options.minRows)).fill(this.options.defaultRowHeight);
            this.columnWidths = new Array(Math.max(newCols, this.options.minColumns)).fill(this.options.defaultColumnWidth);
            
            // 停止当前编辑
            if (this.editingCell) {
                this.cancelEditing();
            }
            
            // 重置滚动位置到左上角
            const wrapper = document.querySelector('.excel-wrapper');
            wrapper.scrollTop = 0;
            wrapper.scrollLeft = 0;
            this.scrollTop = 0;
            this.scrollLeft = 0;
            
            this.calculateOffsets();
            this.updateContainerSize();
            this.calculateVisibleRange();
            this.renderVisibleCells();
            this.updateHeadersPosition();
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
        
        // 更新数据 - 需要调整现有数据的行索引
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
        
        // 更新修改过的单元格记录
        const newModifiedCells = new Set();
        this.modifiedCells.forEach(cellKey => {
            const [row, col] = cellKey.split('-').map(Number);
            if (row >= insertIndex) {
                newModifiedCells.add(`${row + 1}-${col}`);
            } else {
                newModifiedCells.add(cellKey);
            }
        });
        this.modifiedCells = newModifiedCells;
        
        // 插入新行高度
        this.rowHeights.splice(insertIndex, 0, this.options.defaultRowHeight);
        
        // 重新计算偏移量并渲染
        this.calculateOffsets();
        this.updateContainerSize();
        this.calculateVisibleRange();
        this.renderVisibleCells();
        this.updateHeadersPosition();
    }
    
    /**
     * 删除行
     */
    deleteRow(rowIndex) {
        if (this.rowHeights.length <= 1) return; // 至少保留一行
        
        // 更新数据 - 删除该行数据并调整其他行索引
        const newData = new Map();
        this.data.forEach((value, key) => {
            const [row, col] = key.split('-').map(Number);
            if (row === rowIndex) {
                // 删除该行的数据
                return;
            } else if (row > rowIndex) {
                newData.set(`${row - 1}-${col}`, value);
            } else {
                newData.set(key, value);
            }
        });
        this.data = newData;
        
        // 更新修改过的单元格记录
        const newModifiedCells = new Set();
        this.modifiedCells.forEach(cellKey => {
            const [row, col] = cellKey.split('-').map(Number);
            if (row === rowIndex) {
                // 删除该行的记录
                return;
            } else if (row > rowIndex) {
                newModifiedCells.add(`${row - 1}-${col}`);
            } else {
                newModifiedCells.add(cellKey);
            }
        });
        this.modifiedCells = newModifiedCells;
        
        // 删除行高度
        this.rowHeights.splice(rowIndex, 1);
        
        // 清除选择
        this.selectedCells.clear();
        
        // 重新计算偏移量并渲染
        this.calculateOffsets();
        this.updateContainerSize();
        this.calculateVisibleRange();
        this.renderVisibleCells();
        this.updateHeadersPosition();
    }
    
    /**
     * 插入列
     */
    insertColumn(colIndex, position = 'right') {
        const insertIndex = position === 'left' ? colIndex : colIndex + 1;
        
        // 更新数据 - 需要调整现有数据的列索引
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
        
        // 更新修改过的单元格记录
        const newModifiedCells = new Set();
        this.modifiedCells.forEach(cellKey => {
            const [row, col] = cellKey.split('-').map(Number);
            if (col >= insertIndex) {
                newModifiedCells.add(`${row}-${col + 1}`);
            } else {
                newModifiedCells.add(cellKey);
            }
        });
        this.modifiedCells = newModifiedCells;
        
        // 插入新列宽度
        this.columnWidths.splice(insertIndex, 0, this.options.defaultColumnWidth);
        
        // 重新计算偏移量并渲染
        this.calculateOffsets();
        this.updateContainerSize();
        this.calculateVisibleRange();
        this.renderVisibleCells();
        this.updateHeadersPosition();
    }
    
    /**
     * 删除列
     */
    deleteColumn(colIndex) {
        if (this.columnWidths.length <= 1) return; // 至少保留一列
        
        // 更新数据 - 删除该列数据并调整其他列索引
        const newData = new Map();
        this.data.forEach((value, key) => {
            const [row, col] = key.split('-').map(Number);
            if (col === colIndex) {
                // 删除该列的数据
                return;
            } else if (col > colIndex) {
                newData.set(`${row}-${col - 1}`, value);
            } else {
                newData.set(key, value);
            }
        });
        this.data = newData;
        
        // 更新修改过的单元格记录
        const newModifiedCells = new Set();
        this.modifiedCells.forEach(cellKey => {
            const [row, col] = cellKey.split('-').map(Number);
            if (col === colIndex) {
                // 删除该列的记录
                return;
            } else if (col > colIndex) {
                newModifiedCells.add(`${row}-${col - 1}`);
            } else {
                newModifiedCells.add(cellKey);
            }
        });
        this.modifiedCells = newModifiedCells;
        
        // 删除列宽度
        this.columnWidths.splice(colIndex, 1);
        
        // 清除选择
        this.selectedCells.clear();
        
        // 重新计算偏移量并渲染
        this.calculateOffsets();
        this.updateContainerSize();
        this.calculateVisibleRange();
        this.renderVisibleCells();
        this.updateHeadersPosition();
    }
}

// 初始化虚拟Excel模拟器
document.addEventListener('DOMContentLoaded', () => {
    const container = document.querySelector('.excel-container');
    window.virtualExcelSimulator = new VirtualExcelSimulator(container, {
        minRows: 100,
        minColumns: 26
    });
});
