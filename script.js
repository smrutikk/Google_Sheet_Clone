class Spreadsheet {
    constructor() {
        this.data = {};
        this.selectedCell = null;
        this.columnCount = 26; // A to Z
        this.rowCount = 100;
        this.initializeSpreadsheet();
        this.setupEventListeners();
    }

    initializeSpreadsheet() {
        this.createColumnHeaders();
        this.createRowHeaders();
        this.createCells();
    }

    createColumnHeaders() {
        const columnHeaders = document.getElementById('columnHeaders');
        // Add corner cell
        const cornerCell = document.createElement('div');
        cornerCell.className = 'header-cell';
        columnHeaders.appendChild(cornerCell);

        // Add column headers (A-Z)
        for (let i = 0; i < this.columnCount; i++) {
            const header = document.createElement('div');
            header.className = 'header-cell';
            header.textContent = String.fromCharCode(65 + i);
            columnHeaders.appendChild(header);
        }
    }

    createRowHeaders() {
        const rowHeaders = document.getElementById('rowHeaders');
        for (let i = 1; i <= this.rowCount; i++) {
            const header = document.createElement('div');
            header.className = 'header-cell';
            header.textContent = i;
            rowHeaders.appendChild(header);
        }
    }

    createCells() {
        const cellsContainer = document.getElementById('cellsContainer');
        for (let row = 1; row <= this.rowCount; row++) {
            const rowDiv = document.createElement('div');
            rowDiv.style.display = 'flex';

            for (let col = 0; col < this.columnCount; col++) {
                const cell = document.createElement('div');
                cell.className = 'cell';
                cell.contentEditable = true;
                cell.dataset.row = row;
                cell.dataset.col = String.fromCharCode(65 + col);
                
                // Initialize cell data
                const cellId = `${String.fromCharCode(65 + col)}${row}`;
                this.data[cellId] = {
                    value: '',
                    formula: '',
                    formatting: {
                        bold: false,
                        italic: false,
                        color: '#000000',
                        fontSize: '12px',
                        fontFamily: 'Arial'
                    }
                };

                rowDiv.appendChild(cell);
            }
            cellsContainer.appendChild(rowDiv);
        }
    }

    setupEventListeners() {
        // Cell selection
        document.getElementById('cellsContainer').addEventListener('click', (e) => {
            if (e.target.classList.contains('cell')) {
                this.selectCell(e.target);
            }
        });

        // Formula input
        document.getElementById('formulaInput').addEventListener('input', (e) => {
            if (this.selectedCell) {
                const cellId = this.getCellId(this.selectedCell);
                this.data[cellId].formula = e.target.value;
                this.updateCell(this.selectedCell);
            }
        });

        // Formatting buttons
        document.getElementById('bold').addEventListener('click', () => this.toggleBold());
        document.getElementById('italic').addEventListener('click', () => this.toggleItalic());
        document.getElementById('textColor').addEventListener('input', (e) => this.setTextColor(e.target.value));

        // Function buttons
        document.getElementById('sumButton').addEventListener('click', () => this.insertFunction('SUM'));
        document.getElementById('averageButton').addEventListener('click', () => this.insertFunction('AVERAGE'));
        document.getElementById('maxButton').addEventListener('click', () => this.insertFunction('MAX'));
        document.getElementById('minButton').addEventListener('click', () => this.insertFunction('MIN'));
        document.getElementById('countButton').addEventListener('click', () => this.insertFunction('COUNT'));

        // Data quality buttons
        document.getElementById('trimButton').addEventListener('click', () => this.applyDataQuality('TRIM'));
        document.getElementById('upperButton').addEventListener('click', () => this.applyDataQuality('UPPER'));
        document.getElementById('lowerButton').addEventListener('click', () => this.applyDataQuality('LOWER'));
        document.getElementById('removeDuplicatesButton').addEventListener('click', () => this.removeDuplicates());
        document.getElementById('findReplaceButton').addEventListener('click', () => this.showFindReplaceDialog());
    }

    selectCell(cell) {
        if (this.selectedCell) {
            this.selectedCell.classList.remove('selected');
        }
        this.selectedCell = cell;
        cell.classList.add('selected');
        
        // Update formula bar
        const cellId = this.getCellId(cell);
        document.getElementById('formulaInput').value = this.data[cellId].formula;
        document.querySelector('.cell-reference').textContent = cellId;
    }

    getCellId(cell) {
        return `${cell.dataset.col}${cell.dataset.row}`;
    }

    updateCell(cell) {
        const cellId = this.getCellId(cell);
        const cellData = this.data[cellId];
        
        if (cellData.formula.startsWith('=')) {
            // Calculate formula result
            cell.textContent = this.evaluateFormula(cellData.formula.substring(1));
        } else {
            cell.textContent = cellData.formula;
        }

        // Apply formatting
        cell.style.fontWeight = cellData.formatting.bold ? 'bold' : 'normal';
        cell.style.fontStyle = cellData.formatting.italic ? 'italic' : 'normal';
        cell.style.color = cellData.formatting.color;
        cell.style.fontSize = cellData.formatting.fontSize;
        cell.style.fontFamily = cellData.formatting.fontFamily;
    }

    evaluateFormula(formula) {
        // Implement formula evaluation logic here
        // This is a simplified version
        try {
            if (formula.includes('SUM')) {
                return this.calculateSum(formula);
            } else if (formula.includes('AVERAGE')) {
                return this.calculateAverage(formula);
            } else if (formula.includes('MAX')) {
                return this.calculateMax(formula);
            } else if (formula.includes('MIN')) {
                return this.calculateMin(formula);
            } else if (formula.includes('COUNT')) {
                return this.calculateCount(formula);
            }
            return eval(formula);
        } catch (error) {
            return '#ERROR!';
        }
    }

    // Mathematical Functions
    calculateSum(range) {
        const values = this.getRangeValues(range);
        return values.reduce((sum, val) => sum + (parseFloat(val) || 0), 0);
    }

    calculateAverage(range) {
        const values = this.getRangeValues(range);
        const sum = values.reduce((sum, val) => sum + (parseFloat(val) || 0), 0);
        return sum / values.length;
    }

    calculateMax(range) {
        const values = this.getRangeValues(range);
        return Math.max(...values.map(val => parseFloat(val) || 0));
    }

    calculateMin(range) {
        const values = this.getRangeValues(range);
        return Math.min(...values.map(val => parseFloat(val) || 0));
    }

    calculateCount(range) {
        const values = this.getRangeValues(range);
        return values.filter(val => !isNaN(parseFloat(val))).length;
    }

    // Helper function to get values from a range (e.g., "A1:B5")
    getRangeValues(range) {
        const rangeMatch = range.match(/([A-Z]+\d+):([A-Z]+\d+)/);
        if (!rangeMatch) return [];

        const [start, end] = [rangeMatch[1], rangeMatch[2]];
        const startCol = start.match(/[A-Z]+/)[0];
        const startRow = parseInt(start.match(/\d+/)[0]);
        const endCol = end.match(/[A-Z]+/)[0];
        const endRow = parseInt(end.match(/\d+/)[0]);

        const values = [];
        for (let row = startRow; row <= endRow; row++) {
            for (let col = startCol.charCodeAt(0); col <= endCol.charCodeAt(0); col++) {
                const cellId = `${String.fromCharCode(col)}${row}`;
                const cellValue = this.data[cellId]?.value || '';
                values.push(cellValue);
            }
        }
        return values;
    }

    // Formatting Functions
    toggleBold() {
        if (!this.selectedCell) return;
        const cellId = this.getCellId(this.selectedCell);
        this.data[cellId].formatting.bold = !this.data[cellId].formatting.bold;
        this.updateCell(this.selectedCell);
    }

    toggleItalic() {
        if (!this.selectedCell) return;
        const cellId = this.getCellId(this.selectedCell);
        this.data[cellId].formatting.italic = !this.data[cellId].formatting.italic;
        this.updateCell(this.selectedCell);
    }

    setTextColor(color) {
        if (!this.selectedCell) return;
        const cellId = this.getCellId(this.selectedCell);
        this.data[cellId].formatting.color = color;
        this.updateCell(this.selectedCell);
    }

    // Data Quality Functions
    applyDataQuality(operation) {
        if (!this.selectedCell) return;
        const cellId = this.getCellId(this.selectedCell);
        const currentValue = this.data[cellId].value;

        switch (operation) {
            case 'TRIM':
                this.data[cellId].value = currentValue.trim();
                break;
            case 'UPPER':
                this.data[cellId].value = currentValue.toUpperCase();
                break;
            case 'LOWER':
                this.data[cellId].value = currentValue.toLowerCase();
                break;
        }
        this.updateCell(this.selectedCell);
    }

    removeDuplicates() {
        if (!this.selectedCell) return;
        const selection = this.getSelectedRange();
        const values = new Set();
        const duplicates = new Set();

        // First pass: identify duplicates
        selection.forEach(cell => {
            const value = cell.textContent;
            if (values.has(value)) {
                duplicates.add(value);
            }
            values.add(value);
        });

        // Second pass: remove duplicates
        selection.forEach(cell => {
            if (duplicates.has(cell.textContent)) {
                const cellId = this.getCellId(cell);
                this.data[cellId].value = '';
                this.data[cellId].formula = '';
                this.updateCell(cell);
            }
        });
    }

    showFindReplaceDialog() {
        const dialog = document.getElementById('findReplaceDialog');
        dialog.style.display = 'block';

        const findNext = document.getElementById('findNext');
        const replaceAll = document.getElementById('replaceAll');
        const closeDialog = document.getElementById('closeDialog');

        findNext.onclick = () => this.findNext();
        replaceAll.onclick = () => this.replaceAll();
        closeDialog.onclick = () => dialog.style.display = 'none';
    }

    findNext() {
        const findText = document.getElementById('findText').value;
        let found = false;
        
        const cells = document.querySelectorAll('.cell');
        for (let cell of cells) {
            if (cell.textContent.includes(findText)) {
                this.selectCell(cell);
                found = true;
                break;
            }
        }

        if (!found) {
            alert('Text not found!');
        }
    }

    replaceAll() {
        const findText = document.getElementById('findText').value;
        const replaceText = document.getElementById('replaceText').value;
        
        const cells = document.querySelectorAll('.cell');
        cells.forEach(cell => {
            if (cell.textContent.includes(findText)) {
                const cellId = this.getCellId(cell);
                this.data[cellId].value = cell.textContent.replace(new RegExp(findText, 'g'), replaceText);
                this.data[cellId].formula = this.data[cellId].value;
                this.updateCell(cell);
            }
        });
    }

    // Drag functionality
    enableDragSelection() {
        let isSelecting = false;
        let startCell = null;
        let currentSelection = new Set();

        document.getElementById('cellsContainer').addEventListener('mousedown', (e) => {
            if (e.target.classList.contains('cell')) {
                isSelecting = true;
                startCell = e.target;
                this.clearSelection();
                this.selectCell(startCell);
            }
        });

        document.getElementById('cellsContainer').addEventListener('mousemove', (e) => {
            if (!isSelecting) return;
            
            if (e.target.classList.contains('cell')) {
                this.updateSelection(startCell, e.target);
            }
        });

        document.addEventListener('mouseup', () => {
            isSelecting = false;
        });
    }

    updateSelection(startCell, endCell) {
        this.clearSelection();
        const selection = this.getCellsBetween(startCell, endCell);
        selection.forEach(cell => cell.classList.add('selected'));
    }

    getCellsBetween(startCell, endCell) {
        const startRow = parseInt(startCell.dataset.row);
        const endRow = parseInt(endCell.dataset.row);
        const startCol = startCell.dataset.col.charCodeAt(0);
        const endCol = endCell.dataset.col.charCodeAt(0);

        const minRow = Math.min(startRow, endRow);
        const maxRow = Math.max(startRow, endRow);
        const minCol = Math.min(startCol, endCol);
        const maxCol = Math.max(startCol, endCol);

        const cells = [];
        for (let row = minRow; row <= maxRow; row++) {
            for (let col = minCol; col <= maxCol; col++) {
                const cell = document.querySelector(
                    `[data-row="${row}"][data-col="${String.fromCharCode(col)}"]`
                );
                if (cell) cells.push(cell);
            }
        }
        return cells;
    }

    clearSelection() {
        document.querySelectorAll('.cell.selected').forEach(cell => {
            cell.classList.remove('selected');
        });
    }
}

// Initialize the spreadsheet
document.addEventListener('DOMContentLoaded', () => {
    const spreadsheet = new Spreadsheet();
    spreadsheet.enableDragSelection();
});