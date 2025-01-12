class GoogleSheetsClone {
  constructor() {
      this.rows = 100;
      this.cols = 26;
      this.selectedCell = null;
      this.data = {};
      this.init();
  }

  init() {
      this.createColumnHeaders();
      this.createRowHeaders();
      this.createGrid();
      this.setupEventListeners();
  }

  createColumnHeaders() {
      const headers = document.getElementById('columnHeaders');
      for (let i = 0; i < this.cols; i++) {
          const header = document.createElement('div');
          header.className = 'col-header';
          header.textContent = this.columnIndexToLetter(i);
          
          const resizeHandle = document.createElement('div');
          resizeHandle.className = 'resize-handle';
          header.appendChild(resizeHandle);
          
          headers.appendChild(header);
      }
  }

  createRowHeaders() {
      const headers = document.getElementById('rowHeaders');
      for (let i = 0; i < this.rows; i++) {
          const header = document.createElement('div');
          header.className = 'row-header';
          header.textContent = i + 1;
          headers.appendChild(header);
      }
  }

  createGrid() {
      const grid = document.getElementById('grid');
      for (let i = 0; i < this.rows; i++) {
          const row = document.createElement('div');
          row.className = 'row';
          
          for (let j = 0; j < this.cols; j++) {
              const cell = document.createElement('div');
              cell.className = 'cell';
              cell.contentEditable = true;
              cell.dataset.row = i;
              cell.dataset.col = j;
              row.appendChild(cell);
          }
          
          grid.appendChild(row);
      }
  }

  setupEventListeners() {
      document.getElementById('grid').addEventListener('click', (e) => {
          const cell = e.target.closest('.cell');
          if (cell) {
              this.selectCell(cell);
          }
      });

      document.getElementById('grid').addEventListener('input', (e) => {
          const cell = e.target.closest('.cell');
          if (cell) {
              this.updateCellData(cell);
          }
      });

      // Column resize handling
      document.addEventListener('mousedown', (e) => {
          if (e.target.className === 'resize-handle') {
              this.startResize(e);
          }
      });

      document.addEventListener('mousemove', (e) => {
          if (this.resizing) {
              this.handleResize(e);
          }
      });

      document.addEventListener('mouseup', () => {
          this.resizing = false;
      });
  }

  columnIndexToLetter(index) {
      let letter = '';
      index++;
      while (index > 0) {
          let remainder = (index - 1) % 26;
          letter = String.fromCharCode(65 + remainder) + letter;
          index = Math.floor((index - 1) / 26);
      }
      return letter;
  }

  selectCell(cell) {
      if (this.selectedCell) {
          this.selectedCell.classList.remove('selected');
      }
      this.selectedCell = cell;
      cell.classList.add('selected');
      
      // Update formula bar
      const address = this.getCellAddress(cell);
      document.querySelector('.cell-address').textContent = address;
      document.querySelector('.formula-input').value = cell.textContent;
  }

  getCellAddress(cell) {
      const col = this.columnIndexToLetter(parseInt(cell.dataset.col));
      const row = parseInt(cell.dataset.row) + 1;
      return `${col}${row}`;
  }

  updateCellData(cell) {
      const value = cell.textContent;
      const address = this.getCellAddress(cell);
      this.data[address] = value;

      if (value.startsWith('=')) {
          this.evaluateFormula(cell, value);
      }
      document.querySelector('.formula-input').value = value;
  }

  startResize(e) {
      this.resizing = true;
      this.resizeStartX = e.pageX;
      const headerCell = e.target.closest('.col-header');
      this.resizeColumnIndex = Array.from(headerCell.parentNode.children).indexOf(headerCell);
      this.initialWidth = headerCell.offsetWidth;
  }

  handleResize(e) {
      if (!this.resizing) return;
      
      const diff = e.pageX - this.resizeStartX;
      const newWidth = Math.max(50, this.initialWidth + diff);
      
      // Update column header width
      const headers = document.querySelectorAll('.col-header');
      headers[this.resizeColumnIndex].style.width = `${newWidth}px`;
      
      // Update all cells in the column
      const cells = document.querySelectorAll(`.cell:nth-child(${this.resizeColumnIndex + 1})`);
      cells.forEach(cell => {
          cell.style.width = `${newWidth}px`;
      });
  }
}

// Initialize the spreadsheet
new GoogleSheetsClone();
