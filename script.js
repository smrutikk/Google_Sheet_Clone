// Active Cell Selection
const spreadsheet = document.getElementById('spreadsheet');
let activeCell = null;
let activeCellData = "";

spreadsheet.addEventListener('click', (e) => {
  if (e.target.tagName === 'TD') {
    // Remove the active class from all cells
    document.querySelectorAll('.cell').forEach(cell => cell.classList.remove('active-cell'));
    
    // Set the active class on the clicked cell
    e.target.classList.add('active-cell');
    activeCell = e.target;
    activeCellData = e.target.innerText;
    document.getElementById('formula-bar').value = activeCell.innerText;
  }
});

// Update active cell when formula bar is modified
document.getElementById('formula-bar').addEventListener('input', (e) => {
  if (activeCell) {
    activeCell.innerText = e.target.value;
    updateDependencies(activeCell);
  }
});

// Resize functionality for columns and rows (Adding/Deleting)
document.getElementById('add-row').addEventListener('click', () => {
  const table = document.getElementById('spreadsheet');
  const rowCount = table.rows.length;
  const newRow = table.insertRow(rowCount - 1);  // Insert before the last (empty row)
  
  // Add a row of cells
  for (let i = 0; i < 26; i++) {
    const newCell = newRow.insertCell(i);
    newCell.setAttribute("contenteditable", "true");
    newCell.classList.add("cell");
  }
});

document.getElementById('add-col').addEventListener('click', () => {
  const table = document.getElementById('spreadsheet');
  const rows = table.rows;
  
  // Add a new column to each row
  for (let i = 0; i < rows.length; i++) {
    const newCell = rows[i].insertCell(rows[i].cells.length - 1);
    newCell.setAttribute("contenteditable", "true");
    newCell.classList.add("cell");
  }

  // Add a new header for the new column
  const header = table.rows[0].insertCell(0);
  header.innerText = String.fromCharCode(65 + table.rows[0].cells.length - 2);  // New header like A, B, etc.
});

document.getElementById('delete-row').addEventListener('click', () => {
  const table = document.getElementById('spreadsheet');
  if (table.rows.length > 2) {  // Prevent deletion of the last row
    table.deleteRow(table.rows.length - 2);  // Delete the second last row (because of the header row)
  }
});

document.getElementById('delete-col').addEventListener('click', () => {
  const table = document.getElementById('spreadsheet');
  if (table.rows[0].cells.length > 1) {  // Prevent deletion of the last column
    for (let i = 0; i < table.rows.length; i++) {
      table.rows[i].deleteCell(table.rows[i].cells.length - 2);  // Delete the second last column
    }
  }
});

// Handling formulas (e.g., SUM, AVERAGE)
function updateDependencies(cell) {
  const formula = cell.innerText.trim();
  if (formula.startsWith('=')) {
    const result = calculateFormula(formula);
    cell.innerText = result;
  }
}

function calculateFormula(formula) {
  const match = formula.match(/=(SUM|AVERAGE|MAX|MIN|COUNT)\(([^)]+)\)/i);
  if (!match) return formula; // Return as is if not a valid formula

  const [, func, range] = match;
  const values = getRangeValues(range);

  switch (func.toUpperCase()) {
    case 'SUM':
      return values.reduce((a, b) => a + b, 0);
    case 'AVERAGE':
      return values.length ? values.reduce((a, b) => a + b, 0) / values.length : 0;
    case 'MAX':
      return Math.max(...values);
    case 'MIN':
      return Math.min(...values);
    case 'COUNT':
      return values.filter(v => !isNaN(v)).length;
    default:
      return formula; // Fallback
  }
}

function getRangeValues(range) {
  const [start, end] = range.split(':');
  const { row: startRow, col: startCol } = parseCellId(start);
  const { row: endRow, col: endCol } = parseCellId(end);
  const values = [];

  for (let i = startRow; i <= endRow; i++) {
    for (let j = startCol; j <= endCol; j++) {
      values.push(getCellValue(i, j));
    }
  }
  return values;
}

function parseCellId(cellId) {
  const col = cellId.charCodeAt(0) - 65; // 'A' -> 0, 'B' -> 1, etc.
  const row = parseInt(cellId.slice(1)) - 1; // '1' -> 0, '2' -> 1, etc.
  return { row, col };
}

function getCellValue(row, col) {
  const table = document.getElementById('spreadsheet');
  const cell = table.rows[row + 1]?.cells[col + 1]; // Offset for headers
  return cell ? parseFloat(cell.textContent) || 0 : 0;
}
