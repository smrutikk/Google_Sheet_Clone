const spreadsheet = document.getElementById('spreadsheet');
let dragSource = null;
let selectedCells = [];
let isDraggingSelection = false;

// Utility to get cell ID
function getCellId(cell) {
  const col = cell.cellIndex - 1; // Column index (exclude row header)
  const row = cell.parentElement.rowIndex - 1; // Row index (exclude column header)
  return `${String.fromCharCode(65 + col)}${row + 1}`; // Convert to "A1", "B2", etc.
}

// Drag Start Event
spreadsheet.addEventListener('dragstart', (e) => {
  if (e.target.tagName === 'TD' && e.target.classList.contains('cell')) {
    dragSource = e.target;

    if (selectedCells.length > 0 && selectedCells.includes(dragSource)) {
      // Dragging a selection
      isDraggingSelection = true;
    } else {
      // Dragging a single cell
      isDraggingSelection = false;
    }

    e.dataTransfer.setData('text/plain', dragSource.innerText);
    e.target.classList.add('dragging');
  }
});

// Drag Over Event
spreadsheet.addEventListener('dragover', (e) => {
  e.preventDefault();
  if (e.target.tagName === 'TD' && e.target.classList.contains('cell')) {
    e.target.classList.add('drag-over');
  }
});

// Drag Leave Event
spreadsheet.addEventListener('dragleave', (e) => {
  if (e.target.tagName === 'TD' && e.target.classList.contains('cell')) {
    e.target.classList.remove('drag-over');
  }
});

// Drop Event
spreadsheet.addEventListener('drop', (e) => {
  e.preventDefault();

  if (e.target.tagName === 'TD' && e.target.classList.contains('cell')) {
    const dropTarget = e.target;
    const draggedData = e.dataTransfer.getData('text/plain');

    if (isDraggingSelection) {
      // Handle dragging of selected cells
      selectedCells.forEach((cell) => {
        const targetCol = dropTarget.cellIndex + (cell.cellIndex - dragSource.cellIndex);
        const targetRow = dropTarget.parentElement.rowIndex + (cell.parentElement.rowIndex - dragSource.parentElement.rowIndex);
        const targetCell = spreadsheet.rows[targetRow]?.cells[targetCol];

        if (targetCell) {
          targetCell.innerText = cell.innerText;
        }
      });

      // Clear source cells (optional)
      selectedCells.forEach((cell) => (cell.innerText = ''));
    } else {
      // Handle dragging of a single cell
      dropTarget.innerText = draggedData;
      dragSource.innerText = ''; // Optional: Clear source cell
    }

    dropTarget.classList.remove('drag-over');
  }
});

// Drag End Event
spreadsheet.addEventListener('dragend', (e) => {
  if (e.target.tagName === 'TD' && e.target.classList.contains('cell')) {
    e.target.classList.remove('dragging');
    document.querySelectorAll('.drag-over').forEach((cell) => cell.classList.remove('drag-over'));
  }
});

// Select Cells (Click + Drag)
spreadsheet.addEventListener('mousedown', (e) => {
  if (e.target.tagName === 'TD' && e.target.classList.contains('cell')) {
    selectedCells = [];
    spreadsheet.addEventListener('mousemove', selectCell);
  }
});

document.addEventListener('mouseup', () => {
  spreadsheet.removeEventListener('mousemove', selectCell);
});

function selectCell(e) {
  if (e.target.tagName === 'TD' && e.target.classList.contains('cell')) {
    if (!selectedCells.includes(e.target)) {
      selectedCells.push(e.target);
      e.target.classList.add('selected-cell');
    }
  }
}
