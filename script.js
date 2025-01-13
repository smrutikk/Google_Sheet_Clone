/* Global Variables */
let wb = [];
let db = [];
let lsc = null;
let currentCell = document.querySelector('.cell');
let sheetsDiv = document.querySelector('.sheets');

/* Functions */
function initUI() {
    let allrows = document.querySelectorAll('.rows');
    for (let row of allrows) {
        let cells = row.querySelectorAll('.cell');
        for (let cell of cells) {
            cell.innerHTML = "";
            cell.style.fontFamily = "arial";
            cell.style.fontWeight = "normal";
            cell.style.fontStyle = "normal";
            cell.style.textDecoration = "none";
            cell.style.justifyContent = "center";
            cell.style.fontSize = "16px";
            cell.style.color = "black";
            cell.style.backgroundColor = "transparent";
            cell.style.alignItems = "flex-start";
        }
    }
}

function createNewSheet() {
    let sheetIndex = wb.length;
    let sheetDB = [];
    for (let i = 0; i < 100; i++) {
        let rows = [];
        for (let j = 0; j < 26; j++) {
            rows.push({
                value: "",
                fontName: "arial",
                fontSize: "16px",
                textColor: "black",
                background: "white",
                cellAlignment: "center",
                vAlignment: "flex-start"
            });
        }
        sheetDB.push(rows);
    }
    wb.push(sheetDB);
}

function loadSheet(sheetIndex) {
    db = wb[sheetIndex];
    initUI();
}

function handleCellClick(event) {
    lsc = event.target;
    lsc.style.backgroundColor = '#e6e6e6';
}

/* Button Event Listeners */
document.querySelector('.new').addEventListener('click', function () {
    wb = [];
    createNewSheet();
    loadSheet(0);
});

document.querySelector('.save').addEventListener('click', function () {
    let data = JSON.stringify(wb);
    let blob = new Blob([data], { type: 'application/json' });
    let link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'spreadsheet.json';
    link.click();
});

document.querySelector('.open').addEventListener('click', function () {
    let input = document.createElement('input');
    input.type = 'file';
    input.accept = '.json';
    input.addEventListener('change', function (e) {
        let file = e.target.files[0];
        let reader = new FileReader();
        reader.onload = function () {
            wb = JSON.parse(reader.result);
            loadSheet(0);
        };
        reader.readAsText(file);
    });
    input.click();
});

document.querySelector('.addSheet').addEventListener('click', function () {
    createNewSheet();
    let sheetIndex = wb.length - 1;
    let sheetDiv = document.createElement('div');
    sheetDiv.classList.add('sheet');
    sheetDiv.textContent = 'Sheet ' + (sheetIndex + 1);
    sheetDiv.addEventListener('click', function () {
        loadSheet(sheetIndex);
    });
    sheetsDiv.appendChild(sheetDiv);
});

document.querySelector('#font').addEventListener('change', function () {
    if (lsc) {
        lsc.style.fontFamily = this.value;
    }
});

document.querySelector('#bold').addEventListener('click', function () {
    if (lsc) {
        lsc.style.fontWeight = lsc.style.fontWeight === 'bold' ? 'normal' : 'bold';
    }
});

document.querySelector('#italic').addEventListener('click', function () {
    if (lsc) {
        lsc.style.fontStyle = lsc.style.fontStyle === 'italic' ? 'normal' : 'italic';
    }
});

document.querySelector('#underline').addEventListener('click', function () {
    if (lsc) {
        lsc.style.textDecoration = lsc.style.textDecoration === 'underline' ? 'none' : 'underline';
    }
});

document.querySelector('#fontsize').addEventListener('change', function () {
    if (lsc) {
        lsc.style.fontSize = this.value + 'px';
    }
});

document.querySelector('#fontcolor').addEventListener('input', function () {
    if (lsc) {
        lsc.style.color = this.value;
    }
});

document.querySelector('#cellcolor').addEventListener('input', function () {
    if (lsc) {
        lsc.style.backgroundColor = this.value;
    }
});

/* Dynamic Cell Creation */
function createCells() {
    let rowsContainer = document.querySelector('.cells');
    for (let i = 0; i < 100; i++) {
        let row = document.createElement('div');
        row.classList.add('rows');
        for (let j = 0; j < 26; j++) {
            let cell = document.createElement('div');
            cell.classList.add('cell');
            cell.setAttribute('r-id', i);
            cell.setAttribute('c-id', j);
            cell.addEventListener('click', handleCellClick);
            row.appendChild(cell);
        }
        rowsContainer.appendChild(row);
    }
}

createCells(); // Initialize the cells

  // Cell click event
cell.addEventListener("click", function (e) {
    let rid = parseInt(cell.getAttribute("r-Id"));
    let cid = parseInt(cell.getAttribute("c-Id"));
    let address = String.fromCharCode(65 + cid) + (rid + 1);
    addressBar.value = address;

    let obj = getCellobj(cell);
    formulaBar.value = obj.formula;
    let { vAlignment, cellFormatting, cellAlignment } = db[rid][cid];
    
    // Apply formatting
    toggleButtonState(boldBtn, cellFormatting.bold);
    toggleButtonState(italicBtn, cellFormatting.italic);
    toggleButtonState(underlineBtn, cellFormatting.underline);
    setAlignment(alignBtn, cellAlignment);
    setVerticalAlignment(vAlign, vAlignment);
});

// Helper function to toggle button active state
function toggleButtonState(button, isActive) {
    if (isActive) {
        button.classList.add("activeBtn");
    } else {
        button.classList.remove("activeBtn");
    }
}

// Helper function to set alignment state
function setAlignment(alignButtons, alignment) {
    alignButtons.forEach((button, index) => {
        if (button.classList.contains(alignment)) {
            button.classList.add("activeBtn");
        } else {
            button.classList.remove("activeBtn");
        }
    });
}

// Helper function to set vertical alignment state
function setVerticalAlignment(alignButtons, vAlignment) {
    alignButtons.forEach((button, index) => {
        if (button.classList.contains(vAlignment)) {
            button.classList.add("activeBtn");
        } else {
            button.classList.remove("activeBtn");
        }
    });
}

// Cell blur event
cell.addEventListener("blur", function () {
    lsc = cell;
    let newVal = cell.innerHTML;
    let obj = getCellobj(cell);
    
    let NN_num = Number(newVal);
    if (newVal === "") {
        obj.value = "";
        obj.parent = [];
        obj.formula = "";
    } else if (!NN_num) {
        obj.value = newVal;
        if (obj.child.length > 0) {
            obj.child.forEach(child => {
                let { rowId, colId } = get_rId_cId_fromAddress(child);
                let childobj = db[rowId][colId];
                showErr(childobj);
            });
        }
    } else if (newVal !== obj.value) {
        obj.value = newVal;
        if (obj.formula) {
            removeFormula(obj);
            formulaBar.value = "";
        }
        updateChildren(obj);
    }
});

// Set top-most & left-most scroll position
contentDiv.addEventListener("scroll", function () {
    let topOffset = contentDiv.scrollTop;
    let leftOffset = contentDiv.scrollLeft;

    topMost.style.top = topOffset + "px";
    topLeft.style.top = topOffset + "px";
    leftMost.style.left = leftOffset + "px";
    topLeft.style.left = leftOffset + "px";
});

// Change height of cells
topMostAll.forEach(topCell => {
    topCell.addEventListener("mousedown", function (e) {
        if ((topCell.getBoundingClientRect().width - e.offsetX) < 5) {
            let rid = topCell.getAttribute("c-Id");
            changeWidth = rid;
            changeCell = topCell;
            changeDirection = 'H';
            contentDiv.style.cursor = "col-resize";
        }
    });

    window.addEventListener("mousemove", function (e) {
        if (changeWidth && changeCell !== "" && changeDirection === 'H') {
            let finalX = e.clientX;
            let dx = finalX - changeCell.getBoundingClientRect().left;

            let allCells = document.querySelectorAll(`.cell[c-Id="${changeWidth}"]`);
            let topcell = topMostAll[changeWidth];

            topcell.style.minWidth = dx + "px";
            topcell.style.maxWidth = dx + "px";

            allCells.forEach(cell => {
                cell.style.minWidth = dx + "px";
                cell.style.maxWidth = dx + "px";
            });

            for (let i = 0; i < 100; i++) {
                db[i][changeWidth].width = dx;
            }
        }
    });

    window.addEventListener("mouseup", function () {
        changeWidth = "";
        changeCell = "";
        changeDirection = "";
        contentDiv.style.cursor = "auto";
    });
});

leftMostAll.forEach(leftCell => {
    leftCell.addEventListener("mousedown", function (e) {
        if ((leftCell.getBoundingClientRect().height - e.offsetY) < 5) {
            let cid = leftCell.getAttribute("r-Id");
            changeHeight = cid;
            changeCell = leftCell;
            changeDirection = 'V';
            contentDiv.style.cursor = "row-resize";
        }
    });

    window.addEventListener("mousemove", function (e) {
        if (changeHeight && changeCell !== "" && changeDirection === 'V') {
            let finalY = e.clientY;
            let dy = finalY - changeCell.getBoundingClientRect().top;

            let allCells = document.querySelectorAll(`.cell[r-Id="${changeHeight}"]`);
            let leftcell = leftMostAll[changeHeight];

            leftcell.style.minHeight = dy + "px";
            leftcell.style.minHeight = dy + "px";

            allCells.forEach(cell => {
                cell.style.minHeight = dy + "px";
                cell.style.minHeight = dy + "px";
            });
            
            for (let i = 0; i < 26; i++) {
                db[changeHeight][i].height = dy;
            }
        }
    });

    window.addEventListener("mouseup", function () {
        changeHeight = "";
        changeCell = "";
        changeDirection = "";
        contentDiv.style.cursor = "auto";
    });
});

// Formula blur event
formulaBar.addEventListener("blur", function () {
    let formula = formulaBar.value;
    let obj = getCellobj(lsc);
    if (obj.formula.length > 0 && formula.length === 0) {
        lsc.innerHTML = "";
        obj.formula = "";
        obj.value = "";
        removeFormula(obj);
        obj.child = [];
    } else if (obj.formula !== formula) {
        removeFormula(obj);
        if (addFormula(formula)) {
            updateChildren(obj);
        }
    }
});

// Add formula to cell
function addFormula(formula) {
    let obj = getCellobj(lsc);
    let formulaArray = formula.split(" ");
    let children = obj.child;

    for (let i = 0; i < children.length; i++) {
        for (let j = 0; j < formulaArray.length; j++) {
            if (children[i] === formulaArray[j] || formulaArray[j] === obj.address) {
                let { rowId, colId } = get_rId_cId_fromAddress(obj.address);
                let currCell = document.querySelector(`.cell[r-Id= "${rowId}"][c-Id="${colId}"]`);
                currCell.innerHTML = '';
                currCell.value = "";
                promptDiv.style.display = "block";
                showErr(obj);
                return false;
            }
        }
    }

    for (let j = 0; j < formulaArray.length; j++) {
        if (formulaArray[j] === obj.address) {
            let { rowId, colId } = get_rId_cId_fromAddress(obj.address);
            let currCell = document.querySelector(`.cell[r-Id= "${rowId}"][c-Id="${colId}"]`);
            currCell.innerHTML = '';
            currCell.value = "";
            promptDiv.style.display = "block";
            showErr(obj);
            return false;
        }
    }

    obj.formula = formula;
    solveFormula(obj);
    return true;
}

// Solve the formula
function solveFormula(obj) {
    let formula = obj.formula;
    let fcomps = formula.split(" ");

    for (let i = 0; i < fcomps.length; i++) {
        let fcomp = fcomps[i];
        let initial = fcomp[0];

        if (initial >= "A" && initial <= "Z") {
            let { rowId, colId } = get_rId_cId_fromAddress(fcomp);
            let parentobj = db[rowId][colId];
            obj.parent.push(fcomp);
            parentobj.child.push(obj.address);
            let value = Number(parentobj.value);
            formula = formula.replace(fcomp, value);
        }
    }

    let value = evaluateFormula(formula);
    obj.value = value;
    let { rowId, colId } = get_rId_cId_fromAddress(obj.address);
    let currCell = document.querySelector(`.cell[r-Id= "${rowId}"][c-Id="${colId}"]`);
    currCell.innerHTML = value;
}

// Update children of a cell
function updateChildren(obj) {
    let children = obj.child;
    for (let i = 0; i < children.length; i++) {
        let child = children[i];
        let { rowId, colId } = get_rId_cId_fromAddress(child);
        let childobj = db[rowId][colId];
        recalculate(childobj);
        updateChildren(childobj);
    }
}

// Recalculate a cell's value
function recalculate(obj) {
    let formula = obj.formula;
    let fcomps = formula.split(" ");

    for (let i = 0; i < fcomps.length; i++) {
        let fcomp = fcomps[i];
        let initial = fcomp[0];

        if (initial >= "A" && initial <= "Z") {
            let { rowId, colId } = get_rId_cId_fromAddress(fcomp);
            let parentobj = db[rowId][colId];
            let value = Number(parentobj.value);
            formula = formula.replace(fcomp, value);
        }
    }

    let value = evaluateFormula(formula);
    obj.value = value;
    let { rowId, colId } = get_rId_cId_fromAddress(obj.address);
    let currCell = document.querySelector(`.cell[r-Id= "${rowId}"][c-Id="${colId}"]`);
    currCell.innerHTML = value;
}

// Remove formula from a cell
function removeFormula(obj) {
    obj.formula = "";

    for (let i = 0; i < obj.parent.length; i++) {
        let { rowId, colId } = get_rId_cId_fromAddress(obj.parent[i]);
        let parentobj = db[rowId][colId];
        let newParentChild = parentobj.child.filter(function (child) {
            return child !== obj.address;
        });

        parentobj.child = newParentChild;
    }
    obj.parent = [];
}

// Show error in cell
function showErr(obj) {
    let presentobj = get_rId_cId_fromAddress(obj.address);
    db[presentobj.rowId][presentobj.colId].value = "err";
    let currCell = document.querySelector(`.cell[r-Id="${presentobj.rowId}"][c-Id="${presentobj.colId}"]`);
    currCell.innerHTML = 'err';
    obj.child.forEach(child => {
        let childobj = db[child.rowId][child.colId];
        showErr(childobj);
    });
}

// Get cell object
function getCellobj(ele) {
    let rid = Number(ele.getAttribute("r-Id"));
    let cid = Number(ele.getAttribute("c-Id"));
    return db[rid][cid];
}

// Get rowId and colId from address
function get_rId_cId_fromAddress(address) {
    let rId = Number(address.substring(1)) - 1;
    let cId = Number(address.charCodeAt(0) - 65);
    return { rowId: rId, colId: cId };
}

// Get address from rowId and colId
function get_Address_from_rId_cId(rId, cId) {
    let rowId = rId + 1;
    let colId = String.fromCharCode(65 + cId);
    return `${colId}${rowId}`;
}

// Initialize
function init() {
    let newButton = document.querySelector(".new");
    newButton.click();
}
init();

// Set width of the column
function setWidth(elem, width) {
    let cid = elem.getAttribute("c-Id");
    let leftcol = document.querySelectorAll(".topMostCell")[cid];
    leftcol.style.minWidth = width + "px";
    leftcol.style.maxWidth = width + "px";
    let cells = document.querySelectorAll(`.cell[c-Id="${cid}"]`);
    cells.forEach(cell => {
        cell.style.minWidth = width + "px";
        cell.style.maxWidth = width + "px";
    });
    for (let i = 0; i < 100; i++) {
        db[i][cid].width = width;
    }
}

// Set height of the row
function setHeight(elem, height) {
    let rid = elem.getAttribute("r-Id");
    let leftcell = document.querySelectorAll(".leftMostCell")[rid];
    leftcell.style.minHeight = height + "px";
    leftcell.style.maxHeight = height + "px";
    let cells = document.querySelectorAll(`.cell[r-Id="${rid}"]`);
    cells.forEach(cell => {
        cell.style.minHeight = height + "px";
        cell.style.maxHeight = height + "px";
    });
    for (let i = 0; i < 26; i++) {
        db[rid][i].height = height;
    }
}

//  ======================== remove prompt ===================================================

okBtn.addEventListener("click",function(){
    promptDiv.style.display = "none";
})