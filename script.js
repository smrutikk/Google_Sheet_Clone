/*===================================IMPORT FILES===========================================*/
let $ = require("jquery");
let fs = require("fs");
let dialog = require("electron").remote.dialog;

let newBtn = document.querySelector(".new");
let saveBtn = document.querySelector(".save");
let openBtn = document.querySelector(".open");

let fileBtn = document.querySelector(".file");
let homeBtn = document.querySelector(".home");

let fontBtn = document.querySelector("#font");
let boldBtn = document.querySelector("#bold");
let italicBtn = document.querySelector("#italic");
let cellColor = document.querySelector("#cellcolor");
let fontColor = document.querySelector("#fontcolor");
let fontSizeBtn = document.querySelector("#fontsize");
let underlineBtn = document.querySelector("#underline");
let alignBtn = document.querySelectorAll(".alignment div");
let vAlign = document.querySelectorAll(".vAlignment div")

let homeMenuOptions = document.querySelector(".homeOptions")
let fileMenuOptions = document.querySelector(".fileOptions")

let contentDiv = document.querySelector(".content");
let topMost = document.querySelector(".topMost");
let topMostAll = document.querySelectorAll(".topMost div")
let leftMost = document.querySelector(".leftMost");
let leftMostAll = document.querySelectorAll(".leftMost div");
let topLeft = document.querySelector(".topLeft");

let addressBar = document.querySelector(".addressInput");
let formulaBar = document.querySelector(".FormulaInput");

let cells = document.querySelectorAll(".cell")

let addSheet = document.querySelector(".addSheet");
let sheetsDiv = document.querySelector(".sheets");
let allSheet = document.querySelectorAll(".sheet");
allSheet[0].addEventListener("click", handleActiveSheet);

let firstCell = document.querySelector('.cell[r-id="0"][c-id="0"]');

let promptDiv = document.querySelector(".prompt");
let okBtn = document.querySelector(".okBtn");


/*===================================GLOBAL VARIABLES===========================================*/
let wb = [];
let db = [];
let lsc = firstCell;
let changeWidth = "";
let changeCell = ""
let changeDirection = "";
let changeHeight = "";

/*===================================Menu Btns===========================================*/

fileBtn.addEventListener("click", function () {
    homeMenuOptions.classList.remove("active");
    fileMenuOptions.classList.add("active");

});

homeBtn.addEventListener("click", function () {
    homeMenuOptions.classList.add("active");
    fileMenuOptions.classList.remove("active");

})

/*===================================FILE MENU OPTIONS====================================*/
newBtn.addEventListener("click", function () {
    wb = [];
    db = [];
    let allrows = document.querySelectorAll(".cells .rows");
    for (let i = 0; i < allrows.length; i++) {
        let rows = [];
        let allColsinRow = allrows[i].querySelectorAll(".cell");
        for (let j = 0; j < 26; j++) {
            let address = get_Address_from_rId_cId(i, j);
            let obj = {
                address: address,
                value: "",
                formula: "",
                child: [],
                parent: [],
                fontName: "arial",
                cellFormatting: { bold: false, underline: false, italic: false },
                cellAlignment: "center",
                fontSize: "16px",
                textColor: "black",
                background: "white",
                vAlignment: "flex-start",
                height: 25,
                width: 70
            };
            rows.push(obj);
        }
        db.push(rows);
    }
    let sheets = document.querySelectorAll(".sheet");
    for (let i = 0; i < sheets.length; i++)sheets[i].remove();
    let span = document.createElement("span");
    span.classList.add("sheet");
    span.classList.add("active-sheet");
    span.setAttribute("sheetIdx", 0);
    span.innerHTML = `Sheet ${1}`
    sheetsDiv.appendChild(span);
    span.addEventListener("click", handleActiveSheet);

    initUI();
    wb.push(db);

    document.querySelector(".addFormula, .address").value = "";
    firstCell.click();

});

saveBtn.addEventListener("click", function () {
    let path = dialog.showSaveDialogSync();
    if (!path) return;
    let data = JSON.stringify(wb);
    fs.writeFileSync(path, data);
    // alert("file saved");
})

openBtn.addEventListener("click", function () {
    let path = dialog.showOpenDialogSync();
    if (!path) return;
    path = path[0];

    let data = fs.readFileSync(path);
    data = JSON.parse(data);
    wb = data;
    db = wb[0];

    let sheets = document.querySelectorAll(".sheet");
    for (let i = 0; i < sheets.length; i++)sheets[i].remove();
    for (let i = 0; i < wb.length; i++) {
        let span = document.createElement("span");
        span.classList.add("sheet");
        if (i == 0) span.classList.add("active-sheet");
        span.setAttribute("sheetIdx", i);
        span.innerHTML = `Sheet ${i + 1}`
        sheetsDiv.appendChild(span);
        span.addEventListener("click", handleActiveSheet);
    }

    for (let i = 0; i < db[0].length; i++) {
        let cell = document.querySelector(`.cell[r-id="0"][c-id="${i}"]`);
        setWidth(cell, db[0][i].width)
    }
    for (let i = 0; i < db.length; i++) {
        let cell = document.querySelector(`.cell[r-id="${i}"][c-id="0"]`);
        setHeightByDB(cell, db[i][0].height)
    }

    let allrows = document.querySelectorAll(".cells .rows");
    for (let i = 0; i < allrows.length; i++) {
        let allCols = allrows[i].querySelectorAll(".cell")
        for (let j = 0; j < allCols.length; j++) {
            let cell = allCols[j];
            let { value, width, vAlignment, fontName, cellFormatting, cellAlignment, fontSize, textColor, background } = db[i][j];
            cell.innerHTML = value;
            cell.style.fontFamily = fontName;
            cell.style.fontWeight = cellFormatting.bold ? "bold" : "normal";
            cell.style.fontStyle = cellFormatting.italic ? "italic" : "normal";
            cell.style.textDecoration = cellFormatting.underline ? "underline" : "none";
            cell.style.justifyContent = cellAlignment;
            cell.style.fontSize = `${fontSize}px`;
            cell.style.color = textColor;
            cell.style.backgroundColor = background;
            cell.style.alignItems = vAlignment;

        }
    }
});

function initUI() {
    for (let i = 0; i < db[0].length; i++) {
        let cell = document.querySelector(`.cell[r-id="${0}"][c-id="${i}"]`);
        setWidth(cell, 70)
    }
    for (let i = 0; i < db.length; i++) {
        let cell = document.querySelector(`.cell[r-id="${i}"][c-id="0"]`);
        setHeightByDB(cell, 25)
    }
    let allrows = document.querySelectorAll(".cells .rows");
    for (let i = 0; i < allrows.length; i++) {
        let allColsinRow = allrows[i].querySelectorAll(".cell");
        for (let j = 0; j < 26; j++) {
            allColsinRow[j].innerHTML = "";
            allColsinRow[j].style.fontFamily = "arial";
            allColsinRow[j].style.fontWeight = "normal";
            allColsinRow[j].style.fontStyle = "normal";
            allColsinRow[j].style.textDecoration = "none";
            allColsinRow[j].style.justifyContent = "center";
            allColsinRow[j].style.fontSize = "16px";
            allColsinRow[j].style.color = "black";
            allColsinRow[j].style.backgroundColor = "transparent";
            allColsinRow[j].style.alignItems = "flex-start";

        }
    }
}

/*===================================Add Sheet ===========================================*/

addSheet.addEventListener("click", function () {
    let sheets = document.querySelectorAll(".sheet");
    let span = document.createElement("span");
    span.classList.add("sheet");
    span.setAttribute("sheetIdx", sheets.length);
    span.innerHTML = `Sheet ${sheets.length + 1}`
    sheetsDiv.appendChild(span);
    span.addEventListener("click", handleActiveSheet);
    addDBtoWB();
    initUI();
    firstCell.click();
    db = wb[wb.length - 1];
    let sheetsArr = document.querySelectorAll(".sheet");
    sheetsArr.forEach(function (sheet) {
        sheet.classList.remove("active-sheet");
    })
    span.classList.add("active-sheet");
    
    
})

function addDBtoWB() {
    let sheetDB = [];
    for (let i = 0; i < 100; i++) {
        let rows = [];
        for (let j = 0; j < 26; j++) {
            let address = get_Address_from_rId_cId(i, j);
            let obj = {
                address: address,
                value: "",
                formula: "",
                child: [],
                parent: [],
                fontName: "arial",
                cellFormatting: { bold: false, underline: false, italic: false },
                cellAlignment: "center",
                fontSize: "16px",
                textColor: "black",
                background: "white",
                vAlignment: "fle-start",
                width: 70,
                height: 25
            };
            rows.push(obj);
        }
        sheetDB.push(rows);
    }
    wb.push(sheetDB);
}

function handleActiveSheet(e) {
    let MySheet = e.currentTarget;
    let sheetsArr = document.querySelectorAll(".sheet");
    sheetsArr.forEach(function (sheet) {
        sheet.classList.remove("active-sheet");
    })
    if (!MySheet.classList[1]) {
        MySheet.classList.add("active-sheet");
    }
    //  index
    let sheetIdx = MySheet.getAttribute("sheetIdx");
    db = wb[sheetIdx];
    // get data from that and set ui
    setUI();
    firstCell.click();
}

function setUI() {
    for (let i = 0; i < db[0].length; i++) {
        let cell = document.querySelector(`.cell[r-id="${0}"][c-id="${i}"]`);
        setWidth(cell, db[0][i].width)
    }
    for (let i = 0; i < db.length; i++) {
        let cell = document.querySelector(`.cell[r-id="${i}"][c-id="0"]`);
        setHeightByDB(cell, db[i][0].height)
    }
    for (let i = 0; i < db.length; i++) {
        let height = -1;
        for (let j = 0; j < db[i].length; j++) {
            let cell = document.querySelector(`.cell[r-id="${i}"][c-id="${j}"]`);
            let { value, fontName, cellFormatting, cellAlignment, fontSize, textColor, background, vAlignment, width } = db[i][j];
            cell.innerHTML = value;
            cell.style.fontFamily = fontName;
            cell.style.fontWeight = cellFormatting.bold ? "bold" : "normal";
            cell.style.fontStyle = cellFormatting.italic ? "italic" : "normal";
            cell.style.textDecoration = cellFormatting.underline ? "underline" : "none";
            cell.style.justifyContent = cellAlignment;
            cell.style.fontSize = `${fontSize}px`;
            cell.style.color = textColor;
            cell.style.backgroundColor = background;
            cell.style.alignItems = vAlignment;

        }
    }
}
/*===================================HOME MENU OPTIONS===========================================*/

boldBtn.addEventListener("click", function () {
    let cellObj = getCellobj(lsc);

    lsc.style.fontWeight = cellObj.cellFormatting.bold ? "normal" : "bold";
    cellObj.cellFormatting.bold = !cellObj.cellFormatting.bold;
    if (cellObj.cellFormatting.bold) {
        boldBtn.classList.add("activeBtn");
    } else {
        boldBtn.classList.remove("activeBtn");
    }

    setHeight(lsc);
})

italicBtn.addEventListener("click", function () {
    let cellObj = getCellobj(lsc);
    lsc.style.fontStyle = cellObj.cellFormatting.italic ? "normal" : "italic";
    cellObj.cellFormatting.italic = !cellObj.cellFormatting.italic;

    if (cellObj.cellFormatting.italic) {
        italicBtn.classList.add("activeBtn");
    } else {
        italicBtn.classList.remove("activeBtn");
    }
    setHeight(lsc);
})

underlineBtn.addEventListener("click", function () {
    let cellObj = getCellobj(lsc);
    lsc.style.textDecoration = cellObj.cellFormatting.underline ? "none" : "underline";
    cellObj.cellFormatting.underline = !cellObj.cellFormatting.underline;
    setHeight(lsc);

    if (cellObj.cellFormatting.underline) {
        underlineBtn.classList.add("activeBtn");
    } else {
        underlineBtn.classList.remove("activeBtn");
    }
})

fontColor.addEventListener("change", function (e) {
    let cellObj = getCellobj(lsc);
    let value = fontColor.value;
    lsc.style.color = value + "";
    cellObj.textColor = `${value}`;

})

cellColor.addEventListener("change", function () {
    let cellObj = getCellobj(lsc);
    let value = cellColor.value;
    lsc.style.backgroundColor = value + "";
    cellObj.background = `${value}`;

})
alignBtn.forEach(btn => {
    btn.addEventListener("click", function (e) {
        let cellObj = getCellobj(lsc);
        let value = e.currentTarget.classList[1];
        lsc.style.justifyContent = value;
        cellObj.cellAlignment = `${value}`;
        if (value == 'flex-start') {
            alignBtn[0].classList.add("activeBtn");
            alignBtn[1].classList.remove("activeBtn");
            alignBtn[2].classList.remove("activeBtn");
        } else if (value == 'center') {
            alignBtn[0].classList.remove("activeBtn");
            alignBtn[1].classList.add("activeBtn");
            alignBtn[2].classList.remove("activeBtn");

        } else if (value == 'flex-end') {
            alignBtn[0].classList.remove("activeBtn");
            alignBtn[1].classList.remove("activeBtn");
            alignBtn[2].classList.add("activeBtn");

        }

    })
})

vAlign.forEach(btn => {
    btn.addEventListener("click", function (e) {
        let cellObj = getCellobj(lsc);
        let value = e.currentTarget.classList[1];
        lsc.style.alignItems = value;
        cellObj.vAlignment = `${value}`;
        if (value == 'flex-start') {
            vAlign[0].classList.add("activeBtn");
            vAlign[1].classList.remove("activeBtn");
            vAlign[2].classList.remove("activeBtn");
        } else if (value == 'center') {
            vAlign[0].classList.remove("activeBtn");
            vAlign[1].classList.add("activeBtn");
            vAlign[2].classList.remove("activeBtn");

        } else if (value == 'flex-end') {
            vAlign[0].classList.remove("activeBtn");
            vAlign[1].classList.remove("activeBtn");
            vAlign[2].classList.add("activeBtn");

        }


    })
})

fontSizeBtn.addEventListener("change", function () {

    let cellObj = getCellobj(lsc);
    let value = fontSizeBtn.value;
    lsc.style.fontSize = `${value}px`;
    cellObj.fontSize = `${value}`;
    setHeight(lsc);

})

fontBtn.addEventListener("click", function () {

    let cellObj = getCellobj(lsc);
    let value = fontBtn.value;
    lsc.style.fontFamily = value;
    cellObj.fontName = `${value}`;
    setHeight(lsc);
})

/*===================================CELL CLICK and BLUR ===========================================*/
cells.forEach(cell => {
    /*===================================SET HEIGHT==========================================*/
    cell.addEventListener("keyup", function () {
        setHeight(this);
    })

    cell.addEventListener("click", function (e) {
        let rid = parseInt(cell.getAttribute("r-Id"));
        let cid = parseInt(cell.getAttribute("c-Id"));
        let address = String.fromCharCode(65 + cid) + (rid + 1);
        addressBar.value = address;

        let obj = getCellobj(cell);
        formulaBar.value = obj.formula;
        let { vAlignment, cellFormatting, cellAlignment } = db[rid][cid];
        if (cellFormatting.bold) {
            boldBtn.classList.add("activeBtn");
        } else {
            boldBtn.classList.remove("activeBtn");
        }

        if (cellFormatting.italic) {
            italicBtn.classList.add("activeBtn");
        } else {
            italicBtn.classList.remove("activeBtn");
        }

        if (cellFormatting.underline) {
            underlineBtn.classList.add("activeBtn");
        } else {
            underlineBtn.classList.remove("activeBtn");
        }

        if (cellAlignment == 'flex-start') {
            alignBtn[0].classList.add("activeBtn");
            alignBtn[1].classList.remove("activeBtn");
            alignBtn[2].classList.remove("activeBtn");
        } else if (cellAlignment == 'center') {
            alignBtn[0].classList.remove("activeBtn");
            alignBtn[1].classList.add("activeBtn");
            alignBtn[2].classList.remove("activeBtn");

        } else if (cellAlignment == 'flex-end') {
            alignBtn[0].classList.remove("activeBtn");
            alignBtn[1].classList.remove("activeBtn");
            alignBtn[2].classList.add("activeBtn");

        }
        if (vAlignment == 'flex-start') {
            vAlign[0].classList.add("activeBtn");
            vAlign[1].classList.remove("activeBtn");
            vAlign[2].classList.remove("activeBtn");
        } else if (vAlignment == 'center') {
            vAlign[0].classList.remove("activeBtn");
            vAlign[1].classList.add("activeBtn");
            vAlign[2].classList.remove("activeBtn");

        } else if (vAlignment == 'flex-end') {
            vAlign[0].classList.remove("activeBtn");
            vAlign[1].classList.remove("activeBtn");
            vAlign[2].classList.add("activeBtn");

        }

    })

    /*===================================CELL BLUR===========================================*/

    cell.addEventListener("blur", function () {
        lsc = cell;
        let newVal = cell.innerHTML;
        let obj = getCellobj(cell);


        let NN_num = Number(newVal);
        if (newVal == "") {
            obj.value = "";
            obj.parent = [];
            obj.formula = "";
        } else if (!NN_num) {           //when your cell have child but you put some string, then children cell will show error
            obj.value = newVal;
            if (obj.child.length > 0) {
                for (let i = 0; i < obj.child.length; i++) {
                    let { rowId, colId } = get_rId_cId_fromAddress(obj.child[i]);
                    let childobj = db[rowId][colId];

                    showErr(childobj);
                }
            }
        } else if (newVal != obj.value) {
            obj.value = newVal;
            if (obj.formula) {
                removeFormula(obj);
               
                formulaBar.value = ""
            }
            updateChildren(obj);
        }
    });

})

/*===================================SET TOP MOST & LEFT MOST =========================================*/

contentDiv.addEventListener("scroll", function () {

    let topOffset = contentDiv.scrollTop;
    let leftOffset = contentDiv.scrollLeft;

    topMost.style.top = topOffset + "px";
    topLeft.style.top = topOffset + "px";
    leftMost.style.left = leftOffset + "px";
    topLeft.style.left = leftOffset + "px";
})

/*===================================CHANGE HEIGHT ===========================================*/

topMostAll.forEach(topCell => {
    topCell.addEventListener("mousedown", function (e) {
        if ((topCell.getBoundingClientRect().width - e.offsetX) < 5) {
            let rid = topCell.getAttribute("c-Id");
            changeWidth = rid;
            changeCell = topCell;
            changeDirection = 'H';
            contentDiv.style.cursor = "col-resize";
        }
    })


    window.addEventListener("mousemove", function (e) {
        if (changeWidth && changeCell != "" && changeDirection == 'H') {
            let finalX = e.clientX;

            let dx = finalX - changeCell.getBoundingClientRect().left;

            let allCells = this.document.querySelectorAll(`.cell[c-Id="${changeWidth}"]`);
            let topcell = topMostAll[changeWidth];

            topcell.style.minWidth = dx + "px";
            topcell.style.maxWidth = dx + "px";

            allCells.forEach(x => {
                x.style.minWidth = dx + "px";
                x.style.maxWidth = dx + "px";

            })

            for (let i = 0; i < 100; i++) {
                // console.log(changeWidth);
                db[i][changeWidth].width = dx;
            }
        }
    })

    window.addEventListener("mouseup", function () {
        changeWidth = "";
        changeCell = "";
        changeDirection = "";
        contentDiv.style.cursor = "auto";
    })
})

leftMostAll.forEach(leftCell => {
    leftCell.addEventListener("mousedown", function (e) {
        if ((leftCell.getBoundingClientRect().height - e.offsetY) < 5) {
            let cid = leftCell.getAttribute("r-Id");
            changeHeight = cid;
            changeCell = leftCell;
            changeDirection = 'V';
            contentDiv.style.cursor = "row-resize";
        }
    })


    window.addEventListener("mousemove", function (e) {
        if (changeHeight && changeCell != "" && changeDirection == 'V') {
            let finalY = e.clientY;
            let dy = finalY - changeCell.getBoundingClientRect().top;

            let allCells = this.document.querySelectorAll(`.cell[r-Id="${changeHeight}"]`);
            let leftcell = leftMostAll[changeHeight];

            leftcell.style.minHeight = dy + "px";
            leftcell.style.minHeight = dy + "px";

            allCells.forEach(x => {
                x.style.minHeight = dy + "px";
                x.style.minHeight = dy + "px";

            })
            for (let i = 0; i < 26; i++) {
                // console.log(changeWidth);
                db[changeHeight][i].height = dy;
            }
        }
    })

    window.addEventListener("mouseup", function () {
        changeHeight = "";
        changeCell = "";
        changeDirection = "";
        contentDiv.style.cursor = "auto";
    })
})


/*===================================FORMULA BLUR===========================================*/

formulaBar.addEventListener("blur", function () {
    let formula = formulaBar.value;
    let obj = getCellobj(lsc);
    if (obj.formula.length > 0 && formula.length == 0) {
        lsc.innerHTML = "";
        obj.formula = "";
        obj.value = "";
        removeFormula(obj)
        obj.child = [];
    } else if (obj.formula != formula) {

        removeFormula(obj);

        if (addFormula(formula)) {
            updateChildren(obj);
        }
    }
})

/*===================================ADD FORMULA===========================================*/

function addFormula(formula) {
    let obj = getCellobj(lsc);
    // A1
    let formulaArray = formula.split(" ");
    let children = obj.child;

    for (let i = 0; i < children.length; i++) {
        for (let j = 0; j < formulaArray.length; j++) {
            if (children[i] == formulaArray[j] || formulaArray[j] == obj.address) {
                let { rowId, colId } = get_rId_cId_fromAddress(obj.address);
                let currCell = document.querySelector(`.cell[r-Id= "${rowId}"][c-Id="${colId}"]`);
                currCell.innerHTML = ''
                currCell.value = ""
                promptDiv.style.display = "block"
                showErr(obj)
                console.log("cycle");

                return false;
            }
        }
    }

    for (let j = 0; j < formulaArray.length; j++) {
        if (formulaArray[j] == obj.address) {
            let { rowId, colId } = get_rId_cId_fromAddress(obj.address);
            let currCell = document.querySelector(`.cell[r-Id= "${rowId}"][c-Id="${colId}"]`);
            currCell.innerHTML = ''
            currCell.value = ""
            promptDiv.style.display = "block"
            showErr(obj)
            console.log("cycle");
            return false;
        }
    }

    obj.formula = formula;
    solveFormula(obj);
    return true;
}

/*===================================SOLVE FORMULA===========================================*/


function solveFormula(obj) {
    let formula = obj.formula;
    let fcomps = formula.split(" ");

    for (let i = 0; i < fcomps.length; i++) {
        fcomp = fcomps[i];

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
    console.log(formula);
    let value = evaluateFormula(formula);
    obj.value = value;
    let { rowId, colId } = get_rId_cId_fromAddress(obj.address);
    let currCell = document.querySelector(`.cell[r-Id= "${rowId}"][c-Id="${colId}"]`);
    currCell.innerHTML = value

}

/*===================================UPDATE CHILDREN===========================================*/
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
/*===================================RECALCULATE===========================================*/

function recalculate(obj) {
    let formula = obj.formula;
    let fcomps = formula.split(" ");

    for (let i = 0; i < fcomps.length; i++) {
        fcomp = fcomps[i];

        let initial = fcomp[0];

        if (initial >= "A" && initial <= "Z") {
            let { rowId, colId } = get_rId_cId_fromAddress(fcomp);
            let parentobj = db[rowId][colId];
            let value = Number(parentobj.value);
            formula = formula.replace(fcomp, value);
        }
    }
    console.log(formula);
    let value = evaluateFormula(formula);
    obj.value = value;
    let { rowId, colId } = get_rId_cId_fromAddress(obj.address);
    let currCell = document.querySelector(`.cell[r-Id= "${rowId}"][c-Id="${colId}"]`);
    currCell.innerHTML = value
}

/*===================================REMOVE FORMULA===========================================*/
function removeFormula(obj) {
    obj.formula = "";

    for (let i = 0; i < obj.parent.length; i++) {
        let { rowId, colId } = get_rId_cId_fromAddress(obj.parent[i]);
        let parentobj = db[rowId][colId];

        let newParentChild = parentobj.child.filter(function (child) {
            return child != obj.address;
        });

        parentobj.child = newParentChild;

    }
    obj.parent = [];
}

/*===================================SHOW ERROR===========================================*/
function showErr(obj) {

    let presentobj = get_rId_cId_fromAddress(obj.address);
    db[presentobj.rowId][presentobj.colId].value = "err";
    let currCell = document.querySelector(`.cell[r-Id="${presentobj.rowId}"][c-Id="${presentobj.colId}"]`);
    currCell.innerHTML = 'err'
    for (let i = 0; i < obj.child.length; i++) {
        let child = get_rId_cId_fromAddress(obj.child[i]);
        let childobj = db[child.rowId][child.colId];

        showErr(childobj);
    }


}

/*===================================GET CELL OBJ===========================================*/

function getCellobj(ele) {
    let rid = Number($(ele).attr("r-Id"));
    let cid = Number($(ele).attr("c-Id"));
    return db[rid][cid]
}

/*===================================GET ID'S FROM ADDRESS===========================================*/
function get_rId_cId_fromAddress(address) {
    let rId = Number(address.substring(1)) - 1;
    let cId = Number(address.charCodeAt(0) - 65)

    return {
        rowId: rId,
        colId: cId
    };
}
/*===================================GET ADDRESS FROM ID'S===========================================*/
function get_Address_from_rId_cId(rId, cId) {
    let rowId = rId + 1;
    let colId = String.fromCharCode(65 + cId);

    return `${colId}${rowId}`;
}

/*===================================FUNCTION RUN AT BOOT TIME===========================================*/
function init() {
    $(".new").trigger("click");
}
init();

//  ======================== set height fn===================================================
function setWidth(elem, width) {
    // console.log(width);
    let cid = $(elem).attr("c-Id");
    let leftcol = $(".topMostCell")[cid];
    $(leftcol).css({
        "min-width": width + "px",
        "max-width": width + "px"
    });
    $(`.cell[c-Id="${cid}"]`).css({
        "min-width": width + "px",
        "max-width": width + "px"
    });
    
}
function setHeightByDB(elem, height) {
    let rid = $(elem).attr("r-Id");
    let leftcol = $(".leftMostCell")[rid];
    $(leftcol).css({
        "min-height": height + "px",
        "height": height + "px"
    });
    $(`.cell[r-Id="${rid}"]`).css({
        "min-height": height + "px",
        "height": height + "px"
    });
    
}

function setHeight(elem) {
    let height = $(elem).height() + 2;
    let rid = $(elem).attr("r-Id");
    db[rid][0].height = height;
    let leftcol = $(".leftMostCell")[rid];
    $(leftcol).css({
        "min-height": height + "px",
        "max-height": height + "px"
    });
}



//  ======================== remove prompt ===================================================

okBtn.addEventListener("click",function(){
    promptDiv.style.display = "none";
})





// ==============================================================================================================================================

// Formula evaluation function
function evaluateFormula(formula, currentCell) {
    let result;
    try {
        // Check if formula starts with '='
        if (!formula.startsWith('=')) {
            result = eval(formula);
        } else {
            // Remove the '=' and get the function name and arguments
            let cleanFormula = formula.substring(1).trim();
            let funcName = cleanFormula.split('(')[0].toLowerCase();
            
            // Extract arguments between the first and last parentheses
            let argsString = cleanFormula.substring(cleanFormula.indexOf('(') + 1, cleanFormula.lastIndexOf(')'));
            
            // Handle special case for REPLACE which has specific argument requirements
            let args;
            if (funcName === 'replace') {
                args = argsString.split(',').map(arg => {
                    // Remove extra quotes for string arguments
                    return arg.trim().replace(/^"|"$/g, '').replace(/^'|'$/g, '');
                });
            } else {
                args = argsString.split(',').map(arg => arg.trim());
            }

            switch (funcName) {
                case 'sum':
                    result = calculateSum(args);
                    break;
                case 'average':
                    result = calculateAverage(args);
                    break;
                case 'min':
                    result = calculateMin(args);
                    break;
                case 'max':
                    result = calculateMax(args);
                    break;
                case 'count':
                    result = calculateCount(args);
                    break;
                case 'trim':
                    result = trimText(args[0]);
                    break;
                case 'upper':
                    result = upperText(args[0]);
                    break;
                case 'lower':
                    result = lowerText(args[0]);
                    break;
                case 'remove_duplicate':
                    result = removeDuplicates(args);
                    break;
                case 'replace':
                    result = findAndReplace(args[0], args[1], args[2]);
                    break;
                default:
                    throw new Error("Unknown function: " + funcName);
            }
        }

        // Update the cell with the calculated result
        updateCellValue(currentCell, result);
        return result;

    } catch (error) {
        console.error("Error evaluating formula:", error);
        updateCellValue(currentCell, "#ERROR!");
        return "#ERROR!";
    }
}

// Function to get values from a cell range
function getRangeValues(range) {
    let values = [];
    let [start, end] = range.split(':');
    let startCell = get_rId_cId_fromAddress(start);
    let endCell = get_rId_cId_fromAddress(end);

    for (let i = startCell.rowId; i <= endCell.rowId; i++) {
        for (let j = startCell.colId; j <= endCell.colId; j++) {
            let value = Number(db[i][j].value);
            if (!isNaN(value)) {
                values.push(value);
            }
        }
    }
    return values;
}

// Function to update cell value and display
function updateCellValue(cellAddress, value) {
    // Format numbers to a reasonable number of decimal places
    if (typeof value === 'number') {
        value = Number.isInteger(value) ? value : Number(value.toFixed(2));
    }

    // Get row and column IDs from the cell address
    const { rowId, colId } = get_rId_cId_fromAddress(cellAddress);

    // Update the cell's value in the database
    if (db[rowId] && db[rowId][colId]) {
        db[rowId][colId].value = value;
        
        // Find and update the actual cell element
        const cellElement = document.querySelector(`.cell[r-id="${rowId}"][c-id="${colId}"]`);
        if (cellElement) {
            cellElement.innerHTML = value;
        }

        // Update formula bar if this is the currently selected cell
        if (lsc && lsc.getAttribute('r-id') === rowId.toString() && 
            lsc.getAttribute('c-id') === colId.toString()) {
            const formulaBar = document.querySelector('.FormulaInput');
            if (formulaBar) {
                formulaBar.value = db[rowId][colId].formula || '';
            }
        }
    }
}

// Event handler for formula evaluation
function handleFormulaEvaluation(event) {
    if (event.key === 'Enter' && lsc) {
        const formula = event.target.value;
        
        // Get the address of the currently selected cell
        const rowId = lsc.getAttribute('r-id');
        const colId = lsc.getAttribute('c-id');
        const cellAddress = get_Address_from_rId_cId(parseInt(rowId), parseInt(colId));

        // Store the formula in the cell's data
        if (db[rowId] && db[rowId][colId]) {
            db[rowId][colId].formula = formula;
            
            // Evaluate the formula and update the display
            evaluateFormula(formula, cellAddress);
        }

        // Remove focus from the formula bar
        event.target.blur();
    }
}

// Add this event listener to your formula bar
document.querySelector('.FormulaInput').addEventListener('keydown', handleFormulaEvaluation);

// Implementation of each function
function calculateSum(args) {
    let sum = 0;
    args.forEach(arg => {
        if (arg.includes(':')) {
            // Handle range
            let values = getRangeValues(arg);
            sum += values.reduce((a, b) => a + b, 0);
        } else if (!isNaN(arg)) {
            // Handle direct numeric values
            sum += parseFloat(arg);
        }else {
            // Handle single cell
            let cell = get_rId_cId_fromAddress(arg);
            let value = Number(db[cell.rowId][cell.colId].value);
            if (!isNaN(value)) sum += value;
        }
    });
    return sum;
}

function calculateAverage(args) {
    let sum = 0;
    let count = 0;
    args.forEach(arg => {
        if (arg.includes(':')) {
            let values = getRangeValues(arg);
            sum += values.reduce((a, b) => a + b, 0);
            count += values.length;
        } else if (!isNaN(arg)) {
            // Handle direct numeric values
            sum += parseFloat(arg);
            count++;
        } else {
            let cell = get_rId_cId_fromAddress(arg);
            let value = Number(db[cell.rowId][cell.colId].value);
            if (!isNaN(value)) {
                sum += value;
                count++;
            }
        }
    });
    return count > 0 ? sum / count : 0;
}

function calculateMin(args) {
    let values = [];
    args.forEach(arg => {
        if (arg.includes(':')) {
            values = values.concat(getRangeValues(arg));
        } else if (!isNaN(arg)) {
            // Handle direct numeric values
            values.push(parseFloat(arg));
        } else {
            let cell = get_rId_cId_fromAddress(arg);
            let value = Number(db[cell.rowId][cell.colId].value);
            if (!isNaN(value)) values.push(value);
        }
    });
    return values.length > 0 ? Math.min(...values) : 0;
}

function calculateMax(args) {
    let values = [];
    args.forEach(arg => {
        if (arg.includes(':')) {
            values = values.concat(getRangeValues(arg));
        } else if (!isNaN(arg)) {
            // Handle direct numeric values
            values.push(parseFloat(arg));
        } else {
            let cell = get_rId_cId_fromAddress(arg);
            let value = Number(db[cell.rowId][cell.colId].value);
            if (!isNaN(value)) values.push(value);
        }
    });
    return values.length > 0 ? Math.max(...values) : 0;
}

function calculateCount(args) {
    let count = 0;
    args.forEach(arg => {
        if (arg.includes(':')) {
            let values = getRangeValues(arg);
            count += values.length;
        } else if (!isNaN(arg)) {
            // Handle direct numeric values
            count++;
        } else {
            let cell = get_rId_cId_fromAddress(arg);
            let value = db[cell.rowId][cell.colId].value;
            if (value !== "") count++;
        }
    });
    return count;
}

function trimText(arg) {
    if (!arg.includes(':')) {
        let cell = get_rId_cId_fromAddress(arg);
        let value = db[cell.rowId][cell.colId].value;
        return typeof value === 'string' ? value.trim() : value;
    } else if (typeof value !== 'string') {
        return "ERROR: Value is not a string";
    } else {
    return "ERROR: TRIM only works with single cells";
    }
}

function upperText(arg) {
    if (!arg.includes(':')) {
        let cell = get_rId_cId_fromAddress(arg);
        let value = db[cell.rowId][cell.colId].value;
        return typeof value === 'string' ? value.toUpperCase() : value;
    } else if (typeof value !== 'string') {
        return "ERROR: Value is not a string";
    } else {
        return "ERROR: UPPER only works with single cells";
    }
}

function lowerText(arg) {
    if (!arg.includes(':')) {
        let cell = get_rId_cId_fromAddress(arg);
        let value = db[cell.rowId][cell.colId].value;
        return typeof value === 'string' ? value.toLowerCase() : value;
    } else if (typeof value !== 'string') {
        return "ERROR: Value is not a string";
    } else {
        return "ERROR: LOWER only works with single cells";
    }
}


function removeDuplicates(args) {
    let values = [];
    args.forEach(arg => {
        if (arg.includes(':')) {
            let [start, end] = arg.split(':');
            let startCell = get_rId_cId_fromAddress(start);
            let endCell = get_rId_cId_fromAddress(end);

            for (let i = startCell.rowId; i <= endCell.rowId; i++) {
                for (let j = startCell.colId; j <= endCell.colId; j++) {
                    let value = db[i][j].value;
                    if (value !== "" && !values.includes(value)) {
                        values.push(value);
                    }
                }
            }
        }
    });
    return values.join(', ');
}

function findAndReplace(range, find, replace) {
    if (range.includes(':')) {
        let [start, end] = range.split(':');
        let startCell = get_rId_cId_fromAddress(start);
        let endCell = get_rId_cId_fromAddress(end);

        for (let i = startCell.rowId; i <= endCell.rowId; i++) {
            for (let j = startCell.colId; j <= endCell.colId; j++) {
                let value = db[i][j].value;
                if (typeof value === 'string' && value.includes(find)) {
                    db[i][j].value = value.replace(new RegExp(find, 'g'), replace);
                    let cell = document.querySelector(`.cell[r-id="${i}"][c-id="${j}"]`);
                    cell.innerHTML = db[i][j].value;
                }
            }
        }
        //return "Replacement complete";
    }
    return "ERROR: REPLACE needs a range";
}


document.getElementById('insert-function').addEventListener('click', function() {
    const functionSelect = document.getElementById('spreadsheet-functions');
    const selectedFunction = functionSelect.value;
    
    if (selectedFunction === 'replace') {
        // Show the find & replace dialog
        document.querySelector('.find-replace-dialog').style.display = 'flex';
    } else {
        // Insert the function template into the formula bar
        const formulaBar = document.querySelector('.FormulaInput');
        formulaBar.value = `=${selectedFunction.toUpperCase()}()`;
        formulaBar.focus();
        // Place cursor between the parentheses
        const cursorPos = formulaBar.value.length - 1;
        formulaBar.setSelectionRange(cursorPos, cursorPos);
    }
});

// Handle find & replace dialog
document.getElementById('replace-btn').addEventListener('click', function() {
    const range = document.getElementById('range-text').value;
    const findText = document.getElementById('find-text').value;
    const replaceText = document.getElementById('replace-text').value;
    
    const formulaBar = document.querySelector('.FormulaInput');
    formulaBar.value = `=REPLACE(${range},"${findText}","${replaceText}")`;
    
    // Close the dialog
    document.querySelector('.find-replace-dialog').style.display = 'none';
    
    // Clear the inputs
    document.getElementById('range-text').value = '';
    document.getElementById('find-text').value = '';
    document.getElementById('replace-text').value = '';
});

document.getElementById('cancel-btn').addEventListener('click', function() {
    document.querySelector('.find-replace-dialog').style.display = 'none';
});

document.querySelector('.close-dialog').addEventListener('click', function() {
    document.querySelector('.find-replace-dialog').style.display = 'none';
});

//====================================================== SELECTION FUNCTION ==========================================================

/*let isSelecting = false;
let selectionStart = null;
let selectionEnd = null;
let selectionBox = null;

function createSelectionBox() {
    if (!selectionBox) {
        selectionBox = document.createElement('div');
        selectionBox.classList.add('selection-box');
        document.querySelector('.cells').appendChild(selectionBox);
    }
    return selectionBox;
}

// Mouse down event to start selection
document.querySelectorAll('.cell').forEach(cell => {
    cell.addEventListener('mousedown', function(e) {
        if (e.button === 0) { // Left click only
            isSelecting = true;
            selectionStart = {
                row: parseInt(this.getAttribute('r-id')),
                col: parseInt(this.getAttribute('c-id'))
            };
            selectionEnd = { ...selectionStart }; // Initialize selectionEnd at the start
            updateSelection(); // Update the selection visually
            e.preventDefault(); // Prevent text selection
        }
    });

    // Mouse over event to update the selection
    cell.addEventListener('mouseover', function(e) {
        if (isSelecting) {
            selectionEnd = {
                row: parseInt(this.getAttribute('r-id')),
                col: parseInt(this.getAttribute('c-id'))
            };
            updateSelection(); // Update selection on hover
        }
    });
});

// Mouse up event to finish selection
window.addEventListener('mouseup', function(e) {
    if (isSelecting) {
        isSelecting = false;
        if (selectionStart && selectionEnd) {
            updateFormulaBarWithRange(); // Update formula bar when selection is done
        }
    }
});

// Function to update the selection display
function updateSelection() {
    if (!selectionStart || !selectionEnd) return;

    let startRow = Math.min(selectionStart.row, selectionEnd.row);
    let endRow = Math.max(selectionStart.row, selectionEnd.row);
    let startCol = Math.min(selectionStart.col, selectionEnd.col);
    let endCol = Math.max(selectionStart.col, selectionEnd.col);

    let startCell = document.querySelector(`.cell[r-id="${startRow}"][c-id="${startCol}"]`);
    let endCell = document.querySelector(`.cell[r-id="${endRow}"][c-id="${endCol}"]`);

    let startRect = startCell.getBoundingClientRect();
    let endRect = endCell.getBoundingClientRect();
    
    let contentDiv = document.querySelector('.content');
    let scrollLeft = contentDiv.scrollLeft;
    let scrollTop = contentDiv.scrollTop;

    let box = createSelectionBox();
    box.style.display = 'block';
    box.style.left = (startRect.left + scrollLeft - contentDiv.getBoundingClientRect().left) + 'px';
    box.style.top = (startRect.top + scrollTop - contentDiv.getBoundingClientRect().top) + 'px';
    box.style.width = (endRect.right - startRect.left) + 'px';
    box.style.height = (endRect.bottom - startRect.top) + 'px';

    // Highlight selected cells
    document.querySelectorAll('.cell.selected').forEach(cell => {
        cell.classList.remove('selected');
    });
    for (let i = startRow; i <= endRow; i++) {
        for (let j = startCol; j <= endCol; j++) {
            let cell = document.querySelector(`.cell[r-id="${i}"][c-id="${j}"]`);
            cell.classList.add('selected');
        }
    }
}

// Function to update formula bar with the selected range
function updateFormulaBarWithRange() {
    if (!selectionStart || !selectionEnd) return;

    let startAddress = get_Address_from_rId_cId(selectionStart.row, selectionStart.col);
    let endAddress = get_Address_from_rId_cId(selectionEnd.row, selectionEnd.col);

    if (startAddress === endAddress) {
        document.querySelector('.FormulaInput').value = startAddress;
    } else {
        document.querySelector('.FormulaInput').value = `${startAddress}:${endAddress}`;
    }
}*/
