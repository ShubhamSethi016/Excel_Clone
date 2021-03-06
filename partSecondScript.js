const ps = new PerfectScrollbar("#cells", {
    wheelSpeed: 15
});

for (let i = 1; i <= 100; i++) {
    let str = "";
    let n = i;

    while (n > 0) {
        let rem = n % 26;
        if (rem == 0) {
            str = "Z" + str;
            n = Math.floor(n / 26) - 1;
        } else {
            str = String.fromCharCode((rem - 1) + 65) + str;
            n = Math.floor(n / 26);
        }
    }
    $("#columns").append(`<div class="column-name">${str}</div>`)
    $("#rows").append(`<div class="row-name">${i}</div>`)
}

// main container for storing the data.
//cellData will store the particular sheet cell content whose properties are not default.MEANS it will store the particular cell's row and col whose properties are changed from default.
let cellData = {
    "Sheet1": {} //empty object for sheet1.
};

let save = true; // variable used for save file command.

let selectedSheet = "Sheet1"; //default selected
let totalSheets = 1;  //default no of sheets.
let lastlyAddedSheet = 1; // default no of sheets.

//default properties for each cell.
let defaultProperties = {
    "font-family": "Noto Sans",
    "font-size": 14,
    "text": "",
    "bold": false,
    "italic": false,
    "underlined": false,
    "alignment": "left",
    "color": "#444",
    "bgcolor": "#fff"
};

//creating of rows and columns.
for (let i = 1; i <= 100; i++) {
    let row = $('<div class="cell-row"></div>');
    for (let j = 1; j <= 100; j++) {
        row.append(`<div id="row-${i}-col-${j}" class="input-cell" contenteditable="false"></div>`);
    }
    $("#cells").append(row);
}

// scrolling rows(1-100) and columns(A-CV) with scrollbar.
$("#cells").scroll(function (e) {
    $("#columns").scrollLeft(this.scrollLeft);
    $("#rows").scrollTop(this.scrollTop);
});

//by double clicking, cells can become contenteditable. 
$(".input-cell").dblclick(function (e) {
    $(".input-cell.selected").removeClass("selected top-selected bottom-selected left-selected right-selected");
    $(this).addClass("selected");
    $(this).attr("contenteditable", "true");
    $(this).focus();
});


$(".input-cell").blur(function (e) {
    $(this).attr("contenteditable", "false");
    updateCellData("text", $(this).text());
});

// function used to access specific row and column of a cell by using id which we uniquely specified to each cell in the time of creation.
function getRowCol(ele) {
    let id = $(ele).attr("id");
    let idArray = id.split("-");
    let rowId = parseInt(idArray[1]);
    let colId = parseInt(idArray[3]);
    return [rowId, colId];
}

//function used to get topCell's,bottomCell's,leftCell's and rightCell's row and col of a particular cell.
function getTopLeftBottomRightCell(rowId, colId) {
    let topCell = $(`#row-${rowId - 1}-col-${colId}`);
    let bottomCell = $(`#row-${rowId + 1}-col-${colId}`);
    let leftCell = $(`#row-${rowId}-col-${colId - 1}`);
    let rightCell = $(`#row-${rowId}-col-${colId + 1}`);
    return [topCell, bottomCell, leftCell, rightCell];
}
$(".input-cell").click(function (e) {
    let [rowId, colId] = getRowCol(this);
    let [topCell, bottomCell, leftCell, rightCell] = getTopLeftBottomRightCell(rowId, colId);
    if ($(this).hasClass("selected") && e.ctrlKey) {
        unselectCell(this, e, topCell, bottomCell, leftCell, rightCell);
    } else {
        selectCell(this, e, topCell, bottomCell, leftCell, rightCell);
    }
});

function unselectCell(ele, e, topCell, bottomCell, leftCell, rightCell) {
    if ($(ele).attr("contenteditable") == "false") {
        if ($(ele).hasClass("top-selected")) {
            topCell.removeClass("bottom-selected");
        }

        if ($(ele).hasClass("bottom-selected")) {
            bottomCell.removeClass("top-selected");
        }

        if ($(ele).hasClass("left-selected")) {
            leftCell.removeClass("right-selected");
        }

        if ($(ele).hasClass("right-selected")) {
            rightCell.removeClass("left-selected");
        }

        $(ele).removeClass("selected top-selected bottom-selected left-selected right-selected");
    }

}

function selectCell(ele, e, topCell, bottomCell, leftCell, rightCell) {
    if (e.ctrlKey) {

        // top selected or not
        let topSelected;
        if (topCell) {
            topSelected = topCell.hasClass("selected");
        }

        // bottom selected or not
        let bottomSelected;
        if (bottomCell) {
            bottomSelected = bottomCell.hasClass("selected");
        }


        // left selected or not
        let leftSelected;
        if (leftCell) {
            leftSelected = leftCell.hasClass("selected");
        }

        // right selected or not
        let rightSelected;
        if (rightCell) {
            rightSelected = rightCell.hasClass("selected");
        }

        if (topSelected) {
            $(ele).addClass("top-selected");
            topCell.addClass("bottom-selected");
        }

        if (bottomSelected) {
            $(ele).addClass("bottom-selected");
            bottomCell.addClass("top-selected");
        }

        if (leftSelected) {
            $(ele).addClass("left-selected");
            leftCell.addClass("right-selected");
        }

        if (rightSelected) {
            $(ele).addClass("right-selected");
            rightCell.addClass("left-selected");
        }
    } else {
        $(".input-cell.selected").removeClass("selected top-selected bottom-selected left-selected right-selected");
    }
    $(ele).addClass("selected");
    changeHeader(getRowCol(ele));
}


function changeHeader([rowId, colId]) {
    let data;
    // if row and col in cellData exists then run if part otherwise run else part.
    if(cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1]) {
        data = cellData[selectedSheet][rowId - 1][colId - 1];
    } else {
        data = defaultProperties;
    }
    $(".alignment.selected").removeClass("selected");
    $(`.alignment[data-type=${data.alignment}]`).addClass("selected");
    addRemoveSelectFromFontStyle(data, "bold");
    addRemoveSelectFromFontStyle(data, "italic");
    addRemoveSelectFromFontStyle(data, "underlined");
    $("#fill-color").css("border-bottom", `4px solid ${data.bgcolor}`);
    $("#text-color").css("border-bottom", `4px solid ${data.color}`);
    $("#font-family").val(data["font-family"]);
    $("#font-size").val(data["font-size"]);
    $("#font-family").css("font-family", data["font-family"]);
}

function addRemoveSelectFromFontStyle(data, property) {
    if (data[property]) {
        $(`#${property}`).addClass("selected");
    } else {
        $(`#${property}`).removeClass("selected");
    }
}

let count = 0;
let startcellSelected = false;
let startCell = {};
let endCell = {};
let scrollXRStarted = false;
let scrollXLStarted = false;
$(".input-cell").mousemove(function (e) {
    e.preventDefault();
    if (e.buttons == 1) {
        if (e.pageX > ($(window).width() - 10) && !scrollXRStarted) {
            scrollXR();
        } else if (e.pageX < (10) && !scrollXLStarted) {
            scrollXL();
        }
        if (!startcellSelected) {
            let [rowId, colId] = getRowCol(this);
            startCell = { "rowId": rowId, "colId": colId };
            selectAllBetweenCells(startCell, startCell);
            startcellSelected = true;
        }
    } else {
        startcellSelected = false;
    }
});

$(".input-cell").mouseenter(function (e) {
    if (e.buttons == 1) {
        if (e.pageX < ($(window).width() - 10) && scrollXRStarted) {
            clearInterval(scrollXRInterval);
            scrollXRStarted = false;
        }

        if (e.pageX > 10 && scrollXLStarted) {
            clearInterval(scrollXLInterval);
            scrollXLStarted = false;
        }
        let [rowId, colId] = getRowCol(this);
        endCell = { "rowId": rowId, "colId": colId };
        selectAllBetweenCells(startCell, endCell);
    }
})

function selectAllBetweenCells(start, end) {
    $(".input-cell.selected").removeClass("selected top-selected bottom-selected left-selected right-selected");
    for (let i = Math.min(start.rowId, end.rowId); i <= Math.max(start.rowId, end.rowId); i++) {
        for (let j = Math.min(start.colId, end.colId); j <= Math.max(start.colId, end.colId); j++) {
            let [topCell, bottomCell, leftCell, rightCell] = getTopLeftBottomRightCell(i, j);
            selectCell($(`#row-${i}-col-${j}`)[0], { "ctrlKey": true }, topCell, bottomCell, leftCell, rightCell);
        }
    }
}

let scrollXRInterval;
let scrollXLInterval;
function scrollXR() {
    scrollXRStarted = true;
    scrollXRInterval = setInterval(() => {
        $("#cells").scrollLeft($("#cells").scrollLeft() + 100);
    }, 100);
}


function scrollXL() {
    scrollXLStarted = true;
    scrollXLInterval = setInterval(() => {
        $("#cells").scrollLeft($("#cells").scrollLeft() - 100);
    }, 100);
}

$(".data-container").mousemove(function (e) {
    e.preventDefault();
    if (e.buttons == 1) {
        if (e.pageX > ($(window).width() - 10) && !scrollXRStarted) {
            scrollXR();
        } else if (e.pageX < (10) && !scrollXLStarted) {
            scrollXL();
        }
    }
});

$(".data-container").mouseup(function (e) {
    clearInterval(scrollXRInterval);
    clearInterval(scrollXLInterval);
    scrollXRStarted = false;
    scrollXLStarted = false;
});

$(".alignment").click(function (e) {
    let alignment = $(this).attr("data-type");
    $(".alignment.selected").removeClass("selected");
    $(this).addClass("selected");
    $(".input-cell.selected").css("text-align", alignment);
    updateCellData("alignment",alignment);
});

$("#bold").click(function (e) {
    setStyle(this, "bold", "font-weight", "bold");
});

$("#italic").click(function (e) {
    setStyle(this, "italic", "font-style", "italic");
});

$("#underlined").click(function (e) {
    setStyle(this, "underlined", "text-decoration", "underline");
});

function setStyle(ele, property, key, value) {
    if ($(ele).hasClass("selected")) {
        $(ele).removeClass("selected");
        $(".input-cell.selected").css(key, "");
        // $(".input-cell.selected").each(function (index, data) {
        //     let [rowId, colId] = getRowCol(data);
        //     cellData[rowId - 1][colId - 1][property] = false;
        // });
        updateCellData(property,false);
    } else {
        $(ele).addClass("selected");
        $(".input-cell.selected").css(key, value);
        // $(".input-cell.selected").each(function (index, data) {
        //     let [rowId, colId] = getRowCol(data);
        //     cellData[rowId - 1][colId - 1][property] = true;
        // });
        updateCellData(property,true);
    }
}


$(".pick-color").colorPick({
    'initialColor': "#abcd",
    'allowRecent': true,
    'recentMax': 5,
    'allowCustomColor': true,
    'palette': ["#1abc9c", "#16a085", "#2ecc71", "#27ae60", "#3498db", "#2980b9", "#9b59b6", "#8e44ad", "#34495e", "#2c3e50", "#f1c40f", "#f39c12", "#e67e22", "#d35400", "#e74c3c", "#c0392b", "#ecf0f1", "#bdc3c7", "#95a5a6", "#7f8c8d"],
    'onColorSelected': function () {
        if (this.color != "#ABCD") {
            if ($(this.element.children()[1]).attr("id") == "fill-color") {
                $(".input-cell.selected").css("background-color", this.color);
                $("#fill-color").css("border-bottom", `4px solid ${this.color}`);
                // $(".input-cell.selected").each((index, data) => {
                //     let [rowId, colId] = getRowCol(data);
                //     cellData[rowId - 1][colId - 1].bgcolor = this.color;
                // });
                updateCellData("bgcolor",this.color)
            }
            if ($(this.element.children()[1]).attr("id") == "text-color") {
                $(".input-cell.selected").css("color", this.color);
                $("#text-color").css("border-bottom", `4px solid ${this.color}`);
                // $(".input-cell.selected").each((index, data) => {
                //     let [rowId, colId] = getRowCol(data);
                //     cellData[rowId - 1][colId - 1].color = this.color;
                // });
                updateCellData("color",this.color);
            }
        }
    }
});


$("#fill-color").click(function (e) {
    setTimeout(() => {
        $(this).parent().click();
    }, 10);
});

$("#text-color").click(function (e) {
    setTimeout(() => {
        $(this).parent().click();
    }, 10);
});

$(".menu-selector").change(function (e) {
    let value = $(this).val();
    let key = $(this).attr("id");
    if (key == "font-family") {
        $("#font-family").css(key, value);
    }
    if (!isNaN(value)) {
        value = parseInt(value);
    }

    $(".input-cell.selected").css(key, value);
    // $(".input-cell.selected").each((index, data) => {
    //     let [rowId, colId] = getRowCol(data);
    //     cellData[rowId - 1][colId - 1][key] = value;
    // })
    updateCellData(key,value);
});

function updateCellData(property,value) {
    let currCellData = JSON.stringify(cellData); //
    if (value != defaultProperties[property]) {
        $(".input-cell.selected").each(function (index, data) {
            let [rowId, colId] = getRowCol(data);
            if (cellData[selectedSheet][rowId - 1] == undefined) { //this line checks if the row exits or not.If not then follow if part and exits then follow else part. 
                cellData[selectedSheet][rowId - 1] = {}; //creation of row.

                // { ...defaultProperties } this command makes the copy of defaultProperties OBJECT.
                // [ ...defaultProperties] this command makes the copy of defaultProperties ARRAY.
                cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties }; //creation of column and storing default properties in the cell.
                cellData[selectedSheet][rowId - 1][colId - 1][property] = value; //updating value in the passed property in selectedSheet cellData.
            } else {
                //else part runs when the row is defined means it exists.
                if (cellData[selectedSheet][rowId - 1][colId - 1] == undefined) { //checks if the column is defined(exists).
                    cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties }; // if col not exists,then creating the col and storing default properties to the cell.
                    cellData[selectedSheet][rowId - 1][colId - 1][property] = value; //updating value in the passed property in selectedSheet cellData.
                } else { //if col exists then just simply updating the given value in property in selectedSheet cellData.
                    cellData[selectedSheet][rowId - 1][colId - 1][property] = value;
                }
            }

        });
    } else {
        $(".input-cell.selected").each(function (index, data) {
            let [rowId, colId] = getRowCol(data);
            if (cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1]) { //this if checks that if row and col exits previously then enter the if part. 
                cellData[selectedSheet][rowId - 1][colId - 1][property] = value;
                //the next if checks that the row and col we got above has the default properties or not. If it has default properties then delete that particular cell.
                if (JSON.stringify(cellData[selectedSheet][rowId - 1][colId - 1]) == JSON.stringify(defaultProperties)) {
                    delete cellData[selectedSheet][rowId - 1][colId - 1]; //will delete that particular cell from selectedSheet cellData.
                    if(Object.keys(cellData[selectedSheet][rowId - 1]).length == 0) {
                        delete cellData[selectedSheet][rowId - 1];
                    }
                }
            }
        });
    }
    //
    if (save && currCellData != JSON.stringify(cellData)) {
        save = false;
    }
}
//
$(".container").click(function (e) {
    $(".sheet-options-modal").remove();
});


function addSheetEvents() {
    $(".sheet-tab.selected").on("contextmenu", function (e) {
        e.preventDefault(); // this will make sure that our inspect element modal does not open by clicking with mouse.
        selectSheet(this);
        $(".sheet-options-modal").remove(); // this will make sure that one single modal should open at a single time.This is done so that when we have more than one sheet in our excel then ,it can open modal for any sheet any number of times.Hence we removed our modal before adding another modal into our html. 
        // creating HTML content (here modal) through JS.
        let modal = $(`<div class="sheet-options-modal">
                        <div class="option sheet-rename">Rename</div>
                        <div class="option sheet-delete">Delete</div>
                    </div>`);
        modal.css({ "left": e.pageX }); // e.pageX defines our cursor position in pixels.
        $(".container").append(modal); // adding modal to our container.
        $(".sheet-rename").click(function (e) { // will work on clicking of sheet-rename class element.
            // adding HTML content through JS.
            let renameModal = $(`<div class="sheet-modal-parent">
                                    <div class="sheet-rename-modal">
                                        <div class="sheet-modal-title">Rename Sheet</div>
                                        <div class="sheet-modal-input-container">
                                            <span class="sheet-modal-input-title">Rename Sheet to:</span>
                                            <input class="sheet-modal-input" type="text" />
                                        </div>
                                        <div class="sheet-modal-confirmation">
                                            <div class="button yes-button">OK</div>
                                            <div class="button no-button">Cancel</div>
                                        </div>
                                    </div>
                                </div>`);

            $(".container").append(renameModal); //adding to container.
            $(".sheet-modal-input").focus(); // input tag will be put into focus after opening of renameModal. 
            $(".no-button").click(function (e) { // when clicked on cancel button(.no-button),renameModal will be removed.
                $(".sheet-modal-parent").remove();
            });
            $(".yes-button").click(function (e) { //when clicked on yes button renameSheet() function will be called.
                renameSheet();
            });
            // the next is EVENTLISTENER of keyboard i,e;keypress.MEANS after entering the new name to the sheet in input tag ,then by keypressing the enter keyword, renameSheet() function gets called as same as clicking the yes button with mouse. 
            $(".sheet-modal-input").keypress(function (e) {
                if (e.key == "Enter") {
                    renameSheet();
                }
            })
        });

        $(".sheet-delete").click(function (e) {
            if (totalSheets > 1) { // delete command will work when there are more than 1 sheets present in the excel. 
                let deleteModal = $(`<div class="sheet-modal-parent">
                                    <div class="sheet-delete-modal">
                                        <div class="sheet-modal-title">Sheet Name</div>
                                        <div class="sheet-modal-detail-container">
                                            <span class="sheet-modal-detail-title">Are you sure?</span>
                                        </div>
                                        <div class="sheet-modal-confirmation">
                                            <div class="button yes-button">
                                                <div class="material-icons delete-icon">delete</div>
                                                Delete
                                            </div>
                                            <div class="button no-button">Cancel</div>
                                        </div>
                                    </div>
                                </div>`);

                $(".container").append(deleteModal);
                $(".no-button").click(function (e) {
                    $(".sheet-modal-parent").remove();
                });
                $(".yes-button").click(function (e) {
                    deleteSheet();
                });
            } else {
                alert("Not possible");
            }
        })
    });

    $(".sheet-tab.selected").click(function (e) {
        selectSheet(this);
    });
}

addSheetEvents();


// will add new sheet to our sheet-bar.
$(".add-sheet").click(function (e) {
    //
    save = false; 
    lastlyAddedSheet++; // will increase the sheet number index when we want to add a new sheet.MEANS by default our sheet number is Sheet1 and by clicking add sheet class element it will add sheet number 2 to our sheet bar i,e;Sheet2 and so on. hence we increased the lastlyaddedsheet index everytime when a new sheet is added. 
    totalSheets++; // increasing the count of total no of sheets by one everytime when a new sheet is added to our excel.
    cellData[`Sheet${lastlyAddedSheet}`] = {}; //adding sheetName to our cellData container like Sheet2 or Sheet3 ,etc and making an empty object to store the defaultProperties.
    $(".sheet-tab.selected").removeClass("selected"); //removing selected class from our sheet tab which were previously selected.
    $(".sheet-tab-container").append(`<div class="sheet-tab selected">Sheet${lastlyAddedSheet}</div>`); //adding new sheet tag html to our sheet-tab-container with selected class.MEANS whenver any new sheet is added,By Default that sheet is our selected sheet.
    selectSheet();
    addSheetEvents(); //adding events to our newly added sheets.(EVENTS are making MODAL=>RENAME,DELETE) which are applied to each sheet. 
    //
    $(".sheet-tab.selected")[0].scrollIntoView();
});

function selectSheet(ele) {
    if (ele && !$(ele).hasClass("selected")) {
        $(".sheet-tab.selected").removeClass("selected");
        $(ele).addClass("selected");
    }
    emptyPreviousSheet();
    selectedSheet = $(".sheet-tab.selected").text();
    loadCurrentSheet();
    $("#row-1-col-1").click(); // this single command will apply our focus of each cell on 1st cell(row1 & col1) of every new sheet added.
}

// this function empties our previous sheet content.
function emptyPreviousSheet() {
    let data = cellData[selectedSheet]; //this will store our sheet Name(previous sheet selected) form our cellData.
    let rowKeys = Object.keys(data); // will store the rows from previously selected sheet.ROWS stored are that rows where our default properties(like font-family,alignment,bgcolor,etc) are changed(modified).
    for (let i of rowKeys) { // passing loop through the rows.
        let rowId = parseInt(i); //converting string into integer. i,e; each row we obtain from the rowsKey are by default of string type.HENCE converting into an integer. 
        let colKeys = Object.keys(data[rowId]); //storing col of each row where are default properties are changed. 
        for (let j of colKeys) {
            let colId = parseInt(j); //converting string into integer.
            let cell = $(`#row-${rowId + 1}-col-${colId + 1}`); // storing position of row and col of changed defaultProperties into cell variable.
            cell.text(""); //making css empty of that cell where default properties have been changed.
            cell.css({ // giving that cell our previously defaultProperties so that when adding a new sheet ,the complete cells are empty which is required for making a new sheet.
                "font-family": "NotoSans",
                "font-size": 14,
                "background-color": "#fff",
                "color": "#444",
                "font-weight": "",
                "font-style": "",
                "text-decoration": "",
                "text-align": "left"
            });
        }
    }
}

// will load a new sheet after making previous sheet empty.
// same functionality is performed.
function loadCurrentSheet() {
    let data = cellData[selectedSheet];
    let rowKeys = Object.keys(data);
    for (let i of rowKeys) {
        let rowId = parseInt(i);
        let colKeys = Object.keys(data[rowId]);
        for (let j of colKeys) {
            let colId = parseInt(j);
            let cell = $(`#row-${rowId + 1}-col-${colId + 1}`);
            cell.text(data[rowId][colId].text);
            cell.css({
                "font-family": data[rowId][colId]["font-family"],
                "font-size": data[rowId][colId]["font-size"],
                "background-color": data[rowId][colId]["bgcolor"],
                "color": data[rowId][colId].color,
                "font-weight": data[rowId][colId].bold ? "bold" : "",
                "font-style": data[rowId][colId].italic ? "italic" : "",
                "text-decoration": data[rowId][colId].underlined ? "underline" : "",
                "text-align": data[rowId][colId].alignment
            });
        }
    }
}

function renameSheet() {
    let newSheetName = $(".sheet-modal-input").val(); //will store the new name given to the sheet through input tag.
    // the next if checks =>1st condition=> the input tag val entered is present or not means any new name is given or it is empty.
    //2nd condition of if statement=> if newName of the sheet is already present in cellData as other's sheetName,then return false.LIKE=>present sheetname--Hello and newSheetName entered -- Hello.Both the names are same means already present ,return false.  
    if (newSheetName && !Object.keys(cellData).includes(newSheetName)) { 
        let newCellData = {}; // empty objectto store the sheet names.
        for (let i of Object.keys(cellData)) { //looping through cellData i,e;Sheet1,Sheet2,Sheet3,etc.(no of sheets present in excel)
            if (i == selectedSheet) { // this means if we are renaming the sheet is the selected sheet,then rename the sheet and store it into an object.
                newCellData[newSheetName] = cellData[selectedSheet]; //will rename the currentsheet name to newSheetName that we entered.
            } else {
                newCellData[i] = cellData[i]; //else storing the same name of our sheets to our object.
            }
        }

        cellData = newCellData; //assigning newCellData into cellData.

        selectedSheet = newSheetName; 
        $(".sheet-tab.selected").text(newSheetName);
        $(".sheet-modal-parent").remove(); //closing the modal after renaming the sheet.

    } else { // IF the above if statement becomes false from the above 2 conditions,then
        $(".rename-error").remove();  // removing the previous same errors as it will be popped again and again after the if staement gets false.
        //alerting the message(error) to the user ont he screen
        $(".sheet-modal-input-container").append(`
            <div class="rename-error"> Sheet Name is not valid or Sheet already exists! </div>
        `)
    }
}

function deleteSheet() {
    $(".sheet-modal-parent").remove();
    let sheetIndex = Object.keys(cellData).indexOf(selectedSheet); //will give the index of sheet from cellData so that deletion can be done.
    let currSelectedSheet = $(".sheet-tab.selected"); //selecting the current selected sheet which is to be deleted.
    if (sheetIndex == 0) {  // if sheet of index is 0 which is to be deleted, then make the next sheet selected.
        selectSheet(currSelectedSheet.next()[0]);
    } else { // else select the previous sheet.
        selectSheet(currSelectedSheet.prev()[0]);
    }
    delete cellData[currSelectedSheet.text()];  // deleting the current selected sheet from cellData.
    currSelectedSheet.remove(); // removing the sheet from sheet bar which is deleted.
    totalSheets--; // decreasing the count of total no of sheets by 1 after deleting the sheet.
}

// this will help in scrolling the sheets from sheet bar.
$(".left-scroller,.right-scroller").click(function(e){
    let keysArray = Object.keys(cellData);  // storing the values in keysArray from cellData which are sheetNames. 
    let selectedSheetIndex = keysArray.indexOf(selectedSheet); // getting the indexes of sheetNames from keysArray.
    if(selectedSheetIndex != 0 && $(this).text() == "arrow_left") { // if sheet index is !=0 means there are already sheets present on the left side of selected sheet, then select left sheet from the sheetbar.
        selectSheet($(".sheet-tab.selected").prev()[0]); // means will scroll and select the left side sheet(previous sheet from sheetbar).
    // else if means if there are already sheets present on the right side of the selected sheet,then scroll to the right sheet i,e; next sheet from the sheetbar. 
    } else if(selectedSheetIndex != (keysArray.length - 1) && $(this).text() == "arrow_right") { // $(this).text() == "arrow_right" => this command checks that the clicked scroller is right arrow or left arrow.If clicked scroller' text value is arrow_right ,then scroll to the right sheet(next sheet).
        selectSheet($(".sheet-tab.selected").next()[0]); //will scroll to the next sheet.
    }
    
    $(".sheet-tab.selected")[0].scrollIntoView(); // this will help in scrolling the sheetname as well as scrollbar when no of sheets are more which will make scrollbar.
});


// FILE TASKBAR COMMANDS
$("#menu-file").click(function(e) {
    // adding HTML content through JS.
    let fileModal = $(`<div class="file-modal">
                        <div class="file-options-modal">
                            <div class="close">
                                <div class="material-icons close-icon">arrow_circle_down</div>
                                <div>Close</div>
                            </div>
                            <div class="new">
                                <div class="material-icons new-icon">insert_drive_file</div>
                                <div>New</div>
                            </div>
                            <div class="open">
                                <div class="material-icons open-icon">folder_open</div>
                                <div>Open</div>
                            </div>
                            <div class="save">
                                <div class="material-icons save-icon">save</div>
                                <div>Save</div>
                            </div>
                        </div>
                        <div class="file-recent-modal"></div>
                        <div class="file-transparent"></div>
                    </div>`);
    $(".container").append(fileModal);
    // animations are applied on the file modal when it is clicked and removed.
    fileModal.animate({
        width: "100vw"
    },300);
    $(".close,.file-transparent").click(function(e) {
        fileModal.animate({
            width: "0vw"
        },300);
        setTimeout(() => {
            fileModal.remove();
        }, 250);
    });
});

// when clicked on new class element from fileModal, perform these operations.
$(".new").click(function (e) {
    if (save) { //if save == true, MEANS no changes are done to the existing excel file.
        newFile(); // make a new excel sheet file.
    } else {
        $(".container").append(`<div class="sheet-modal-parent">
                                    <div class="sheet-delete-modal">
                                        <div class="sheet-modal-title">${$(".title").text()}</div>
                                        <div class="sheet-modal-detail-container">
                                            <span class="sheet-modal-detail-title">Do you want to save changes?</span>
                                        </div>
                                        <div class="sheet-modal-confirmation">
                                            <div class="button yes-button">
                                                Yes
                                            </div>
                                            <div class="button no-button">No</div>
                                        </div>
                                    </div>
                                </div>`);
        $(".yes-button").click(function (e) {
            // save function
        });
        $(".no-button,.yes-button").click(function (e) {
            $(".sheet-modal-parent").remove();
            newFile();
        });
    }
});
$(".save").click(function(e){
    saveFile();
})

function newFile() {
emptyPreviousSheet();
cellData = { "Sheet1": {} };
$(".sheet-tab").remove();
$(".sheet-tab-container").append(`<div class="sheet-tab selected">Sheet1</div>`);
addSheetEvents();
selectedSheet = "Sheet1";
totalSheets = 1;
lastlyAddedSheet = 1;
$(".title").text("Excel - Book");
$("#row-1-col-1").click();
}

function saveFile() {
$(".container").append(`<div class="sheet-modal-parent">
                            <div class="sheet-rename-modal">
                                <div class="sheet-modal-title">Save File</div>
                                <div class="sheet-modal-input-container">
                                    <span class="sheet-modal-input-title">File Name:</span>
                                    <input class="sheet-modal-input" value="${$(".title").text()}" type="text" />
                                </div>
                                <div class="sheet-modal-confirmation">
                                    <div class="button yes-button">Save</div>
                                    <div class="button no-button">Cancel</div>
                                </div>
                            </div>
                        </div>`);
$(".yes-button").click(function(e) {
    $(".title").text($(".sheet-modal-input").val());
    let a = document.createElement("a");
    a.href = `data:application/json,${JSON.stringify(cellData)}`;
    a.download = $(".title").text() + ".json";
    $(".container").append(a);
    a.click();
    a.remove();
    save = true;
});
$(".no-button,.yes-button").click(function (e) {
    $(".sheet-modal-parent").remove();
});
}