const ps = new PerfectScrollbar("#cells", {  //applying PerfectScrollBar on cells 
    wheelSpeed: 15  // changing speed of scrollbar, default value is 1
});


// for column and row entry i,e; column from A(i=1) to CV(i=100)  and row from (i=1) to (i=100)
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

let cellData = []; // will store the default value for each cell its backgroundcolor,font style,font size,textcolor,cells alignment whether it is left aligned,or center or right aligned.

// for adding cells in excel sheet
for (let i = 1; i <= 100; i++) {
    let row = $('<div class="cell-row"></div>')
    let rowArray = [];
    for (let j = 1; j <= 100; j++) {
        // id is unique for each cell and contenteditable means we can write in the cells
        row.append(`<div id="row-${i}-col-${j}" class="input-cell" contenteditable="false"></div>`);
        rowArray.push({
            "font-family": "Noto Sans",
            "font-size": 14,
            "text": "",
            "bold": false,
            "italic": false,
            "underlined": false,
            "alignment": "left",
            "color": "#fff",
            "bgcolor": "#444"
        })
    }
    cellData.push(rowArray);
    $("#cells").append(row);
}

// applying scroll property on the cells 
$("#cells").scroll(function(e) {
    $("#columns").scrollLeft(this.scrollLeft); //columns will also scroll with the scrollbar i,e;columns A - CV.
    //this.scrollLeft scrolls the scrollbar horizontally and hence columns.scrollLeft will scroll columns side by side. 
    $("#rows").scrollTop(this.scrollTop); // rows will also scroll with the scrollbar i,e; rows 1 - 100.
});

$(".input-cell").dblclick(function(e) {
    $(".input-cell.selected").removeClass("selected top-selected bottom-selected left-selected right-selected");
    $(this).addClass("selected");
    $(this).attr("contenteditable", "true");
    $(this).focus();
});

// on blurring, the cells conteneteditable becomes false means we deselected the cell on which we were writing previously.
$(".input-cell").blur(function(e) {
    $(this).attr("contenteditable", "false");
});


// will get index of row and column selected
function getRowCol(ele) { 
    let id = $(ele).attr("id"); // will give id from div class => "row-${i}-col-${j}" .
    let idArray = id.split("-"); // will split above id on the basis of "-" hiphen. idArray=>[row,uniqueRowId,col,uniqueColId] .
    let rowId = parseInt(idArray[1]); // will pass index 1 element present from id into rowId as unique row is present on index 1 for each cell.
    let colId = parseInt(idArray[3]); // will pass index 3 element present from id into colId as unique column is present on index 3 for each cell.
    return [rowId,colId];
}

//will give adjacent cells position i,e; row and column of top,bottom,left and right cell of the given selected cell.
function getTopLeftBottomRightCell(rowId,colId) {
    let topCell = $(`#row-${rowId - 1}-col-${colId}`);
    let bottomCell = $(`#row-${rowId + 1}-col-${colId}`);
    let leftCell = $(`#row-${rowId}-col-${colId - 1}`);
    let rightCell = $(`#row-${rowId}-col-${colId + 1}`);
    return [topCell,bottomCell,leftCell,rightCell];
}
$(".input-cell").click(function(e) {
    let [rowId,colId] = getRowCol(this);
    let [topCell,bottomCell,leftCell,rightCell] = getTopLeftBottomRightCell(rowId,colId);
    if($(this).hasClass("selected") && e.ctrlKey) {
        unselectCell(this,e,topCell,bottomCell,leftCell,rightCell);
    } else {
        selectCell(this,e,topCell,bottomCell,leftCell,rightCell);
    }
});

function unselectCell(ele,e,topCell,bottomCell,leftCell,rightCell) {
    if($(ele).attr("contenteditable") == "false") {
        if($(ele).hasClass("top-selected")) {
            topCell.removeClass("bottom-selected");
        }
    
        if($(ele).hasClass("bottom-selected")) {
            bottomCell.removeClass("top-selected");
        }
    
        if($(ele).hasClass("left-selected")) {
            leftCell.removeClass("right-selected");
        }
    
        if($(ele).hasClass("right-selected")) {
            rightCell.removeClass("left-selected");
        }
    
        $(ele).removeClass("selected top-selected bottom-selected left-selected right-selected");
    }

}

function selectCell(ele,e,topCell,bottomCell,leftCell,rightCell) { // ele means current cell and e means event 
    if(e.ctrlKey) { //if control key is pressed

        //.hasClass returns true if the given class passed to it is present otherwise return false.

        // top selected or not
        let topSelected;
        if(topCell) {
            topSelected = topCell.hasClass("selected");
        }

        // bottom selected or not
        let bottomSelected;
        if(bottomCell) {
            bottomSelected = bottomCell.hasClass("selected");
        }


        // left selected or not
        let leftSelected;
        if(leftCell) {
            leftSelected = leftCell.hasClass("selected");
        }

        // right selected or not
        let rightSelected;
        if(rightCell) {
            rightSelected = rightCell.hasClass("selected");
        }
        
        // if topSelected is true then,
        if(topSelected) {
            $(ele).addClass("top-selected"); // add class top-selected to the current selected cell(ele). 
            topCell.addClass("bottom-selected"); //now add bottom-selected class to topCell for which border-bottom is provided.
        }
        
        // Same case for all the other cells i,e; bottomSelected and leftSelected and rightSelected

        if(bottomSelected) {
            $(ele).addClass("bottom-selected");
            bottomCell.addClass("top-selected");
        }

        if(leftSelected) {
            $(ele).addClass("left-selected");
            leftCell.addClass("right-selected");
        }

        if(rightSelected) {
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
    let data = cellData[rowId - 1][colId - 1]; // acquired the cell on which alignment position is to be applied.
    $(".alignment.selected").removeClass("selected"); 
    $(`.alignment[data-type=${data.alignment}]`).addClass("selected");  // added the selected class on the cell's alignment factor which were acquired previously.
    addRemoveSelectFromFontStyle(data, "bold");
    addRemoveSelectFromFontStyle(data, "italic");
    addRemoveSelectFromFontStyle(data, "underlined");

    // will give default value to all the below selectors their previous value which were alloted to them.
    //like for font size it was 14. and for font-family it was Noto Sans. and same for other selectors. 
    $("#fill-color").css("border-bottom", `4px solid ${data.bgcolor}`); 
    $("#text-color").css("border-bottom", `4px solid ${data.color}`);
    $("#font-family").val(data["font-family"]);
    $("#font-size").val(data["font-size"]);
    $("#font-family").css("font-family",data["font-family"]);
}

// this function applies and removes selected class from taskbar BOLD,ITALIC AND UNDERLINE BUTTONS accordingly. 
function addRemoveSelectFromFontStyle(data, property) {
    if (data[property]) {
        $(`#${property}`).addClass("selected");
    } else {
        $(`#${property}`).removeClass("selected");
    }
}

let count = 0;
let startcellSelected = false;
let startCell = {};   //empty object to store starting row id and col id of the cells.
let endCell = {};   //empty object to store ending row id and col id of the cells.
let scrollXRStarted = false; //defines that scrolling has not done previously in right direction of x axes
let scrollXLStarted = false; //defines that scrolling has not done previously in left direction of x axes

$(".input-cell").mousemove(function (e) {  // mousemove is an event handler for mouse in JS.        
    e.preventDefault();  // will prevent default action to happen for mouse related activities.
    // e.buttons is an object property of event for which 1 is Left click mouse button and 2 is right click mouse button and 0 is for none.
    if (e.buttons == 1) {

        //e.pageX defines the position of the mouse where it is present 
        //window.width returns the size of the screen in pixels
        if(e.pageX > ($(window).width() - 10) && !scrollXRStarted) {
            scrollXR();
        } else if(e.pageX < (10) && !scrollXLStarted) {
            scrollXL();
        }
        if (!startcellSelected) {  // if mouse event has started means we have not made some cells as a single cell.
            let [rowId, colId] = getRowCol(this);  // will get rowId and colId for the starting point for making multiple cells into one cell.
            startCell = { "rowId": rowId, "colId": colId };  // will store rowId and colId into an object. In startCell rowId and colId are the ones where we clicked the mouse for dragging multiple cells. 
            selectAllBetweenCells(startCell, startCell);
            startcellSelected = true;  //making it as true as for making start cell fixed.
        }
    } else {
        startcellSelected = false; // making it as false so another multiple cells can be combined into one cell.  
    }
});


//mouseenter is better than mousemove as it reduces the no of for loops in selectAllBetweenCells function which is accessed in making multiple cells into one cell.

$(".input-cell").mouseenter(function (e) { // mouseenter is an event handler for mouse in JS.
    if (e.buttons == 1) {

        // this if and else is used so that => agr hum right ko ja rhe h aur beech mai hi left ko scroll kre aur vice-versa isliye use kr rhe h.
        if(e.pageX < ($(window).width() - 10) && scrollXRStarted) {
            clearInterval(scrollXRInterval);
            scrollXRStarted = false;
        }

        if(e.pageX > 10 && scrollXLStarted) {
            clearInterval(scrollXLInterval);
            scrollXLStarted = false;
        }
        let [rowId, colId] = getRowCol(this);
        endCell = { "rowId": rowId, "colId": colId }; // storing rowId and colId into endCell object.In endCell rowId and colId are the ones where we left the mouse after dragging multiple cells.
        selectAllBetweenCells(startCell, endCell);
    }
})

// function used to select cells between the dragged area from startCell to endCell.This function will provide us the cells which are in between startCell and endCell which were dragged for combining multiple cells.
function selectAllBetweenCells(start, end) {
    $(".input-cell.selected").removeClass("selected top-selected bottom-selected left-selected right-selected");

    // the next 2 for loops will access the cells from minimum position to maximum position dragged in between startCell and endCell.
    //i,e; 1st loop will start from min rowId and will loop through max rowId.
    // and 2nd loop will start from min colId and will loop till max colID.
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
    scrollXRStarted = true; // making it true as selecting of cells is being done by the mouse in right direction.
    scrollXRInterval =  setInterval(() => {
        $("#cells").scrollLeft($("#cells").scrollLeft() + 100);  // will scroll cells to x axes(right) direction by 100px every 100 millisec. 
    }, 100);
}


function scrollXL() {
    scrollXLStarted = true;
    scrollXLInterval =  setInterval(() => {
        $("#cells").scrollLeft($("#cells").scrollLeft() - 100); // will scroll cells to x axes(left) direction by 100px every 100ms.
    }, 100);
}

$(".data-container").mousemove(function(e){
    e.preventDefault();
    if (e.buttons == 1) {
        //e.pageX defines the position of the mouse where it is present 
        //window.width returns the size of the screen in pixels
        if(e.pageX > ($(window).width() - 10) && !scrollXRStarted) {
            scrollXR();
        } else if(e.pageX < (10) && !scrollXLStarted) {
            scrollXL();
        }
    }
});

$(".data-container").mouseup(function(e) {  // mouseup is an event handler which will work when clicking of mouse is left. 
    clearInterval(scrollXRInterval); // will stop the right scrolling functioning
    clearInterval(scrollXLInterval);  // will stop the left scrolling functioning
    scrollXRStarted = false; // making it back to false so that selection of cells can be done again by clicking the mouse. 
    scrollXLStarted = false; // making it back to false so that selection of cells can be done again by clicking the mouse.
});

$(".alignment").click(function (e) { // on clicking the alignment class ,perform the next operations.
    let alignment = $(this).attr("data-type"); // alignment variable will store the clicked alignment from the taskbar on which we clicked.
    $(".alignment.selected").removeClass("selected");  // will remove selected class from the selected alignment button from taskbar.
    $(this).addClass("selected"); // will add selected class to the selected taskbar alignment button
    $(".input-cell.selected").css("text-align", alignment); // will apply the selected alignment to the input cells which we selected .
    $(".input-cell.selected").each(function (index, data) {
        let [rowId, colId] = getRowCol(data); // acquiring cell's position 
        cellData[rowId - 1][colId - 1].alignment = alignment; //updating cell's data from cellData i,e; updating its alignment factor.
    });
});

 // used to bold the cell's content
$("#bold").click(function (e) {
    setStyle(this, "bold", "font-weight", "bold");
});
// used to italisize the cell's content
$("#italic").click(function (e) {
    setStyle(this, "italic", "font-style", "italic");
});
// used to undeline the cell's content
$("#underlined").click(function (e) {
    setStyle(this, "underlined", "text-decoration", "underline");
});

// here ele is element on which we want to apply properties.
//here property is id means ITALIC ,BOLD OR UNDERLINE.
//here key is css property key.
//here value is css property value.
function setStyle(ele, property, key, value) {
    if ($(ele).hasClass("selected")) { //hasClass returns true if the given passed parameter(here in this case it is class=>selected) is present on the element 
        // if part=> us selected cell se property remove krta hai whether it is bold property or italic or underline.
        $(ele).removeClass("selected"); 
        $(".input-cell.selected").css(key, ""); // removing css property from the given key by passing empty value("").
        $(".input-cell.selected").each(function (index, data) {
            let [rowId, colId] = getRowCol(data);
            cellData[rowId - 1][colId - 1][property] = false; //making cell's selected(bold,italic or underline) property from cellData false.
        });
    } else { 
        // else part=> us selected cell(element) pr property(bold ,italic or underline) lagata hai. 
        $(ele).addClass("selected");
        $(".input-cell.selected").css(key, value); //passing css property of key(bold,italic or underline) and value to the selected cell.
        $(".input-cell.selected").each(function (index, data) {
            let [rowId, colId] = getRowCol(data);
            cellData[rowId - 1][colId - 1][property] = true; //making cell's selected(bold,italic or underline) property from cellData true.
        });
    }
}


$(".pick-color").colorPick({
    'initialColor': "#abcd", //default value.
    'allowRecent': true, //default value
    'recentMax': 5, //default value
    'allowCustomColor': true, //default value
    //default value
    'palette': ["#1abc9c", "#16a085", "#2ecc71", "#27ae60", "#3498db", "#2980b9", "#9b59b6", "#8e44ad", "#34495e", "#2c3e50", "#f1c40f", "#f39c12", "#e67e22", "#d35400", "#e74c3c", "#c0392b", "#ecf0f1", "#bdc3c7", "#95a5a6", "#7f8c8d"],
    
    'onColorSelected': function () {
        if(this.color != "#ABCD") {
            if($(this.element.children()[1]).attr("id") == "fill-color") {
                $(".input-cell.selected").css("background-color",this.color);
                $("#fill-color").css("border-bottom",`4px solid ${this.color}`);
                // in the next 4 lines of code , we are updating the cell's content(here bgcolor)(default values)in cellData which we selected from the above code. 
                $(".input-cell.selected").each((index,data) => {
                    let [rowId, colId] = getRowCol(data);
                    cellData[rowId - 1][colId - 1].bgcolor = this.color;
                });
            }
            if($(this.element.children()[1]).attr("id") == "text-color") {
                $(".input-cell.selected").css("color",this.color);
                $("#text-color").css("border-bottom",`4px solid ${this.color}`);
                // in the next 4 lines of code , we are updating the cell's content(here color)(default values)in cellData which we selected from the above code.
                $(".input-cell.selected").each((index,data) => {
                    let [rowId, colId] = getRowCol(data);
                    cellData[rowId - 1][colId - 1].color = this.color;
                });
            }
        }
    }
});

// works on image of pick-color. i,e; if we click on the image of pick-color ,it will call its parent() element which is div(pick-color).
$("#fill-color").click(function(e) {
    setTimeout(() => {
        $(this).parent().click();
    }, 10);
});

// works on image of text-color. i,e; if we click on the image of text-color ,it will call its parent() element which is div(text-color).
$("#text-color").click(function(e) {
    setTimeout(() => {
        $(this).parent().click();
    }, 10);
});

$(".menu-selector").change(function(e) {
    let value = $(this).val();  // will store the changed value which we want to chage. LIKE for font size if we want to change 14(defalut value) into 18, then it will store 18.SAME for font-family.
    let key = $(this).attr("id"); // will acquire the key(property) on which we want to change like font-family or font-size. 
    if(key == "font-family") { 
        $("#font-family").css(key,value);  // will apply css property i,e; font-family to selected value which we have acquired previously on line 406.
     }

    if(!isNaN(value)) { //if we want to change font-size i,e; value we stored on line 406 is number, then follow these if steps.
        value = parseInt(value);  //changing string into integer.MEANS the font-size value which are present in html like 12,14,16,18 etc are present in the form of string.HENCE changing into integer.
    }

    $(".input-cell.selected").css(key,value);  // applying css property on the selected input cells.
    $(".input-cell.selected").each((index,data) => {
        let [rowId,colId] = getRowCol(data);
        cellData[rowId-1][colId -1][key] = value; //updating the cell's content in cellData.
    })
})