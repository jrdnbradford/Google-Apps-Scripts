/* 
   Google Sheet-bound script that creates a column that allows sorting rows by color
   License: MIT (c) 2019 Jordan Bradford
   GitHub: jrdnbradford
   
   Recommended OAuth Scopes:
   https://www.googleapis.com/auth/spreadsheets.currentonly
   https://www.googleapis.com/auth/script.container.ui
*/


var ss = SpreadsheetApp.getActiveSpreadsheet();
var userCache = CacheService.getUserCache();
var Ui = SpreadsheetApp.getUi();


function onOpen(e) {
    Ui.createMenu("Row Color Sort")
        .addItem("Add Sorting Column", "hasHeaderPrompt")
        .addToUi();
} // Create menu


function hasHeaderPrompt() {
    userCache.put("Processing", true);
    
    var response = Ui.alert("Row Color Sort","Does this Sheet have a header?", Ui.ButtonSet.YES_NO);
    var activeSheet = SpreadsheetApp.getActiveSheet();

    if (response == Ui.Button.YES) {
        rowColorSort(true, activeSheet);
    } else if (response == Ui.Button.NO) {
        rowColorSort(false, activeSheet);
    } else {
        ss.toast("You clicked the close button.");
    }
} // Raise Yes/No prompt for header row


function openProgressDialog() {
    var html = HtmlService
                 .createHtmlOutputFromFile("index")
                 .setWidth(200)
                 .setHeight(50);
  
    Ui.showModalDialog(html, "Processing...");
} // Open client side progress screen 


function isProcessing() {
    return userCache.get("Processing");
} // Called from client side to check completion of loop in rowColorSort


function getHexColor(cell) { 
    return ss.getRange(cell).getBackgrounds(); 
} // Return hex value of cell background color
  

function rowColorSort(hasHeader, sheet) {
    /* Main function. Hides a column, adds the background color of each row as the text content 
    of the column/row intersection, sorts the Sheet, then removes the column content and 
    unhides the column.*/
    
    var sortColumn = sheet.getLastColumn() + 1; // Empty column after last column with content
    var emptyRow = sheet.getLastRow() + 1; // Empty row after last row with content
    sheet.hideColumns(sortColumn); // Hide column while adding hex values and sorting

    // Assumes header is only first row
    if (hasHeader) {
        var startRow = 2;
        sheet.setFrozenRows(1);
        sheet.getRange(1, sortColumn).setValue("Sort Column"); 
    } else {
        var startRow = 1;
    } 
    
    openProgressDialog();
    for (var row = startRow; row < emptyRow; row++) {
        /* hexCell -- cell that needs hex value added
        cellA1 -- A1 notation of each used row in first column
        
        The background color of first cell in each row of the first 
        column is used to identify the entire row's background color */
        var hexCell = sheet.getRange(row, sortColumn);
        var cellA1 = sheet.getRange(row, 1).getA1Notation();
        hexCell.setValue(getHexColor(cellA1));
    } // Iterate over cells in column, add hex values
    
    userCache.put("Processing", false);
    
    // Sort sheet and clear/unhide column
    /* sheet.sort(sortColumn);
       var rangeToClear = sheet.getRange(1, sortColumn, emptyRow);
       rangeToClear.clear({contentsOnly: true}); */
    
    sheet.showColumns(sortColumn);
    ss.toast("All done!");
} 