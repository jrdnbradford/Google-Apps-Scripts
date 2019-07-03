/*
SYNOPSIS:
Google Sheet-bound script that logs access, permissions, 
editors, and viewers of all files in user's Drive.

LICENSE: MIT (c) 2019 Jordan Bradford
GITHUB: jrdnbradford

RECOMMENDED SCOPES:
https://www.googleapis.com/auth/drive.readonly
https://www.googleapis.com/auth/gmail.send
https://www.googleapis.com/auth/spreadsheets.currentonly
*/


var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();
var rows = [];


function onOpen() {
    SpreadsheetApp.getUi()
                  .createMenu("Audit Drive")
                  .addItem("Audit Drive", "auditDrive")
                  .addToUi();
}



function auditDrive() {   
    var root = DriveApp.getRootFolder();  
    createSheetRows(root.getFiles(), "Root Drive");
    getChildFolders(root); 
    
    sheet.clear();
    sheet.appendRow(["File", "Folder", "Access", "Permission", "Editors", "Viewers"]);
    sheet.getRange(2, 1, rows.length, 6)
         .setValues(rows);
    
    formatSheet(); 
    var url = ss.getUrl();
    sendEmail(url);
}



function getChildFolders(parentFolder) {
    var childFolders = parentFolder.getFolders();
    while (childFolders.hasNext()) {
        var childFolder = childFolders.next();
        var childFolderName = childFolder.getName();
        var files = childFolder.getFiles();  
        createSheetRows(files, childFolderName);
        getChildFolders(childFolder);  
    }
}



function createSheetRows(fileIter, folderName) {
    while (fileIter.hasNext()) {
        var file = fileIter.next();
        var fileName = file.getName();
        var access = file.getSharingAccess();
        var permission = file.getSharingPermission();
        var editorEmails = getEmails(file.getEditors());
        var viewerEmails = getEmails(file.getViewers());
        //Logger.log(fileName + folderName + access + permission + editorEmails + viewerEmails);
        rows.push([fileName, folderName, access, permission, editorEmails, viewerEmails]);
    }
} 



function formatSheet() {
    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    
    var headerRange = sheet.getRange(1, 1, 1, lastCol);
    headerRange.setFontWeight("bold")
               .setFontSize(24);
    
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, lastCol);
        
    var dataRange = sheet.getRange(1, 1, lastRow, lastCol);
    dataRange.setHorizontalAlignment("left");
}



function getEmails(users) {
    if (users.length >= 1) {
        var emails = [];
        for (var i = 0; i < users.length; i++) {
        emails.push(users[i].getEmail());
        } return emails.join(",");
    } else {return "None";}
}



function sendEmail(url) {
    var yourGmailAddress = Session.getActiveUser().getEmail();
    var subject = "Drive Audit Sheet";
    var body = url;
    GmailApp.sendEmail(yourGmailAddress, subject, body);
} 