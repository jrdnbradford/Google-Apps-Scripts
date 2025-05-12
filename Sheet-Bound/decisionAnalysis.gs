/*
SYNOPSIS:
Google Sheet-bound script that 
performs basic decision analysis.

LICENSE: MIT (c) 2019 Jordan Bradford

GITHUB: jrdnbradford

RECOMMENDED SCOPES:
https://www.googleapis.com/auth/spreadsheets.currentonly
*/

var app = SpreadsheetApp;
var decisionSheet = app.getActiveSpreadsheet();
var dataWorksheetName = decisionSheet.getSheetName();
var activeRange = app.getActiveRange();

var numCols = activeRange.getNumColumns();
var numRows = activeRange.getNumRows();

var firstCol = activeRange.getColumn();
var lastCol = firstCol + numCols;

var firstRow = activeRange.getRow();
var lastRow = firstRow + numRows;
    
var Ui = SpreadsheetApp.getUi();

var dataSheet = decisionSheet.getSheetByName(dataWorksheetName);


function onOpen(e) {
    Ui.createMenu("Decision Analysis")
        .addItem("Analyze Range", "analyze")
        .addItem("Create Profit Regret Table", "buildProfitRegretTable")
        .addItem("Delete All Conditional Formatting", "deleteAllConditionalFormatRules")
        .addToUi();
} 



function analyze() { 
    // Main
    var probRowRange = seekProbRowRange();
    if (probRowRange) {
        var probRowIndex = probRowRange.getRow();
        var probRowRangeA1 = probRowRange.getA1Notation();
            
        setLabels();
        setMaxMinEVRowFormulas(probRowRangeA1);
        setMaxMinEVSolutionFormulas(); 
        setExpectValSolutionFormulas(probRowIndex);
        addConditionalFormat();
    }
}



function seekProbRowRange() {
    var sumOfProbs = 0;
    var rowIdxToCheck = firstRow;

    do {
        if (rowIdxToCheck == 1) {
            decisionSheet.toast(
                "Could not identify a row with \
                 probabilities. Either format \
                 the row as percentages or ensure \
                 that the numbers add to 1."
            );    
            return;
        }
        rowIdxToCheck--;
        var possibleProbRowRange = dataSheet.getRange(rowIdxToCheck, firstCol, 1, numCols);
        var rowFormats = possibleProbRowRange.getNumberFormats()[0];
        var rowVals = possibleProbRowRange.getValues()[0];
        sumOfProbs = possibleProbRowRange.getValues()[0].reduce(function(a, b) {return a + b;}, 0);
        
        if (rowFormats.every(checkPercent) || rowVals.every(checkNumber) || sumOfProbs == 1) {break;}
        
    } while (true);
    return possibleProbRowRange;
}



function buildProfitRegretTable() {
    var regretTableLabel = dataSheet.getRange(lastRow + 1, firstCol);
    formatRange(regretTableLabel, "PROFIT REGRET TABLE");
    for (var i = 0; i < numCols; i++) {   
        var regretTableColFormulas = []; 
        for (var j = 0; j < numRows; j++) {
            var tableColRange = dataSheet.getRange(firstRow, firstCol + i, numRows, 1);                      
            var tableColRangeA1 = tableColRange.getA1Notation();
            var dataCellA1 = dataSheet.getRange(firstRow + j, firstCol + i).getA1Notation();
            var formula = "=MAX(" + tableColRangeA1 + ") - " + dataCellA1;
            regretTableColFormulas.push([formula]);
        }  
        var regretTableColRange = dataSheet.getRange(lastRow + 2, firstCol + i, numRows, 1);
        regretTableColRange.setFormulas(regretTableColFormulas);
    } 
    
    // Regret table border
    var regretTableRange = dataSheet.getRange(lastRow + 2, firstCol, numRows, numCols).setBorder(true, true, true, true, false, false);
    
    // Regret table MAX
    for (var i = 0; i < numRows; i++) {   
        var regretTableRowRangeA1 = dataSheet.getRange(lastRow + i + 2, firstCol, 1, numCols).getA1Notation();                                 
        dataSheet.getRange(lastRow + i + 2, lastCol + 2).setFormula("=MAX(" + regretTableRowRangeA1 + ")");
    }
            
    // Minimax regret cell
    var regretMaxValRangeA1 = dataSheet.getRange(lastRow + 2, lastCol + 2, numRows, 1).getA1Notation();      
    var minimaxRegretSolutionCell = dataSheet.getRange(lastRow + numRows + 2, lastCol + 2).setBackground("lightblue");                                                                                                       
    minimaxRegretSolutionCell.setFormula("=MIN(" + regretMaxValRangeA1 + ")");
    setThickSolutionBorder(minimaxRegretSolutionCell);
    
    var minimaxRegretLabel = dataSheet.getRange(lastRow + numRows + 2, lastCol + 1);
    formatRange(minimaxRegretLabel, "MINIMAX REGRET");   
}



function checkPercent(item) {
    // Used to find range/row with the probabilities
    return item[item.length - 1] == "%"; 
}



function checkNumber(item) {
    // Used to find range/row with the probabilities
    return typeof(item) == "number";
}



function setMaxMinEVRowFormulas(probRowRangeA1) {
    var maxMinEVFormulas = [];
    for (var i = 0; i < numRows; i++) {
        var tableRowRangeA1 = dataSheet.getRange(firstRow + i, firstCol, 1, numCols).getA1Notation();  
        maxMinEVFormulas.push(
                [
                    "=MAX(" + tableRowRangeA1 + ")", 
                    "=MIN(" + tableRowRangeA1 + ")", 
                    "=SUMPRODUCT(" + tableRowRangeA1 + "," + probRowRangeA1 + ")"
                ]
            );       
    }
    // Add MAX, MIN, and EV formulas to range
    var maxMinEVRange = dataSheet.getRange(firstRow, lastCol + 2, numRows, 3);
    maxMinEVRange.setFormulas(maxMinEVFormulas);
}  



function setMaxMinEVSolutionFormulas() {
    // Maximax, minimax cells
    var maxColRangeA1 = dataSheet.getRange(firstRow, lastCol + 2, numRows, 1).getA1Notation();
    
    var maximaxSolutionCell = dataSheet.getRange(lastRow, lastCol + 2);
    maximaxSolutionCell.setFormula("=MAX(" + maxColRangeA1 + ")")
    
    var minimaxSolutionCell = dataSheet.getRange(lastRow + 1, lastCol + 2);
    minimaxSolutionCell.setFormula("=MIN(" + maxColRangeA1 + ")");
    
    // Maximin, minimin cells
    var minColRangeA1 = dataSheet.getRange(firstRow, lastCol + 3, numRows, 1).getA1Notation();    
    
    var maximinSolutionCell = dataSheet.getRange(lastRow, lastCol + 3);
    maximinSolutionCell.setFormula("=MAX(" + minColRangeA1 + ")");
    
    var miniminSolutionCell = dataSheet.getRange(lastRow + 1, lastCol + 3);
    miniminSolutionCell.setFormula("=MIN(" + minColRangeA1 + ")");
    
    // EV cells
    var expectValColRangeA1 = dataSheet.getRange(firstRow, lastCol + 4, numRows, 1).getA1Notation();    
    var maxExpectValSolutionCell = dataSheet.getRange(lastRow, lastCol + 4);
    maxExpectValSolutionCell.setFormula("=MAX(" + expectValColRangeA1 + ")");
    
    var minExpectValSolutionCell = dataSheet.getRange(lastRow + 1, lastCol + 4);
    minExpectValSolutionCell.setFormula("=MIN(" + expectValColRangeA1 + ")");
    
    // Highlight
    var maxSolutionRange = dataSheet.getRange(lastRow, lastCol + 2, 1, 3).setBackground("lightgreen");
    var minSolutionRange = dataSheet.getRange(lastRow + 1, lastCol + 2, 1, 3).setBackground("lightblue");
    
    // Border
    var fullMaxMinEVSolutionRange = dataSheet.getRange(lastRow, lastCol + 2, 2, 3);
    setThickSolutionBorder(fullMaxMinEVSolutionRange);
}



function setExpectValSolutionFormulas(probRowIndex) {
    // Create formula for expected value with perfect information    
    var expectValPerfInfoFormulas = [];
    for (var i = 0; i < numCols; i++) {  
        var tableColRange = dataSheet.getRange(firstRow, firstCol + i, numRows, 1);                            
        var tableColRangeA1 = tableColRange.getA1Notation();                             
        var colProbabilityA1 = dataSheet.getRange(probRowIndex, firstCol + i, 1, 1).getA1Notation(); 
        var formula = "(MAX(" + tableColRangeA1 + ") * " + colProbabilityA1 + ")";
        expectValPerfInfoFormulas.push(formula);
    }
    // Expected value with perfect information
    var expectValPerfInfoRange = dataSheet.getRange(firstRow, lastCol + 5);
    var expectValPerfInfoFormula = "=" + expectValPerfInfoFormulas.join(" + ");
    expectValPerfInfoRange.setFormula(expectValPerfInfoFormula).setBackground("lightgreen");
    
    // Expected value of perfect information
    var expectValPerfInfoRangeA1 = expectValPerfInfoRange.getA1Notation();  
    var expectValNoPerfInfoA1 = dataSheet.getRange(firstRow, lastCol + 4, numRows, 1).getA1Notation();                                
    var expectValOfPerfInfoSolutionCell = dataSheet.getRange(firstRow, lastCol + 6);
    var formula = "=" + expectValPerfInfoRangeA1 + " - MAX(" + expectValNoPerfInfoA1 + ")";
    expectValOfPerfInfoSolutionCell.setFormula(formula).setBackground("lightgreen");
    
    setThickSolutionBorder(dataSheet.getRange(firstRow, lastCol + 5, 1, 2));
}



function setLabels() {
    // Column labels
    var colLabels = ["MAX", "MIN", "EV", "EV w/ PI", "EV of PI"];
    dataSheet.getRange(firstRow - 1, lastCol + 2, 1, colLabels.length)
                     .setValues([colLabels])
                     .setHorizontalAlignment("center")
                     .setFontWeight("bold");
    
    // Row labels
    var maxLabel = dataSheet.getRange(lastRow, lastCol + 1);
    formatRange(maxLabel, "MAX");
    
    var minLabel = dataSheet.getRange(lastRow + 1, lastCol + 1);
    formatRange(minLabel, "MIN");                             
}



function formatRange(range, val) {
    range.setValue(val)
        .setHorizontalAlignment("center")
        .setFontWeight("bold");
}



function addConditionalFormat() {
    var rules = dataSheet.getConditionalFormatRules();
    for (var i = 0; i < 3; i++) {
        for (var j = 0; j < numRows; j++) { 
            var cellToFormatRange = dataSheet.getRange(firstRow + j, lastCol + i + 2);
            var cellToFormatA1 = cellToFormatRange.getA1Notation();
            var maxRowSolutionA1 = dataSheet.getRange(lastRow, lastCol + i + 2).getA1Notation();
            var minRowSolutionA1 = dataSheet.getRange(lastRow + 1, lastCol + i + 2).getA1Notation();
                    
            var rule = buildCondFormatRule(cellToFormatA1, maxRowSolutionA1, "lightgreen", cellToFormatRange);
            rules.push(rule);
            var rule = buildCondFormatRule(cellToFormatA1, minRowSolutionA1, "lightblue", cellToFormatRange);
            rules.push(rule);
        }   
    } 
    dataSheet.setConditionalFormatRules(rules);
}



function buildCondFormatRule(cellA1, solution, color, range) {
    var rule = app.newConditionalFormatRule()
                    .whenFormulaSatisfied("=EQ(" + cellA1 + "," + solution + ")")
                    .setBackground(color)
                    .setRanges([range])
                    .build();

    return rule;
}



function deleteAllConditionalFormatRules () {
    dataSheet.setConditionalFormatRules([]);
}



function setThickSolutionBorder(range) {
    range.setBorder(true, true, true, true, false, false, null, app.BorderStyle.SOLID_THICK);
}
