// Contribution by Matthew Pepper

function CreateNewBallot() {

var ss = SpreadsheetApp.getActiveSpreadsheet();           // Finds the active spreadSHEET
var templatesheet = ss.getSheetByName('Check In');        // Uses Check In as the template
var candidates = ss.getRange("H11").getValue();           // Finds the number of candidates
var open_positions = ss.getRange("H9").getValue();        // Finds the number of open positions
var name = ss.getRange("H13").getValue() + " Ballot";     // creates the name of the new sheet as POSTION Ballot
ss.insertSheet(name,1,{template: templatesheet});         // creates the new sheet for the ballot
var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // turns the new sheet into the active one

// Declare Global Constant
var TOTAL_ROW_START = 16;     // start of the precinct names
var TOTAL_ROW_END = sheet.getLastRow();       // end of the precinct names
var NEW_DATA_COL = 9;         // which column we want the new data to start on
var NEW_DATA_ROW = 15;        // which row we want the new data to start on

// Initial variables - given dummy data to start
var act_row = NEW_DATA_ROW;
var offset = 0;
var act_col = NEW_DATA_COL;
var data_cell = 0;
var sum_raw_votes = 0;
var linked_cell = "A1";
  
// hide extraneous info from the ballot
  sheet.hideRows(1,13);
  sheet.hideColumns(1);
  sheet.hideColumns(3,2);
  sheet.hideColumns(6,2);

  
// format the new ballot sheet
  var range = sheet.getRange('B15');
  range.copyFormatToRange(sheet, 2, 2, 14, 14);
  range.copyFormatToRange(sheet, 9, 9, 14, 14);
  range.copyFormatToRange(sheet, (act_col + candidates + 1), (act_col + candidates + 1), 14, 14);
  sheet.getRange('B14').setValue(name); // adds the title
  sheet.getRange('I14').setValue("Raw Votes");
  sheet.getRange(14, (act_col + candidates + 1)).setValue("Weighted Votes");
  sheet.getRange(TOTAL_ROW_END+2,2).setValue("Total");
  sheet.setColumnWidth(8,35);
  
// Populate the new ballot with candidates and totals formulas
  for (var i = 0; i < candidates; i++) { 
    sheet.getRange(act_row,act_col).setValue("Candidate " + (i+1));
    offset = act_col + candidates + 1;
    var sum_rows = sheet.getRange(TOTAL_ROW_START,offset,(TOTAL_ROW_END-TOTAL_ROW_START+1),1).getA1Notation();
    sheet.getRange((TOTAL_ROW_END+2),offset).setFormula("=SUM("+sum_rows+")");
    linked_cell = sheet.getRange(act_row,act_col).getA1Notation();
    sheet.getRange(act_row,offset).setFormula("=" + linked_cell); // link the names of the weighted vote to raw vote
    sheet.autoResizeColumn(act_col);
    sheet.autoResizeColumn(offset);
    act_col++;
  } 

act_col = NEW_DATA_COL;  // reset active column for the new loops
  
// Populate the new ballot with formulas for calculating total vote
  for (var i = TOTAL_ROW_START; i < (TOTAL_ROW_END + 1); i++) { // loop for rows 
    sum_raw_votes = sheet.getRange(i, act_col, 1, candidates).getA1Notation();
    for (var j = 0; j < candidates; j++) { // loop for columns
      sheet.getRange(i,act_col).setValue(0);
      offset = act_col + candidates + 1;
      data_cell = sheet.getRange(i,act_col).getA1Notation();
      sheet.getRange(i,offset).setFormula("=IF(SUM("+sum_raw_votes+")=0,0,PRODUCT(E"+i+",DIVIDE("+data_cell+",SUM("+sum_raw_votes+")),"+open_positions +"))");
      act_col++;
    }
    act_col = NEW_DATA_COL;
  }

  return 

}
