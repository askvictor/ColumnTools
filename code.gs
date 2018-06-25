function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Convert Columns to Named Ranges', 'NameRanges')
    .addItem('Conditionally Format selection one column at a time', 'ColourColumns')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Sort Horizontally')
          .addItem('Sort Sheet Horizontally A-Z', 'SortSheetHorizontallyAZ')
          .addItem('Sort Sheet Horizontally Z-A', 'SortSheetHorizontallyZA')
          .addItem('Sort Range Horizontally A-Z', 'SortRangeHorizontallyAZ')
          .addItem('Sort Range Horizontally Z-A', 'SortRangeHorizontallyZA')
          )
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function NameRanges() {
  var ui = SpreadsheetApp.getUi(); 
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var rangeList = ss.getActiveRangeList()
  var ranges = rangeList.getRanges()
  
  if(rangeList == null || ranges.length == 0 || (ranges.length == 1 && ranges[0].getNumColumns() == 1 && ranges[0].getNumRows() == 1)){ //check if anything has been selected
    ui.alert(
     'Please select one more more columns to convert to named ranges.',
      ui.ButtonSet.OK)
    return
  } else { //something has been selected
    var result = ui.alert( //check the user wants to do this thing
      'Continue?',
      'This will convert each selected Column to a Named Range, using the top cell as the name. Any named ranges with the same name will be over-written. Are you sure you want to continue?',
      ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {  //user wants to do this thing
      for (var i = 0; i < ranges.length; i++) {
        for(var j=0; j<ranges[i].getNumColumns(); j++) {
          var col = ranges[i].offset(0,j,ranges[i].getNumRows(), 1)
          var name = col.getValue()
          // clean up name to be a valid NamedRange name
          name = String(name)  //ensure it's a string not a number
            .replace(/\s/g,"_")  //replace spaces with _
            .replace(/[^0-9a-zA-Z_]/g,"")  //get rid of non alphanum or _ 
            .replace(/^([0-9])/,"_$1") // if starts with number, put underscore in front
            .replace(/^(true|false)/i,"_$1") //if starts with true/false, put underscore in front
          if(name == ""){  //skip over empty named columns
            continue;
          }
          ss.setNamedRange(name, col)
        }
      }
    }
  }
}

//find the conditional formatting rules applied to a given range
function getCondFmtRulesForRange(range, sheet){
  var condFmtRules = sheet.getConditionalFormatRules()
  var rulesForRange = []
  for (var i = 0; i < condFmtRules.length; i++) {
    var rulesRanges = condFmtRules[i].getRanges()
    for(var j=0; j<rulesRanges.length; j++){
      //check if the col is wholly within a conditional formatting rule
      if(range.getColumn() >= rulesRanges[j].getColumn() && range.getLastColumn() <= rulesRanges[j].getLastColumn()) {
        if(range.getRow() >= rulesRanges[j].getRow() && range.getLastRow() <= rulesRanges[j].getLastRow()) {
          rulesForRange.push(condFmtRules[i])
        }
      }
    }
  }
  return rulesForRange
}
  

function ColourColumns() {
  var ui = SpreadsheetApp.getUi(); 
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getActiveSheet()
  var conditionalFormatRules = sheet.getConditionalFormatRules();
  
  var rangeList = ss.getActiveRangeList()
  var ranges = rangeList.getRanges()
  
  if(rangeList == null || ranges.length == 0 || (ranges.length == 1 && ranges[0].getNumColumns() == 1 && ranges[0].getNumRows() == 1)){ //check if anything has been selected
    ui.alert(
      'Please select one more more columns to format.',
      ui.ButtonSet.OK)
    return
  } else { //something has been selected
    var result = ui.alert( //check the user wants to do this thing
      'Continue?',
      'This will conditionally format the selection column-by-column based on the rules set for the first column. Are you sure you want to continue?',
      ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {  //user wants to do this thing
      var rules = getCondFmtRulesForRange(ranges[0].offset(0,0,ranges[0].getNumRows(),1), sheet)
      if(rules.length == 0){
        ui.alert("No conditional formatting has been set for the first column of your selection. Please set conditional formatting for the first column of you selection then try again", ui.ButtonSet.OK)
        return
      }
            
      for (var i = 0; i < ranges.length; i++) {
        for(var j=0; j<ranges[i].getNumColumns(); j++) {
          if(i==0 && j==0){
            continue  //skip first column as that's our source of formatting; no need to apply it back to itself
          }
          var col = ranges[i].offset(0,j,ranges[i].getNumRows(), 1)

          for(var k=0; k<rules.length; k++){
             var rulebuilder = rules[k].copy()
             conditionalFormatRules.push(rulebuilder.setRanges([col]).build());
          }
        }
      }
      sheet.setConditionalFormatRules(conditionalFormatRules);
    }
  }
}

function SortSheetHorizontallyAZ(){
  SortHorizontally('sheet', true)

}
function SortRangeHorizontallyAZ(){
  SortHorizontally('range', true)
}

function SortSheetHorizontallyZA(){
  SortHorizontally('sheet', false)

}
function SortRangeHorizontallyZA(){
  SortHorizontally('range', false)
}


function SortHorizontally(type, order){ 
  var ui = SpreadsheetApp.getUi(); 
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getActiveSheet()
  var range = null
  switch(type) {
    case 'sheet':
      range = sheet.getDataRange()
      break;
    case 'range':
      range = sheet.getActiveRange()
      break;
    default:
      throw("error: unknown sort type")
  }

  var response = ui.prompt("Sort by which row?", "Specify the absolute row number from the left side of the spreadsheet", ui.ButtonSet.OK_CANCEL)
  if(response.getSelectedButton() == ui.Button.CANCEL) {
    return
  }
  
  if(response.getResponseText().match(/^\d+$/)){
    var sort_row = parseInt(response.getResponseText())
  }else{
    throw("invalid row entered")
  }
  sort_row = sort_row - range.getRow() + 1
  if(sort_row < 1 || sort_row > range.getNumRows()){
    throw("invalid row entered")
  }

  var tempsheet = ss.insertSheet('SortTemp');

  range.copyTo(tempsheet.getRange("A1"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, true)
  var temprange = tempsheet.getRange(1,1,range.getNumColumns(), range.getNumRows())

  temprange.sort({column: sort_row, ascending: order})
  temprange.copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, true)
  
  ss.deleteSheet(tempsheet)
  ss.setActiveSheet(sheet, true)  // go back to sheet and selection where we started
}
