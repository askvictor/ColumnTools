var dialogMsg = ""

function getDialogMsg(){
  return dialogMsg
}

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Convert Columns to Named Ranges', 'NameRangesCols')
    .addItem('Convert Rows to Named Ranges', 'NameRangesRows')
    .addItem('Conditionally Format selection one column at a time', 'ColourCols')
    .addItem('Conditionally Format selection one row at a time', 'ColourRows')
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

//function menuLink(){
//showURL("http://www.google.com")
//}
//
//function showURL(href){
//  var app = UiApp.createApplication().setHeight(500).setWidth(500);
//  app.setTitle("Open Link");
//  var link = app.createAnchor('Google.com ', href).setId("link");
//  app.add(link);  
//  var doc = SpreadsheetApp.getActive();
//  doc.show(app);
//  }

var OPTYPE = {COLS: 0, ROWS: 1}

function NameRangesCols(){
  return NameRanges(OPTYPE.COLS)
}
function NameRangesRows(){
  return NameRanges(OPTYPE.ROWS)
}

function NameRanges(type) {
  var ui = SpreadsheetApp.getUi(); 
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var rangeList = ss.getActiveRangeList()
  var ranges = rangeList.getRanges()
  if(type != OPTYPE.ROWS && type != OPTYPE.COLS){
    throw("Invalid type for NameRanges()")
  }
  if(rangeList == null || ranges.length == 0 || (ranges.length == 1 && ranges[0].getNumColumns() == 1 && ranges[0].getNumRows() == 1)){ //check if anything has been selected
    ui.alert(
     'Please select some cells to convert to named ranges.',
      ui.ButtonSet.OK)
    return
  } else { //something has been selected
    switch(type){
      case OPTYPE.COLS: //col
        dialogMsg = 'This will convert each selected Column to a Named Range, using the top cell as the name. Any named ranges with the same name will be over-written. Are you sure you want to continue?'
        break
      case OPTYPE.ROWS: //row
        dialogMsg = 'This will convert each selected Row to a Named Range, using the left cell as the name. Any named ranges with the same name will be over-written. Are you sure you want to continue?'
        break
    }
    
    var result = ui.alert( //check the user wants to do this thing
      'Continue?',
      dialogMsg,
      ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {  //user wants to do this thing
      for (var i = 0; i < ranges.length; i++) {
        switch(type){
          case OPTYPE.COLS://col
            var num_blocks = ranges[i].getNumColumns()
            break
          case OPTYPE.ROWS://row
            var num_blocks = ranges[i].getNumRows()
            break
        }
        for(var j=0; j<num_blocks; j++) {
          switch(type){
            case OPTYPE.COLS://col
              var block = ranges[i].offset(1,j,ranges[i].getNumRows()-1, 1)
              var name = ranges[i].offset(0,j).getValue()

              break
            case OPTYPE.ROWS://row
              var block = ranges[i].offset(j,1,1,ranges[i].getNumColumns()-1)
              var name = ranges[i].offset(j,0).getValue()

              break
          }
          
          // clean up name to be a valid NamedRange name
          name = String(name)  //ensure it's a string not a number
            .replace(/\s/g,"_")  //replace spaces with _
            .replace(/[^0-9a-zA-Z_]/g,"")  //get rid of non alphanum or _ 
            .replace(/^([0-9])/,"_$1") // if starts with number, put underscore in front
            .replace(/^(true|false)/i,"_$1") //if starts with true/false, put underscore in front
          if(name == ""){  //skip over empty named columns
            continue;
          }
          ss.setNamedRange(name, block)
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
      //check if the range is wholly within a conditional formatting rule
      if(range.getColumn() >= rulesRanges[j].getColumn() && range.getLastColumn() <= rulesRanges[j].getLastColumn()) {
        if(range.getRow() >= rulesRanges[j].getRow() && range.getLastRow() <= rulesRanges[j].getLastRow()) {
          rulesForRange.push(condFmtRules[i])
        }
      }
    }
  }
  return rulesForRange
}
  

function ColourCols() {
  return ColourRanges(OPTYPE.COLS)
}
function ColourRows() {
  return ColourRanges(OPTYPE.ROWS)
}
function ColourRanges(type) {
  var ui = SpreadsheetApp.getUi(); 
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getActiveSheet()
  var conditionalFormatRules = sheet.getConditionalFormatRules();
  
  var rangeList = ss.getActiveRangeList()
  var ranges = rangeList.getRanges()
  
  if(rangeList == null || ranges.length == 0 || (ranges.length == 1 && ranges[0].getNumColumns() == 1 && ranges[0].getNumRows() == 1)){ //check if anything has been selected
    ui.alert(
      'Please select some cells to format.',
      ui.ButtonSet.OK)
    return
  } else { //something has been selected
    switch(type){
      case OPTYPE.COLS: //col
        dialogMsg = 'This will conditionally format the selection column-by-column based on the rules set for the first column. Are you sure you want to continue?'
        break
      case OPTYPE.ROWS: //row
        dialogMsg = 'This will conditionally format the selection row-by-row based on the rules set for the first row. Are you sure you want to continue?'
        break
    }
    var result = ui.alert( //check the user wants to do this thing
      'Continue?',
      dialogMsg,
      ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {  //user wants to do this thing
      switch(type){
        case OPTYPE.COLS:
          var rules = getCondFmtRulesForRange(ranges[0].offset(0,0,ranges[0].getNumRows(),1), sheet)
          break
        case OPTYPE.ROWS:
          var rules = getCondFmtRulesForRange(ranges[0].offset(0,0,1, ranges[0].getNumColumns()), sheet)
          break
      }
      
      if(rules.length == 0){
        ui.alert("No conditional formatting has been set for the first row/column of your selection. Please set conditional formatting for the first column of you selection then try again", ui.ButtonSet.OK)
        return
      }

      for (var i = 0; i < ranges.length; i++) {
        switch(type){
          case OPTYPE.COLS:
            var numrowscols = ranges[i].getNumColumns()
            break
          case OPTYPE.ROWS:
            var numrowscols = ranges[i].getNumRows()
            break
        }
        for(var j=0; j<numrowscols; j++) {
          if(i==0 && j==0){
            continue  //skip first row/column as that's our source of formatting; no need to apply it back to itself
          }
          switch(type){
            case OPTYPE.COLS:
              var range = ranges[i].offset(0,j,ranges[i].getNumRows(), 1)              
              break
            case OPTYPE.ROWS:
              var range = ranges[i].offset(j,0,1, ranges[i].getNumColumns())
              break
          }

          for(var k=0; k<rules.length; k++){
             var rulebuilder = rules[k].copy()
             conditionalFormatRules.push(rulebuilder.setRanges([range]).build());
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
  var  tempSheetName = 'ColumnTools Temporary Sorting Sheet'
  while(ss.getSheetByName(tempSheetName) != null){
    tempSheetName += Math.floor(Math.random() * 9).toString()
  }
  var tempsheet = ss.insertSheet(tempSheetName);

  range.copyTo(tempsheet.getRange("A1"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, true)
  var temprange = tempsheet.getRange(1,1,range.getNumColumns(), range.getNumRows())

  temprange.sort({column: sort_row, ascending: order})
  temprange.copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, true)
  
  ss.deleteSheet(tempsheet)
  ss.setActiveSheet(sheet, true)  // go back to sheet and selection where we started
}
