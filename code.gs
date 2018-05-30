function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Convert Columns to Named Ranges', 'NameRanges')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function NameRanges() {
  var ui = SpreadsheetApp.getUi(); 
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var rl = ss.getActiveRangeList()
  rs = rl.getRanges()
  
  if(rl == null || rs.length == 0 || (rs.length == 1 && rs[0].getNumColumns() == 1 && rs[0].getNumRows() == 1)){ //check if anything has been selected
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
      for (var i = 0; i < rs.length; i++) {
        for(var j=0; j<rs[i].getNumColumns(); j++) {
          var col = rs[i].offset(0,j,rs[i].getNumRows(), 1)
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
