function readSettingsToJson(sheet) {
  var lastRow = sheet.getLastRow();  
  var json = Object();
  
  for(var rowIndex=1; rowIndex<=lastRow; rowIndex++) {
    var colStartIndex = 1;
    var rowNum = 1;
    var range = sheet.getRange(rowIndex, colStartIndex, rowNum, sheet.getLastColumn());
    var values = range.getValues();
    json[values[0][0]]=values[0][1]
  }  
  return json;
}
