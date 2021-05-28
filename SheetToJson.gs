function readSettingsToJson(sheet) {
  var lastRow = sheet.getLastRow();
  var json = Object();

  for (var rowIndex = 1; rowIndex <= lastRow; rowIndex++) {
    var colStartIndex = 1;
    var rowNum = 1;
    var range = sheet.getRange(rowIndex, colStartIndex, rowNum, sheet.getLastColumn());
    var values = range.getValues();
    json[values[0][0]] = values[0][1]
  }
  return json;
}


function readTransToJson(sheet, _indexOfFirstEmptyColumnNameToStartRangeWith) {
  //_startFromFirstEmptyRowName - range should start where first empty column with _indexOfFirstEmptyColumnNameToStartRangeWith name is found
  var colStartIndex = 1;
  var rowNum = 1;
  var firstRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  var firstRowValues = firstRange.getValues();
  var titleColumns = firstRowValues[0];
  titleColumns.push('rowIndex')

  // after the second line(data)
  var lastRow = sheet.getLastRow();
  var rowValues = [];



  var rangeForFilter = sheet.getRange(2, _indexOfFirstEmptyColumnNameToStartRangeWith, sheet.getLastRow() - 1, 1)
  var range = sheet.getRange(2, sheet.getLastColumn(), sheet.getLastRow() - 1, 1)

  var values = rangeForFilter.getValues()


  for (var rowArrIdx = 0; rowArrIdx < rangeForFilter.getLastRow() - 1; rowArrIdx++) {
    if (values[rowArrIdx][0] == '') {
      var rowRangeIdx = rowArrIdx + 2 //1 for array (arrays start from 0, sheets indexes from 1) then +1 for eliminating heading row
      var rowValuesToPush = sheet.getRange(rowRangeIdx, 1, 1, range.getLastColumn()).getValues()
      rowValuesToPush[0].push(rowRangeIdx)
      rowValues.push(rowValuesToPush[0]);
    }
  }


  // create json
  var jsonMap = Object();
  for (var i = 0; i < rowValues.length; i++) {
    var line = rowValues[i];
    var json = new Object();
    for (var j = 0; j < titleColumns.length; j++) {
      json[titleColumns[j]] = line[j];
    }

    jsonMap[i] = json;
  }
  return jsonMap;
}
