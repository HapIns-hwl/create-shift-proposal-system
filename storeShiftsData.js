function storeShiftsDataTo() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = spreadsheet.getActiveSheet();
  var employeeSheet = spreadsheet.getSheetByName('employee');
  var targetTableSheet = spreadsheet.getSheetByName('confirmed_shifts');

  spreadsheet.toast('処理中...', 'シフトを作成してます', -1);

  // ActiveSheetがシフト案シートかどうか確認
  var sheetName = activeSheet.getName();
  if (!sheetName.endsWith('月のシフト案')) {
    SpreadsheetApp.getUi().alert("このシートはシフト案シートではありません。");
    spreadsheet.toast('処理が完了しました', '処理完了', 3);
    return;
  }

  var monthMatch = sheetName.match(/^\d+\/(\d+)月/);
  if (!monthMatch) {
    SpreadsheetApp.getUi().alert("シート名から月を解析できませんでした。正しいフォーマットか確認してください。");
    spreadsheet.toast('処理が完了しました', '処理完了', 3);
    return;
  }

  var month = parseInt(monthMatch[1], 10);
  var year = new Date().getFullYear();
  var confirmedDataVals = activeSheet.getDataRange().getValues();
  var employeeData = employeeSheet.getDataRange().getValues();
  var daysInMonth = new Date(year, month, 0).getDate();

  // 対象月データの既存確認
  var startDate = new Date(year, month - 1, 1);
  var endDate = new Date(year, month - 1, daysInMonth);
  
  var lastRow = targetTableSheet.getLastRow();
  var existingData = lastRow > 1 ? targetTableSheet.getRange(2, 3, targetTableSheet.getLastRow() - 1, 1).getValues() : [];
  var monthExists = existingData.some(function (row) {
    var date = new Date(row[0]);
    return date >= startDate && date <= endDate;
  });

  if (monthExists) {
    SpreadsheetApp.getUi().alert(sheetName + "のデータは既に格納されています。");
    spreadsheet.toast('処理が完了しました', '処理完了', 3);
    return;
  }

  // var lastRow = targetTableSheet.getLastRow();
  var rowIndex = lastRow > 1 ? parseInt(targetTableSheet.getRange(lastRow, 1).getValue().substring(1)) + 1 : 1;
  var dataToWrite = [];

  var employeeMap = {};
  for (var k = 0; k < employeeData.length; k++) {
    // 名前からIDへのmapping and Role information
    employeeMap[employeeData[k][1]] = {
      id: employeeData[k][0],
      role: employeeData[k][3]
    };
  }

  // Helper function to generate the prefixed row index
  function getPrefixedRowIndex(index) {
    var maxIndex = 9999999;
    var alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    var prefixIndex = Math.floor((index - 1) / maxIndex);
    var prefix = alphabet.charAt(prefixIndex % alphabet.length);
    var numericPart = (index - 1) % maxIndex + 1;
    return prefix + ('000000' + numericPart).slice(-7);
  }

  for (var i = 0; i < confirmedDataVals.length; i++) {
    var empName = confirmedDataVals[i][3];
    if (empName) {
      var empData = employeeMap[empName];
      if (empData) {
        for (var day = 1; day <= daysInMonth; day++) {
          var shiftType = confirmedDataVals[i][day + 5];
          var date = year + '/' + month + '/' + day;
          dataToWrite.push([getPrefixedRowIndex(rowIndex), empData.id, date, shiftType, empData.role]);
          rowIndex++;
        }
      }
    }
  }

  if (dataToWrite.length > 0) {
    targetTableSheet.getRange(lastRow + 1, 1, dataToWrite.length, 5).setValues(dataToWrite);
  }

  // シートの右上に「シフト格納済み」を表示
  var lastColumn = activeSheet.getLastColumn();
  activeSheet.getRange(1, lastColumn - 1).setValue('シフト格納済み');

  // 格納完了メッセージ
  SpreadsheetApp.getUi().alert(sheetName + "のシフトデータをテーブルに格納完了しました。");
  spreadsheet.toast('処理が完了しました', '処理完了', 5);
}



