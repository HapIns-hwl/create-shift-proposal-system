/**
 * 
 * 文字色取得
 */
function getFontColor(sheet, row, column) {
  var cell = sheet.getRange(row, column);
  return cell.getFontColor();
}


/**
 * 
 * 空欄行削除
 */
function deleteEmptyRows(newsheet) {

  // シートのデータ範囲を取得 (D列からAJ列)
  var startColumn = 4; // D列は4番目
  var endColumn = 36; // AJ列は36番目
  var dataRange = newsheet.getRange(1, startColumn, newsheet.getLastRow(), endColumn - startColumn + 1);
  var data = dataRange.getValues();

  // 空欄行を削除
  for (var i = data.length - 1; i >= 0; i--) {
    var isRowEmpty = data[i].every(function (cell) {
      return cell === '';
    });

    if (isRowEmpty) {
      newsheet.deleteRow(i + 1);
    }
  }
}


/**
 * 最新のデータを取得する関数
 * @param {SpreadsheetApp.Sheet} sheet - 対象のシート
 * @return {Array} 最新のデータ
 */
function getLatestSheetData(sheet) {
  SpreadsheetApp.flush();
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  // 最後の行と列のセルデータを取得
  var range = sheet.getRange(1, 1, lastRow, lastColumn);
  var cellData = range.getValues();

  return cellData;
  // return sheet.getDataRange().getValues();
}


/**
 * 
 * 日付形式を変更
 */
function formatDate(date) {

  var month = date.getMonth() + 1;
  var day = date.getDate();
  var year = date.getFullYear();

  var formattedMonth = month < 10 ? + month : month;
  var formattedDay = day < 10 ? + day : day;

  return year + '/' + formattedMonth + '/' + formattedDay;
}


/**
 * 
 * 公休日の場合
 */
function highlightWeekendsAndHolidays(year, month, daysInMonth, newSheet) {
  var lastRow = newSheet.getLastRow();
  var numCols = daysInMonth;
  var backgrounds = new Array(lastRow - 1).fill(null).map(() => new Array(numCols).fill(null));

  for (var i = 1; i <= daysInMonth; i++) {
    var dayOfW = new Date(year, month - 1, i).getDay();
    if (dayOfW === 6 || dayOfW === 0 || isHolidayOfJOrC(`${year}/${month}/${i}`, true)) {
      for (var row = 2; row <= lastRow; row++) {
        backgrounds[row - 2][i - 1] = '#C9DAF8';
      }
    }
  }

  var range = newSheet.getRange(2, 7, lastRow - 1, numCols);
  range.setBackgrounds(backgrounds);
}




