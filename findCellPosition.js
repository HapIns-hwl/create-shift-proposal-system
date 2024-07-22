// /**
//  * 
//  * 希望の勤務を与えるため社員名の行と日付の列を取得
//  */
// function findCellPosition(newSheet, empName, dateKey) {

//   var shiftsData = getLatestSheetData(newSheet);

//   var currentMonth = new Date().getMonth() + 1;
//   var currentDate = new Date();
//   var year = currentDate.getFullYear();
//   var daysInMonth = new Date(year, currentMonth, 0).getDate();

//   var nameRow = -1;
//   var dateColumn = -1;

//   for (var i = 2; i < shiftsData.length; i++) {
//     var rowData = shiftsData[i];
//     var name = rowData[3];

//     if (name === empName) {
//       nameRow = i + 1;
//       break;
//     }
//   }

//   for (var day = 0; day <= daysInMonth; day++) {
//     var dateColumnVal = newSheet.getRange(2, day + 7).getValue();
//     var formattedDate = formatDate(dateColumnVal);

//     if (formattedDate === dateKey) {
//       dateColumn = day + 7;
//       break;
//     }
//   }

//   return { row: nameRow, column: dateColumn };
// }

/**
 * 希望の勤務を与えるため社員名の行と日付の列を取得
 */
function findCellPosition(shiftsData, dateColumnMapping, empName, dateKey) {
  var nameRow = -1;
  var dateColumn = dateColumnMapping[dateKey] || -1;

  for (var i = 2; i < shiftsData.length; i++) {
    var rowData = shiftsData[i];
    var name = rowData[3];

    if (name === empName) {
      nameRow = i + 1;
      break;
    }
  }

  return { row: nameRow, column: dateColumn };
}

/**
 * シートの全ての日付と列のマッピングを取得
 */
function getDateColumnMapping(newSheet) {
  var currentMonth = new Date().getMonth() + 1;
  var currentDate = new Date();
  var year = currentDate.getFullYear();
  var daysInMonth = new Date(year, currentMonth, 0).getDate();

  var dateColumnMapping = {};
  for (var day = 0; day < daysInMonth; day++) {
    var dateColumnVal = newSheet.getRange(2, day + 7).getValue();
    var formattedDate = formatDate(dateColumnVal);

    dateColumnMapping[formattedDate] = day + 7;
  }

  return dateColumnMapping;
}