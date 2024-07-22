/**
 * 夜勤シフト割り振り
 */
function createNightShiftsProposal() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var employeeSheet = ss.getSheetByName('employee');
  var scheduleSheet = ss.getSheetByName('desired_shifts');
  var confirmedSheet = ss.getSheetByName('confirmed_shifts');
  var shiftRuleSheet = ss.getSheetByName('shift_rules');
  var newSheet = ss.getActiveSheet();

  // ユーザーに月を入力させる
  var ui = SpreadsheetApp.getUi();

  try {

  // 処理開始メッセージを表示
  ss.toast('処理中...', 'シフトを作成してます', -1);

  var maleCareworkerStartRow = PropertiesService.getScriptProperties().getProperty('maleCareworkerStartRow');
  var maleCareworkerEndRow = PropertiesService.getScriptProperties().getProperty('maleCareworkerEndRow');
  var femaleCareworkerStartRow = PropertiesService.getScriptProperties().getProperty('femaleCareworkerStartRow');
  var femaleCareworkerEndRow = PropertiesService.getScriptProperties().getProperty('femaleCareworkerEndRow');
  var professionsStartRow = PropertiesService.getScriptProperties().getProperty('professionsStartRow');
  var maleTherapistEndRow = PropertiesService.getScriptProperties().getProperty('maleTherapistEndRow');

  // schedule・employeeテーブルのデータの取得
  var employeeData = employeeSheet.getDataRange().getValues();
  var desiredShiftData = scheduleSheet.getDataRange().getValues();
  var confirmedData = confirmedSheet.getDataRange().getValues();

  // ActiveSheetがシフト案シートかどうか確認
  var sheetName = newSheet.getName();
  if (!sheetName.endsWith('月のシフト案')) {
    SpreadsheetApp.getUi().alert("このシートはシフト案シートではありません。");
    spreadsheet.toast('処理が完了しました', '処理完了', 3);
    return;
  }

  var pattern = /(\d{4})\/(\d{1,2})月/;

  // Extracting year and month using regular expression
  var match = sheetName.match(pattern);
  var year, month;

  if (match) {
    year = parseInt(match[1]);
    month = parseInt(match[2]);
  }


  var daysInMonth = new Date(year, month, 0).getDate();

  // 前月を取得
  var lastMonth = month - 1 > 0 ? month - 1 : 12;

  // 先月が12月の場合は前年になる
  if (lastMonth === 12) {
    year -= 1;
  }

  // 先月の最後日を取得
  var lastDayOfLastMonth = new Date(year, lastMonth, 0);
  var formattedLastDay = formatDate(lastDayOfLastMonth);

  var idsWithShift = [];
  for (var indx = 1; indx < confirmedData.length; indx++) {
    var row = confirmedData[indx];
    var shiftDate = new Date(row[2]);
    var shiftType = row[3];
    var pastShiftMonth = shiftDate.getMonth() + 1;
    var pastLastDay = formatDate(shiftDate);

    if (pastShiftMonth === lastMonth && formattedLastDay === pastLastDay) {
      if (shiftType === 'A' || shiftType === 'A13' || shiftType === 'A22') {
        idsWithShift.push(row[1]);
      }
    }
  }

  // 先月の最終日に条件を満たすシフトが見つかった場合、対応するIDの今月の1日に 'F' を配置
  for (var k = 0; k < idsWithShift.length; k++) {
    var currentID = idsWithShift[k];

    for (var j = 1; j < desiredShiftData.length; j++) {
      var desiredRow = desiredShiftData[j];
      var desiredShiftDate = new Date(row[2]);
      var desiredShiftMonth = desiredShiftDate.getMonth() + 1;

      var id = desiredRow[1];
      if (id === currentID && desiredShiftMonth === month) {
        scheduleSheet.getRange(j + 1, 4).setValue('F');
        break;
      }
    }
  }



  // Fetch data ranges only once
  var shiftsData = getLatestSheetData(newSheet);
  var emptyShifts = ["A,H,N,O"];

  var employeeInfo = {};
  var omittedEmployeeIDs = [];

  // Populate employeeInfo and omittedEmployeeIDs
  for (var k = 1; k < employeeData.length; k++) {
    var id = employeeData[k][0];
    var employeeName = employeeData[k][1];
    var gender = employeeData[k][2];
    var empRole = employeeData[k][3];
    var omit = employeeData[k][5];
    var displayOrder = employeeData[k][7];

    employeeInfo[id] = {
      name: employeeName,
      gender: gender,
      role: empRole,
      omit: omit,
      displayOrder: displayOrder
    };

    omittedEmployeeIDs.push(id);
  }

  // Processing scheduleData
  var maleNightShiftsResults = {};
  var femaleNightShiftsResults = {};
  var inputShift = {};
  var femaleInputShift = {};

  for (var day = 0; day < daysInMonth; day++) {
    var date = new Date(2024, month - 1, day + 1);
    var dateKey = date.getFullYear() + '/' + (date.getMonth() + 1) + '/' + date.getDate();

    if (!maleNightShiftsResults[dateKey]) {
      maleNightShiftsResults[dateKey] = { count: 0, maleNightShiftName: [] };
    }

    if (!femaleNightShiftsResults[dateKey]) {
      femaleNightShiftsResults[dateKey] = { femaleCount: 0, femaleNightShiftName: [] };
    }

    maleCareworkerStartRow = Math.round(maleCareworkerStartRow);
    professionsStartRow = Math.round(professionsStartRow);
    maleTherapistEndRow = Math.round(maleTherapistEndRow);

    for (var j = maleCareworkerStartRow; j < maleCareworkerEndRow; j++) {

      if (!shiftsData[j]) {
        continue;
      }
      var rowDataMale = shiftsData[j];
      var shiftType = rowDataMale[day + 6];
      var shiftTypeOrg = rowDataMale[day + 6];
      var employeeName = rowDataMale[3];


      var fontColor = getFontColor(newSheet, j, day + 7);

      if (!shiftType) {
        shiftType = emptyShifts[0];
      }

      countNigthShiftsMale(employeeName, shiftType, dateKey, maleNightShiftsResults, inputShift);
      getInputShift(employeeName, shiftTypeOrg, dateKey, inputShift, fontColor);
    }

    femaleCareworkerStartRow = Math.round(femaleCareworkerStartRow);
    femaleCareworkerEndRow = Math.round(femaleCareworkerEndRow);
    for (var j = femaleCareworkerStartRow; j < femaleCareworkerEndRow; j++) {

      var rowDataFemale = shiftsData[j];
      if (!rowDataFemale) {
        continue;
      }
      var shiftType = rowDataFemale[day + 6];
      var shiftTypeOrg = rowDataFemale[day + 6];
      var employeeName = rowDataFemale[3];

      var fontColor = getFontColor(newSheet, j, day + 7);

      if (!shiftType) {
        shiftType = emptyShifts[0];
      }

      countNigthShiftsFemale(employeeName, shiftType, dateKey, femaleNightShiftsResults, inputShift);
      getInputShift(employeeName, shiftTypeOrg, dateKey, femaleInputShift, fontColor);
    }
  }

  //過去データっ整理・シフト割り振り
  getMaleEmpNightShiftDataProcess(newSheet, shiftRuleSheet, employeeInfo, maleNightShiftsResults, confirmedData, inputShift, month);
  getFemaleEmpNightShiftDataProcess(newSheet, shiftRuleSheet, employeeInfo, femaleNightShiftsResults, confirmedData, femaleInputShift, month);

  ui.alert('夜勤シフトを作成完了しました。');
  ss.toast('処理が完了しました', '処理完了', 5);
  } catch (error) {
    // エラーメッセージを表示
    ss.toast('エラーが発生しました: ' + error.message, 'エラー', 5);
  }

}


/**
 * 
 * 男性夜勤勤務整理・データ配置
 */
function getMaleEmpNightShiftDataProcess(newSheet, shiftRuleSheet, employeeInfo, maleNightShiftsResults, confirmedData, inputShift, month) {
  // シートからMaxValueを取得
  var maxMaleNightShiftVal = shiftRuleSheet.getRange('I4').getValue();

  var tempShift = {};

  // 対象メンバーを抽出
  var targetEmployeeList = [];
  var substituteEmployeeList = {};
  var totalCountList = {};
  for (var id in employeeInfo) {
    if (employeeInfo[id].role === "ケアワーカー" && employeeInfo[id].gender === "男") {
      targetEmployeeList[id] = employeeInfo[id];
    } else if (employeeInfo[id].role === "その他専門職") {
      substituteEmployeeList[id] = employeeInfo[id];
    }
  }

  // 過去2ヶ月のA,Hカウント
  for (var id in targetEmployeeList) {
    var malePastACount = 0;
    var malePastHCount = 0;
    for (var k = 1; k < confirmedData.length; k++) {
      var pastShiftDate = new Date(confirmedData[k][2]);
      var pastShiftMonth = pastShiftDate.getMonth() + 1;
      var pastShiftType = confirmedData[k][3];

      // 現在の月と前月、前々月を取得
      var currentMonth = month;
      var lastMonth = currentMonth - 1 > 0 ? currentMonth - 1 : 12;
      var lastTwoMonths = lastMonth - 1 > 0 ? lastMonth - 1 : 12;

      if (String(confirmedData[k][1]) !== id) continue

      if ((pastShiftMonth === lastMonth || pastShiftMonth === lastTwoMonths) && (pastShiftType === 'A' || pastShiftType === 'A13' || pastShiftType === 'A22')) {
        malePastACount++;
      }
      if ((pastShiftMonth === lastMonth || pastShiftMonth === lastTwoMonths) && pastShiftType === 'H') {
        malePastHCount++;
      }
    }

    // 結果を格納
    totalCountList[targetEmployeeList[id].name] = 0;
    totalCountList[targetEmployeeList[id].name] = malePastACount + malePastHCount;
  }

  var remainingAllocations = 0;
  var lastDateKey = '';
  for (var dateKey in inputShift) {
    remainingAllocations = maxMaleNightShiftVal;
    tempShift[dateKey] = {};

    for (var employeeName in inputShift[dateKey]) {
      tempShift[dateKey][employeeName] = {};

      // 記入済みの確認
      var shiftType = inputShift[dateKey][employeeName];
      if (shiftType !== null) {
        tempShift[dateKey][employeeName] = shiftType;

        if (shiftType === 'A') {
          remainingAllocations--;
        }
      }

      // 前日の確認
      if (lastDateKey !== '') {
        var lastShiftType = tempShift[lastDateKey][employeeName];

        if (lastShiftType === 'A') {
          tempShift[dateKey][employeeName] = 'F';
        }
        if (lastShiftType === 'F') {
          tempShift[dateKey][employeeName] = '/';
        }
      }
    }

    if (remainingAllocations > 0) {
      // 候補者のソート+割り当て済み除外
      var sortedCandidates = {};
      for (var j = 0; j < maleNightShiftsResults[dateKey].maleNightShiftName.length; j++) {
        var candidatesName = maleNightShiftsResults[dateKey].maleNightShiftName[j];

        var tempShiftType = tempShift[dateKey][candidatesName];
        if (Object.keys(tempShiftType).length !== 0) continue;

        var totalCount = totalCountList[candidatesName];

        if (!sortedCandidates[totalCount]) {
          sortedCandidates[totalCount] = {};
        }
        sortedCandidates[totalCount][candidatesName] = {};
      }

      // ソート順にAを割り当て
      for (var i in sortedCandidates) {
        for (var sortedEmployeeName in sortedCandidates[i]) {
          tempShift[dateKey][sortedEmployeeName] = 'A';

          totalCountList[sortedEmployeeName]++;
          remainingAllocations--;

          if (remainingAllocations <= 0) break;
        }
        if (remainingAllocations <= 0) break;
      }
    }
    lastDateKey = dateKey;
  }

  // 描画
  var shiftsData = getLatestSheetData(newSheet);
  var dateColumnMapping = getDateColumnMapping(newSheet);

  for (var dateKey in tempShift) {
    for (var employeeName in tempShift[dateKey]) {
      var drawShiftType = tempShift[dateKey][employeeName];
      if (Object.keys(drawShiftType).length === 0) continue;

      var cellPosition = findCellPosition(shiftsData, dateColumnMapping, employeeName, dateKey);
      if (cellPosition.row !== -1 && cellPosition.column !== -1) {
        var cell = newSheet.getRange(cellPosition.row, cellPosition.column);
        cell.setValue(drawShiftType)
          .setFontColor('#000000')
          .setFontSize(10)
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');
      }
    }
  }
  // for (var dateKey in tempShift) {
  //   for (var employeeName in tempShift[dateKey]) {
  //     var drawShiftType = tempShift[dateKey][employeeName];
  //     if (Object.keys(drawShiftType).length === 0) continue;

  //     var cellPosition = findCellPosition(newSheet, employeeName, dateKey);
  //     if (cellPosition.row !== -1 && cellPosition.column !== -1) {
  //       var cell = newSheet.getRange(cellPosition.row, cellPosition.column);
  //       cell.setValue(drawShiftType).setFontColor('#000000').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
  //     }
  //   }
  // }

}


/**
 * 
 * 男性夜勤勤務整理・データ配置
 */
function getFemaleEmpNightShiftDataProcess(newSheet, shiftRuleSheet, employeeInfo, femaleNightShiftsResults, confirmedData, femaleInputShift, month) {
  var maxFemaleNightShiftVal = shiftRuleSheet.getRange('I6').getValue();

  var tempShift = {};

  // 対象メンバーを抽出
  var targetEmployeeList = [];
  var totalCountList = {};
  for (var id in employeeInfo) {
    if (employeeInfo[id].role === "ケアワーカー" && employeeInfo[id].gender === "女") {
      targetEmployeeList[id] = employeeInfo[id];
    }
  }

  // 過去2ヶ月のA,Hカウント
  for (var id in targetEmployeeList) {
    var femalePastACount = 0;
    var femalePastHCount = 0;
    for (var k = 1; k < confirmedData.length; k++) {
      var pastShiftDate = new Date(confirmedData[k][2]);
      var pastShiftMonth = pastShiftDate.getMonth() + 1;
      var pastShiftType = confirmedData[k][3];

      // 現在の月と前月、前々月を取得
      var currentMonth = month;
      var lastMonth = currentMonth - 1 > 0 ? currentMonth - 1 : 12;
      var lastTwoMonths = lastMonth - 1 > 0 ? lastMonth - 1 : 12;

      if (String(confirmedData[k][1]) !== id) continue

      if ((pastShiftMonth === lastMonth || pastShiftMonth === lastTwoMonths) && (pastShiftType === 'A' || pastShiftType === 'A13' || pastShiftType === 'A22')) {
        femalePastACount++;
      }
      if ((pastShiftMonth === lastMonth || pastShiftMonth === lastTwoMonths) && pastShiftType === 'H') {
        femalePastHCount++;
      }
    }

    // 結果を格納
    totalCountList[targetEmployeeList[id].name] = 0;
    totalCountList[targetEmployeeList[id].name] = femalePastACount + femalePastHCount;
  }

  var remainingAllocations = 0;
  var lastDateKey = '';
  for (var dateKey in femaleInputShift) {
    remainingAllocations = maxFemaleNightShiftVal;
    tempShift[dateKey] = {};

    for (var employeeName in femaleInputShift[dateKey]) {
      tempShift[dateKey][employeeName] = {};

      // 記入済みの確認
      var shiftType = femaleInputShift[dateKey][employeeName];
      // Logger.log("ID:" + dateKey + ' ' + employeeName + ' ' + shiftType);
      if (shiftType !== null) {
        tempShift[dateKey][employeeName] = shiftType;

        if (shiftType === 'A') {
          remainingAllocations--;
        }
      }

      // 前日の確認
      if (lastDateKey !== '') {
        var lastShiftType = tempShift[lastDateKey][employeeName];

        if (lastShiftType === 'A') {
          tempShift[dateKey][employeeName] = 'F';
        }
        if (lastShiftType === 'F') {
          tempShift[dateKey][employeeName] = '/';
        }
      }
    }

    if (remainingAllocations > 0) {
      // 候補者のソート+割り当て済み除外
      var sortedCandidates = {};
      for (var j = 0; j < femaleNightShiftsResults[dateKey].femaleNightShiftName.length; j++) {
        var candidatesName = femaleNightShiftsResults[dateKey].femaleNightShiftName[j];

        var tempShiftType = tempShift[dateKey][candidatesName];
        if (Object.keys(tempShiftType).length !== 0) continue;

        var totalCount = totalCountList[candidatesName];
        //  Logger.log(candidatesName + ' ' + totalCount );

        if (!sortedCandidates[totalCount]) {
          sortedCandidates[totalCount] = {};
        }
        sortedCandidates[totalCount][candidatesName] = {};
      }

      // ソート順にAを割り当て
      for (var i in sortedCandidates) {
        for (var sortedEmployeeName in sortedCandidates[i]) {
          tempShift[dateKey][sortedEmployeeName] = 'A';

          totalCountList[sortedEmployeeName]++;
          remainingAllocations--;

          if (remainingAllocations <= 0) break;
        }
        if (remainingAllocations <= 0) break;
      }
    }
    lastDateKey = dateKey;
  }

  // 描画
  // for (var dateKey in tempShift) {
  //   for (var employeeName in tempShift[dateKey]) {
  //     var drawShiftType = tempShift[dateKey][employeeName];
  //     if (Object.keys(drawShiftType).length === 0) continue;

  //     var cellPosition = findCellPosition(newSheet, employeeName, dateKey);
  //     Logger.log(cellPosition)
  //     if (cellPosition.row !== -1 && cellPosition.column !== -1) {
  //       var cell = newSheet.getRange(cellPosition.row, cellPosition.column);
  //       cell.setValue(drawShiftType).setFontColor('#000000').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
  //     }
  //   }
  // }

  var shiftsData = getLatestSheetData(newSheet);
  var dateColumnMapping = getDateColumnMapping(newSheet);

  for (var dateKey in tempShift) {
    for (var employeeName in tempShift[dateKey]) {
      var drawShiftType = tempShift[dateKey][employeeName];
      if (Object.keys(drawShiftType).length === 0) continue;

      var cellPosition = findCellPosition(shiftsData, dateColumnMapping, employeeName, dateKey);
      if (cellPosition.row !== -1 && cellPosition.column !== -1) {
        var cell = newSheet.getRange(cellPosition.row, cellPosition.column);
        cell.setValue(drawShiftType)
          .setFontColor('#000000')
          .setFontSize(10)
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');
      }
    }
  }

}




/**
 * 男性夜勤シフト処理
 */
function countNigthShiftsMale(employeeName, shiftType, dateKey, maleNightShiftsResults) {
  if (shiftType.includes('A') || shiftType.includes('A12') || shiftType.includes('A22')) {
    maleNightShiftsResults[dateKey].count++;
    if (maleNightShiftsResults[dateKey].maleNightShiftName.indexOf(employeeName) === -1) {
      maleNightShiftsResults[dateKey].maleNightShiftName.push(employeeName);
    }
  }
}


/**
 * 女性夜勤シフト処理
 */
function countNigthShiftsFemale(employeeName, shiftType, dateKey, femaleNightShiftsResults) {
  if (shiftType.includes('A') || shiftType.includes('A12') || shiftType.includes('A22')) {
    femaleNightShiftsResults[dateKey].femaleCount++;
    if (femaleNightShiftsResults[dateKey].femaleNightShiftName.indexOf(employeeName) === -1) {
      femaleNightShiftsResults[dateKey].femaleNightShiftName.push(employeeName);
    }

  }
}

