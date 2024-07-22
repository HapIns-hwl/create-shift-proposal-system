
/**
 * 遅出シフト割り振り
 */
function createLateShiftsProposal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var newSheet = ss.getActiveSheet();
  var employeeSheet = ss.getSheetByName('employee');
  var scheduleSheet = ss.getSheetByName('desired_shifts');
  var confirmedSheet = ss.getSheetByName('confirmed_shifts');
  var shiftRuleSheet = ss.getSheetByName('shift_rules');

  // ユーザーに月を入力させる
  var ui = SpreadsheetApp.getUi();

  try {

    // 処理開始メッセージを表示
    ss.toast('処理中...', 'シフトを作成してます', -1);

    var maleCareworkerStartRow = PropertiesService.getScriptProperties().getProperty('maleCareworkerStartRow');
    var maleCareworkerEndRow = PropertiesService.getScriptProperties().getProperty('maleCareworkerEndRow');
    var femaleCareworkerStartRow = PropertiesService.getScriptProperties().getProperty('femaleCareworkerStartRow');
    var femaleCareworkerEndRow = PropertiesService.getScriptProperties().getProperty('femaleCareworkerEndRow');

    var maleTherapistStartRow = PropertiesService.getScriptProperties().getProperty('maleTherapistStartRow');
    var maleTherapistEndRow = PropertiesService.getScriptProperties().getProperty('maleTherapistEndRow');

    var femaleTherapistStartRow = PropertiesService.getScriptProperties().getProperty('femaleTherapistStartRow');
    var femaleTherapistFemaleEndRow = PropertiesService.getScriptProperties().getProperty('femaleTherapistFemaleEndRow');

    var professionsStartRow = PropertiesService.getScriptProperties().getProperty('professionsStartRow');
    var professionsEndRow = PropertiesService.getScriptProperties().getProperty('professionsEndRow');


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
      var year = parseInt(match[1]);
      var month = parseInt(match[2]);
    }

    var daysInMonth = new Date(year, month, 0).getDate();

    // 新しいシート上でのデータ整理
    var shiftsData = getLatestSheetData(newSheet);
    var emptyShifts = ["A,H,N,O"];

    var employeeInfo = {};
    var omittedEmployeeIDs = [];

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

    var results = {};
    var resultsProfession = {};
    var resultsTherapist = {};
    var femaleResults = {};
    var inputShift = {};
    var femaleInputShift = {};
    var femaleResultsTherapist = {};

    maleCareworkerStartRow = Math.round(maleCareworkerStartRow);
    femaleCareworkerStartRow = Math.round(femaleCareworkerStartRow);

    professionsStartRow = Math.round(professionsStartRow);
    professionsEndRow = Math.round(professionsEndRow);

    maleTherapistStartRow = Math.round(maleTherapistStartRow);
    maleTherapistEndRow = Math.round(maleTherapistEndRow);

    femaleTherapistStartRow = Math.round(femaleTherapistStartRow);
    femaleTherapistFemaleEndRow = Math.round(femaleTherapistFemaleEndRow);

    maleCareworkerStartRow = Math.round(maleCareworkerStartRow);
    femaleCareworkerStartRow = Math.round(femaleCareworkerStartRow);


    for (var day = 0; day < daysInMonth; day++) {
      var date = new Date(2024, month - 1, day + 1);
      var dateKey = date.getFullYear() + '/' + (date.getMonth() + 1) + '/' + date.getDate();

      // 日付ごとに男性遅出データ格納
      if (!results[dateKey]) {
        results[dateKey] = { count: 0, names: [] };
      }

      // 日付ごとに専門職遅出データ格納
      if (!resultsProfession[dateKey]) {
        resultsProfession[dateKey] = { count: 0, names: [] };
      }

      // 日付ごとに男性セラピスト遅出データ格納
      if (!resultsTherapist[dateKey]) {
        resultsTherapist[dateKey] = { count: 0, names: [] };
      }


      // 日付ごとに女性遅出データ格納
      if (!femaleResults[dateKey]) {
        femaleResults[dateKey] = { femaleCount: 0, femaleOnames: [] };
      }

      // 日付ごとに女性遅出データ格納
      if (!femaleResultsTherapist[dateKey]) {
        femaleResultsTherapist[dateKey] = { femaleCount: 0, femaleOnames: [] };
      }

      let rowData, shiftType, employeeName, shiftTypeOrg, fontColor;

      for (var j = maleCareworkerStartRow; j < maleCareworkerEndRow; j++) {

        if (!shiftsData[j]) continue;

        rowData = shiftsData[j];
        shiftType = rowData[day + 6] || emptyShifts[0];
        employeeName = rowData[3];
        shiftTypeOrg = rowData[day + 6];
        fontColor = getFontColor(newSheet, j, day + 7);

        countLateShiftsMale_careworker(employeeName, shiftType, dateKey, results);
        getInputShift(employeeName, shiftTypeOrg, dateKey, inputShift, fontColor);
      }

      for (var j = professionsStartRow; j < professionsEndRow; j++) {

        if (!shiftsData[j]) continue;

        rowData = shiftsData[j];
        shiftType = rowData[day + 6] || emptyShifts[0];
        employeeName = rowData[3];
        shiftTypeOrg = rowData[day + 6];
        fontColor = getFontColor(newSheet, j, day + 7);

        countNightShiftsMale_profession(employeeName, shiftType, dateKey, resultsProfession);
        getInputShift(employeeName, shiftTypeOrg, dateKey, inputShift, fontColor);

      }

      for (var j = maleTherapistStartRow; j < maleTherapistEndRow; j++) {

        if (!shiftsData[j]) continue;

        rowData = shiftsData[j];
        shiftType = rowData[day + 6] || emptyShifts[0];
        employeeName = rowData[3];
        shiftTypeOrg = rowData[day + 6];
        fontColor = getFontColor(newSheet, j, day + 7);

        countLateShiftsMale_therapist(employeeName, shiftType, dateKey, resultsTherapist);
        getInputShift(employeeName, shiftTypeOrg, dateKey, inputShift, fontColor);

      }



      //　女性遅出
      for (var j = femaleCareworkerStartRow; j < femaleCareworkerEndRow; j++) {

        if (!shiftsData[j]) continue;

        rowData = shiftsData[j];
        shiftType = rowData[day + 6] || emptyShifts[0];
        employeeName = rowData[3];
        shiftTypeOrg = rowData[day + 6];
        fontColor = getFontColor(newSheet, j, day + 7);

        countLateShiftsFemale_careworker(employeeName, shiftType, dateKey, femaleResults);
        getInputShift(employeeName, shiftTypeOrg, dateKey, femaleInputShift, fontColor);
      }

      for (var j = femaleTherapistStartRow; j < femaleTherapistFemaleEndRow; j++) {

        if (!shiftsData[j]) continue;

        rowData = shiftsData[j];
        shiftType = rowData[day + 6] || emptyShifts[0];
        employeeName = rowData[3];
        shiftTypeOrg = rowData[day + 6];
        fontColor = getFontColor(newSheet, j, day + 7);

        countLateShiftsFemale_therapist(employeeName, shiftType, dateKey, femaleResultsTherapist);
        getInputShift(employeeName, shiftTypeOrg, dateKey, femaleInputShift, fontColor);
      }

    }

    // 男性遅出ケアワーカー過去データ整理・配置メソッド
    getMaleEmpLateShiftDataProcess(newSheet, shiftRuleSheet, employeeInfo, results, resultsProfession, resultsTherapist, confirmedData, inputShift, month);

    // 休みを配置
    var lastDataForMale = getLatestSheetData(newSheet);
    for (var day = 0; day < daysInMonth; day++) {

      for (var j = maleCareworkerStartRow; j < maleCareworkerEndRow; j++) {
        var rowData = lastDataForMale[j];
        var shiftType = rowData[day + 6];

        if (shiftType === "") {
          newSheet.getRange(j + 1, day + 7).setValue('/');
        }
      }
    }

    // 女性遅出ケアワーカー過去データ整理・配置メソッド
    getFemaleEmpLateShiftDataProcess(newSheet, shiftRuleSheet, employeeInfo, femaleResults, femaleResultsTherapist, confirmedData, femaleInputShift, month);

    // 休みを配置
    var lastDataForFemale = getLatestSheetData(newSheet);
    for (var day = 0; day < daysInMonth; day++) {

      for (var j = femaleCareworkerStartRow; j < femaleCareworkerEndRow; j++) {
        var rowData = lastDataForFemale[j];
        var shiftType = rowData[day + 6];

        if (shiftType === "") {
          newSheet.getRange(j + 1, day + 7).setValue('/');
        }
      }
    }


    deleteEmptyRows(newSheet);

    ui.alert('遅出シフトを作成完了しました。');
    ss.toast('処理が完了しました', '処理完了', 5);
  } catch (error) {
    // エラーメッセージを表示
    Logger.log(error.message);
    ss.toast('エラーが発生しました: ' + error.message, 'エラー', 5);
  }

}



/**
 * 
 * 男性遅出勤務整理・データ配置
 */
function getMaleEmpLateShiftDataProcess(newSheet, shiftRuleSheet, employeeInfo, results, resultsProfession, resultsTherapist, confirmedData, inputShift, month) {
  // シートからMaxValueを取得
  var maxMaleLateShiftVal = shiftRuleSheet.getRange('H4').getValue();

  var tempShift = {};

  // 対象メンバーを抽出
  var targetEmployeeList = [];
  var totalCountList = {};
  for (var id in employeeInfo) {
    if (employeeInfo[id].role === "ケアワーカー" && employeeInfo[id].gender === "男") {
      targetEmployeeList[id] = employeeInfo[id];
    }
    if (employeeInfo[id].role === "専門職") {
      targetEmployeeList[id] = employeeInfo[id];
    }
    if (employeeInfo[id].role === "セラピスト" && employeeInfo[id].gender === "男") {
      targetEmployeeList[id] = employeeInfo[id];
    }
  }

  // 過去2ヶ月のA,Hカウント
  for (var id in targetEmployeeList) {
    var malePastOCount = 0;
    var malePastJCount = 0;
    for (var k = 1; k < confirmedData.length; k++) {
      var pastShiftDate = new Date(confirmedData[k][2]);
      var pastShiftMonth = pastShiftDate.getMonth() + 1;
      var pastShiftType = confirmedData[k][3];

      // 現在の月と前月、前々月を取得
      var currentMonth = month;
      var lastMonth = currentMonth - 1 > 0 ? currentMonth - 1 : 12;
      var lastTwoMonths = lastMonth - 1 > 0 ? lastMonth - 1 : 12;

      if (String(confirmedData[k][1]) !== id) continue

      if ((pastShiftMonth === lastMonth || pastShiftMonth === lastTwoMonths) && (pastShiftType === 'O')) {
        malePastOCount++;
      }
      if ((pastShiftMonth === lastMonth || pastShiftMonth === lastTwoMonths) && pastShiftType === 'J') {
        malePastJCount++;
      }
    }

    // 結果を格納
    totalCountList[targetEmployeeList[id].name] = 0;
    totalCountList[targetEmployeeList[id].name] = malePastOCount + malePastJCount;
  }

  var remainingAllocations = 0;
  for (var dateKey in inputShift) {
    remainingAllocations = maxMaleLateShiftVal;
    tempShift[dateKey] = {};

    for (var employeeName in inputShift[dateKey]) {
      tempShift[dateKey][employeeName] = {};

      // 記入済みの確認
      var shiftType = inputShift[dateKey][employeeName];
      // Logger.log("ID:" + dateKey + ' ' + employeeName + ' ' + shiftType);
      if (shiftType !== null) {
        tempShift[dateKey][employeeName] = shiftType;

        if (shiftType === 'O') {
          remainingAllocations--;
        }
      }

    }


    if (remainingAllocations > 0) {
      // 候補者のソート+割り当て済み除外
      var sortedCandidates = {};
      for (var j = 0; j < results[dateKey].names.length; j++) {
        var candidatesName = results[dateKey].names[j];

        var tempShiftType = tempShift[dateKey][candidatesName];
        if (Object.keys(tempShiftType).length !== 0) continue;

        var totalCount = totalCountList[candidatesName];
        // Logger.log(candidatesName + ' ' + totalCount);

        if (!sortedCandidates[totalCount]) {
          sortedCandidates[totalCount] = {};
        }
        sortedCandidates[totalCount][candidatesName] = {};
      }

      //　専門職
      for (var j = 0; j < resultsProfession[dateKey].names.length; j++) {
        var candidatesName = resultsProfession[dateKey].names[j];

        var tempShiftType = tempShift[dateKey][candidatesName];
        if (Object.keys(tempShiftType).length !== 0) continue;

        var totalCount = totalCountList[candidatesName] + 10000;
        // Logger.log(candidatesName + ' ' + totalCount);

        if (!sortedCandidates[totalCount]) {
          sortedCandidates[totalCount] = {};
        }
        sortedCandidates[totalCount][candidatesName] = {};
      }

      //　セラピスト
      for (var j = 0; j < resultsTherapist[dateKey].names.length; j++) {
        var candidatesName = resultsTherapist[dateKey].names[j];

        var tempShiftType = tempShift[dateKey][candidatesName];
        if (Object.keys(tempShiftType).length !== 0) continue;

        var totalCount = totalCountList[candidatesName] + 20000;

        if (!sortedCandidates[totalCount]) {
          sortedCandidates[totalCount] = {};
        }
        sortedCandidates[totalCount][candidatesName] = {};
      }

      // ソート順にAを割り当て
      for (var i in sortedCandidates) {
        for (var sortedEmployeeName in sortedCandidates[i]) {
          tempShift[dateKey][sortedEmployeeName] = 'O';

          totalCountList[sortedEmployeeName]++;
          remainingAllocations--;

          if (remainingAllocations <= 0) break;
        }
        if (remainingAllocations <= 0) break;
      }
    }
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
        cell.setValue(drawShiftType).setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontColor('#000000');
      }
    }
  }

}




/**
 * 
 * 女性遅出勤務整理・データ配置
 */
function getFemaleEmpLateShiftDataProcess(newSheet, shiftRuleSheet, employeeInfo, femaleResults, femaleResultsTherapist, confirmedData, femaleInputShift, month) {
  // シートからMaxValueを取得
  var maxFemaleLateShiftVal = shiftRuleSheet.getRange('H6').getValue();

  var tempShift = {};

  // 対象メンバーを抽出
  var targetEmployeeList = [];
  var totalCountList = {};
  for (var id in employeeInfo) {
    if (employeeInfo[id].role === "ケアワーカー" && employeeInfo[id].gender === "女") {
      targetEmployeeList[id] = employeeInfo[id];
    }
    if (employeeInfo[id].role === "セラピスト" && employeeInfo[id].gender === "女") {
      targetEmployeeList[id] = employeeInfo[id];
    }
  }

  // 過去2ヶ月のA,Hカウント
  for (var id in targetEmployeeList) {
    var malePastOCount = 0;
    var malePastJCount = 0;
    for (var k = 1; k < confirmedData.length; k++) {
      var pastShiftDate = new Date(confirmedData[k][2]);
      var pastShiftMonth = pastShiftDate.getMonth() + 1;
      var pastShiftType = confirmedData[k][3];

      // 現在の月と前月、前々月を取得
      var currentMonth = month;
      var lastMonth = currentMonth - 1 > 0 ? currentMonth - 1 : 12;
      var lastTwoMonths = lastMonth - 1 > 0 ? lastMonth - 1 : 12;

      if (String(confirmedData[k][1]) !== id) continue

      if ((pastShiftMonth === lastMonth || pastShiftMonth === lastTwoMonths) && (pastShiftType === 'O')) {
        malePastOCount++;
      }
      if ((pastShiftMonth === lastMonth || pastShiftMonth === lastTwoMonths) && pastShiftType === 'J') {
        malePastJCount++;
      }
    }

    // 結果を格納
    totalCountList[targetEmployeeList[id].name] = 0;
    totalCountList[targetEmployeeList[id].name] = malePastOCount + malePastJCount;
  }

  var remainingAllocations = 0;
  for (var dateKey in femaleInputShift) {
    remainingAllocations = maxFemaleLateShiftVal;
    tempShift[dateKey] = {};

    for (var employeeName in femaleInputShift[dateKey]) {
      tempShift[dateKey][employeeName] = {};

      // 記入済みの確認
      var shiftType = femaleInputShift[dateKey][employeeName];
      // Logger.log("ID:" + dateKey + ' ' + employeeName + ' ' + shiftType);
      if (shiftType !== null) {
        tempShift[dateKey][employeeName] = shiftType;

        if (shiftType === 'O') {
          remainingAllocations--;
        }
      }
    }

    if (remainingAllocations > 0) {
      // 候補者のソート+割り当て済み除外
      var sortedCandidates = {};
      for (var j = 0; j < femaleResults[dateKey].femaleOnames.length; j++) {
        var candidatesName = femaleResults[dateKey].femaleOnames[j];

        var tempShiftType = tempShift[dateKey][candidatesName];
        if (Object.keys(tempShiftType).length !== 0) continue;

        var totalCount = totalCountList[candidatesName];
        //  Logger.log(candidatesName + ' ' + totalCount );

        if (!sortedCandidates[totalCount]) {
          sortedCandidates[totalCount] = {};
        }
        sortedCandidates[totalCount][candidatesName] = {};
      }

      // セラピスト
      for (var j = 0; j < femaleResultsTherapist[dateKey].femaleOnames.length; j++) {
        var candidatesName = femaleResultsTherapist[dateKey].femaleOnames[j];

        var tempShiftType = tempShift[dateKey][candidatesName];
        if (Object.keys(tempShiftType).length !== 0) continue;

        var totalCount = totalCountList[candidatesName] + 10000;

        if (!sortedCandidates[totalCount]) {
          sortedCandidates[totalCount] = {};
        }
        sortedCandidates[totalCount][candidatesName] = {};
      }

      // ソート順にOを割り当て
      for (var i in sortedCandidates) {
        for (var sortedEmployeeName in sortedCandidates[i]) {
          tempShift[dateKey][sortedEmployeeName] = 'O';

          totalCountList[sortedEmployeeName]++;
          remainingAllocations--;

          if (remainingAllocations <= 0) break;
        }
        if (remainingAllocations <= 0) break;
      }
    }
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
        cell.setValue(drawShiftType).setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontColor('#000000');
      }
    }
  }

  return remainingAllocations;
}



/**
 * 男遅出（0）希望のケアワーカー、専門職、セラピストの割り当て方
 */
function countLateShiftsMale_careworker(employeeName, shiftType, dateKey, results) {
  if (shiftType.includes('O')) {
    results[dateKey].count++;
    if (results[dateKey].names.indexOf(employeeName) === -1) {
      results[dateKey].names.push(employeeName);
    }
  }
}


/**
 * 男遅出（0）希望のケアワーカー、専門職、セラピストの割り当て方
 */
function countNightShiftsMale_profession(employeeName, shiftType, dateKey, resultsProfession) {
  if (shiftType.includes('O')) {
    resultsProfession[dateKey].count++;
    if (resultsProfession[dateKey].names.indexOf(employeeName) === -1) {
      resultsProfession[dateKey].names.push(employeeName);
    }
  }
}



/**
 * 男遅出（0）希望のケアワーカー、専門職、セラピストの割り当て方
 */
function countLateShiftsMale_therapist(employeeName, shiftType, dateKey, resultsTherapist) {
  if (shiftType.includes('O')) {
    resultsTherapist[dateKey].count++;
    if (resultsTherapist[dateKey].names.indexOf(employeeName) === -1) {
      resultsTherapist[dateKey].names.push(employeeName);
    }
  }
}



/**
 * 女遅出（0）希望のケアワーカー、専門職、セラピストの割り当て方
 */
function countLateShiftsFemale_careworker(employeeName, shiftType, dateKey, femaleResults) {
  if (shiftType.includes('O')) {
    femaleResults[dateKey].femaleCount++;
    if (femaleResults[dateKey].femaleOnames.indexOf(employeeName) === -1) {
      femaleResults[dateKey].femaleOnames.push(employeeName);
    }
  }
}


/**
 * 女遅出（0）希望のケアワーカー、専門職、セラピストの割り当て方
 */
function countLateShiftsFemale_therapist(employeeName, shiftType, dateKey, femaleResultsTherapist) {
  if (shiftType.includes('O')) {
    femaleResultsTherapist[dateKey].femaleCount++;
    if (femaleResultsTherapist[dateKey].femaleOnames.indexOf(employeeName) === -1) {
      femaleResultsTherapist[dateKey].femaleOnames.push(employeeName);
    }
  }
}








