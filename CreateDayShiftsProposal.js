/**
 * 日勤・早出シフト割り振り
 */
function createDayShiftsProposal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var employeeSheet = ss.getSheetByName('employee');
  var scheduleSheet = ss.getSheetByName('desired_shifts');
  var confirmedSheet = ss.getSheetByName('confirmed_shifts');
  var shiftRuleSheet = ss.getSheetByName('shift_rules');
  var newSheet = ss.getActiveSheet();

  // ユーザーに月を入力させる
  var ui = SpreadsheetApp.getUi();

  // try {

    // 処理開始メッセージを表示
    ss.toast('処理中...', 'シフトを作成してます', -1);

    var maleCareworkerStartRow = PropertiesService.getScriptProperties().getProperty('maleCareworkerStartRow');
    var maleCareworkerEndRow = PropertiesService.getScriptProperties().getProperty('maleCareworkerEndRow');

    var maleTherapistStartRow = PropertiesService.getScriptProperties().getProperty('maleTherapistStartRow');
    var maleTherapistEndRow = PropertiesService.getScriptProperties().getProperty('maleTherapistEndRow');

    var femaleCareworkerStartRow = PropertiesService.getScriptProperties().getProperty('femaleCareworkerStartRow');
    var femaleCareworkerEndRow = PropertiesService.getScriptProperties().getProperty('femaleCareworkerEndRow');

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
    var scheduleData = scheduleSheet.getDataRange().getValues();
    var shiftsData = getLatestSheetData(newSheet);
    var emptyShifts = ["A,H,N,O"];

    /**** */

    var employeeInfo = {};

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
    }

    var resultsHN = {};
    var femaleDaysShiftsResults = {};
    var bothGenHResults = {};
    var inputDaysShift = {};
    var femaleInputDaysShift = {};
    var bothGenInputShift = {};
    var resultsProfession = {};
    var resultsTherapist = {}
    var femaleDaysShiftsResultsTherapist = {};


    maleCareworkerStartRow = Math.round(maleCareworkerStartRow);
    maleCareworkerEndRow = Math.round(maleCareworkerEndRow);

    femaleCareworkerStartRow = Math.round(femaleCareworkerStartRow);
    femaleCareworkerEndRow = Math.round(femaleCareworkerEndRow);

    professionsStartRow = Math.round(professionsStartRow);
    professionsEndRow = Math.round(professionsEndRow);

    maleTherapistStartRow = Math.round(maleTherapistStartRow);
    maleTherapistEndRow = Math.round(maleTherapistEndRow);

    femaleTherapistStartRow = Math.round(femaleTherapistStartRow);
    femaleTherapistFemaleEndRow = Math.round(femaleTherapistFemaleEndRow);

    for (var day = 0; day < daysInMonth; day++) {
      var date = new Date(year, month - 1, day + 1);
      var dateKey = date.getFullYear() + '/' + (date.getMonth() + 1) + '/' + date.getDate();

      // 日付ごとに男性日勤データ格納
      if (!resultsHN[dateKey]) {
        resultsHN[dateKey] = { count: 0, maleDaysShiftName: [] };
      }

      // 日付ごとに男性日勤データ格納
      if (!resultsProfession[dateKey]) {
        resultsProfession[dateKey] = { count: 0, maleDaysShiftName: [] };
      }

      // 日付ごとに男性日勤データ格納
      if (!resultsTherapist[dateKey]) {
        resultsTherapist[dateKey] = { count: 0, maleDaysShiftName: [] };
      }

      // 日付ごとに女性日勤データ格納
      if (!femaleDaysShiftsResults[dateKey]) {
        femaleDaysShiftsResults[dateKey] = { femaleCount: 0, femaleDaysShiftName: [] };
      }

      // 日付ごとに女性日勤データ格納
      if (!femaleDaysShiftsResultsTherapist[dateKey]) {
        femaleDaysShiftsResultsTherapist[dateKey] = { femaleCount: 0, femaleDaysShiftName: [] };
      }

      // 日付ごとに男性日勤データ格納
      let rowData, shiftType, employeeName, shiftTypeOrg, fontColor;

      // 男性日勤ケアワーカーシフト割り振り
      for (var j = maleCareworkerStartRow; j < maleCareworkerEndRow; j++) {

        if (!shiftsData[j]) continue;

        rowData = shiftsData[j];
        shiftType = rowData[day + 6] || emptyShifts[0];
        employeeName = rowData[3];
        shiftTypeOrg = rowData[day + 6];
        fontColor = getFontColor(newSheet, j, day + 7);

        countDayShiftsMale_careworker(employeeName, shiftType, dateKey, resultsHN);
        getInputShift(employeeName, shiftTypeOrg, dateKey, inputDaysShift, fontColor);
      }

      for (var j = professionsStartRow; j < professionsEndRow; j++) {

        if (!shiftsData[j]) continue;

        rowData = shiftsData[j];
        shiftType = rowData[day + 6] || emptyShifts[0];
        employeeName = rowData[3];
        shiftTypeOrg = rowData[day + 6];
        fontColor = getFontColor(newSheet, j, day + 7);

        countDayShiftsMale_profession(employeeName, shiftType, dateKey, resultsProfession);
        getInputShift(employeeName, shiftTypeOrg, dateKey, inputDaysShift, fontColor);

      }

      for (var j = maleTherapistStartRow; j < maleTherapistEndRow; j++) {

        if (!shiftsData[j]) continue;

        rowData = shiftsData[j];
        shiftType = rowData[day + 6] || emptyShifts[0];
        employeeName = rowData[3];
        shiftTypeOrg = rowData[day + 6];
        fontColor = getFontColor(newSheet, j, day + 7);

        countDayShiftsMale_therapist(employeeName, shiftType, dateKey, resultsTherapist);
        getInputShift(employeeName, shiftTypeOrg, dateKey, inputDaysShift, fontColor);
      }


      // 女性日勤シフト割り振り
      for (var j = femaleCareworkerStartRow; j < femaleCareworkerEndRow; j++) {
        if (!shiftsData[j]) continue;

        rowData = shiftsData[j];
        shiftType = rowData[day + 6] || emptyShifts[0];
        employeeName = rowData[3];
        shiftTypeOrg = rowData[day + 6];
        fontColor = getFontColor(newSheet, j, day + 7);

        countDayShiftsFemale_careworker(employeeName, shiftType, dateKey, femaleDaysShiftsResults);
        getInputShift(employeeName, shiftTypeOrg, dateKey, femaleInputDaysShift, fontColor);
      }

      for (var j = femaleTherapistStartRow; j < femaleTherapistFemaleEndRow; j++) {

        if (!shiftsData[j]) continue;

        rowData = shiftsData[j];
        shiftType = rowData[day + 6] || emptyShifts[0];
        employeeName = rowData[3];
        shiftTypeOrg = rowData[day + 6];
        fontColor = getFontColor(newSheet, j, day + 7);

        countDayShiftsFemale_therapist(employeeName, shiftType, dateKey, femaleDaysShiftsResultsTherapist);
        getInputShift(employeeName, shiftTypeOrg, dateKey, femaleInputDaysShift, fontColor);

      }

    }
    // 男性日勤ケアワーカー過去データ整理・シフト割り振り
    getMaleEmpDaysShiftDataProcess(newSheet, shiftRuleSheet, employeeInfo, resultsHN, resultsProfession, resultsTherapist, confirmedData, inputDaysShift, year, month);


    // 女性日勤ケアワーカー過去データ整理・シフト割り振り
    getFemaleEmpDaysShiftDataProcess(newSheet, shiftRuleSheet, employeeInfo, femaleDaysShiftsResults, femaleDaysShiftsResultsTherapist, confirmedData, femaleInputDaysShift, year, month);


    var lastShiftsData = getLatestSheetData(newSheet);

    for (var day = 0; day < daysInMonth; day++) {
      var date = new Date(year, month - 1, day + 1);
      var dateKey = date.getFullYear() + '/' + (date.getMonth() + 1) + '/' + date.getDate();

      // 日付ごとに男女早出（H）データ格納
      if (!bothGenHResults[dateKey]) {
        bothGenHResults[dateKey] = { bothGenCount: 0, bothGenNames: [], bothGenPastCounts: {} };
      }


      // 男女合わせ早出シフト割り振り
      for (var j = maleCareworkerStartRow; j < femaleCareworkerEndRow; j++) {

        var rowData = lastShiftsData[j];

        var shiftType = rowData[day + 6];
        var employeeName = rowData[3];
        var shiftTypeOrg = rowData[day + 6];

        var fontColor = getFontColor(newSheet, j, day + 7);

        countBothGenDayShift(employeeName, shiftType, dateKey, bothGenHResults);
        getInputShift(employeeName, shiftTypeOrg, dateKey, bothGenInputShift, fontColor);

      }
    }

    // 男女早出（H）過去データ整理・シフト割り振り
    getBothGenHShiftDataProcess(newSheet, shiftRuleSheet, employeeInfo, bothGenHResults, confirmedData, bothGenInputShift, month);

    // 早出色付け
    var lastDataForEarlyShift = getLatestSheetData(newSheet);
    for (var day = 0; day < daysInMonth; day++) {

      // 男性早出色付け
      for (var j = maleCareworkerStartRow; j < maleCareworkerEndRow; j++) {
        var rowData = lastDataForEarlyShift[j];
        var shiftType = rowData[day + 6];

        if (shiftType === "H") {
          newSheet.getRange(j + 1, day + 7).setBackground("#d1e5c9");
        }
      }

      // 女性早出色付け
      for (var j = femaleCareworkerStartRow; j < femaleCareworkerEndRow; j++) {
        var rowData = lastDataForEarlyShift[j];
        var shiftType = rowData[day + 6];

        if (shiftType === "H") {
          newSheet.getRange(j + 1, day + 7).setBackground("#f0c0c1");
        }
      }
    }

    ui.alert('日勤シフトを作成完了しました。');
    ss.toast('処理が完了しました', '処理完了', 5);
  // } catch (error) {
  //   // エラーメッセージを表示
  //   Logger.log(error.message);
  //   ss.toast('エラーが発生しました: ' + error.message, 'エラー', 5);
  // }

}



/**
 * 
 * 男性日勤（平日・休日）勤務整理・データ配置
 */
function getMaleEmpDaysShiftDataProcess(newSheet, shiftRuleSheet, employeeInfo, resultsHN, resultsProfession, resultsTherapist, confirmedData, inputDaysShift, year, month) {
  var maxMaleWeekdaysShiftVal = shiftRuleSheet.getRange('F4').getValue();
  var maxMaleholidayShiftVal = shiftRuleSheet.getRange('G4').getValue();

  var tempShiftDays = {};
  var holidaysCache = {};

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
    var malePastNCount = 0;
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

      if ((pastShiftMonth === lastMonth || pastShiftMonth === lastTwoMonths) && (pastShiftType === 'N')) {
        malePastNCount++;
      }
      if ((pastShiftMonth === lastMonth || pastShiftMonth === lastTwoMonths) && pastShiftType === 'J') {
        malePastJCount++;
      }
    }

    // 結果を格納
    totalCountList[targetEmployeeList[id].name] = 0;
    totalCountList[targetEmployeeList[id].name] = malePastNCount + malePastJCount;
  }

  var remainingAllocations = 0;

  for (var dateKey in inputDaysShift) {

    // Calculate or retrieve isHoliday from holidaysCache
    var isHoliday = holidaysCache[dateKey];
    if (isHoliday === undefined) {
      // Calculate isHoliday if it's not already cached
      var date = new Date(year, month - 1, parseInt(dateKey.split('/')[2]));
      var dayOfW = date.getDay();
      isHoliday = (dayOfW === 6 || dayOfW === 0 || isHolidayOfJOrC(dateKey, true));
      holidaysCache[dateKey] = isHoliday; // Cache the calculated isHoliday
    }

    if (isHoliday) {
      remainingAllocations = maxMaleholidayShiftVal;
    } else {
      remainingAllocations = maxMaleWeekdaysShiftVal;
    }

    tempShiftDays[dateKey] = {};

    for (var employeeName in inputDaysShift[dateKey]) {
      tempShiftDays[dateKey][employeeName] = {};

      // 記入済みの確認
      var shiftType = inputDaysShift[dateKey][employeeName];
      if (shiftType !== null) {
        tempShiftDays[dateKey][employeeName] = shiftType;

        if (shiftType === 'N') {
          remainingAllocations--;
        }
      }
    }

    if (remainingAllocations > 0) {
      // 候補者のソート+割り当て済み除外
      var sortedCandidates = {};
      for (var j = 0; j < resultsHN[dateKey].maleDaysShiftName.length; j++) {
        var candidatesName = resultsHN[dateKey].maleDaysShiftName[j];

        var tempShiftType = tempShiftDays[dateKey][candidatesName];
        if (Object.keys(tempShiftType).length !== 0) continue;

        var totalCount = totalCountList[candidatesName];

        if (!sortedCandidates[totalCount]) {
          sortedCandidates[totalCount] = {};
        }
        sortedCandidates[totalCount][candidatesName] = {};
      }

      //専門職
      for (var j = 0; j < resultsProfession[dateKey].maleDaysShiftName.length; j++) {
        var candidatesName = resultsProfession[dateKey].maleDaysShiftName[j];

        var tempShiftType = tempShiftDays[dateKey][candidatesName];
        if (Object.keys(tempShiftType).length !== 0) continue;

        var totalCount = totalCountList[candidatesName] + 10000;

        if (!sortedCandidates[totalCount]) {
          sortedCandidates[totalCount] = {};
        }
        sortedCandidates[totalCount][candidatesName] = {};
      }

      //セラピスト
      for (var j = 0; j < resultsTherapist[dateKey].maleDaysShiftName.length; j++) {
        var candidatesName = resultsTherapist[dateKey].maleDaysShiftName[j];

        var tempShiftType = tempShiftDays[dateKey][candidatesName];
        if (Object.keys(tempShiftType).length !== 0) continue;

        var totalCount = totalCountList[candidatesName] + 20000;

        if (!sortedCandidates[totalCount]) {
          sortedCandidates[totalCount] = {};
        }
        sortedCandidates[totalCount][candidatesName] = {};
      }

      // ソート順にNを割り当て
      for (var i in sortedCandidates) {
        for (var sortedEmployeeName in sortedCandidates[i]) {
          tempShiftDays[dateKey][sortedEmployeeName] = 'N';

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

  for (var dateKey in tempShiftDays) {
    for (var employeeName in tempShiftDays[dateKey]) {
      var drawShiftType = tempShiftDays[dateKey][employeeName];
      if (Object.keys(drawShiftType).length === 0) continue;

      var cellPosition = findCellPosition(shiftsData, dateColumnMapping, employeeName, dateKey);
      if (cellPosition.row !== -1 && cellPosition.column !== -1) {
        var cell = newSheet.getRange(cellPosition.row, cellPosition.column);
        cell.setValue(drawShiftType).setFontColor('#000000');

        // 背景色を緑に設定
        if (drawShiftType === 'N') {
          cell.setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#d1e5c9');
        }
      }
    }
  }
  // for (var dateKey in tempShiftDays) {
  //   for (var employeeName in tempShiftDays[dateKey]) {
  //     var drawShiftType = tempShiftDays[dateKey][employeeName];
  //     if (Object.keys(drawShiftType).length === 0) continue;

  //     var cellPosition = findCellPosition(newSheet, employeeName, dateKey);
  //     if (cellPosition.row !== -1 && cellPosition.column !== -1) {
  //       var cell = newSheet.getRange(cellPosition.row, cellPosition.column);
  //       cell.setValue(drawShiftType).setFontColor('#000000');

  //       // 背景色を緑に設定
  //       if (drawShiftType === 'N') {
  //         cell.setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#d1e5c9');
  //       }
  //     }
  //   }
  // }
}



/**
 * 
 * 男性日勤(平日・休日)勤務整理・データ配置
 */
function getFemaleEmpDaysShiftDataProcess(newSheet, shiftRuleSheet, employeeInfo, femaleDaysShiftsResults, femaleDaysShiftsResultsTherapist, confirmedData, femaleInputShift, year, month) {
  var maxFemaleWeekdaysShiftVal = shiftRuleSheet.getRange('F6').getValue();
  var maxFemaleholidayShiftVal = shiftRuleSheet.getRange('G6').getValue();

  var tempShiftDaysFemale = {};
  var holidaysCache = {};

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
    var malePastNCount = 0;
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

      if ((pastShiftMonth === lastMonth || pastShiftMonth === lastTwoMonths) && (pastShiftType === 'N')) {
        malePastNCount++;
      }
      if ((pastShiftMonth === lastMonth || pastShiftMonth === lastTwoMonths) && pastShiftType === 'J') {
        malePastJCount++;
      }
    }

    // 結果を格納
    totalCountList[targetEmployeeList[id].name] = 0;
    totalCountList[targetEmployeeList[id].name] = malePastNCount + malePastJCount;
  }

  var remainingAllocations = 0;
  for (var dateKey in femaleInputShift) {

    // 休日かの確認
    var isHoliday = holidaysCache[dateKey];
    if (isHoliday === undefined) {
      // Calculate isHoliday if it's not already cached
      var date = new Date(year, month - 1, parseInt(dateKey.split('/')[2]));
      var dayOfW = date.getDay();
      isHoliday = (dayOfW === 6 || dayOfW === 0 || isHolidayOfJOrC(dateKey, true));
      holidaysCache[dateKey] = isHoliday; // Cache the calculated isHoliday
    }

    // 休日であれば休日シフト数、そうでなければ平日シフト数を設定
    remainingAllocations = isHoliday ? maxFemaleholidayShiftVal : maxFemaleWeekdaysShiftVal;

    tempShiftDaysFemale[dateKey] = {};

    for (var employeeName in femaleInputShift[dateKey]) {
      tempShiftDaysFemale[dateKey][employeeName] = {};

      // 記入済みの確認
      var shiftType = femaleInputShift[dateKey][employeeName];
      if (shiftType !== null) {
        tempShiftDaysFemale[dateKey][employeeName] = shiftType;

        if (shiftType === 'N') {
          remainingAllocations--;
        }
      }
    }


    if (remainingAllocations > 0) {
      // 候補者のソート+割り当て済み除外
      var sortedCandidates = {};
      for (var j = 0; j < femaleDaysShiftsResults[dateKey].femaleDaysShiftName.length; j++) {
        var candidatesName = femaleDaysShiftsResults[dateKey].femaleDaysShiftName[j];

        var tempShiftType = tempShiftDaysFemale[dateKey][candidatesName];
        if (Object.keys(tempShiftType).length !== 0) continue;

        var totalCount = totalCountList[candidatesName];

        if (!sortedCandidates[totalCount]) {
          sortedCandidates[totalCount] = {};
        }
        sortedCandidates[totalCount][candidatesName] = {};
      }

      //　セラピスト
      for (var j = 0; j < femaleDaysShiftsResultsTherapist[dateKey].femaleDaysShiftName.length; j++) {
        var candidatesName = femaleDaysShiftsResultsTherapist[dateKey].femaleDaysShiftName[j];

        var tempShiftType = tempShiftDaysFemale[dateKey][candidatesName];
        if (Object.keys(tempShiftType).length !== 0) continue;

        var totalCount = totalCountList[candidatesName] + 10000;

        if (!sortedCandidates[totalCount]) {
          sortedCandidates[totalCount] = {};
        }
        sortedCandidates[totalCount][candidatesName] = {};
      }

      // ソート順にNを割り当て
      for (var i in sortedCandidates) {
        for (var sortedEmployeeName in sortedCandidates[i]) {
          tempShiftDaysFemale[dateKey][sortedEmployeeName] = 'N';

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

  for (var dateKey in tempShiftDaysFemale) {
    for (var employeeName in tempShiftDaysFemale[dateKey]) {
      var drawShiftType = tempShiftDaysFemale[dateKey][employeeName];
      if (Object.keys(drawShiftType).length === 0) continue;

      var cellPosition = findCellPosition(shiftsData, dateColumnMapping, employeeName, dateKey);
      if (cellPosition.row !== -1 && cellPosition.column !== -1) {
        var cell = newSheet.getRange(cellPosition.row, cellPosition.column);
        cell.setValue(drawShiftType).setFontColor('#000000');

        // 背景色を緑に設定
        if (drawShiftType === 'N') {
          cell.setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#f0c0c1');
        }
      }
    }
  }

  // for (var dateKey in tempShiftDaysFemale) {
  //   for (var employeeName in tempShiftDaysFemale[dateKey]) {
  //     var drawShiftType = tempShiftDaysFemale[dateKey][employeeName];
  //     if (Object.keys(drawShiftType).length === 0) continue;

  //     var cellPosition = findCellPosition(newSheet, employeeName, dateKey);
  //     if (cellPosition.row !== -1 && cellPosition.column !== -1) {
  //       var cell = newSheet.getRange(cellPosition.row, cellPosition.column);
  //       cell.setValue(drawShiftType).setFontColor('#000000');

  //       // 背景色を緑に設定
  //       if (drawShiftType === 'N') {
  //         cell.setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#f0c0c1');
  //       }
  //     }
  //   }
  // }

}



/**
 * 
 * 早出（H）勤務男女合わせて一名整理・データ配置
 */
function getBothGenHShiftDataProcess(newSheet, shiftRuleSheet, employeeInfo, bothGenHResults, confirmedData, bothGenInputShift, month) {
  // シートからMaxValueを取得
  var maxBothGenHVal = shiftRuleSheet.getRange('E8').getValue();

  var tempShift = {};

  // 対象メンバーを抽出
  var targetEmployeeList = [];
  var totalCountList = {};
  for (var id in employeeInfo) {
    if (employeeInfo[id].role === "ケアワーカー" && (employeeInfo[id].gender === "男" || employeeInfo[id].gender === "女")) {
      targetEmployeeList[id] = employeeInfo[id];
    }
  }

  // 過去2ヶ月のA,Hカウント
  for (var id in targetEmployeeList) {
    var bothGenACount = 0;
    var bothGenHCount = 0;
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
        bothGenACount++;
      }
      if ((pastShiftMonth === lastMonth || pastShiftMonth === lastTwoMonths) && pastShiftType === 'H') {
        bothGenHCount++;
      }
    }

    // 結果を格納
    totalCountList[targetEmployeeList[id].name] = 0;
    totalCountList[targetEmployeeList[id].name] = bothGenACount + bothGenHCount;
  }

  var remainingAllocations = 0;
  for (var dateKey in bothGenInputShift) {
    remainingAllocations = maxBothGenHVal;
    tempShift[dateKey] = {};

    for (var employeeName in bothGenInputShift[dateKey]) {
      tempShift[dateKey][employeeName] = {};

      // 記入済みの確認
      var shiftType = bothGenInputShift[dateKey][employeeName];
      if (shiftType !== null) {
        tempShift[dateKey][employeeName] = shiftType;

        if (shiftType === 'H') {
          remainingAllocations--;
        }
      }

    }

    if (remainingAllocations > 0) {
      // 候補者のソート+割り当て済み除外
      var sortedCandidates = {};
      for (var j = 0; j < bothGenHResults[dateKey].bothGenNames.length; j++) {
        var candidatesName = bothGenHResults[dateKey].bothGenNames[j];

        var tempShiftType = tempShift[dateKey][candidatesName];
        // if (Object.keys(tempShiftType).length !== 0) continue;

        var totalCount = totalCountList[candidatesName];

        if (!sortedCandidates[totalCount]) {
          sortedCandidates[totalCount] = {};
        }
        sortedCandidates[totalCount][candidatesName] = {};
      }

      // ソート順にHを割り当て
      for (var i in sortedCandidates) {
        for (var sortedEmployeeName in sortedCandidates[i]) {
          tempShift[dateKey][sortedEmployeeName] = 'H';

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
        cell.setValue(drawShiftType).setFontColor('#000000');

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
  //       cell.setValue(drawShiftType).setFontColor('#000000');

  //       // 背景色を緑に設定
  //       // if (drawShiftType === 'H') {
  //       //   cell.setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
  //       // }
  //     }
  //   }
  // }
}




/**
 * 男性日勤シフト割り振り
 */
function countDayShiftsMale_careworker(employeeName, shiftType, dateKey, resultsHN) {

  if (shiftType.includes('N')) {
    resultsHN[dateKey].count++;
    if (resultsHN[dateKey].maleDaysShiftName.indexOf(employeeName) === -1) {
      resultsHN[dateKey].maleDaysShiftName.push(employeeName);
    }
  }
}


/**
 * 男性日勤シフト割り振り
 */
function countDayShiftsMale_profession(employeeName, shiftType, dateKey, resultsProfession) {

  if (shiftType.includes('N')) {
    resultsProfession[dateKey].count++;
    if (resultsProfession[dateKey].maleDaysShiftName.indexOf(employeeName) === -1) {
      resultsProfession[dateKey].maleDaysShiftName.push(employeeName);
    }
  }
}


/**
 * 男性日勤シフト割り振り
 */
function countDayShiftsMale_therapist(employeeName, shiftType, dateKey, resultsTherapist) {

  if (shiftType.includes('N')) {
    resultsTherapist[dateKey].count++;
    if (resultsTherapist[dateKey].maleDaysShiftName.indexOf(employeeName) === -1) {
      resultsTherapist[dateKey].maleDaysShiftName.push(employeeName);
    }
  }
}



/**
 * 女性日勤シフト割り振り
 */
function countDayShiftsFemale_careworker(employeeName, shiftType, dateKey, femaleDaysShiftsResults) {
  if (shiftType.includes('N')) {
    femaleDaysShiftsResults[dateKey].femaleCount++;
    if (femaleDaysShiftsResults[dateKey].femaleDaysShiftName.indexOf(employeeName) === -1) {
      femaleDaysShiftsResults[dateKey].femaleDaysShiftName.push(employeeName);
    }
  }
}

/**
 * 女性日勤シフト割り振り
 */
function countDayShiftsFemale_therapist(employeeName, shiftType, dateKey, femaleDaysShiftsResultsTherapist) {
  if (shiftType.includes('N')) {
    femaleDaysShiftsResultsTherapist[dateKey].femaleCount++;
    if (femaleDaysShiftsResultsTherapist[dateKey].femaleDaysShiftName.indexOf(employeeName) === -1) {
      femaleDaysShiftsResultsTherapist[dateKey].femaleDaysShiftName.push(employeeName);
    }
  }
}

/**
 * 男女性早出シフト割り振り
 */
function countBothGenDayShift(employeeName, shiftType, dateKey, bothGenHResults) {
  if (shiftType === "N") {
    bothGenHResults[dateKey].bothGenCount++;
    if (bothGenHResults[dateKey].bothGenNames.indexOf(employeeName) === -1) {
      bothGenHResults[dateKey].bothGenNames.push(employeeName);
    }
  }
}






