
/**
 * その他シフトシフト割り振り
 */
function createOtherShiftsProposal() {
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

  // get data from PropertiesService
  var cookingStartRow = PropertiesService.getScriptProperties().getProperty('cookingStartRow');
  var cookingEndRow = PropertiesService.getScriptProperties().getProperty('cookingEndRow');
  var officeStartRow = PropertiesService.getScriptProperties().getProperty('officeStartRow');
  var officeEndRow = PropertiesService.getScriptProperties().getProperty('officeEndRow');

  var nurseStartRow = PropertiesService.getScriptProperties().getProperty('nurseStartRow');
  var nurseEndRow = PropertiesService.getScriptProperties().getProperty('nurseEndRow');
  var generalAffairsStartRow = PropertiesService.getScriptProperties().getProperty('generalAffairsStartRow');
  var generalAffairsEndRow = PropertiesService.getScriptProperties().getProperty('generalAffairsEndRow');
  var janitorialStartRow = PropertiesService.getScriptProperties().getProperty('janitorialStartRow');
  var janitorialEndRow = PropertiesService.getScriptProperties().getProperty('janitorialEndRow');

  var managementStartRow = PropertiesService.getScriptProperties().getProperty('managementStartRow');
  var managementEndRow = PropertiesService.getScriptProperties().getProperty('managementEndRow');
  var othersProfessionsStartRow = PropertiesService.getScriptProperties().getProperty('othersProfessionsStartRow');
  var othersProfessionsEndRow = PropertiesService.getScriptProperties().getProperty('othersProfessionsEndRow');

  var maleTherapistStartRow = PropertiesService.getScriptProperties().getProperty('maleTherapistStartRow');
  var maleTherapistEndRow = PropertiesService.getScriptProperties().getProperty('maleTherapistEndRow');

  var femaleTherapistStartRow = PropertiesService.getScriptProperties().getProperty('femaleTherapistStartRow');
  var femaleTherapistFemaleEndRow = PropertiesService.getScriptProperties().getProperty('femaleTherapistFemaleEndRow');

  var professionsStartRow = PropertiesService.getScriptProperties().getProperty('professionsStartRow');
  var professionsEndRow = PropertiesService.getScriptProperties().getProperty('professionsEndRow');




  managementStartRow = Math.round(managementStartRow);
  officeStartRow = Math.round(officeStartRow);
  generalAffairsStartRow = Math.round(generalAffairsStartRow);
  maleTherapistStartRow = Math.round(maleTherapistStartRow);
  femaleTherapistStartRow = Math.round(femaleTherapistStartRow);
  professionsStartRow = Math.round(professionsStartRow);

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

  var maxOthersRoleVal = shiftRuleSheet.getRange('K8').getValue();

  var shiftTypeDataCooking = ["C,N,Y"];
  var shiftTypeDataOffice = ["N"];
  var shiftTypeDataOtherRoles = ["N"];
  var shiftTypeDataOtherProfessionals = ["R"];

  var inputShift = {};
  var inputShiftY = {};
  var cookingDataInputShift = {};
  var nursingDataInputShift = {};
  var generalDataInputShift = {};
  var janitoDataInputShift = {};

  var managementDataInputShift = {};
  var otherProfessionalsDataInputShift = {};

  var cookingRoleCResults = {};
  var cookingRoleYResults = {};
  var officeResults = {};
  var nurseRoleResults = {};
  var generalAffairsRoleResults = {};
  var janitorialRoleResults = {};

  var managementRoleResults = {};
  var otherProfessionalsRoleResults = {};

  var employeeInfo = {};

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
  }


  calculateAndDisplayHolidays(newSheet, confirmedData, employeeInfo, year, month, daysInMonth);

  var shiftsData = getLatestSheetData(newSheet);

  for (var day = 0; day < daysInMonth; day++) {
    var date = new Date(year, month - 1, day + 1);
    var dateKey = date.getFullYear() + '/' + (date.getMonth() + 1) + '/' + date.getDate();

    // 日付ごとに調理Cデータ格納
    if (!cookingRoleCResults[dateKey]) {
      cookingRoleCResults[dateKey] = { cCcount: 0, cCnames: [], cCPastCounts: {} };
    }

    // 日付ごとに調理Yデータ格納
    if (!cookingRoleYResults[dateKey]) {
      cookingRoleYResults[dateKey] = { cYcount: 0, cYnames: [], cYPastCounts: {} };
    }

    // 日付ごとに事務データ格納
    if (!officeResults[dateKey]) {
      officeResults[dateKey] = { offiCount: 0, offiNames: [], offiPastCounts: {} };
    }

    // 日付ごとに総務データ格納
    if (!generalAffairsRoleResults[dateKey]) {
      generalAffairsRoleResults[dateKey] = { generalCount: 0, generalNames: [], generalPastCounts: {} };
    }



    // 調理処理
    cookingStartRow = Math.round(cookingStartRow);
    for (var j = cookingStartRow; j < cookingEndRow; j++) {
      if (!shiftsData[j]) {
        continue;
      }
      var rowDataCooking = shiftsData[j];
      var shiftType = rowDataCooking[day + 6];
      var shiftTypeOrg = rowDataCooking[day + 6];
      var employeeName = rowDataCooking[3];

      var fontColor = getFontColor(newSheet, j, day + 7);

      if (!shiftType) {
        shiftType = shiftTypeDataCooking[0];
      }

      countCookingShift(employeeName, shiftType, dateKey, cookingRoleCResults, cookingRoleYResults);
      getInputShift(employeeName, shiftTypeOrg, dateKey, inputShift, fontColor);
    }

    // 事務処理
    for (var j = officeStartRow; j < officeEndRow; j++) {
      if (!shiftsData[j]) {
        continue;
      }
      var rowDataOffice = shiftsData[j];
      var shiftType = rowDataOffice[day + 6];
      var shiftTypeOrg = rowDataOffice[day + 6];
      var employeeName = rowDataOffice[3];

      var fontColor = getFontColor(newSheet, j, day + 7);

      if (!shiftType) {
        shiftType = shiftTypeDataOffice[0];
      }

      countOfficeShift(employeeName, shiftType, dateKey, officeResults);
      getInputShift(employeeName, shiftTypeOrg, dateKey, cookingDataInputShift, fontColor);
    }



    // 総務処理
    for (var j = generalAffairsStartRow; j < generalAffairsEndRow; j++) {
      var rowDataGeneral = shiftsData[j];
      if (!shiftsData[j]) {
        continue;
      }
      var shiftType = rowDataGeneral[day + 6];
      var shiftTypeOrg = rowDataGeneral[day + 6];
      var employeeName = rowDataGeneral[3];

      var fontColor = getFontColor(newSheet, j, day + 7);

      if (!shiftType) {
        shiftType = shiftTypeDataOtherRoles[0];
      }

      countGeneralShift(employeeName, shiftType, dateKey, generalAffairsRoleResults);
      getInputShift(employeeName, shiftTypeOrg, dateKey, generalDataInputShift, fontColor);
    }



  }

  // 総務過去データ整理・配置メソッド
  getGeneralEmpShiftDataProcess(newSheet, employeeInfo, generalAffairsRoleResults, confirmedData, generalDataInputShift, month, maxOthersRoleVal);

  // 事務過去データ整理・配置メソッド
  getOfficeEmpShiftDataProcess(newSheet, shiftRuleSheet, employeeInfo, officeResults, confirmedData, cookingDataInputShift, month);

  // 調理C過去データ整理・配置メソッド
  getCookingRoleCShiftDataProcess(newSheet, shiftRuleSheet, employeeInfo, cookingRoleCResults, cookingRoleYResults, confirmedData, inputShift, month);

  // 休みを配置
  var lastDataForCookingRole = getLatestSheetData(newSheet);
  for (var day = 0; day < daysInMonth; day++) {

    for (var j = cookingStartRow; j < cookingEndRow; j++) {
      var rowData = lastDataForCookingRole[j];
      var shiftType = rowData[day + 6];

      if (shiftType === "") {
        newSheet.getRange(j + 1, day + 7).setValue('/');
      }
    }
  }


  ui.alert('その他シフトを作成完了しました。');
  ss.toast('処理が完了しました', '処理完了', 5);
}





/**
 * 
 * 総務シフト勤務整理・データ配置
 */
function getGeneralEmpShiftDataProcess(newSheet, employeeInfo, generalAffairsRoleResults, confirmedData, generalDataInputShift, month, maxOthersRoleVal) {
  var tempShift = {};

  // 対象メンバーを抽出
  var targetEmployeeList = [];
  var totalCountList = {};
  for (var id in employeeInfo) {
    if (employeeInfo[id].role === "総務") {
      targetEmployeeList[id] = employeeInfo[id];
    }
  }

  // 過去2ヶ月のA,Hカウント
  for (var id in targetEmployeeList) {
    var pastGeneralCount = 0;
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
        pastGeneralCount++;
      }
    }

    // 結果を格納
    totalCountList[targetEmployeeList[id].name] = 0;
    totalCountList[targetEmployeeList[id].name] = pastGeneralCount;
  }

  var remainingAllocations = 0;
  for (var dateKey in generalDataInputShift) {
    remainingAllocations = maxOthersRoleVal;
    tempShift[dateKey] = {};

    for (var employeeName in generalDataInputShift[dateKey]) {
      tempShift[dateKey][employeeName] = {};

      // 記入済みの確認
      var shiftType = generalDataInputShift[dateKey][employeeName];
      if (shiftType !== null) {
        tempShift[dateKey][employeeName] = shiftType;

        if (shiftType === 'N') {
          remainingAllocations--;
        }
      }

    }


    if (remainingAllocations > 0) {
      // 候補者のソート+割り当て済み除外
      var sortedCandidates = {};
      for (var j = 0; j < generalAffairsRoleResults[dateKey].generalNames.length; j++) {
        var candidatesName = generalAffairsRoleResults[dateKey].generalNames[j];

        var tempShiftType = tempShift[dateKey][candidatesName];
        if (Object.keys(tempShiftType).length !== 0) continue;

        var totalCount = totalCountList[candidatesName];

        if (!sortedCandidates[totalCount]) {
          sortedCandidates[totalCount] = {};
        }
        sortedCandidates[totalCount][candidatesName] = {};
      }

      // ソート順にNを割り当て
      for (var i in sortedCandidates) {
        for (var sortedEmployeeName in sortedCandidates[i]) {
          tempShift[dateKey][sortedEmployeeName] = 'N';

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
}




/**
 * 
 * 事務シフト勤務整理・データ配置
 */
function getOfficeEmpShiftDataProcess(newSheet, shiftRuleSheet, employeeInfo, officeResults, confirmedData, cookingDataInputShift, month) {
  var maxofficeVal = shiftRuleSheet.getRange('J8').getValue();

  var tempShift = {};

  // 対象メンバーを抽出
  var targetEmployeeList = [];
  var totalCountList = {};
  for (var id in employeeInfo) {
    if (employeeInfo[id].role === "事務") {
      targetEmployeeList[id] = employeeInfo[id];
    }
  }

  // 過去2ヶ月のA,Hカウント
  for (var id in targetEmployeeList) {
    var pastoffiCount = 0;
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
        pastoffiCount++;
      }
    }

    // 結果を格納
    totalCountList[targetEmployeeList[id].name] = 0;
    totalCountList[targetEmployeeList[id].name] = pastoffiCount;
  }

  var remainingAllocations = 0;
  for (var dateKey in cookingDataInputShift) {
    remainingAllocations = maxofficeVal;
    tempShift[dateKey] = {};

    for (var employeeName in cookingDataInputShift[dateKey]) {
      tempShift[dateKey][employeeName] = {};

      // 記入済みの確認
      var shiftType = cookingDataInputShift[dateKey][employeeName];
      if (shiftType !== null) {
        tempShift[dateKey][employeeName] = shiftType;

        if (shiftType === 'N') {
          remainingAllocations--;
        }
      }

    }


    if (remainingAllocations > 0) {
      // 候補者のソート+割り当て済み除外
      var sortedCandidates = {};
      for (var j = 0; j < officeResults[dateKey].offiNames.length; j++) {
        var candidatesName = officeResults[dateKey].offiNames[j];

        var tempShiftType = tempShift[dateKey][candidatesName];
        if (Object.keys(tempShiftType).length !== 0) continue;

        var totalCount = totalCountList[candidatesName];
        // Logger.log(candidatesName + ' ' + totalCount);

        if (!sortedCandidates[totalCount]) {
          sortedCandidates[totalCount] = {};
        }
        sortedCandidates[totalCount][candidatesName] = {};
      }

      // ソート順にAを割り当て
      for (var i in sortedCandidates) {
        for (var sortedEmployeeName in sortedCandidates[i]) {
          tempShift[dateKey][sortedEmployeeName] = 'N';

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
        cell.setValue(drawShiftType).setFontColor('#000000').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
      }
    }
  }
}



/**
 * 
 * 調理Cシフト勤務整理・データ配置
 */
function getCookingRoleCShiftDataProcess(newSheet, shiftRuleSheet, employeeInfo, cookingRoleCResults, cookingRoleYResults, confirmedData, inputShift, month) {
  var maxCookingCVal = shiftRuleSheet.getRange('L8').getValue();
  var maxCookingYVal = shiftRuleSheet.getRange('M7').getValue();

  var tempShift = {};
  var tempShiftY = {};

  // 対象メンバーを抽出
  var targetEmployeeList = [];
  var totalCountList = {};
  for (var id in employeeInfo) {
    if (employeeInfo[id].role === "調理") {
      targetEmployeeList[id] = employeeInfo[id];
    }
  }

  // 過去2ヶ月のA,Hカウント
  for (var id in targetEmployeeList) {
    var pastCCount = 0;
    for (var k = 1; k < confirmedData.length; k++) {
      var pastShiftDate = new Date(confirmedData[k][2]);
      var pastShiftMonth = pastShiftDate.getMonth() + 1;
      var pastShiftType = confirmedData[k][3];

      // 現在の月と前月、前々月を取得
      var currentMonth = month;
      var lastMonth = currentMonth - 1 > 0 ? currentMonth - 1 : 12;
      var lastTwoMonths = lastMonth - 1 > 0 ? lastMonth - 1 : 12;

      if (String(confirmedData[k][1]) !== id) continue

      if ((pastShiftMonth === lastMonth || pastShiftMonth === lastTwoMonths) && (pastShiftType === 'C')) {
        pastCCount++;
      }
    }

    // 結果を格納
    totalCountList[targetEmployeeList[id].name] = 0;
    totalCountList[targetEmployeeList[id].name] = pastCCount;
  }

  var remainingAllocations = 0;
  var remainingAllocationsY = 0;

  for (var dateKey in inputShift) {
    remainingAllocations = maxCookingCVal;
    remainingAllocationsY = maxCookingYVal;

    tempShift[dateKey] = {};
    tempShiftY[dateKey] = {};

    for (var employeeName in inputShift[dateKey]) {
      tempShift[dateKey][employeeName] = {};
      tempShiftY[dateKey][employeeName] = {};

      // 記入済みの確認
      var shiftType = inputShift[dateKey][employeeName];
      // Logger.log("ID:" + dateKey + ' ' + employeeName + ' ' + shiftType);
      if (shiftType !== null) {
        tempShift[dateKey][employeeName] = shiftType;
        tempShiftY[dateKey][employeeName] = shiftType;

        if (shiftType === 'C') {
          remainingAllocations--;
        }
        if (shiftType === 'Y') {
          remainingAllocationsY--;
        }
      }

    }

    if (remainingAllocations > 0) {
      // 候補者のソート+割り当て済み除外
      var sortedCandidates = {};
      for (var j = 0; j < cookingRoleCResults[dateKey].cCnames.length; j++) {
        var candidatesName = cookingRoleCResults[dateKey].cCnames[j];

        var tempShiftType = tempShift[dateKey][candidatesName];
        if (Object.keys(tempShiftType).length !== 0) continue;

        var totalCount = totalCountList[candidatesName];

        if (!sortedCandidates[totalCount]) {
          sortedCandidates[totalCount] = {};
        }
        sortedCandidates[totalCount][candidatesName] = {};
      }

      // ソート順にCを割り当て
      for (var i in sortedCandidates) {
        for (var sortedEmployeeName in sortedCandidates[i]) {
          tempShift[dateKey][sortedEmployeeName] = 'C';

          totalCountList[sortedEmployeeName]++;
          remainingAllocations--;

          if (remainingAllocations <= 0) break;
        }
        if (remainingAllocations <= 0) break;
      }
    }


    if (remainingAllocationsY > 0) {
      // 候補者のソート+割り当て済み除外
      var sortedCandidates = {};
      for (var j = 0; j < cookingRoleYResults[dateKey].cYnames.length; j++) {
        var candidatesName = cookingRoleYResults[dateKey].cYnames[j];

        var tempShiftType = tempShiftY[dateKey][candidatesName];
        if (Object.keys(tempShiftType).length !== 0) continue;

        var totalCount = totalCountList[candidatesName];

        if (!sortedCandidates[totalCount]) {
          sortedCandidates[totalCount] = {};
        }
        sortedCandidates[totalCount][candidatesName] = {};
      }

      // ソート順にYを割り当て
      for (var i in sortedCandidates) {
        for (var sortedEmployeeName in sortedCandidates[i]) {
          tempShiftY[dateKey][sortedEmployeeName] = 'Y';

          totalCountList[sortedEmployeeName]++;
          remainingAllocationsY--;

          if (remainingAllocationsY <= 0) break;
        }
        if (remainingAllocationsY <= 0) break;
      }
    }

  }

  var shiftsData = getLatestSheetData(newSheet);
  var dateColumnMapping = getDateColumnMapping(newSheet);

  // 描画
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

  // 描画
  for (var dateKey in tempShiftY) {
    for (var employeeName in tempShiftY[dateKey]) {
      var drawShiftType = tempShiftY[dateKey][employeeName];
      if (Object.keys(drawShiftType).length === 0) continue;

      var cellPosition = findCellPosition(shiftsData, dateColumnMapping, employeeName, dateKey);
      if (cellPosition.row !== -1 && cellPosition.column !== -1) {
        var cell = newSheet.getRange(cellPosition.row, cellPosition.column);
        cell.setValue(drawShiftType).setFontColor('#000000');
      }
    }
  }
}





/**
 * 休日計算
 */
function calculateAndDisplayHolidays(newSheet, confirmedData, employeeInfo, year, month, daysInMonth) {

  // 前月を取得
  var lastMonth = month - 1 > 0 ? month - 1 : 12;

  // 先月が12月の場合は前年になる
  if (lastMonth === 12) {
    year -= 1;
  }

  // 先月の最後日を取得
  var lastDayOfLastMonth = new Date(year, lastMonth, 0);
  var formattedLastDay = formatDate(lastDayOfLastMonth);

  var idCounts = {};

  // 対象メンバーを抽出
  var targetEmployeeList = [];
  var name = {};

  for (var id in employeeInfo) {
    if (employeeInfo[id].role === "総務" || employeeInfo[id].role === "事務") {
      targetEmployeeList[id] = employeeInfo[id];
      name[id] = employeeInfo[id].name;
      idCounts[id] = { count: 0, shiftFound: false };
    }
  }


  // 先月の休み計算
  for (var id in targetEmployeeList) {

    for (var k = 1; k < confirmedData.length; k++) {

      var row = confirmedData[k];
      var empId = row[1];
      var shiftDate = new Date(row[2]);
      var shiftType = row[3];

      if (String(empId) !== id) continue;

      // 先月の日付を確認
      if (shiftDate.getMonth() + 1 === lastMonth && shiftDate.getFullYear() === year) {
        // 「/」シフトが見つかった場合
        if (shiftType === '/') {
          idCounts[empId].shiftFound = true;
          idCounts[empId].count = Math.abs(Math.floor((shiftDate - lastDayOfLastMonth) / (1000 * 60 * 60 * 24)));
        } else if (!idCounts[empId].shiftFound) {
          // 「/」シフトが見つかっていない場合、日数をカウント
          idCounts[empId].count++;

        }
      }
    }
  }

  var holidays = [];


  // 当月の休み配置
  for (var day = 1; day <= daysInMonth; day++) {
    var maxCountId = null;
    var maxCount = -1;

    // 最もカウントが多い人を見つける
    for (var id in idCounts) {
      if (idCounts[id].count > maxCount) {
        maxCount = idCounts[id].count;
        maxCountId = id;
      }
    }


    if (maxCountId !== null) {
      // 休みを配置
      holidays.push({
        day: day,
        empId: maxCountId,
        name: name[maxCountId],
        shiftType: '/'
      });

      // カウントをリセット
      idCounts[maxCountId].count = 0;
      idCounts[maxCountId].shiftFound = true;

      // 他の人のカウントを増加
      for (var id in idCounts) {
        if (id !== maxCountId) {
          idCounts[id].count++;
        }
      }
    }
  }

  holidays.forEach(function (holiday) {
    var employeeName = holiday.name;
    var day = holiday.day;
    var date = new Date(year, month - 1, day);
    var dateKey = date.getFullYear() + '/' + (date.getMonth() + 1) + '/' + date.getDate();

    var drawShiftType = holiday.shiftType;
    var shiftsData = getLatestSheetData(newSheet);
    var dateColumnMapping = getDateColumnMapping(newSheet);

    var cellPosition = findCellPosition(shiftsData, dateColumnMapping, employeeName, dateKey);
    if (cellPosition.row !== -1 && cellPosition.column !== -1) {
      var cell = newSheet.getRange(cellPosition.row, cellPosition.column);
      cell.setValue(drawShiftType).setFontColor('#000000');
    }
  });

}


/**
 * 調理CYシフトの割り当て方
 */
function countCookingShift(employeeName, shiftType, dateKey, cookingRoleCResults, cookingRoleYResults) {
  if (shiftType.includes("C")) {
    cookingRoleCResults[dateKey].cCcount++;
    if (cookingRoleCResults[dateKey].cCnames.indexOf(employeeName) === -1) {
      cookingRoleCResults[dateKey].cCnames.push(employeeName);
    }
  }

  if (shiftType.includes("Y")) {
    cookingRoleYResults[dateKey].cYcount++;
    if (cookingRoleYResults[dateKey].cYnames.indexOf(employeeName) === -1) {
      cookingRoleYResults[dateKey].cYnames.push(employeeName);
    }
  }
}




/**
 * 事務シフトの割り当て方
 */
function countOfficeShift(employeeName, shiftType, dateKey, officeResults) {
  if (shiftType.includes('N')) {
    officeResults[dateKey].offiCount++;
    if (officeResults[dateKey].offiNames.indexOf(employeeName) === -1) {
      officeResults[dateKey].offiNames.push(employeeName);
    }
  }
}


/**
 * 総務シフトの割り当て方
 */
function countGeneralShift(employeeName, shiftType, dateKey, generalAffairsRoleResults) {
  if (shiftType.includes("N")) {
    generalAffairsRoleResults[dateKey].generalCount++;
    if (generalAffairsRoleResults[dateKey].generalNames.indexOf(employeeName) === -1) {
      generalAffairsRoleResults[dateKey].generalNames.push(employeeName);
    }
  }
}







