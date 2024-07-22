let managementStartRow = 3;
let managementEndRow = -1;
let othersProfessionsStartRow = -1;
let othersProfessionsEndRow = -1;
let professionsStartRow = -1;
let professionsEndRow;
let maleCareworkerStartRow;
let maleCareworkerEndRow;
let maleTherapistStartRow;
let maleTherapistEndRow;
let femaleCareworkerStartRow;
let femaleCareworkerEndRow;
let femaleTherapistStartRow;
let femaleTherapistFemaleEndRow;
let partTimeStartRow = -1;
let partTimeEndRow = -1;
let nurseStartRow = -1;
let nurseEndRow = -1;
let cookingStartRow = -1;
let cookingEndRow = -1
let generalAffairsStartRow = -1;
let generalAffairsEndRow = -1;
let janitorialStartRow = -1;
let janitorialEndRow = -1;
let officeStartRow = -1;
let officeEndRow = -1;



/**
 * 原紙作成
 */
function createSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var employeeSheet = ss.getSheetByName('employee');
  var scheduleSheet = ss.getSheetByName('desired_shifts');
  var confirmedSheet = ss.getSheetByName('confirmed_shifts');
  var shiftRuleSheet = ss.getSheetByName('shift_rules');

  // ユーザーに月を入力させる
  var ui = SpreadsheetApp.getUi();

  var response = ui.prompt('シフト案シート作成', 'シフト案シートを作成する月を入力してください（例：2024/05）', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() != ui.Button.OK) {
    ss.toast('処理が完了しました', '処理完了', 3);
    return;
  }

  var input = response.getResponseText();
  var dateParts = input.split('/');
  if (dateParts.length != 2) {
    ui.alert('入力形式が正しくありません。例：2024/05');
    ss.toast('処理が完了しました', '処理完了', 3);
    return;
  }

  var year = parseInt(dateParts[0]);
  var month = parseInt(dateParts[1]);

  if (isNaN(year) || isNaN(month) || month < 1 || month > 12) {
    ui.alert('入力された年月が正しくありません。');
    ss.toast('処理が完了しました', '処理完了', 3);
    return;
  }

  try {

    // 処理開始メッセージを表示
    ss.toast('処理中...', 'シフトを作成してます', -1);

    // シート名を作成
    var sheetName = year + '/' + month + '月のシフト案';

    // 既存のシートがあるかどうか確認
    var existingSheet = ss.getSheetByName(sheetName);
    if (existingSheet) {
      ui.alert(sheetName + 'のシートは既に存在しています！');
      ss.toast('処理が完了しました', '処理完了', 3);
      return;
    }


    // schedule・employeeテーブルのデータの取得
    var employeeData = employeeSheet.getDataRange().getValues();
    var confirmedData = confirmedSheet.getDataRange().getValues();


    var newSheet = ss.insertSheet(sheetName);

    // merge the cells and set data
    var rangeToMergeYM = newSheet.getRange('G1:M1');
    rangeToMergeYM.merge();

    var monthHeader = year + '年' + month + '月';
    newSheet.getRange('G1')
      .setValue(monthHeader)
      .setFontWeight('bold')
      .setFontSize(15)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    var rangeToMergeTitle = newSheet.getRange('N1:AE1');
    rangeToMergeTitle.merge();
    newSheet.getRange('N1')
      .setValue('職　員　勤　務　表　')
      .setFontWeight('bold')
      .setFontSize(15)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // 名前
    var mergeNameCol = newSheet.getRange('D2:E3');
    mergeNameCol.merge();
    newSheet.getRange('D2').setHorizontalAlignment('center').setVerticalAlignment('middle');

    var daysInMonth = new Date(year, month, 0).getDate();

    var headers = [''];
    headers.push('');
    headers.push('');
    headers.push('名前');
    headers.push('');

    headers.push('日');
    for (var i = 1; i <= daysInMonth; i++) {
      headers.push(month + '/' + i);
    }
    headers.push('公\n休', '半\n休', '有\n休', '特\n休', '宿\n直', '遅\n出', '日\n勤');
    newSheet.appendRow(headers);
    var headerRange = newSheet.getRange(2, 1, 1, headers.length);
    headerRange
      .setValues([headers])
      .setFontSize(10)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    var getHeaderRange = newSheet.getRange(2, 1, 1, newSheet.getLastColumn());
    var getHeaders = getHeaderRange.getValues()[0];


    // 各休暇タイプの列インデックスを取得
    var leaveTypes = ['公\n休', '半\n休', '有\n休', '特\n休', '宿\n直', '遅\n出', '日\n勤'];
    var leaveTypeIndexes = [];
    leaveTypes.forEach(function (type) {
      leaveTypeIndexes.push(getHeaders.indexOf(type) + 1);
    });

    // 各休暇タイプごとに背景色をつける
    leaveTypeIndexes.forEach(function (index, i) {
      if (index > 0) {
        var getlastRow = newSheet.getLastRow();
        var leaveTypeRange = newSheet.getRange(2, index, getlastRow - 1);

        // セルに背景色をつける
        var color;
        switch (i) {
          case 0:
            color = '#f0c0c1';
            break;
          case 1:
            color = '#c5dbef';
            break;
          case 2:
            color = '#d1e5c9';
            break;
          case 3:
            color = '#ffffcc';
            break;
          case 4:
            color = '#d0c5e2';
            break;
          case 5:
            color = '#c5dbef';
            break;
          case 6:
            color = '#d1e5c9';
            break;
          default:
            color = 'white';
        }
        leaveTypeRange.setBackground(color);
      }
    });


    var dayOfWeekHeader = [''];
    dayOfWeekHeader.push('曜');
    for (var i = 1; i <= daysInMonth; i++) {
      var dayOfWeek = ['日', '月', '火', '水', '木', '金', '土'][new Date(year, month - 1, i).getDay()];
      dayOfWeekHeader.push(dayOfWeek);
    }
    newSheet.appendRow(dayOfWeekHeader);
    var dayOfWeekRange = newSheet.getRange(3, 5, 1, dayOfWeekHeader.length);
    dayOfWeekRange
      .setValues([dayOfWeekHeader])
      .setFontSize(10)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    var rangeToClearC3 = newSheet.getRange('C3');
    var rangeToClearB3 = newSheet.getRange('B3');
    rangeToClearC3.clear();
    rangeToClearB3.clear();

    var newSheetLastColumn = newSheet.getLastColumn();

    // 列の幅を設定
    for (var column = 1; column <= newSheetLastColumn; column++) {
      newSheet.setColumnWidth(column, 30);
    }

    var scheduleData = scheduleSheet.getDataRange().getValues();

    /**** */

    var employeeInfo = {};
    var omittedEmployeeIDs = [];
    var othersProfessionsNames = [];
    var professionsNames = [];
    var managementNames = [];
    var janitorialNames = [];
    var partTimeNames = [];
    var officeNames = [];
    var generalAffairsNames = [];
    var nurseNames = [];
    var maleCareWNames = [];
    var femaleCareWNames = [];
    var maleTherapistNames = [];
    var femaleTherapistNames = [];
    var cookingStaffNames = [];

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

      if (omit === false) {
        var employeeObj = { id: id, name: employeeName, displayOrder: displayOrder };
        if (empRole === "管理職") {
          managementNames.push(employeeObj);
        } else if (empRole === "その他専門職") {
          othersProfessionsNames.push(employeeObj);
        } else if (empRole === "専門職") {
          professionsNames.push(employeeObj);
        } else if (empRole === "用務") {
          janitorialNames.push(employeeObj);
        } else if (empRole === "非常勤") {
          partTimeNames.push(employeeObj);
        } else if (empRole === "事務") {
          officeNames.push(employeeObj);
        } else if (empRole === "総務") {
          generalAffairsNames.push(employeeObj);
        } else if (empRole === "看護師") {
          nurseNames.push(employeeObj);
        } else if (gender === "男" && empRole === "ケアワーカー") {
          maleCareWNames.push(employeeObj);
        } else if (gender === "女" && empRole === "ケアワーカー") {
          femaleCareWNames.push(employeeObj);
        } else if (gender === "男" && empRole === "セラピスト") {
          maleTherapistNames.push(employeeObj);
        } else if (gender === "女" && empRole === "セラピスト") {
          femaleTherapistNames.push(employeeObj);
        } else if (empRole === "調理") {
          cookingStaffNames.push(employeeObj);
        }
      }
    }

    managementNames = sortByDisplayOrder(managementNames);
    othersProfessionsNames = sortByDisplayOrder(othersProfessionsNames);
    professionsNames = sortByDisplayOrder(professionsNames);
    janitorialNames = sortByDisplayOrder(janitorialNames);
    partTimeNames = sortByDisplayOrder(partTimeNames);
    officeNames = sortByDisplayOrder(officeNames);
    generalAffairsNames = sortByDisplayOrder(generalAffairsNames);
    nurseNames = sortByDisplayOrder(nurseNames);
    maleCareWNames = sortByDisplayOrder(maleCareWNames);
    femaleCareWNames = sortByDisplayOrder(femaleCareWNames);
    maleTherapistNames = sortByDisplayOrder(maleTherapistNames);
    femaleTherapistNames = sortByDisplayOrder(femaleTherapistNames);
    cookingStaffNames = sortByDisplayOrder(cookingStaffNames);

    // 管理職名前配置
    managementEndRow = setEmployeeData(newSheet, managementStartRow, managementNames);

    // その他専門職名前配置
    othersProfessionsStartRow = managementEndRow;
    othersProfessionsEndRow = setEmployeeData(newSheet, othersProfessionsStartRow, othersProfessionsNames);

    // 専門職名前配置
    professionsStartRow = othersProfessionsEndRow;
    professionsEndRow = setEmployeeData(newSheet, professionsStartRow, professionsNames);

    // 男ケアワーカー名前配置
    maleCareworkerStartRow = professionsEndRow;
    maleCareworkerEndRow = setEmployeeData(newSheet, maleCareworkerStartRow, maleCareWNames);

    // 男セラピスト名前配置
    maleTherapistStartRow = maleCareworkerEndRow;
    maleTherapistEndRow = setEmployeeData(newSheet, maleTherapistStartRow, maleTherapistNames);

    // 男 宿直 と 遅出 を追加
    var additionalRowsStart = maleTherapistEndRow;
    newSheet.getRange(additionalRowsStart + 1, 4, 2, 1).merge().setValue('男').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground("#f0c0c1");
    newSheet.getRange(additionalRowsStart + 1, 5, 1, 2).merge().setValue('宿直').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#f0c0c1');
    newSheet.getRange(additionalRowsStart + 2, 5, 1, 2).merge().setValue('遅出').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#c5dbef');

    // 女セラピスト名前配置
    femaleTherapistStartRow = additionalRowsStart + 2;
    femaleTherapistFemaleEndRow = setEmployeeData(newSheet, femaleTherapistStartRow, femaleTherapistNames);

    // 女ケアワーカー名前配置
    femaleCareworkerStartRow = femaleTherapistStartRow + 2;
    femaleCareworkerEndRow = setEmployeeData(newSheet, femaleCareworkerStartRow, femaleCareWNames);

    // 非常勤名前配置
    partTimeStartRow = femaleCareworkerEndRow;
    partTimeEndRow = setEmployeeData(newSheet, partTimeStartRow, partTimeNames);

    // 男女 宿直 と 遅出 を追加
    var additionalRowsStart2 = partTimeEndRow;
    newSheet.getRange(additionalRowsStart2 + 1, 4, 2, 1).merge().setValue('男').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground("#f0c0c1");
    newSheet.getRange(additionalRowsStart2 + 1, 5, 1, 2).merge().setValue('宿直').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#f0c0c1');
    newSheet.getRange(additionalRowsStart2 + 2, 5, 1, 2).merge().setValue('遅出').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#c5dbef');

    //　女性
    newSheet.getRange(additionalRowsStart2 + 3, 4, 2, 1).merge().setValue('女').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground("#d1e5c9");
    newSheet.getRange(additionalRowsStart2 + 3, 5, 1, 2).merge().setValue('宿直 ').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#d1e5c9');
    newSheet.getRange(additionalRowsStart2 + 4, 5, 1, 2).merge().setValue('遅出 ').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#c5dbef');

    // 男女早出回数
    newSheet.getRange(additionalRowsStart2 + 5, 4, 1, 3).merge().setValue('男女早出').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#ffffcc');

    // 看護師名前配置
    nurseStartRow = additionalRowsStart2 + 5;
    nurseEndRow = setEmployeeData(newSheet, nurseStartRow, nurseNames);

    // 調理名前配置
    cookingStartRow = nurseEndRow;
    cookingEndRow = setEmployeeData(newSheet, cookingStartRow, cookingStaffNames);

    // 総務名前配置
    generalAffairsStartRow = cookingEndRow;
    generalAffairsEndRow = setEmployeeData(newSheet, generalAffairsStartRow, generalAffairsNames);

    // 事務名前配置
    officeStartRow = generalAffairsEndRow;
    officeEndRow = setEmployeeData(newSheet, officeStartRow, officeNames);

    // 用務名前配置
    janitorialStartRow = officeEndRow;
    janitorialEndRow = setEmployeeData(newSheet, janitorialStartRow, janitorialNames);



    var manageTitleStartRow = managementStartRow + 1;
    var othersProfessionsTitleStartRow = othersProfessionsStartRow + 1;
    var professionsTitleStartRow = professionsStartRow + 1;
    var maleCareworkerTitleStartRow = maleCareworkerStartRow + 1;
    var maleCareworkerSeraTitleStartRow = maleCareworkerEndRow - maleCareworkerStartRow;
    var maleTherapistTitleStartRow = maleTherapistStartRow + 1;
    var femaleTherapistTitleStartRow = femaleTherapistStartRow + 1;
    var femaleCareworkerTitleStartRow = femaleCareworkerStartRow + 1;
    var partTimeTitleStartRow = partTimeStartRow + 1;
    var nurseTitleStartRow = nurseStartRow + 1;
    var cookingTitleStartRow = cookingStartRow + 1;
    var generalAffairsTitleStartRow = generalAffairsStartRow + 1;
    var officeTitleStartRow = officeStartRow + 1;
    var janitorialTitleStartRow = janitorialStartRow + 1;


    newSheet.getRange(manageTitleStartRow, 1, managementEndRow - managementStartRow, 1).merge().setValue('管\n理\n職').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#f9dec1');
    newSheet.getRange(othersProfessionsTitleStartRow, 1, othersProfessionsEndRow - othersProfessionsStartRow, 1).merge().setValue('そ\nの\n他\n専\n門\n職').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#c6d8dc');
    newSheet.getRange(professionsTitleStartRow, 1, professionsEndRow - professionsStartRow, 1).merge().setValue('専\n門\n職').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#d1c4e9');

    newSheet.getRange(maleCareworkerTitleStartRow, 1, maleCareworkerSeraTitleStartRow, 1).merge().setValue('直\n接\n処\n遇\n職\n員').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#d1e5c9');
    newSheet.getRange(femaleCareworkerTitleStartRow, 1, femaleCareworkerEndRow - femaleCareworkerStartRow, 1).merge().setValue('直\n接\n処\n遇\n職\n員').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#d1e5c9');

    newSheet.getRange(maleCareworkerTitleStartRow, 2, maleCareworkerSeraTitleStartRow, 1).merge().setValue('男\n性').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#c5dbef');
    newSheet.getRange(femaleCareworkerTitleStartRow, 2, femaleCareworkerEndRow - femaleCareworkerStartRow, 1).merge().setValue('女\n性').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#c5dbef');

    newSheet.getRange(maleCareworkerTitleStartRow, 3, maleCareworkerEndRow - maleCareworkerStartRow, 1).merge().setValue('ケ\nア\nワ\nl\nカ\nl').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');

    newSheet.getRange(maleTherapistTitleStartRow, 3, maleTherapistEndRow - maleTherapistStartRow, 1).merge().setValue('セ').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
    newSheet.getRange(femaleTherapistTitleStartRow, 3, femaleTherapistFemaleEndRow - femaleTherapistStartRow, 1).merge().setValue('セ').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
    newSheet.getRange(femaleCareworkerTitleStartRow, 3, femaleCareworkerEndRow - femaleCareworkerStartRow, 1).merge().setValue('ケ\nア\nワ\nl\nカ\nl').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
    newSheet.getRange(partTimeTitleStartRow, 3, partTimeEndRow - partTimeStartRow, 1).merge().setValue('非\n常\n勤').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');

    newSheet.getRange(nurseTitleStartRow, 1, nurseEndRow - nurseStartRow, 2).merge().setValue('看護師').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');

    newSheet.getRange(cookingTitleStartRow, 1, cookingEndRow - cookingStartRow, 2).merge().setValue('調理').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');

    newSheet.getRange(generalAffairsTitleStartRow, 1, generalAffairsEndRow - generalAffairsStartRow, 2).merge().setValue('総務').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');

    newSheet.getRange(officeTitleStartRow, 1, officeEndRow - officeStartRow, 2).merge().setValue('事務').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');

    newSheet.getRange(janitorialTitleStartRow, 1, janitorialEndRow - janitorialStartRow, 2).merge().setValue('用務').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');



    var allRange = newSheet.getRange(2, 3, newSheet.getLastRow() - 1, newSheet.getLastColumn() - 2);
    allRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID).setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');

    var range1 = newSheet.getRange(2, 3, newSheet.getLastRow() - 1, newSheet.getLastColumn() - 2);
    range1.setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');

    // border only row
    var rowRange = newSheet.getRange(2, 3, newSheet.getLastRow() - 1);
    rowRange.setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');


    // PropertiesServiceに値を保存
    PropertiesService.getScriptProperties().setProperty('maleCareworkerStartRow', maleCareworkerStartRow);
    PropertiesService.getScriptProperties().setProperty('maleCareworkerEndRow', maleCareworkerEndRow);
    PropertiesService.getScriptProperties().setProperty('maleTherapistStartRow', maleTherapistStartRow);
    PropertiesService.getScriptProperties().setProperty('maleTherapistEndRow', maleTherapistEndRow);
    PropertiesService.getScriptProperties().setProperty('femaleCareworkerStartRow', femaleCareworkerStartRow);
    PropertiesService.getScriptProperties().setProperty('femaleCareworkerEndRow', femaleCareworkerEndRow);
    PropertiesService.getScriptProperties().setProperty('femaleTherapistStartRow', femaleTherapistStartRow);
    PropertiesService.getScriptProperties().setProperty('femaleTherapistFemaleEndRow', femaleTherapistFemaleEndRow);
    PropertiesService.getScriptProperties().setProperty('professionsStartRow', professionsStartRow);
    PropertiesService.getScriptProperties().setProperty('professionsEndRow', professionsEndRow);


    PropertiesService.getScriptProperties().setProperty('cookingStartRow', cookingStartRow);
    PropertiesService.getScriptProperties().setProperty('cookingEndRow', cookingEndRow);
    PropertiesService.getScriptProperties().setProperty('officeStartRow', officeStartRow);
    PropertiesService.getScriptProperties().setProperty('officeEndRow', officeEndRow);
    PropertiesService.getScriptProperties().setProperty('nurseStartRow', nurseStartRow);
    PropertiesService.getScriptProperties().setProperty('nurseEndRow', nurseEndRow);
    PropertiesService.getScriptProperties().setProperty('generalAffairsStartRow', generalAffairsStartRow);
    PropertiesService.getScriptProperties().setProperty('generalAffairsEndRow', generalAffairsEndRow);
    PropertiesService.getScriptProperties().setProperty('janitorialStartRow', janitorialStartRow);
    PropertiesService.getScriptProperties().setProperty('janitorialEndRow', janitorialEndRow);

    PropertiesService.getScriptProperties().setProperty('managementStartRow', managementStartRow);
    PropertiesService.getScriptProperties().setProperty('managementEndRow', managementEndRow);
    PropertiesService.getScriptProperties().setProperty('othersProfessionsStartRow', othersProfessionsStartRow);
    PropertiesService.getScriptProperties().setProperty('othersProfessionsEndRow', othersProfessionsEndRow);

    var shiftTypes = {};

    // scheduleDataからシフト情報を取得
    for (var i = 1; i < scheduleData.length; i++) {
      var shiftDate = new Date(scheduleData[i][2]);
      var shiftMonth = shiftDate.getMonth() + 1;

      if (shiftMonth === month) {
        var employeeID = scheduleData[i][1];
        var shiftType1 = scheduleData[i][3];

        if (omittedEmployeeIDs.includes(employeeID)) {
          var employeeName = employeeInfo[employeeID].name;
          var dateKey = shiftDate.getFullYear() + '/' + (shiftDate.getMonth() + 1) + '/' + shiftDate.getDate();

          // 日付ごとにシフト情報を格納
          if (!shiftTypes[dateKey]) {
            shiftTypes[dateKey] = {};
          }

          if (!shiftTypes[dateKey][employeeID]) {
            shiftTypes[dateKey][employeeID] = { shiftType: [], name: employeeName };
          }
          shiftTypes[dateKey][employeeID].shiftType.push(shiftType1);
        }
      }
    }

    // シートのLastRowデータを取得
    var newSheetLastColumn = newSheet.getLastRow();

    // シフトを日付ごとに配置
    for (var dateKey in shiftTypes) {
      var dateShiftData = shiftTypes[dateKey];

      for (var id in dateShiftData) {
        var shiftTypeData = dateShiftData[id].shiftType;
        var employeeName = dateShiftData[id].name;

        // shiftTypeDataが存在しない場合、デフォルト値を設定
        if (shiftTypeData.length === 0) {
          shiftTypeData = ["A", "H", "N", "O", "R"];
        }

        // employeeNameがある行を探す
        var rowIndex = 0;
        for (var row = 1; row <= newSheet.getLastRow(); row++) {
          if (newSheet.getRange(row, 4).getValue() === employeeName) {
            rowIndex = row;
            break;
          }
        }

        if (rowIndex > 0 && shiftTypeData.length > 0) {
          // 日付ごとの列を計算
          var date = new Date(dateKey);
          var colIndex = date.getDate() + 6;

          var destinationRange = newSheet.getRange(rowIndex, colIndex, 1, shiftTypeData.length);
          destinationRange.setValues([shiftTypeData]).setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontColor('#000000');
        }
      }
    }

    //　休日のセルを青くする
    highlightWeekendsAndHolidays(year, month, daysInMonth, newSheet);


    // 新しいシート上でのデータ整理
    // 遅出・夜勤・日勤・休日などの計算、Count
    var newSheetDataValues = newSheet.getDataRange().getValues();

    var nextPubHoliRow = -1;
    var nextHalfHoliRow = -1;
    var nextPaidLeaveRow = -1;
    var specialLeaveRow = -1;
    var nextNigthShiftRow = -1;

    // 公休を探す
    for (var i = 0; i < newSheetDataValues.length - 1; i++) {
      for (var j = 0; j < newSheetDataValues[i].length; j++) {
        if (newSheetDataValues[i][j].toString().indexOf("公\n休") !== -1) {
          nextPubHoliRow = i;
          break;
        }
      }
    }

    var pubHoliShiftIndex = -1;
    for (var col = 0; col < newSheetLastColumn; col++) {
      for (var row = 0; row < newSheetDataValues.length; row++) {
        if (newSheetDataValues[row][col] === "公\n休") {
          pubHoliShiftIndex = col + 1;
          break;
        }
      }
    }

    // 半休を探す
    for (var i = 0; i < newSheetDataValues.length - 1; i++) {
      for (var j = 0; j < newSheetDataValues[i].length; j++) {
        if (newSheetDataValues[i][j].toString().indexOf("半\n休") !== -1) {
          nextHalfHoliRow = i;
          break;
        }
      }
    }

    var halfShiftIndex = -1;
    for (var col = 0; col < newSheetLastColumn; col++) {
      for (var row = 0; row < newSheetDataValues.length; row++) {
        if (newSheetDataValues[row][col] === "半\n休") {
          halfShiftIndex = col + 1;
          break;
        }
      }
    }

    // 有休を探す
    for (var i = 0; i < newSheetDataValues.length - 1; i++) {
      for (var j = 0; j < newSheetDataValues[i].length; j++) {
        if (newSheetDataValues[i][j].toString().indexOf("有\n休") !== -1) {
          nextPaidLeaveRow = i;
          break;
        }
      }
    }
    var paidLeaveColIndex = -1;
    for (var col = 0; col < newSheetLastColumn; col++) {
      for (var row = 0; row < newSheetDataValues.length; row++) {
        if (newSheetDataValues[row][col] === "有\n休") {
          paidLeaveColIndex = col + 1;
          break;
        }
      }
    }

    // 特休を探す
    for (var i = 0; i < newSheetDataValues.length - 1; i++) {
      for (var j = 0; j < newSheetDataValues[i].length; j++) {
        if (newSheetDataValues[i][j].toString().indexOf("特\n休") !== -1) {
          specialLeaveRow = i;
          break;
        }
      }
    }
    var specialLeaveColIndex = -1;
    for (var col = 0; col < newSheetLastColumn; col++) {
      for (var row = 0; row < newSheetDataValues.length; row++) {
        if (newSheetDataValues[row][col] === "特\n休") {
          specialLeaveColIndex = col + 1;
          break;
        }
      }
    }

    // 宿\n直を探す
    for (var i = 0; i < newSheetDataValues.length - 1; i++) {
      for (var j = 0; j < newSheetDataValues[i].length; j++) {
        if (newSheetDataValues[i][j].toString().indexOf("宿\n直") !== -1) {
          nextNigthShiftRow = i;
          break;
        }
      }
    }
    var nightShiftColIndex = -1;
    for (var col = 0; col < newSheetLastColumn; col++) {
      for (var row = 0; row < newSheetDataValues.length; row++) {
        if (newSheetDataValues[row][col] === "宿\n直") {
          nightShiftColIndex = col + 1;
          break;
        }
      }
    }

    // 遅出を探す
    var nextLateShiftRow = -1;
    for (var i = 0; i < newSheetDataValues.length - 1; i++) {
      for (var j = 0; j < newSheetDataValues[i].length; j++) {
        if (newSheetDataValues[i][j].toString().indexOf("遅\n出") !== -1) {
          nextLateShiftRow = i;
          break;
        }
      }
    }
    var lateShiftColIndex = -1;
    for (var col = 0; col < newSheetLastColumn; col++) {
      for (var row = 0; row < newSheetDataValues.length; row++) {
        if (newSheetDataValues[row][col] === "遅\n出") {
          lateShiftColIndex = col + 1;
          break;
        }
      }
    }

    // 日勤を探す
    var nextDayShiftRow = -1;
    for (var i = 0; i < newSheetDataValues.length - 1; i++) {
      for (var j = 0; j < newSheetDataValues[i].length; j++) {
        if (newSheetDataValues[i][j].toString().indexOf("日\n勤") !== -1) {
          nextDayShiftRow = i;
          break;
        }
      }
    }
    var dayShiftColIndex = -1;
    for (var col = 0; col < newSheetLastColumn; col++) {
      for (var row = 0; row < newSheetDataValues.length; row++) {
        if (newSheetDataValues[row][col] === "日\n勤") {
          dayShiftColIndex = col + 1;
          break;
        }
      }
    }

    // 休日Count
    for (var strow = 1; strow < newSheetDataValues.length - 2; strow++) {

      var publiHoliformula = '=COUNTIF(G' + (strow + 3) + ':AK' + (strow + 3) + ', "/")';
      if (nextPubHoliRow !== -1) {
        newSheet.getRange(nextPubHoliRow + 2 + strow, pubHoliShiftIndex).setFormula(publiHoliformula).setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#f0c0c1');
      }

      var halfHoliformula = '=COUNTIF(G' + (strow + 3) + ':AK' + (strow + 3) + ', "D")';
      if (nextHalfHoliRow !== -1) {
        newSheet.getRange(nextHalfHoliRow + 2 + strow, halfShiftIndex).setValue(halfHoliformula).setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#c5dbef');
      }

      var paidLeaveformula = '=COUNTIF(G' + (strow + 3) + ':AK' + (strow + 3) + ', "/U")';
      if (nextPaidLeaveRow !== -1) {
        newSheet.getRange(nextPaidLeaveRow + 2 + strow, paidLeaveColIndex).setValue(paidLeaveformula).setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#d1e5c9');
      }

      var specialLeaveformula = '=COUNTIF(G' + (strow + 3) + ':AK' + (strow + 3) + ', "/特")';
      if (specialLeaveRow !== -1) {
        newSheet.getRange(specialLeaveRow + 2 + strow, specialLeaveColIndex).setValue(specialLeaveformula).setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#ffffcc');
      }

      var nightShiftformula = '=COUNTIF(G' + (strow + 3) + ':AK' + (strow + 3) + ', "A") + COUNTIF(G' + (strow + 3) + ':AK' + (strow + 3) + ', "A22") + COUNTIF(G' + (strow + 3) + ':AK' + (strow + 3) + ', "A13")';
      if (nextNigthShiftRow !== -1) {
        newSheet.getRange(nextNigthShiftRow + 2 + strow, nightShiftColIndex).setValue(nightShiftformula).setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#d0c5e2');
      }

      var lateShiftformula = '=COUNTIF(G' + (strow + 3) + ':AK' + (strow + 3) + ', "O")';
      if (nextLateShiftRow !== -1) {
        newSheet.getRange(nextLateShiftRow + 2 + strow, lateShiftColIndex).setValue(lateShiftformula).setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#c5dbef');
      }

      var dayShifFormula = '=COUNTIF(G' + (strow + 3) + ':AK' + (strow + 3) + ', "H") + COUNTIF(G' + (strow + 3) + ':AK' + (strow + 3) + ', "N")';
      if (nextDayShiftRow !== -1) {
        newSheet.getRange(nextDayShiftRow + 2 + strow, dayShiftColIndex).setValue(dayShifFormula).setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#d1e5c9');
      }
    }


    // 男性の遅出&宿直
    var nightShiftRowIdx = -1;
    for (var i = 1; i < newSheetDataValues.length; i++) {
      for (var j = 1; j < newSheetDataValues[i].length; j++) {
        if (newSheetDataValues[i][j].toString().indexOf("宿直") !== -1) {
          nightShiftRowIdx = i;
          break;
        }
      }
      // 宿直が見つかった場合、処理終了
      if (nightShiftRowIdx !== -1) {
        break;
      }
    }

    var lateShiftRowIdx = -1;
    for (var i = 1; i < newSheetDataValues.length; i++) {
      for (var j = 1; j < newSheetDataValues[i].length; j++) {
        if (newSheetDataValues[i][j].toString().indexOf("遅出") !== -1) {
          lateShiftRowIdx = i;
          break;
        }
      }
      // 宿直が見つかった場合、処理終了
      if (lateShiftRowIdx !== -1) {
        break;
      }
    }

    // 男性宿直・遅出回数計算
    var startRow = professionsStartRow;
    var endRow = maleTherapistEndRow;
    var startColumn = 6;
    for (var day = 1; day <= daysInMonth; day++) {
      var lateShiftFmula = "=COUNTIF(" + getColumnLetter(startColumn + day) + startRow + ":" + getColumnLetter(startColumn + day) + endRow + ", \"O\")";
      var nightShiftFmula = "=COUNTIF(" + getColumnLetter(startColumn + day) + startRow + ":" + getColumnLetter(startColumn + day) + endRow + ", \"A\") + COUNTIF(" + getColumnLetter(startColumn + day) + startRow + ":" + getColumnLetter(startColumn + day) + endRow + ", \"A22\") + COUNTIF(" + getColumnLetter(startColumn + day) + startRow + ":" + getColumnLetter(startColumn + day) + endRow + ", \"A13\")";

      if (nightShiftRowIdx !== -1) {
        newSheet.getRange(nightShiftRowIdx + 1, startColumn + day).setValue(nightShiftFmula).setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#f0c0c1');
      }

      if (lateShiftRowIdx !== -1) {
        newSheet.getRange(lateShiftRowIdx + 1, startColumn + day).setValue(lateShiftFmula).setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#c5dbef');
      }
    }



    // 女性の遅出&宿直
    var nightShiftRowIdx2 = -1;
    for (var i = 1; i < newSheetDataValues.length; i++) {
      for (var j = 1; j < newSheetDataValues[i].length; j++) {
        if (newSheetDataValues[i][j].toString().indexOf("宿直 ") !== -1) {
          nightShiftRowIdx2 = i;
          break;
        }
      }
      // 宿直が見つかった場合、処理終了
      if (nightShiftRowIdx2 !== -1) {
        break;
      }
    }

    var lateShiftRowIdx2 = -1;
    for (var i = 1; i < newSheetDataValues.length; i++) {
      for (var j = 1; j < newSheetDataValues[i].length; j++) {
        if (newSheetDataValues[i][j].toString().indexOf("遅出 ") !== -1) {
          lateShiftRowIdx2 = i;
          break;
        }
      }
      // 宿直が見つかった場合、処理終了
      if (lateShiftRowIdx2 !== -1) {
        break;
      }
    }


    //男女早出を探す
    var mfemaleRowIndx = -1;
    for (var i = 1; i < newSheetDataValues.length; i++) {
      for (var j = 1; j < newSheetDataValues[i].length; j++) {
        if (newSheetDataValues[i][j].toString().indexOf("男女早出") !== -1) {
          mfemaleRowIndx = i;
          break;
        }
      }
      // 宿直が見つかった場合、処理終了
      if (mfemaleRowIndx !== -1) {
        break;
      }
    }


    // 女性宿直・遅出回数計算
    var mfemaleStartRow = maleCareworkerStartRow;
    var mfemaleEndRow = femaleCareworkerEndRow;
    var femaleStartColumn = 6;
    for (var day = 1; day <= daysInMonth; day++) {
      var femaleLateShiftFmula = "=COUNTIF(" + getColumnLetter(femaleStartColumn + day) + femaleTherapistStartRow + ":" + getColumnLetter(startColumn + day) + femaleCareworkerEndRow + ", \"O\")";

      var femaleNightShiftFmula = "=COUNTIF(" + getColumnLetter(femaleStartColumn + day) + femaleTherapistStartRow + ":" + getColumnLetter(femaleStartColumn + day) + femaleCareworkerEndRow + ", \"A\") + COUNTIF(" + getColumnLetter(femaleStartColumn + day) + femaleTherapistStartRow + ":" + getColumnLetter(femaleStartColumn + day) + femaleCareworkerEndRow + ", \"A22\") + COUNTIF(" + getColumnLetter(femaleStartColumn + day) + femaleTherapistStartRow + ":" + getColumnLetter(femaleStartColumn + day) + femaleCareworkerEndRow + ", \"A13\")";

      var mfemaleEarlyShiftFmula = "=COUNTIF(" + getColumnLetter(femaleStartColumn + day) + mfemaleStartRow + ":" + getColumnLetter(startColumn + day) + mfemaleEndRow + ", \"H\")";

      if (nightShiftRowIdx2 !== -1) {
        newSheet.getRange(nightShiftRowIdx2 + 1, femaleStartColumn + day).setValue(femaleNightShiftFmula).setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#d1e5c9');
      }
      if (lateShiftRowIdx2 !== -1) {
        newSheet.getRange(lateShiftRowIdx2 + 1, femaleStartColumn + day).setValue(femaleLateShiftFmula).setFontSize(10).setHorizontalAlignment('center').
          setVerticalAlignment('middle').setBackground('#c5dbef');;
      }
      if (mfemaleRowIndx !== -1) {
        newSheet.getRange(mfemaleRowIndx + 1, femaleStartColumn + day).setValue(mfemaleEarlyShiftFmula).setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#ffffcc');
      }



      var lateShiftFmula2 = "=COUNTIF(" + getColumnLetter(startColumn + day) + professionsStartRow + ":" + getColumnLetter(startColumn + day) + maleTherapistEndRow + ", \"O\")";

      var nightShiftFmula2 = "=COUNTIF(" + getColumnLetter(startColumn + day) + professionsStartRow + ":" + getColumnLetter(startColumn + day) + maleTherapistEndRow + ", \"A\") + COUNTIF(" + getColumnLetter(startColumn + day) + professionsStartRow + ":" + getColumnLetter(startColumn + day) + maleTherapistEndRow + ", \"A22\") + COUNTIF(" + getColumnLetter(startColumn + day) + professionsStartRow + ":" + getColumnLetter(startColumn + day) + maleTherapistEndRow + ", \"A13\")";

      if (nightShiftRowIdx2 !== -1) {
        newSheet.getRange(nightShiftRowIdx2 - 1, startColumn + day).setValue(nightShiftFmula2).setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#f0c0c1');
      }

      if (lateShiftRowIdx2 !== -1) {
        newSheet.getRange(lateShiftRowIdx2 - 1, startColumn + day).setValue(lateShiftFmula2).setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#c5dbef');
      }

    }


    onOpen();


    ui.alert(sheetName + 'のその他のシフトを作成完了しました。');
    ss.toast('処理が完了しました', '処理完了', 5);

  } catch (error) {
    // エラーメッセージを表示
    Logger.log(error.message);
    ss.toast('エラーが発生しました: ' + error.message, 'エラー', 5);
  }

}


/**
 * 職員名前セット
 */
function setEmployeeData(sheet, row, data) {
  for (var i = 0; i < data.length; i++) {
    sheet.getRange(row + i + 1, 4, 1, 3).merge().setValue(data[i].name).setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
    // sheet.getRange(row + i + 1, 45, 1, 1).setValue(data[i].id).setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
  }
  return row + data.length;
}

// displayOrderでソート
function sortByDisplayOrder(arr) {
  return arr.sort(function (a, b) {
    return a.displayOrder - b.displayOrder;
  });
}




/**
 * 
 * 列の文字を取得する関数
 */
function getColumnLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}


