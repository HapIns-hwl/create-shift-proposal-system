function myfunction() {
  // 日付が祝日かどうかを確認する例
  var targetDate = '2024/6/1'; // 2024年5月10日を例として確認
  var isHolidayResult = isHolidayOfJOrC(targetDate, false);
  // 土日を祝日扱いにするならTrue
  Logger.log(targetDate + ' は祝日ですか？ ' + isHolidayResult);
}

//////////////////////////////////////////////
//  日本の祝日か、会社の公休日（学校の休みや給食がない日）かを調べます
//　土日も祝日として返します。
//  日付はyyyy/m/d 形式で渡す
//  includeWeekendが true だと土日も祝日
//  呼び出し元のスプレッドシートブック内に　指定休日　シートが必要です。
///////////////////////////////////////////////

function isHolidayOfJOrC(dateString, includeWeekend) {

  var dateParts = dateString.split('/');
  var year = parseInt(dateParts[0]);
  var month = parseInt(dateParts[1]); // 1月はじまり
  var day = parseInt(dateParts[2]);

  var date = new Date(year, month -1, day); // 月を0から11に変換する必要はありません
  
  var isCompanyHoliday = checkCompanyHoliday(year, month, day); // 月の補正は必要ありません
  var isJapaneseHoliday = checkJapaneseHoliday(date); // 月を0から11に変換
  var isWeekendDay = isWeekend(date);
  
  // 週末を含める場合は、週末も祝日として扱う
  if (includeWeekend) {
  /*
    Logger.log (isCompanyHoliday); 
    Logger.log (isWeekendDay); 
    Logger.log (isJapaneseHoliday);
  */
    return isCompanyHoliday || isWeekendDay || isJapaneseHoliday;
  } else {
    // 週末を含めない場合は、週末を除いた祝日のみを返す
    return isCompanyHoliday || isJapaneseHoliday;
  }
}

function isWeekend(date) {
  // 日曜日: 0, 土曜日: 6
  var dayOfWeek = date.getDay();
  return (dayOfWeek === 0 || dayOfWeek === 6);
}

function checkCompanyHoliday(year, month, day) {
  // スプレッドシートから会社の休みを読み取る
  //var sheet = SpreadsheetApp.openById('1lWDI-yLOE8vzB1Rc8udUN6KWJypzBOGvKw9z08voaQA').getSheetByName('指定休日');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('指定休日');
  if (!sheet) {

    return false; // シートが存在しない場合は指定された日は会社の休日ではないと判断
  }

  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
   // 2行目から開始し、開始日と終了日の間に指定された日付が含まれているかどうかを確認
  for (var i = 1; i < values.length; i++) {
    var startDate = new Date(values[i][0]);
    var endDate = new Date(values[i][1]);
    
    // 日付をyyyymmdd形式の文字列に変換

    var startDateString = startDate.getFullYear() + ('0' + (startDate.getMonth() + 1)).slice(-2) + ('0' + startDate.getDate()).slice(-2);
    var endDateString = endDate.getFullYear() + ('0' + (endDate.getMonth() + 1)).slice(-2) + ('0' + endDate.getDate()).slice(-2);
    var targetDateString = year + ('0' + (month )).slice(-2) + ('0' + day).slice(-2);
    
    // 指定された日付が開始日から終了日の間に含まれるかどうかを確認
    if (targetDateString >= startDateString) {
      if (targetDateString <= endDateString) {
        return true;
      }
    }
  }

  return false;
}

function checkJapaneseHoliday(date) {
  const calendarId = 'en.japanese#holiday@group.v.calendar.google.com';
  const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  const events = CalendarApp.getCalendarById(calendarId).getEventsForDay(date);
  
  return events.some(event => {
    const eventDate = Utilities.formatDate(event.getStartTime(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return eventDate === formattedDate;
  });
}