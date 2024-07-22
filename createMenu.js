
/**
 * メニュー作成
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  // メインメニューの作成
  var mainMenu = ui.createMenu('シフト操作');

  // 夜勤シフトメニューの追加
  mainMenu.addItem('夜勤のシフトを作成', 'createNightShiftsProposal');

  // 日勤+早出シフトメニューの追加
  mainMenu.addItem('日勤+早出のシフトを作成', 'createDayShiftsProposal');

  // 遅出シフトメニューの追加
  mainMenu.addItem('遅出+休みのシフトを作成', 'createLateShiftsProposal');

  // 他シフトメニューの追加
  mainMenu.addItem('その他シフト作成', 'createOtherShiftsProposal');

  // ラインマークを追加
  mainMenu.addSeparator();

  // 確定シフトメニューの追加
  mainMenu.addSubMenu(ui.createMenu('確定シフト')
    .addItem('確定のシフトデータをテーブルに格納する', 'storeShiftsDataTo'));

  // メインメニューをUIに追加
  mainMenu.addToUi();
}