
/**
 * 入力シフト抽出（男性日勤）
 */
function getInputShift(employeeName, shiftType, dateKey, inputShift, fontColor) {
  if (!inputShift[dateKey]) {
    inputShift[dateKey] = {};
  }

  inputShift[dateKey][employeeName] = {};
  if (shiftType !== "" && fontColor === "#000000") {
    inputShift[dateKey][employeeName] = shiftType;
  }
}