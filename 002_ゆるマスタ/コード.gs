/**
 * 起動時の処理
 */
function onOpen() {
  // 最終行番号
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let lastRowIdx = sheet.getLastRow();

  // 最終行選択
  sheet.getRange(lastRowIdx + 3, 1, 1, 1).activate();
}

/**
 * ゆる山適正リストを作成する
 */
function clickYuruyama() {
  executeYuruyama();
}

/**
 * ゆるの登録を行う
 */
function clickYuruRegister() {
  executeYuruRegister();
}