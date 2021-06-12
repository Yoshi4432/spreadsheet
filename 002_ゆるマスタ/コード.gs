/**
 * 起動時の処理
 */
function onOpen(e) {
  /* 最終行のあたりを初期表示 */
  focusLatestYuru();
  
  /* アドオン追加 */
  SpreadsheetApp.getUi().createAddonMenu()
    .addItem('登録', 'clickYuruRegister')
    .addItem('ゆる山', 'clickYuruyama')
    .addToUi();
}

/**
 * アドオン追加時の処理
 */
function onInstall(e) {
  onOpen(e);
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