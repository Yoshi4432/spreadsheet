/** シート名 */
const yuruyamaSheetName = "ゆる山適正";

/**
 * ゆる山適正処理のメイン
 */
function executeYuruyama() {
  /* スプレッドシート等 */
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mstSheet = spreadsheet.getSheetByName(mstSheetName);
  const yuruyamaSheet = spreadsheet.getSheetByName(yuruyamaSheetName);

  /* IDとゆる名をコピーしてベースを作成 */
  // マスタのIDとゆる名をコピー
  const mstRange = mstSheet.getRange(1, 1, mstSheet.getLastRow(), 2);
  mstRange.copyTo(yuruyamaSheet.getRange(1, 1));

  /* マスタデータを取得 */
  const mstAll = mstSheet.getDataRange().getValues();

  /* ループしながら各ゆるがゆる山に向いたか判定・攻撃力計算を行う */
  // ゆる山シートの表を取得して配列へ
  const yuruyamaArr = yuruyamaSheet.getRange(2, 1, yuruyamaSheet.getLastRow(), yuruyamaSheet.getLastColumn()).getValues();
  // ループしながら各判定
  for (const yuru of yuruyamaArr) {
    const id = yuru[0];
    const name = yuru[1];
    const appropriate = yuru[2];

    // 不要な行は処理しない
    if (StringUtils.isTrimEmpty(name)) {
      continue;
    }
    // 判定済みのゆるは処理しない
    if (!StringUtils.isTrimEmpty(appropriate)) {
      continue;
    }

    // 処理中の名前と完全一致するゆるデータを取得
    const mstYuru = new MstYuru(mstAll[id]);
    // 未公開のゆるはスキップ
    if (StringUtils.isTrimEmpty(mstYuru.Cost.toString())) {
      continue;
    }

    /* ゆる山適正の判定 */
    let appropriateFlg = true;
    // 凸攻撃力に記載が無ければ不適合
    const att = mstYuru.凸Power;
    if (StringUtils.isTrimEmpty(att.toString())) {
      appropriateFlg = false;
    }
    // コストが50を超えていたら不適合
    const cost = Number.parseInt(mstYuru.Cost);
    if (cost > 50) {
      appropriateFlg = false;
    }
    // リーダースキル種類の不適合判定
    const lsType = mstYuru.LS種類;
    const lsMulti = Number.parseInt(mstYuru.LS係数);
    if (lsType !== '連続' && lsType !== '全体' && lsType !== '攻撃↑') {
      appropriateFlg = false;
    } else if (lsType === '連続' && lsMulti < 3) {
      appropriateFlg = false;
    } else if (lsType === '攻撃↑' && mstYuru.LS敵範囲 !== '全') {
      appropriateFlg = false;
    }

    // 適合・不適合を記入
    let vals = [];
    let check = appropriateFlg ? '○' : '×';
    vals.push(check);

    // 適合する場合は残りの情報を記入
    if (appropriateFlg) {
      // コスト
      vals.push(mstYuru.Cost);

      // 攻撃力
      vals.push(mstYuru.凸Power);

      // LS種類
      vals.push(mstYuru.LS種類);

      // 倍率
      if (lsType === '全体') {
        vals.push(3);
      } else {
        vals.push(mstYuru.LS係数);
      }

      // 実質攻撃力
      let att = Number.parseInt(mstYuru.凸Power);
      if (lsType === '全体') {
        // 攻撃力＊全体(3)＊フレンド攻撃力(400)
        att = att * 3 * 4;
      } else if (lsType === '連続') {
        // 攻撃力＊連続＊フレンド攻撃力(400)
        att = att * Number.parseInt(mstYuru.LS係数) * 4;
      } else if (lsType === '攻撃↑') {
        // 攻撃力＊倍率＊連続(3)
        let multi = Number.parseInt(mstYuru.LS係数) + 100;
        multi = multi / 100;
        att = Math.floor(mstYuru.凸Power * multi) * 3;
      } else {
        // あれれーおかしいぞー？
        att = 1;
      }
      vals.push(att);
    } else {
      vals.push('');
      vals.push('');
      vals.push('');
      vals.push('');
      vals.push('');
    }

    const range = yuruyamaSheet.getRange(id + 1, 3, 1, 6);
    range.setValues([vals]);
  }

  // フォーマット適用
  setYuruyamaFormat();
}

function setYuruyamaFormat() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(yuruyamaSheetName);
  const lastRowIdx = sheet.getLastRow();
  const lastColIdx = sheet.getLastColumn();

  // 罫線
  const all = sheet.getRange(1, 1, lastRowIdx, lastColIdx);
  all.setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);

  // 中央ぞろえ
  all.setHorizontalAlignment('center');
}

/**
 * 指定のゆる名の指定のカラムを取得する.
 * 
 * @param {string[]} yuru ゆる情報配列
 * @param {number} col カラム番号
 * @return {Range} 指定のゆる名の指定のカラム番号のRange
 */
function getYuruyamaYuruCell(yuru, col) {
  const id = yuru[0];
  const name = yuru[1];

  // 処理中の名前と完全一致するゆるデータを取得
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(yuruyamaSheetName);
  const finder = sheet.createTextFinder(id);
  const cells = finder.findAll();

  let target = null;
  for (const cell of cells) {
    if (cell.getLastColumn() === 1 && cell.getValue() === id) {
      target = cell;
      break;
    }
  }
  const range = sheet.getRange(target.getRow(), col, 1, 1);

  return range;
}
