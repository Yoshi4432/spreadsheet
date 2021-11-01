/** シート名 */
const mstSheetName = "ゆる";

/**
 * 登録フォーム（という名のHTML）を表示
 */
function executeYuruRegister() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var html = HtmlService.createHtmlOutputFromFile('ゆる登録');
  SpreadsheetApp.getUi().showModalDialog(html, 'ゆる登録');
}

/**
 * 登録フォームからの入力値を元に登録を行う
 */
function executeRegister(form) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mstSheet = spreadsheet.getSheetByName(mstSheetName);

  /* 最終行を表示 */
  mstSheet.activate();
  focusLatestYuru();

  /* 登録エンティティ */
  const yuru = new MstYuru();
  yuru.ID = mstSheet.getLastRow();
  yuru.名前 = form.namae;
  yuru.地方 = form.tihou;
  yuru.出身地1 = form.birthplace1;
  yuru.出身地2 = form.birthplace2;
  yuru.火 = (form.attribute.indexOf('火') > -1) ? '○' : '';
  yuru.水 = (form.attribute.indexOf('水') > -1) ? '○' : '';
  yuru.風 = (form.attribute.indexOf('風') > -1) ? '○' : '';
  yuru.レア = form.rare;
  yuru.Cost = form.cost;
  yuru.HP = form.hp;
  yuru.Power = form.power;
  yuru.凸HP = (Number.parseInt(form.hp) + Number.parseInt(form.hp_difference));
  yuru.凸Power = (Number.parseInt(form.power) + Number.parseInt(form.power_difference));
  yuru.LS種類 = form.ls_type;
  yuru.LS味方範囲 = form.ls_ally;
  yuru.LS敵範囲 = form.ls_enemy;
  yuru.LS係数 = form.ls_multiple;
  yuru.YS種類 = form.ys_type;
  yuru.ST = form.st;
  yuru.YS味方範囲 = form.ys_ally;
  yuru.YS敵範囲 = form.ys_enemy;
  yuru.持続 = form.lasting;
  yuru.YS係数 = form.ys_multiple;
  // 最後に実行する必要あり
  yuru.LS = getLeaderSkill(yuru);
  yuru.YS = getYuruSkill(yuru);

  /* 登録 */
  registerYuru(yuru);

  /* フォーマット */
  setYuruFormat();
}

/**
 * ゆるシートのフォーマット設定
 */
function setYuruFormat() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(mstSheetName);
  const lastRowIdx = sheet.getLastRow();
  const lastColIdx = sheet.getLastColumn();

  // 罫線
  const all = sheet.getRange(1, 1, lastRowIdx, lastColIdx);
  all.setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);

  // 中央ぞろえ
  all.setHorizontalAlignment('center');
}

/**
 * ゆるをマスタに書き込む
 * @param {MstYuru} yuru マスタゆる
 */
function registerYuru(yuru) {
  let vals = [];
  vals.push(yuru.ID);
  vals.push(yuru.名前);
  vals.push(yuru.地方);
  vals.push(yuru.出身地1);
  vals.push(yuru.出身地2);
  vals.push(yuru.火);
  vals.push(yuru.水);
  vals.push(yuru.風);
  vals.push(yuru.レア);
  vals.push(yuru.Cost);
  vals.push(yuru.HP);
  vals.push(yuru.Power);
  vals.push(yuru.凸HP);
  vals.push(yuru.凸Power);
  vals.push(yuru.LS);
  vals.push(yuru.LS種類);
  vals.push(yuru.LS味方範囲);
  vals.push(yuru.LS敵範囲);
  vals.push(yuru.LS係数);
  vals.push(yuru.YS);
  vals.push(yuru.YS種類);
  vals.push(yuru.ST);
  vals.push(yuru.YS味方範囲);
  vals.push(yuru.YS敵範囲);
  vals.push(yuru.持続);
  vals.push(yuru.YS係数);

  /* 登録 */
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mstSheet = spreadsheet.getSheetByName(mstSheetName);
  const range = mstSheet.getRange(mstSheet.getLastRow() + 1, 1, 1, mstSheet.getLastColumn());
  range.setValues([vals]);
}

/**
 * リーダースキルの文言を生成する
 * @param {MstYuru} yuru マスタゆる
 * @return {string} リーダースキル文言
 */
function getLeaderSkill(yuru) {
  /* 文言作成 */
  // $1=味方範囲 $2=敵範囲 $3=倍率

  // テンプレート
  let str = '';
  if (yuru.LS種類 === '攻撃↑') {
    // 攻撃
    str = '$1属性の味方全員の$2属性の敵に対する攻撃力を$3%増加させる。';
  } else if (yuru.LS種類 === '連続') {
    // 連続
    str = '$1属性の味方全員の攻撃が$3回連続攻撃になる。';
  } else if (yuru.LS種類 === '全体') {
    // 全体
    str = '自身の攻撃が全体攻撃になる。';
  } else if (yuru.LS種類 === 'HP↑') {
    // HP↑
    str = '$1属性の味方全員のHPを$3%上昇させる。';
  } else if (yuru.LS種類 === '防御↑') {
    // 防御
    str = '$1属性の味方全員へのダメージを$3%カットする。';
  } else if (yuru.LS種類 === 'かばう') {
    // かばう
    str = '敵の攻撃を全て自分に集中させる。';
  } else if (yuru.LS種類 === 'お金↑') {
    // お金↑
    str = 'ゴールドの取得額を$3%増加させる。';
  } else if (yuru.LS種類 === 'ドロ率↑') {
    // ドロ率↑
    str = 'カードのドロップ率を$3%増加させる。';
  } else if (yuru.LS種類 === '経験↑') {
    // 経験↑
    str = '経験値の取得量を$3%増加させる。';
  } else if (yuru.LS種類 === 'なし') {
    // なし
    str = 'なし。';
  }

  // 味方範囲
  str = str.replace('$1', yuru.LS味方範囲);
  str = str.replace('全属性の味方全員', '味方全員');
  str = str.replace('自属性の味方全員', '自身');

  // 敵範囲
  str = str.replace('$2', yuru.LS敵範囲);
  str = str.replace('全属性の敵に対する', '');

  // 倍率
  str = str.replace('$3', yuru.LS係数);

  return str;
}

/**
 * ゆるスキルの文言を生成する
 * @param {MstYuru} yuru マスタゆる
 * @return {string} ゆるスキル文言
 */
function getYuruSkill(yuru) {
  /* 文言作成 */
  // $1=味方範囲 $2=敵範囲 $3=持続 $4=倍率

  // テンプレート
  let str = '';
  if (yuru.YS種類 === '攻撃↑') {
    // 攻撃↑
    str = '$3ターンの間、$2属性の敵に対する$1属性の味方全員の攻撃力を$4%上昇させる。';
  } else if (yuru.YS種類 === '連続') {
    // 連続
    str = '$3ターンの間、$1属性の味方の攻撃が$4連続攻撃になる。';
  } else if (yuru.YS種類 === '全体') {
    // 全体
    str = '$3ターンの間、$1属性の味方全員の攻撃が全体攻撃になる。';
  } else if (yuru.YS種類 === 'シールド') {
    // シールド
    str = '$3ターンの間、$1属性の味方全員に$2属性の敵からのダメージを$4%軽減するシールドを与える。';
  } else if (yuru.YS種類 === 'かばう') {
    // かばう
    str = '$3ターンの間、$2属性の攻撃を自身に集中させる。($4%カット)';
  } else if (yuru.YS種類 === 'かわす') {
    // かわす
    str = '$3ターンの間、$2属性の攻撃をリーダーに集中させる。($4%カット)';
  } else if (yuru.YS種類 === '回復') {
    // 回復
    str = '$1属性の味方全員のHPを$4%回復する。';
  } else if (yuru.YS種類 === '蘇生') {
    // 蘇生
    str = '死亡している$1属性の味方全員を$4%のHPで蘇らせる。';
  } else if (yuru.YS種類 === '遅延') {
    // 遅延
    str = '$2属性の敵全員の攻撃カウントを$4増加させる。';
  } else if (yuru.YS種類 === '遅延') {
    // 短縮
    str = '$1属性の味方全員のスキルカウントを$4減少させる。';
  } else if (yuru.YS種類 === '属変') {
    // 属変
    str = '$3ターンの間、$1属性の味方全員の属性を$4属性に変更する。';
  } else if (yuru.YS種類 === 'なし') {
    str = 'なし。';
  }

  // 味方範囲
  str = str.replace('$1', yuru.YS味方範囲);
  str = str.replace('全属性の味方全員', '味方全員');
  str = str.replace('自属性の味方全員', '自身'); // 攻撃, 回復, 属変
  str = str.replace('自属性の味方', '自身'); // 連続, 全体
  str = str.replace('リ属性の味方全員', 'リーダー');

  // 敵範囲
  str = str.replace('$2', yuru.YS敵範囲);
  str = str.replace('全属性の敵に対する', ''); // 攻撃
  str = str.replace('全属性の敵からの', ''); // シールド
  str = str.replace('全属性の', ''); // かばう, かわす, 遅延

  // 持続
  str = str.replace('$3', yuru.持続);

  // 倍率
  str = str.replace('$4', yuru.YS係数);

  return str;
}

/**
 * IDからゆる情報を取得
 * @param {number} id ID
 * @return {MstYuru} ゆる情報
 */
function getMstYuruById(id) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mstSheet = spreadsheet.getSheetByName(mstSheetName);

  // 順列なので力技で取得！
  const range = mstSheet.getRange(Number.parseInt(id) + 1, 1, 1, mstSheet.getLastColumn());

  // 変換
  const vals = range.getValues();
  const mstYuru = new MstYuru(vals[0]);

  return mstYuru;
}

/**
 * 名前からゆる情報を取得
 * 
 * @param {string} name 名前
 * @return {MstYuru} ゆる情報
 */
function getMstYuruByName(name) {
  // 処理中の名前と完全一致するゆるデータを取得
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mstSheet = spreadsheet.getSheetByName(mstSheetName);
  const finder = mstSheet.createTextFinder(name);
  const cells = finder.findAll();

  let target = null;
  for (const cell of cells) {
    if (cell.getLastColumn() === 2 && cell.getValue() === name) {
      target = cell;
      break;
    }
  }
  const range = mstSheet.getRange(target.getRow(), 1, 1, mstSheet.getLastColumn());

  // ゆるインスタンス取得
  const vals = range.getValues();
  const mstYuru = new MstYuru(vals[0]);

  return mstYuru;
}

/**
 * 最終行（のあたり）をフォーカスする
 */
function focusLatestYuru() {
  /* 最終行のあたりを初期表示 */
  // 最終行番号
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let lastRowIdx = sheet.getLastRow();
  // 最終行選択
  sheet.getRange(lastRowIdx + 3, 1, 1, 1).activate();
}

/**
 * ゆるマスタ
 */
class MstYuru {

  /**
   * @param {Array} vals
   */
  constructor(vals) {
    if (vals === null || vals === undefined || vals.length === 0) {
      return;
    }

    this.ID = vals[0];
    this.名前 = vals[1];
    this.地方 = vals[2];
    this.出身地1 = vals[3];
    this.出身地2 = vals[4];
    this.火 = vals[5];
    this.水 = vals[6];
    this.風 = vals[7];
    this.レア = vals[8];
    this.Cost = vals[9];
    this.HP = vals[10];
    this.Power = vals[11];
    this.凸HP = vals[12];
    this.凸Power = vals[13];
    this.LS = vals[14];
    this.LS種類 = vals[15];
    this.LS味方範囲 = vals[16];
    this.LS敵範囲 = vals[17];
    this.LS係数 = vals[18];
    this.YS = vals[19];
    this.YS種類 = vals[20];
    this.ST = vals[21];
    this.YS味方範囲 = vals[22];
    this.YS敵範囲 = vals[23];
    this.持続 = vals[24];
    this.YS係数 = [25];
  }
}

/**
 * getMstYuru***のテスト
 */
function testGetMstYuru() {
  let yuru = null;

  // うちいり花子？ 
  yuru = getMstYuruById(1612);
  if (yuru.名前 === '討ち入り☆花子') {
    Logger.log('討ち入り☆花子 ok');
  } else {
    Logger.log('討ち入り☆花子 ng');
  }

  // うちいり花子？
  yuru = getMstYuruByName('討ち入り☆花子');
  if (yuru.名前 === '討ち入り☆花子') {
    Logger.log('討ち入り☆花子 ok');
  } else {
    Logger.log('討ち入り☆花子 ng');
  }
}

/**
 * registerYuruのテスト
 */
function testRegisterYuru() {
  /* ID 1610 のゆるが新規登録されることを確認 */
  const yuru = getMstYuruById(1610);
  registerYuru(yuru);
}
