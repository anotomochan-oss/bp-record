/**
 * こけログ（血圧記録システム）サーバ側＝データ担当
 * スプレッドシートの「拡張機能 → Apps Script」から作ったスクリプトとして使う。
 * 画面（GitHub Pages）から POST で呼ばれ、JSON を返す。
 */

const SHEET_RECORDS = '記録';
const SHEET_SETTINGS = '設定';
const DATE_FORMAT = 'yyyy/MM/dd HH:mm';
const FETCH_DAYS = 182; // 画面に返す記録の期間（グラフの最長3か月＋比較用の前の3か月）

// 入力値の許容範囲（入力ミス防止。画面側もこの値を使う）
const RANGE = {
  sys: { min: 50, max: 300 },
  dia: { min: 30, max: 200 },
  pulse: { min: 30, max: 250 },
  temp: { min: 34, max: 42 }, // 体温（小数1桁）
};
const TEMP_COL = 9; // 記録シートの体温の列（I列。後から追加したので一番右）

// 暗証番号を続けて間違えたときのロック
const MAX_PIN_FAILS = 5;          // この回数間違えたらロック
const PIN_LOCK_SECONDS = 15 * 60; // ロックする時間（最後に間違えてから）

/** ブラウザで直接URLを開いたとき用（動作確認） */
function doGet() {
  return ContentService.createTextOutput('こけログ API は動いています');
}

/**
 * 画面からの呼び出しの入口
 * 受け取る JSON：{ action: 'getConfig' | 'addRecord' | 'getRecords', pin: '1234', ... }
 * 返す JSON    ：{ ok: true, data: ... } または { ok: false, code: 'BAD_PIN' など, error: 'メッセージ' }
 */
function doPost(e) {
  let res;
  try {
    const req = JSON.parse(e.postData.contents);
    const actions = {
      getConfig: function () { return getSettings_(); },
      addRecord: function () { return addRecord_(req.record); },
      getRecords: function () { return getRecords_(req.name, req.from, req.to); },
    };
    if (!actions[req.action]) throw apiError_('BAD_REQUEST', '不明な操作です');
    checkPin_(req.pin);
    res = { ok: true, data: actions[req.action]() };
  } catch (err) {
    res = { ok: false, code: err.code || 'ERROR', error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

/** 初期設定：エディタから1回だけ実行する（既にあるシートは変更しない） */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(Session.getScriptTimeZone());

  if (!ss.getSheetByName(SHEET_SETTINGS)) {
    const st = ss.insertSheet(SHEET_SETTINGS);
    // ※ 行の位置（B3〜B6）は記録シートの色付けで参照しているので動かさないこと
    st.getRange(1, 1, 7, 3).setValues([
      ['項目', '値', '説明'],
      ['利用者名', '夫,妻', 'カンマ区切りで入力'],
      ['注意_最高', 135, 'この値以上で黄色'],
      ['注意_最低', 85, 'この値以上で黄色'],
      ['強い注意_最高', 160, 'この値以上で赤'],
      ['強い注意_最低', 100, 'この値以上で赤'],
      ['朝夜の境界(時)', 12, 'この時刻より前を「朝」とする'],
    ]);
    st.getRange('A1:C1').setFontWeight('bold');
    st.setColumnWidth(1, 140);
    st.setColumnWidth(3, 240);
  }

  if (!ss.getSheetByName(SHEET_RECORDS)) {
    const rec = ss.insertSheet(SHEET_RECORDS, 0);
    rec.appendRow(['登録日時', '測定日時', '名前', '区分', '最高血圧', '最低血圧', '脈拍', 'メモ', '体温']);
    rec.getRange('A1:I1').setFontWeight('bold');
    rec.getRange('I:I').setNumberFormat('0.0');
    rec.setFrozenRows(1);
    rec.getRange('A:B').setNumberFormat(DATE_FORMAT);
    rec.setColumnWidths(1, 2, 130);
    rec.setColumnWidth(8, 200);
    addAlertFormatting_(rec);
  }

  // 新規作成時の空シートを削除
  ['シート1', 'Sheet1'].forEach(function (n) {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  if (!PropertiesService.getScriptProperties().getProperty('PIN')) {
    Logger.log('スクリプト プロパティに PIN（4〜8桁の数字の暗証番号）を設定してください');
  }
}

/** 記録シートで基準値を超えた血圧に色を付ける（印刷時にも分かるように） */
function addAlertFormatting_(sheet) {
  const RED = '#f4c7c3';
  const YELLOW = '#fce8b2';
  const rule = function (col, settingCell, color) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER(' + col + '2),' + col + '2>=INDIRECT("' + SHEET_SETTINGS + '!' + settingCell + '"))')
      .setBackground(color)
      .setRanges([sheet.getRange(col + '2:' + col)])
      .build();
  };
  // 先に書いたルールが優先される（赤 → 黄の順）
  sheet.setConditionalFormatRules([
    rule('E', 'B5', RED), rule('F', 'B6', RED),
    rule('E', 'B3', YELLOW), rule('F', 'B4', YELLOW),
  ]);
}

// ===== doPost から呼ばれる関数（暗証番号は doPost で確認済み） =====

/** 記録を1件登録する */
function addRecord_(input) {
  const r = validate_(input, getSettings_());
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_(SHEET_RECORDS);
    ensureTempColumn_(sh);
    sh.appendRow([
      new Date(), r.measuredAt, r.name, r.cat, r.sys, r.dia, r.pulse === null ? '' : r.pulse, r.memo,
      r.temp === null ? '' : r.temp,
    ]);
  } finally {
    lock.releaseLock();
  }
  return true;
}

/** 体温の列がない（体温を追加する前に作った）シートなら、見出しと表示形式を足す */
function ensureTempColumn_(sh) {
  const header = sh.getRange(1, TEMP_COL);
  if (header.getValue() === '体温') return;
  header.setValue('体温').setFontWeight('bold');
  sh.getRange(2, TEMP_COL, sh.getMaxRows() - 1, 1).setNumberFormat('0.0');
}

/**
 * 指定した人の記録を、測定日時の古い順に返す
 * from / to（'yyyy-MM-dd'）を指定するとその期間（両日を含む）、省略すると直近 FETCH_DAYS 日分
 */
function getRecords_(name, from, to) {
  const sh = sheet_(SHEET_RECORDS);
  if (sh.getLastRow() < 2) return [];

  const tz = Session.getScriptTimeZone();
  let since = new Date(Date.now() - FETCH_DAYS * 24 * 60 * 60 * 1000);
  let until = new Date(8640000000000000); // 上限なし
  if (from || to) {
    try {
      since = Utilities.parseDate(String(from), tz, 'yyyy-MM-dd');
      until = Utilities.parseDate(String(to), tz, 'yyyy-MM-dd');
    } catch (e) {
      throw apiError_('BAD_REQUEST', '期間の指定が正しくありません');
    }
    until = new Date(until.getTime() + 24 * 60 * 60 * 1000); // 終了日の翌日0時の手前まで
  }
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, Math.min(TEMP_COL, sh.getMaxColumns())).getValues();
  return rows
    .filter(function (row) { return row[2] === name && row[1] instanceof Date && row[1] >= since && row[1] < until; })
    .sort(function (a, b) { return a[1] - b[1]; })
    .map(function (row) {
      // Date 型はそのまま画面に返せないので文字列と数値にする
      return {
        t: row[1].getTime(),
        dt: Utilities.formatDate(row[1], tz, DATE_FORMAT),
        cat: row[3],
        sys: Number(row[4]),
        dia: Number(row[5]),
        pulse: row[6] === '' ? null : Number(row[6]),
        memo: String(row[7]),
        temp: row[8] === '' || row[8] === undefined ? null : Number(row[8]),
      };
    });
}

// ===== 内部で使う関数 =====

function apiError_(code, message) {
  const err = new Error(message);
  err.code = code;
  return err;
}

/** 暗証番号を確認する。続けて間違えたら一定時間すべて断る（総当たり対策） */
function checkPin_(pin) {
  const expected = PropertiesService.getScriptProperties().getProperty('PIN');
  if (!expected) throw apiError_('NO_PIN', '暗証番号がまだ設定されていません（セットアップ手順を参照）');

  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('pinFails') || 0);
  if (fails >= MAX_PIN_FAILS) {
    throw apiError_('LOCKED', '暗証番号の間違いが続いたため、しばらく使えません。' + (PIN_LOCK_SECONDS / 60) + '分ほど待ってからやり直してください');
  }
  if (String(pin || '') !== expected) {
    cache.put('pinFails', String(fails + 1), PIN_LOCK_SECONDS);
    throw apiError_('BAD_PIN', '暗証番号が違います');
  }
}

/** ロックを今すぐ解除したいときにエディタから実行する */
function unlockPin() {
  CacheService.getScriptCache().remove('pinFails');
}

function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('シート「' + name + '」がありません。setup を実行してください');
  return sh;
}

/** 設定シートの値を読む（項目名で探すので、行を追加しても壊れない） */
function getSettings_() {
  const map = {};
  sheet_(SHEET_SETTINGS).getDataRange().getValues().slice(1).forEach(function (row) {
    map[String(row[0]).trim()] = row[1];
  });
  return {
    users: String(map['利用者名'] || '').split(/[,、，]/).map(function (s) { return s.trim(); }).filter(Boolean),
    warn: { sys: Number(map['注意_最高']) || 135, dia: Number(map['注意_最低']) || 85 },
    danger: { sys: Number(map['強い注意_最高']) || 160, dia: Number(map['強い注意_最低']) || 100 },
    boundaryHour: Number(map['朝夜の境界(時)']) || 12,
    range: RANGE,
  };
}

/** 入力値を確認して、シートに書く形に整える */
function validate_(input, settings) {
  input = input || {};
  const name = String(input.name || '');
  if (settings.users.indexOf(name) < 0) throw new Error('名前「' + name + '」は設定シートにありません');

  let measuredAt;
  try {
    measuredAt = Utilities.parseDate(String(input.measuredAt || ''), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm");
  } catch (e) {
    throw new Error('測定日時が正しくありません');
  }
  if (measuredAt.getTime() > Date.now() + 10 * 60 * 1000) throw new Error('未来の日時は登録できません');

  const cat = String(input.cat || '');
  if (cat !== '朝' && cat !== '夜') throw new Error('朝か夜を選んでください');

  const sys = toInt_(input.sys, '最高血圧', RANGE.sys, true);
  const dia = toInt_(input.dia, '最低血圧', RANGE.dia, true);
  const pulse = toInt_(input.pulse, '脈拍', RANGE.pulse, false);
  if (sys <= dia) throw new Error('最高血圧は最低血圧より大きい値にしてください');
  const temp = toTemp_(input.temp);

  let memo = String(input.memo || '').trim().slice(0, 100);
  if (/^[=+\-@]/.test(memo)) memo = "'" + memo; // 数式として解釈されないようにする

  return { name: name, measuredAt: measuredAt, cat: cat, sys: sys, dia: dia, pulse: pulse, memo: memo, temp: temp };
}

/** 体温（任意）。小数1桁に丸める */
function toTemp_(value) {
  if (value === '' || value === null || value === undefined) return null;
  const n = Number(value);
  if (!isFinite(n) || n < RANGE.temp.min || n > RANGE.temp.max) {
    throw new Error('体温は ' + RANGE.temp.min + '.0〜' + RANGE.temp.max + '.0 で入力してください');
  }
  return Math.round(n * 10) / 10;
}

function toInt_(value, label, range, required) {
  if (value === '' || value === null || value === undefined) {
    if (required) throw new Error(label + 'を入力してください');
    return null;
  }
  const n = Number(value);
  if (!Number.isInteger(n) || n < range.min || n > range.max) {
    throw new Error(label + 'は ' + range.min + '〜' + range.max + ' の整数で入力してください');
  }
  return n;
}
