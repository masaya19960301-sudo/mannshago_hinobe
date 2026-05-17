/**
 * 配送センター日延集計アプリ - サーバーサイドロジック
 */

const SHEET_RECORDS = 'records';
const SHEET_REASONS = 'reasons';
const SHEET_SETTINGS = 'settings';

const RECORDS_HEADERS = [
  'タイムスタンプ',
  '配送日',
  '伝票No',
  '販売金額合計',
  '店舗名',
  '販売担当者',
  'お客様名',
  '理由',
  '日延後の配送日種別',
  '日延後の配送日',
  '販売金額内訳'
];

const POST_DELAY_DECIDED = '決定';
const POST_DELAY_PENDING = '未定';

const DEFAULT_REASONS = [
  '在庫不足',
  '配送車両満車',
  'お客様都合',
  '商品準備遅延',
  'その他'
];

/**
 * ウェブアプリのエントリーポイント
 */
function doGet(e) {
  try {
    initializeSpreadsheet_();
    maybePurgeOldRecords_();
    const allowedPages = ['form', 'search', 'dashboard', 'settings'];
    const requested = (e && e.parameter && e.parameter.page) ? String(e.parameter.page) : 'form';
    const page = allowedPages.indexOf(requested) !== -1 ? requested : 'form';
    const template = HtmlService.createTemplateFromFile('index');
    template.initialPage = page;
    return template.evaluate()
      .setTitle('配送センター日延集計')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput('<h1>エラー</h1><p>' + escapeHtml_(err.message) + '</p>');
  }
}

/**
 * HTMLファイルをインクルードするためのヘルパー
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * スプレッドシートを取得
 */
function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) {
    throw new Error('スクリプトプロパティ SPREADSHEET_ID が設定されていません');
  }
  return SpreadsheetApp.openById(id);
}

/**
 * シートを取得。無ければ作成して返す。
 * getSheetByName と insertSheet の不整合（既存だがgetでnull等）にも耐える。
 */
function getOrCreateSheet_(ss, name) {
  // まず全シートをリストアップして名前一致を確実に探す
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() === name) return sheets[i];
  }
  try {
    return ss.insertSheet(name);
  } catch (err) {
    // 競合などで作成失敗した場合、もう一度リストから探す
    const retry = ss.getSheets();
    for (let i = 0; i < retry.length; i++) {
      if (retry[i].getName() === name) return retry[i];
    }
    throw err;
  }
}

/**
 * 初回実行時にシートとヘッダーを自動生成
 */
function initializeSpreadsheet_() {
  const ss = getSpreadsheet_();

  // records シート
  const recordsSheet = getOrCreateSheet_(ss, SHEET_RECORDS);
  if (recordsSheet.getLastRow() === 0) {
    recordsSheet.getRange(1, 1, 1, RECORDS_HEADERS.length).setValues([RECORDS_HEADERS]);
    recordsSheet.setFrozenRows(1);
    recordsSheet.getRange(1, 1, 1, RECORDS_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#F97316')
      .setFontColor('#FFFFFF');
  } else {
    // 既存スプレッドシートのスキーマ拡張（マイグレーション）
    const lastCol = recordsSheet.getLastColumn();
    if (lastCol < RECORDS_HEADERS.length) {
      const missing = RECORDS_HEADERS.slice(lastCol);
      recordsSheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing])
        .setFontWeight('bold')
        .setBackground('#F97316')
        .setFontColor('#FFFFFF');
    }
  }

  // reasons シート
  const reasonsSheet = getOrCreateSheet_(ss, SHEET_REASONS);
  if (reasonsSheet.getLastRow() === 0) {
    reasonsSheet.getRange(1, 1).setValue('理由').setFontWeight('bold').setBackground('#F97316').setFontColor('#FFFFFF');
    const values = DEFAULT_REASONS.map(r => [r]);
    reasonsSheet.getRange(2, 1, values.length, 1).setValues(values);
    reasonsSheet.setFrozenRows(1);
  }

  // settings シート
  const settingsSheet = getOrCreateSheet_(ss, SHEET_SETTINGS);
  if (settingsSheet.getLastRow() === 0) {
    settingsSheet.getRange(1, 1, 1, 2).setValues([['キー', '値']])
      .setFontWeight('bold').setBackground('#F97316').setFontColor('#FFFFFF');
    settingsSheet.setFrozenRows(1);
  }
}

/**
 * レコードを保存
 */
function saveRecord(data) {
  try {
    if (!data || typeof data !== 'object') {
      return { success: false, error: '入力データが不正です' };
    }
    const validationError = validateRecord_(data);
    if (validationError) {
      return { success: false, error: validationError };
    }

    initializeSpreadsheet_();
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEET_RECORDS);

    // コピー複製モード時は重複チェックをスキップして同一伝票Noも許可
    if (!data.allowDuplicate) {
      const dup = findDuplicateSlip_(data.slips, null);
      if (dup) {
        return { success: false, duplicate: dup };
      }
    }

    const row = buildRow_(data, new Date());
    sheet.appendRow(row);
    return { success: true, message: '登録しました' };
  } catch (err) {
    return { success: false, error: '保存中にエラーが発生しました: ' + err.message };
  }
}

/**
 * 既存レコードを更新
 */
function updateRecord(rowIndex, data) {
  try {
    const idx = Number(rowIndex);
    if (!idx || idx < 2) {
      return { success: false, error: '更新対象の行が不正です' };
    }
    if (!data || typeof data !== 'object') {
      return { success: false, error: '入力データが不正です' };
    }
    const validationError = validateRecord_(data);
    if (validationError) {
      return { success: false, error: validationError };
    }

    initializeSpreadsheet_();
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEET_RECORDS);
    if (idx > sheet.getLastRow()) {
      return { success: false, error: '指定行が存在しません' };
    }

    const dup = findDuplicateSlip_(data.slips, idx);
    if (dup) {
      return { success: false, duplicate: dup };
    }

    // 既存タイムスタンプ（列A）を保持
    const existingTimestamp = sheet.getRange(idx, 1).getValue();
    const ts = existingTimestamp instanceof Date ? existingTimestamp : new Date();
    const row = buildRow_(data, ts);
    sheet.getRange(idx, 1, 1, RECORDS_HEADERS.length).setValues([row]);
    return { success: true, message: '更新しました' };
  } catch (err) {
    return { success: false, error: '更新中にエラーが発生しました: ' + err.message };
  }
}

/**
 * 単一伝票Noの重複チェック（フォーム入力中の即時チェック用）
 */
function checkDuplicateSlip(slipNo, excludeRowIndex) {
  try {
    const s = String(slipNo == null ? '' : slipNo).trim();
    if (!/^\d{6}$/.test(s)) {
      return { success: true, duplicate: null };
    }
    const exclude = excludeRowIndex ? Number(excludeRowIndex) : null;
    const dup = findDuplicateSlip_([{ slipNo: s, amount: 0 }], exclude);
    return { success: true, duplicate: dup };
  } catch (err) {
    return { success: false, error: err.message, duplicate: null };
  }
}

/**
 * 伝票Noの重複を検出する。
 * @param {Array<{slipNo:string,amount:number}>} slips - 入力された伝票
 * @param {number|null} excludeRowIndex - 自身のレコード（編集時）を除外する行番号
 * @return {object|null} 重複が見つかれば { slipNo, record }、なければ null
 */
function findDuplicateSlip_(slips, excludeRowIndex) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_RECORDS);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const numRows = sheet.getLastRow() - 1;
  const values = sheet.getRange(2, 1, numRows, RECORDS_HEADERS.length).getValues();
  const tz = Session.getScriptTimeZone();
  const searchSet = {};
  slips.forEach(s => { searchSet[String(s.slipNo).trim()] = true; });

  for (let i = 0; i < values.length; i++) {
    const rowIndex = i + 2;
    if (excludeRowIndex && rowIndex === excludeRowIndex) continue;
    const slipCell = String(values[i][2] || '');
    const existingSlips = slipCell.split('/');
    for (let j = 0; j < existingSlips.length; j++) {
      const candidate = existingSlips[j].trim();
      if (!candidate) continue;
      if (searchSet[candidate]) {
        return {
          slipNo: candidate,
          record: {
            rowIndex: rowIndex,
            timestamp: values[i][0] instanceof Date ? values[i][0].toISOString() : String(values[i][0] || ''),
            deliveryDate: values[i][1] instanceof Date ? Utilities.formatDate(values[i][1], tz, 'yyyy-MM-dd') : String(values[i][1] || ''),
            slipNumbers: slipCell,
            totalAmount: Number(values[i][3]) || 0,
            storeName: String(values[i][4] == null ? '' : values[i][4]),
            salesPerson: String(values[i][5] || ''),
            customerName: String(values[i][6] || ''),
            reason: String(values[i][7] || ''),
            postDelayDateStatus: String(values[i][8] || ''),
            postDelayDate: values[i][9] instanceof Date ? Utilities.formatDate(values[i][9], tz, 'yyyy-MM-dd') : String(values[i][9] == null ? '' : values[i][9]),
            amountBreakdown: String(values[i][10] || '')
          }
        };
      }
    }
  }
  return null;
}

/**
 * フォームデータから行配列を組み立てる
 */
function buildRow_(data, timestamp) {
  const slipNumbers = data.slips.map(s => String(s.slipNo).trim()).join('/');
  const amountBreakdown = data.slips.map(s => String(Number(s.amount))).join('/');
  const totalAmount = data.slips.reduce((sum, s) => sum + Number(s.amount), 0);
  const status = data.postDelayDateStatus === POST_DELAY_DECIDED ? POST_DELAY_DECIDED : POST_DELAY_PENDING;
  return [
    timestamp,
    data.deliveryDate,
    slipNumbers,
    totalAmount,
    Number(data.storeName),
    String(data.salesPerson).trim(),
    String(data.customerName).trim(),
    String(data.reason).trim(),
    status,
    String(data.postDelayDate || '').trim(),
    amountBreakdown
  ];
}

/**
 * 入力データのバリデーション
 */
function validateRecord_(data) {
  if (!data.deliveryDate) return '配送日が未入力です';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(data.deliveryDate)) return '配送日の形式が不正です';

  if (!Array.isArray(data.slips) || data.slips.length === 0) return '伝票Noが未入力です';
  if (data.slips.length > 10) return '伝票Noは最大10件までです';

  for (let i = 0; i < data.slips.length; i++) {
    const s = data.slips[i];
    if (!s || !s.slipNo) return (i + 1) + '行目の伝票Noが未入力です';
    if (!/^\d{6}$/.test(String(s.slipNo))) return (i + 1) + '行目の伝票Noは6桁の数字で入力してください';
    if (s.amount === undefined || s.amount === null || s.amount === '') return (i + 1) + '行目の販売金額が未入力です';
    if (isNaN(Number(s.amount)) || Number(s.amount) < 0) return (i + 1) + '行目の販売金額が不正です';
  }

  if (!data.storeName) return '店舗名が未入力です';
  if (!/^\d+$/.test(String(data.storeName))) return '店舗名は数字のみで入力してください';

  if (!data.salesPerson) return '販売担当者が未入力です';
  if (containsWhitespace_(data.salesPerson)) return '販売担当者に空白文字を含めることはできません';

  if (!data.customerName) return 'お客様名が未入力です';
  if (containsWhitespace_(data.customerName)) return 'お客様名に空白文字を含めることはできません';

  if (!data.reason) return '理由が未入力です';

  if (data.postDelayDateStatus !== POST_DELAY_DECIDED && data.postDelayDateStatus !== POST_DELAY_PENDING) {
    return '日延後の配送日（決定／未定）を選択してください';
  }
  if (data.postDelayDateStatus === POST_DELAY_DECIDED) {
    if (!data.postDelayDate || !/^\d{4}-\d{2}-\d{2}$/.test(data.postDelayDate)) {
      return '日延後の配送日を選択してください';
    }
  } else {
    if (!data.postDelayDate || String(data.postDelayDate).trim().length === 0) {
      return '日延後の配送日（未定時の内容）を入力してください';
    }
  }

  return null;
}

/**
 * 空白文字（半角・全角スペース・タブ）を含むかチェック
 */
function containsWhitespace_(str) {
  return /[\s　]/.test(String(str));
}

/**
 * 全レコードを取得
 */
function getRecords() {
  try {
    initializeSpreadsheet_();
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEET_RECORDS);
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, records: [] };
    }
    const numRows = sheet.getLastRow() - 1;
    const values = sheet.getRange(2, 1, numRows, RECORDS_HEADERS.length).getValues();
    const tz = Session.getScriptTimeZone();
    const records = values.map((row, i) => {
      const deliveryDate = row[1] instanceof Date
        ? Utilities.formatDate(row[1], tz, 'yyyy-MM-dd')
        : String(row[1] || '');
      const postDelayRaw = row[9];
      const postDelayDate = postDelayRaw instanceof Date
        ? Utilities.formatDate(postDelayRaw, tz, 'yyyy-MM-dd')
        : String(postDelayRaw == null ? '' : postDelayRaw);
      return {
        rowIndex: i + 2,
        timestamp: row[0] instanceof Date ? row[0].toISOString() : String(row[0] || ''),
        deliveryDate: deliveryDate,
        slipNumbers: String(row[2] || ''),
        totalAmount: Number(row[3]) || 0,
        storeName: String(row[4] == null ? '' : row[4]),
        salesPerson: String(row[5] || ''),
        customerName: String(row[6] || ''),
        reason: String(row[7] || ''),
        postDelayDateStatus: String(row[8] || ''),
        postDelayDate: postDelayDate,
        amountBreakdown: String(row[10] || '')
      };
    });
    return { success: true, records: records };
  } catch (err) {
    return { success: false, error: err.message, records: [] };
  }
}

/**
 * 理由の選択肢を取得
 */
function getReasons() {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEET_REASONS);
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, reasons: [] };
    }
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    const reasons = values.map(r => String(r[0])).filter(r => r.length > 0);
    return { success: true, reasons: reasons };
  } catch (err) {
    return { success: false, error: err.message, reasons: [] };
  }
}

/**
 * 理由を追加
 */
function addReason(reason) {
  try {
    if (!reason || typeof reason !== 'string') {
      return { success: false, error: '理由が未入力です' };
    }
    const trimmed = reason.trim();
    if (trimmed.length === 0) {
      return { success: false, error: '空の選択肢は登録できません' };
    }
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEET_REASONS);
    if (!sheet) {
      initializeSpreadsheet_();
    }
    const existing = getReasons().reasons;
    if (existing.indexOf(trimmed) !== -1) {
      return { success: false, error: '既に登録されている理由です' };
    }
    ss.getSheetByName(SHEET_REASONS).appendRow([trimmed]);
    return { success: true, reasons: getReasons().reasons };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * 理由を削除（インデックスは reasons 配列上の 0-based）
 */
function deleteReason(index) {
  try {
    const idx = Number(index);
    if (isNaN(idx) || idx < 0) {
      return { success: false, error: '不正なインデックスです' };
    }
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEET_REASONS);
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: false, error: '削除対象がありません' };
    }
    const rowToDelete = idx + 2; // ヘッダー行 + 1-based index
    if (rowToDelete > sheet.getLastRow()) {
      return { success: false, error: '存在しないインデックスです' };
    }
    sheet.deleteRow(rowToDelete);
    return { success: true, reasons: getReasons().reasons };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * 配送日を基準に丸3年経過したレコードを削除する。
 * @return {number} 削除件数
 */
function purgeOldRecords() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_RECORDS);
  if (!sheet || sheet.getLastRow() < 2) return 0;

  const now = new Date();
  // threshold = 今日からちょうど3年前の同日 0:00。これ以前（同日含む）を削除対象とする。
  const threshold = new Date(now.getFullYear() - 3, now.getMonth(), now.getDate());

  const lastRow = sheet.getLastRow();
  const numRows = lastRow - 1;
  const dateValues = sheet.getRange(2, 2, numRows, 1).getValues();
  const rowsToDelete = [];
  for (let i = 0; i < dateValues.length; i++) {
    const v = dateValues[i][0];
    let d = null;
    if (v instanceof Date) {
      d = v;
    } else if (typeof v === 'string') {
      const m = v.match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (m) d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    }
    if (!d || isNaN(d.getTime())) continue; // 解析できないものは保護
    if (d.getTime() <= threshold.getTime()) {
      rowsToDelete.push(i + 2);
    }
  }
  // インデックスのずれを避けるため後ろから削除
  for (let j = rowsToDelete.length - 1; j >= 0; j--) {
    sheet.deleteRow(rowsToDelete[j]);
  }
  return rowsToDelete.length;
}

/**
 * 1日1回だけ purgeOldRecords を実行する（doGet から呼ばれる）。
 */
function maybePurgeOldRecords_() {
  try {
    const props = PropertiesService.getScriptProperties();
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (props.getProperty('LAST_PURGE_DATE') === today) return;
    purgeOldRecords();
    props.setProperty('LAST_PURGE_DATE', today);
  } catch (err) {
    // パージ失敗で本体UIを止めないよう握りつぶす
  }
}

/**
 * 毎日自動パージするためのインストール型トリガを登録する。
 * Apps Script エディタで一度だけ手動実行してください。
 */
function installDailyPurgeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'purgeOldRecords') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('purgeOldRecords')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();
}

/**
 * HTMLエスケープ
 */
function escapeHtml_(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
