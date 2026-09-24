/**
 * 商品マスタ同期スクリプト
 * - Driveの「マスタデータ」フォルダ内のCSV（ブランドごと）を読み込み
 * - このスプレッドシートに、ブランドごとのシートとして分割書き込み
 * - 途中で失敗・時間切れになっても、次回実行時に続きから再開する
 *
 * アプリ側は、このスプレッドシートをGoogle Sheets APIで直接読みに行く
 * （Googleログイン＋スプレッドシートの共有権限で保護されているため、
 * このスクリプト側での配信・認証チェックは不要）。
 */

const MASTER_FOLDER_ID = '1v8-GAuqztp5eluyMwlM0cO6u0jQPZZX-';

// システムの87列CSVのうち、実際に使う列だけを抜き出す（0始まりの列番号）
const COLS = [
  ['itemNo', 0], ['color', 1], ['colorName', 2], ['size', 3], ['sizeName', 4],
  ['jan', 5], ['nameFull', 6],
  ['brand', 10], ['brandName', 11], ['subBrand', 12], ['subBrandName', 13],
  ['makerItemNo', 14],
  ['item', 17], ['itemName', 18], ['subItem', 19], ['subItemName', 20],
  ['year', 21], ['season', 22],
  ['listPrice', 34], ['costPrice', 35],
];
const HEADER = COLS.map(c => c[0]);
const CHUNK_SIZE = 5000;          // 一度に書き込む行数
const TIME_BUDGET_MS = 4.5 * 60 * 1000; // 1回の実行で使ってよい時間（安全マージンを見て4.5分）

const PROP_CYCLE_FILES = 'cycleFileList'; // 今回のサイクルで処理すべきファイル名一覧
const PROP_PROCESSED = 'processedFiles';  // 処理済みファイル名一覧

/**
 * メイン処理。手動実行、またはトリガーで定期実行する。
 * 時間切れで打ち切られても、次回実行時に続きから再開する。
 */
function refreshMaster() {
  const startTime = Date.now();
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActive();
  const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);

  let cycleFiles = JSON.parse(props.getProperty(PROP_CYCLE_FILES) || 'null');
  let processed = JSON.parse(props.getProperty(PROP_PROCESSED) || '[]');

  if (!cycleFiles) {
    // 新しいサイクルの開始：対象ファイルの一覧をここで確定させる
    cycleFiles = [];
    const files = folder.getFilesByType(MimeType.CSV);
    while (files.hasNext()) cycleFiles.push(files.next().getName());
    processed = [];
    props.setProperty(PROP_CYCLE_FILES, JSON.stringify(cycleFiles));
    props.setProperty(PROP_PROCESSED, JSON.stringify(processed));
    Logger.log('新しいサイクル開始: ' + cycleFiles.join(', '));
  } else {
    Logger.log('前回の続きから再開: 処理済み ' + processed.length + ' / ' + cycleFiles.length);
  }

  for (const fileName of cycleFiles) {
    if (processed.indexOf(fileName) !== -1) continue;

    if (Date.now() - startTime > TIME_BUDGET_MS) {
      Logger.log('時間切れのため中断。次回トリガー実行時に続きから再開します。');
      return;
    }

    Logger.log('処理中: ' + fileName);
    const it = folder.getFilesByName(fileName);
    if (!it.hasNext()) {
      // フォルダから消えていた場合は処理済み扱いにしてスキップ
      processed.push(fileName);
      props.setProperty(PROP_PROCESSED, JSON.stringify(processed));
      continue;
    }
    const file = it.next();

    withRetry(() => processFile(ss, file), 3, 2000);

    processed.push(fileName);
    props.setProperty(PROP_PROCESSED, JSON.stringify(processed));
    Logger.log(fileName + ' 完了');
  }

  // 全ファイル処理完了：フォルダから無くなったブランドの古いシートを掃除
  const currentNames = cycleFiles.map(sheetNameFor);
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (name.startsWith('master_') && currentNames.indexOf(name) === -1) {
      ss.deleteSheet(sheet);
    }
  });

  props.deleteProperty(PROP_CYCLE_FILES);
  props.deleteProperty(PROP_PROCESSED);
  props.setProperty('lastRefresh', new Date().toISOString());
  Logger.log('全ブランド更新完了');
}

function sheetNameFor(fileName) {
  return 'master_' + fileName.replace(/\.csv$/i, '').slice(0, 50);
}

function processFile(ss, file) {
  const csvText = file.getBlob().getDataAsString('Shift_JIS');
  const table = Utilities.parseCsv(csvText);
  if (table.length < 2) return;
  const rows = table.slice(1).map(r => COLS.map(([, idx]) => r[idx] || ''));

  const finalName = sheetNameFor(file.getName());
  const tempName = finalName + '_tmp';

  let tempSheet = ss.getSheetByName(tempName);
  if (tempSheet) ss.deleteSheet(tempSheet);
  tempSheet = ss.insertSheet(tempName);

  writeInChunks(tempSheet, [HEADER].concat(rows));

  // 書き込みが完全に終わってから本番シートと入れ替える
  // （読み取り側 = doGet が中途半端な状態を見ないようにするため）
  const oldSheet = ss.getSheetByName(finalName);
  if (oldSheet) ss.deleteSheet(oldSheet);
  tempSheet.setName(finalName);

  Logger.log(finalName + ': ' + rows.length + '件');
}

function writeInChunks(sheet, rows) {
  for (let i = 0; i < rows.length; i += CHUNK_SIZE) {
    const chunk = rows.slice(i, i + CHUNK_SIZE);
    withRetry(() => {
      sheet.getRange(i + 1, 1, chunk.length, HEADER.length).setValues(chunk);
      SpreadsheetApp.flush();
    }, 3, 1500);
  }
}

/** 一時的なAPIエラー対策：失敗したら間隔をあけて最大retries回まで再試行 */
function withRetry(fn, retries, delayMs) {
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      return fn();
    } catch (err) {
      Logger.log('エラー（' + attempt + '回目）: ' + err);
      if (attempt === retries) throw err;
      Utilities.sleep(delayMs * attempt);
    }
  }
}

