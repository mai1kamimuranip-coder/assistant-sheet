/**
 * ----------------------------------------------------------------
 * ▼▼▼ シートごとの設定をここで行います ▼▼▼
 * ----------------------------------------------------------------
 *
 * 各シートに、複数のプロセス（A, B, ...）を設定できます。
 * プロセスごとに、ステータス列と時刻記録列を指定します。
 *
 * ※列番号の対応： A=1, B=2, ... , I=9, J=10, K=11, L=12, M=13, N=14 ...
 */
const SHEET_CONFIGS = {
  "小野田": {
    "プロセスA": { status: 7, start: 8, end: 9, duration: 10 }, // G列 -> H,I,J
    "プロセスB": { status: 11, start: 12, end: 13, duration: 14 } // K列 -> L,M,N
  },
  "定井": {
    "プロセスA": { status: 7, start: 8, end: 9, duration: 10 }, // G列 -> H,I,J
    "プロセスB": { status: 11, start: 12, end: 13, duration: 14 } // K列 -> L,M,N
  },
  "澤口": {
    "プロセスA": { status: 6, start: 7, end: 8, duration: 9 },   // F列 -> G,H,I
    "プロセスB": { status: 10, start: 11, end: 12, duration: 13 } // J列 -> K,L,M
  },
  "宮脇": {
    "プロセスA": { status: 5, start: 6, end: 7, duration: 8 },   // E列 -> F,G,H
    "プロセスB": { status: 9, start: 10, end: 11, duration: 12 }  // I列 -> J,K,L
  },
  "【試験用】上村": {
    "プロセスA": { status: 6, start: 7, end: 8, duration: 9 }, // G列 -> H,I,J
    "プロセスB": { status: 10, start: 11, end: 12, duration: 13 } // K列 -> L,M,N
  }
};
// ▲▲▲ 設定はここまで ▲▲▲


/**
 * スプレッドシートのセルが編集されたときに自動的に実行されるメイン関数
 * @param {Object} e - イベントオブジェクト
 */
function onEdit(e) {
  // 編集されたシートとセルの情報を取得
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // 編集されたシートが設定に存在するか確認
  const sheetConfig = SHEET_CONFIGS[sheetName];
  if (!sheetConfig || row === 1) {
    // 設定にないシート、またはヘッダー行の編集は無視
    return;
  }

  // 編集された列が、設定されているいずれかのプロセスのステータス列に一致するか探す
  for (const processName in sheetConfig) {
    const processConfig = sheetConfig[processName];
    if (col === processConfig.status) {
      // 一致するプロセスが見つかったら、その設定で処理を実行して終了
      handleStatusChange(e, sheet, row, processConfig);
      return;
    }
  }
}

/**
 * ステータス変更時の共通処理
 * @param {Object} e - イベントオブジェクト
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {number} row - 編集された行番号
 * @param {Object} config - 該当するプロセスの設定オブジェクト
 */
function handleStatusChange(e, sheet, row, config) {
  const statusValue = e.value;

  if (statusValue === "対応中") {
    const startTimeCell = sheet.getRange(row, config.start);
    if (startTimeCell.getValue() === "") {
      startTimeCell.setValue(new Date()).setNumberFormat("yyyy/mm/dd hh:mm:ss");
    }
  } else if (statusValue === "完了") {
    const now = new Date();
    sheet.getRange(row, config.end).setValue(now).setNumberFormat("yyyy/mm/dd hh:mm:ss");
    
    const startTime = sheet.getRange(row, config.start).getValue();
    if (startTime instanceof Date) {
      const durationMillis = now.getTime() - startTime.getTime();
      const durationMinutes = Math.round(durationMillis / 60000);
      sheet.getRange(row, config.duration).setValue(durationMinutes);
    }
  }
}
