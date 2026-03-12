/**
 * スプレッドシートを開いたときに実行される処理（カスタムメニューの追加）
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙管理ツール更新')
    .addItem('終了行を下に移動する', 'moveSupportEndedRows_AllSheets')
    .addToUi();
}

/**
 * 編集時に自動実行されるトリガー（色変更のみ担当）
 */
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();

  // 対象シートとステータス列（インデックスに+1した値）の定義
  const config = {
    "上村②": 3,
    "小野田": 3,
    "澤口": 3,
    "定井": 3,
    "宮脇": 3,
    "下間": 3,
    "GATE": 3
  };

  // 1行目（ヘッダー）や対象外シート・列の場合は何もしない
  if (row < 2 || !config[sheetName] || col !== config[sheetName]) {
    return;
  }

  // 入力されたステータス値を取得
  const statusStr = String(range.getValue() || "").trim();
  
  // 色を変更する範囲（A列〜E列）を取得
  const targetRange = sheet.getRange(row, 1, 1, 5);
  
  // ステータスに応じて背景色を設定
  if (statusStr === "アシサポ終了") {
    targetRange.setBackground("#D1ABAA");
  } else if (statusStr === "SF終了") {
    // SF終了時は既存の色を維持するため何もしない（変更なし）
  } else {
    // その他のステータスに戻した場合は白（リセット）にする
    targetRange.setBackground("#ffffff");
  }
}

/**
 * 終了行を下に移動する処理（手動、またはカスタムメニューから実行）
 */
function moveSupportEndedRows_AllSheets() {
  const config = {
    "上村②": 2, // C列のインデックス
    "小野田": 2,
    "澤口": 2,
    "定井": 2,
    "宮脇": 2,
    "下間": 2,
    "GATE": 2
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  for (const [sheetName, colIndex] of Object.entries(config)) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;

    // A列〜E列（1列目〜5列目）のみを取得
    const dataRange = sheet.getRange(1, 1, lastRow, 5);
    const values = dataRange.getValues();
    const backgrounds = dataRange.getBackgrounds();

    const headers = values[0];
    const bgHeaders = backgrounds[0];

    let rowsNormal = [];
    let rowsAssistEnd = [];
    let rowsSFEnd = [];
    let bgNormal = [];
    let bgAssistEnd = [];
    let bgSFEnd = [];

    // 行の色を保持したまま各配列に振り分ける
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const bgRow = backgrounds[i];
      const statusStr = String(row[colIndex] || "").trim();

      if (statusStr === "SF終了") {
        rowsSFEnd.push(row);
        bgSFEnd.push(bgRow);
      } else if (statusStr === "アシサポ終了") {
        rowsAssistEnd.push(row);
        bgAssistEnd.push(bgRow);
      } else {
        // 通常行（移動しない）
        rowsNormal.push(row);
        bgNormal.push(bgRow);
      }
    }

    // 並び順：ヘッダー → 通常行 → アシサポ終了 → SF終了
    const newValues = [headers].concat(rowsNormal).concat(rowsAssistEnd).concat(rowsSFEnd);
    const newBackgrounds = [bgHeaders].concat(bgNormal).concat(bgAssistEnd).concat(bgSFEnd);

    // 指定したA列〜E列の範囲のみを上書き更新（色も同時に復元）
    sheet.getRange(1, 1, newValues.length, 5).setValues(newValues);
    sheet.getRange(1, 1, newBackgrounds.length, 5).setBackgrounds(newBackgrounds);
  }
}
