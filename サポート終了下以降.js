function moveSupportEndedRows_AllSheets() {
  const config = {
    "上村②": 2,
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

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const bgRow = backgrounds[i];
      const status = row[colIndex];

      if (status === "SF終了") {
        rowsSFEnd.push(row);
        bgSFEnd.push(bgRow);
      } else if (status === "アシサポ終了") {
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

    // 指定したA列〜E列の範囲のみを上書き更新（F列以降はそのまま維持）
    sheet.getRange(1, 1, newValues.length, 5).setValues(newValues);
    sheet.getRange(1, 1, newBackgrounds.length, 5).setBackgrounds(newBackgrounds);
  }
}
