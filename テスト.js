function moveCompletedRowsFromMultipleSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archiveSheet = ss.getSheetByName("アーカイブ");

  // 対象シートとステータス列のインデックスを定義
  const sheetsToCheck = [
    { name: "上村②", statusCol: 2 },   // F列
    { name: "宮脇", statusCol: 2 },   // D列
    { name: "定井", statusCol: 2 },   // C列
    { name: "小野田", statusCol: 2 }, // C列
    { name: "澤口", statusCol: 2 }    // C列
  ];

  sheetsToCheck.forEach(config => {
    const sheet = ss.getSheetByName(config.name);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return; // データがヘッダーだけの場合はスキップ

    const headers = data[0];
    const rowsToKeep = [headers];
    const rowsToArchive = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = row[config.statusCol];

      if (status === "サポート終了") {
        rowsToArchive.push(row);
      } else {
        rowsToKeep.push(row);
      }
    }

    // アーカイブへの追加
    if (rowsToArchive.length > 0) {
      const targetRow = archiveSheet.getLastRow() + 1;
      archiveSheet
        .getRange(targetRow, 1, rowsToArchive.length, rowsToArchive[0].length)
        .setValues(rowsToArchive);
    }

    // 元シート更新（残す行のみ）
    sheet.clearContents();
    sheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
  });
}
