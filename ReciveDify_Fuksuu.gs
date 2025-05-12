function doPost(e) {
  Logger.log(JSON.stringify(e)); // ログ出力

  try {
    const data = JSON.parse(e.postData.contents);
    const spreadsheet = SpreadsheetApp.openById('1aI8nI6CJiIxv8lnzlgwOvtS7o7BdkN02myDK6CaQbkQ');
    const sheet = getOrCreateSheet(spreadsheet, data.arg4);

    if (data.summary !== undefined) {
      // ★ summaryが送られてきた場合だけ、D列2行目のデータを取得
      const previousDValue = sheet.getLastRow() >= 2 ? sheet.getRange(2, 4).getValue() : "データなし";

      return ContentService.createTextOutput(
        JSON.stringify({ result: previousDValue })
      ).setMimeType(ContentService.MimeType.JSON);

    } else if (data.arg1 !== undefined && data.arg2 !== undefined && data.arg3 !== undefined && data.arg5 !== undefined) {
      // ★ 即レスポンス（登録完了）を返す
      const response = ContentService.createTextOutput(
        JSON.stringify({ result: "登録完了" })
      ).setMimeType(ContentService.MimeType.JSON);
      Utilities.sleep(50); // 念のため少し待つ

      // ★ 以下、非同期的にスプレッドシートへ記録
      const currentDate = new Date();
      const formattedDate = Utilities.formatDate(currentDate, 'Asia/Tokyo', 'yyyy/MM/dd');
      const formattedTime = Utilities.formatDate(currentDate, 'Asia/Tokyo', 'HH:mm:ss');
      const tomorrowWork = `【${data.arg5}】`;

      const lastRow = sheet.getLastRow() + 1;
      sheet.getRange(lastRow, 1, 1, 7).setValues([
        [data.arg4, formattedDate, formattedTime, tomorrowWork, data.arg1, data.arg2, data.arg3]
      ]);

      sortSheetByDate(sheet); // ソートも非同期的に継続

      return response; // ★ UI用には即返すだけ

    } else {
      // ★ 不正リクエスト
      return ContentService.createTextOutput(
        JSON.stringify({ result: "error", message: "無効なリクエストです。" })
      ).setMimeType(ContentService.MimeType.JSON);
    }

  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({ result: "error", message: error.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);
  return sheet;
}

function sortSheetByDate(sheet) {
  try {
    const numRows = sheet.getLastRow();
    if (numRows > 1) {
      sheet.getRange(2, 1, numRows - 1, sheet.getLastColumn())
        .sort([{ column: 2, ascending: false }, { column: 3, ascending: false }]);
    }
  } catch (error) {
    // ソート失敗時は無視
  }
}


