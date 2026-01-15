// ▼▼▼▼▼ ユーザー設定項目 ▼▼▼▼▼
// 1. スプレッドシートのID（URLからコピーします）
var SPREADSHEET_ID = "ここにスプレッドシートのIDを入力してください";
// 2. 書き込みたいシート名（例: "シート1"）
var SHEET_NAME = "ここに書き込みたいシート名を入力してください";
// ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

function main() {

  // --- 必須設定のチェック ---
  if (SPREADSHEET_ID === "ここにスプレッドシートのIDを入力してください" || SHEET_NAME === "ここに書き込みたいシート名を入力してください") {
    Logger.log("エラー: スクリプトの3行目または5行目の SPREADSHEET_ID と SHEET_NAME を設定してください。");
    return;
  }

  // 1. WHERE句を削除し、すべてのコンバージョン アクションを取得
  var query = `
    SELECT conversion_action.name, conversion_action.type
    FROM conversion_action
    ORDER BY conversion_action.name
  `;

  var iterator = AdsApp.search(query);

  Logger.log("--- コンバージョン名一覧 絞り込み処理 ---");

  if (!iterator.hasNext()) {
    Logger.log("コンバージョン アクションが見つかりませんでした。");
    return;
  }

  // 2. スプレッドシートに書き出すためのデータ配列を準備
  var sheetData = [
    ["Conversion Action Name", "Conversion Type"]
  ];

  while (iterator.hasNext()) {
    var row = iterator.next();
    var name = row.conversionAction.name;
    var type = row.conversionAction.type;

    // 3. ★★★ 修正点 ★★★
    // タイプが 'GOOGLE_ANALYTICS' で「始まらない」ものだけを絞り込む
    // ( ! は「ではない」(NOT) の意味)
    if ( !type.startsWith('GOOGLE_ANALYTICS') ) {
      // "GOOGLE_ANALYTICS..." で始まらなければ、書き込みデータに追加
      Logger.log("追加: " + name + " (Type: " + type + ")");
      sheetData.push([name, type]);
    } else {
      // "GOOGLE_ANALYTICS..." で始まる場合はスキップ
      Logger.log("スキップ (GA): " + name + " (Type: " + type + ")");
    }
  }

  // 4. 既存のスプレッドシートに書き込む
  if (sheetData.length > 1) { // ヘッダー以外にデータがある場合
    try {
      var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheet = spreadsheet.getSheetByName(SHEET_NAME);

      if (sheet == null) {
        Logger.log("エラー: 指定されたシート名 '" + SHEET_NAME + "' が見つかりません。");
        return;
      }

      sheet.clear();
      sheet.getRange(1, 1, sheetData.length, sheetData[0].length).setValues(sheetData);
      sheet.getRange(1, 1, sheetData.length, sheetData[0].length).setBorder(true, true, true, true, true, true);
      sheet.getRange(1, 1, 1, sheetData[0].length).setFontWeight("bold");

      Logger.log("---------------------------------");
      Logger.log("スプレッドシート '" + spreadsheet.getName() + "' のシート '" + SHEET_NAME + "' に書き込みました。");
      Logger.log("URL: " + spreadsheet.getUrl());

    } catch (e) {
      Logger.log("スプレッドシートへの書き込み中にエラーが発生しました: " + e);
    }

  } else {
    Logger.log("---------------------------------");
    Logger.log("スプレッドシートは更新されませんでした（該当するデータがありません）。");
    // もしデータが0件だった場合、シートをクリアする
    try {
      var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheet = spreadsheet.getSheetByName(SHEET_NAME);
      if (sheet != null) {
        sheet.clear();
        sheet.getRange(1, 1, 1, 2).setValues([["(該当データなし)", "(該当データなし)"]]);
      }
    } catch (e) {
      // エラー処理は省略
    }
  }
}