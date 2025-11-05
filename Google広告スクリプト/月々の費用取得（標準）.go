function main() {
  // ① 書き込みたいGoogleスプレッドシートの情報を設定
  const SPREADSHEET_URL = "スプレッドシートのURLをここに貼り付けてください";
  const SHEET_NAME = "費用"; // 対象のシート名

  // ---ここから自動処理---
  try {
    const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
    }

    let startDate;
    const today = new Date();

    // ② シートのデータ状況を確認
    if (sheet.getLastRow() < 2) {
      // データがない場合
      startDate = "20180101";
      sheet.clear();
      sheet.appendRow(["日付", "費用", "キャンペーン名"]);
    } else {
      // データがある場合

      // --- ▼ 修正箇所 ▼ ---
      // getMonth() が時差で「8月」と誤認識するバグを回避する

      const lastDate = new Date(sheet.getRange(sheet.getLastRow(), 1).getValue());

      // (1) スプレッドシートのタイムゾーンを取得 (例: "Asia/Tokyo")
      const spreadsheetTimeZone = spreadsheet.getSpreadsheetTimeZone();

      // (2) getMonth() を使わず、シートのタイムゾーン基準で「年」と「月」を数値として取得
      const lastYear = parseInt(Utilities.formatDate(lastDate, spreadsheetTimeZone, "yyyy"), 10);
      const lastMonth = parseInt(Utilities.formatDate(lastDate, spreadsheetTimeZone, "MM"), 10); // 9月なら「9」が返る

      Logger.log(`シート最終行の日付を「${lastYear}年 ${lastMonth}月」と認識しました。`);

      // (3) 取得した「月」の数値 (lastMonth) を使って翌月を計算
      let nextYear = lastYear;
      let nextMonth = lastMonth + 1; // 9 + 1 = 10 (10月)

      if (nextMonth > 12) {
        nextMonth = 1; // 1月にリセット
        nextYear += 1; // 翌年へ
      }

      // (4) JavaScriptの Date オブジェクトは月が 0 始まり (0=1月) のため、-1 する
      // 10月(nextMonth=10) を指定する場合、 10-1 = 9 を渡す
      const nextMonthDate = new Date(nextYear, nextMonth - 1, 1);
      // --- ▲ 修正箇所 ▲ ---

      // (5) 元のコードの続き
      // nextMonthDate は「10月1日」のオブジェクトになっている
      const y = nextMonthDate.getFullYear();
      const m = ("0" + (nextMonthDate.getMonth() + 1)).slice(-2); // (9 + 1) = 10
      startDate = y + m + "01"; // "20231001" になる
    }

    // ③ 取得期間の終了日を「先月の末日」に設定 (元のコードのまま)
    const firstDayOfThisMonth = new Date(today.getFullYear(), today.getMonth(), 1);
    const lastDayOfPreviousMonth = new Date(firstDayOfThisMonth.getTime() - 1);
    const y = lastDayOfPreviousMonth.getFullYear();
    const m = ("0" + (lastDayOfPreviousMonth.getMonth() + 1)).slice(-2);
    const d = ("0" + lastDayOfPreviousMonth.getDate()).slice(-2);
    const endDate = y + m + d;

    Logger.log(`レポート取得期間: ${startDate} (開始) から ${endDate} (終了)`);

    if (startDate > endDate) {
      Logger.log("更新する新しいデータはありませんでした。処理を終了します。");
      return;
    }

    // ④ キャンペーン別の費用を取得 (元のコードのまま)
    const campaignCostQuery = `SELECT Month, CampaignName, Cost FROM CAMPAIGN_PERFORMANCE_REPORT DURING ${startDate},${endDate}`;
    const report = AdsApp.report(campaignCostQuery);
    const rows = report.rows();

    // ⑤ 取得したデータを追記用の配列に格納 (元のコードのまま)
    const dataToAppend = [];
    while (rows.hasNext()) {
      const row = rows.next();
      const date = new Date(row["Month"]);
      const cost = parseFloat(row["Cost"].replace(/,/g, ''));
      const campaignName = row["CampaignName"];

      if (cost > 0) {
        dataToAppend.push([date, cost, campaignName]);
      }
    }

    if (dataToAppend.length === 0) {
      Logger.log("期間内に追記する新しい広告費用データが見つかりませんでした。");
      return;
    }

    // ⑥ データをシートに一括で追記 (元のコードのまま)
    sheet.getRange(sheet.getLastRow() + 1, 1, dataToAppend.length, 3).setValues(dataToAppend);

    // ⑦ ソートとフォーマット設定 (元のコードのまま)
    if (sheet.getLastRow() > 1) {
      const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3);
      range.sort({column: 1, ascending: true});
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).setNumberFormat("yyyy/mm/dd");
      sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).setNumberFormat("#,##0");
    }

    Logger.log(`${dataToAppend.length}件のデータをシートに追記し、更新しました。`);

  } catch (e) {
    Logger.log(`エラーが発生しました: ${e.message} (Line: ${e.lineNumber}) スタック: ${e.stack}`);
  }
}