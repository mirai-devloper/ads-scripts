/**
 * 【コンバージョン内訳レポート用・改 v2】
 * 指定した1年分のデータを追記し、シート全体を日付順に並べ替えます。
 */
 function main() {

  // ▼▼【要設定】▼▼ 取得したい年を西暦で指定してください
  const TARGET_YEAR = 2025; // 年数を指定してください

  // ▼▼【要設定】▼▼ 記録したいスプレッドシートのURLを貼り付けてください
  const SPREADSHEET_URL = 'スプレッドシートのURLをここに貼り付けてください';

  // ▼設定▼ 記録先のシート名を指定してください
  const SHEET_NAME = 'コンバージョン内訳データ'; // シート名は変更OK

  // --- スプレッドシートの準備 ---
  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  // --- ヘッダー行の準備 ---
  const japaneseHeaders = ['日付', 'デバイス', 'キャンペーン名', 'コンバージョンアクション名', 'コンバージョン数'];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(japaneseHeaders);
  }

  // --- 取得期間を決定するロジック ---
  const accountTimezone = AdsApp.currentAccount().getTimeZone();
  const startDate = new Date(TARGET_YEAR, 0, 1);
  const endDate = new Date(TARGET_YEAR, 11, 31);
  const startDateString = Utilities.formatDate(startDate, accountTimezone, "yyyyMMdd");
  const endDateString = Utilities.formatDate(endDate, accountTimezone, "yyyyMMdd");
  console.log(`取得期間: ${TARGET_YEAR}年1月1日 から ${TARGET_YEAR}年12月31日`);

  // --- GAQLクエリを作成 ---
  const query = `
    SELECT
      segments.date, segments.device, campaign.name,
      segments.conversion_action_name, metrics.conversions
    FROM campaign
    WHERE segments.date >= '${startDateString}' AND segments.date <= '${endDateString}'
      AND metrics.conversions > 0
    ORDER BY segments.date ASC, campaign.name ASC
  `;

  try {
    const report = AdsApp.report(query);
    const rows = report.rows();
    const dataToWrite = [];

    while (rows.hasNext()) {
      const row = rows.next();
      dataToWrite.push([
        row['segments.date'],
        row['segments.device'],
        row['campaign.name'],
        row['segments.conversion_action_name'],
        row['metrics.conversions']
      ]);
    }

    if (dataToWrite.length > 0) {
      // データをシートの末尾に一括で追記
      sheet.getRange(sheet.getLastRow() + 1, 1, dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);
      console.log(`${dataToWrite.length}件のデータを追記しました。`);

      // ★★★【変更点】シート全体を日付で並べ替え ★★★
      if (sheet.getLastRow() > 1) {
        // ヘッダー行（1行目）を除く全データ範囲を取得
        const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
        // 1列目（日付列）を基準に昇順で並べ替え
        dataRange.sort({column: 1, ascending: true});
        console.log('シート全体を日付順に並べ替えました。');
      }
    } else {
        console.log('期間内に記録対象のデータはありませんでした。');
    }
  } catch (e) {
    console.error('コンバージョン内訳レポートのエラー:', e.toString());
  }
}