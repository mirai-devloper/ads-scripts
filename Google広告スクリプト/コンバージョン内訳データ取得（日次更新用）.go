/**
 * 【コンバージョン内訳レポート用・日次更新版】
 * 毎日、未取得の日別のコンバージョン内訳を追記します。
 */
function main() {

  // ▼▼【要設定】▼▼ 記録したいスプレッドシートのURLを貼り付けてください
  const SPREADSHEET_URL = 'スプレッドシートのURLをここに貼り付けてください';

  // ▼設定▼ 記録先のシート名を指定してください
  const SHEET_NAME = 'コンバージョン内訳データ'; // シート名は変更OK

  // --- スプレッドシートの準備 ---
  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) { sheet = spreadsheet.insertSheet(SHEET_NAME); }

  // --- ヘッダー行の準備 ---
  const japaneseHeaders = ['日付', 'デバイス', 'キャンペーン名', 'コンバージョンアクション名', 'コンバージョン数'];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(japaneseHeaders);
  }

  // --- 取得期間を決定するロジック ---
  const accountTimezone = AdsApp.currentAccount().getTimeZone();

  // 終了日は常に昨日
  const today = new Date();
  const yesterday = new Date();
  yesterday.setDate(today.getDate() - 1);
  const endDateString = Utilities.formatDate(yesterday, accountTimezone, "yyyyMMdd");

  let startDateString;

  // シートの最終行から開始日を決定する
  // データが1行もなければ（ヘッダーのみ）、開始日も昨日とする
  if (sheet.getLastRow() <= 1) {
    console.log('データがないため、昨日1日分のデータを取得します。');
    startDateString = endDateString;
  } else {
    console.log('通常実行：未取得の期間のデータを取得します。');
    const lastDate = new Date(sheet.getRange(sheet.getLastRow(), 1).getValue());
    const startDate = new Date(lastDate);
    startDate.setDate(lastDate.getDate() + 1); // 開始日は記録されている最終日の「翌日」

    // 既にデータが最新の場合は処理を終了
    if (startDate > yesterday) {
      console.log('データは既に最新です。');
      return;
    }
    startDateString = Utilities.formatDate(startDate, accountTimezone, "yyyyMMdd");
  }

  // --- GAQLクエリを作成 ---
  const query = `
    SELECT
      segments.date,
      segments.device,
      campaign.name,
      segments.conversion_action_name,
      metrics.conversions
    FROM campaign
    WHERE
      segments.date >= '${startDateString}'
      AND segments.date <= '${endDateString}'
      AND metrics.conversions > 0
    ORDER BY
      segments.date ASC,
      campaign.name ASC
  `;

  try {
    // --- レポートを取得し、シートに書き込み ---
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
      sheet.getRange(sheet.getLastRow() + 1, 1, dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);
      console.log(`${dataToWrite.length}件のデータを記録しました。`);
    } else {
      console.log('期間内に記録対象のデータはありませんでした。');
    }
  } catch (e) {
    console.error('コンバージョン内訳レポートのエラー:', e.toString());
  }
}