/**
 * 【コンバージョン内訳レポート用・日次更新版】
 * 毎日、未取得の日別のコンバージョン内訳を追記します。
 * 「広告チャネルタイプ」列を追加し、デバイス名・チャネル名の表記を統一。
 */
 function main() {

  // ▼▼【要設定】▼▼ 記録したいスプレッドシートのURLを貼り付けてください
  const SPREADSHEET_URL = '（スプレッドシートのURLを入力）';

  // ▼設定▼ 記録先のシート名を指定してください
  const SHEET_NAME = 'コンバージョンデータ'; // シート名は変更OK

  // --- スプレッドシートの準備 ---
  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) { sheet = spreadsheet.insertSheet(SHEET_NAME); }

  const japaneseHeaders = ['日付', 'デバイス', 'キャンペーン名', 'コンバージョンアクション名', 'コンバージョン数', '広告チャネルタイプ'];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(japaneseHeaders);
  }

  // --- 取得期間を決定するロジック ---
  const accountTimezone = AdsApp.currentAccount().getTimeZone();

  const today = new Date();
  const yesterday = new Date();
  yesterday.setDate(today.getDate() - 1);
  const endDateString = Utilities.formatDate(yesterday, accountTimezone, "yyyyMMdd");

  let startDateString;

  if (sheet.getLastRow() <= 1) {
    console.log('データがないため、昨日1日分のデータを取得します。');
    startDateString = endDateString;
  } else {
    console.log('通常実行：未取得の期間のデータを取得します。');
    const lastDate = new Date(sheet.getRange(sheet.getLastRow(), 1).getValue());
    const startDate = new Date(lastDate);
    startDate.setDate(lastDate.getDate() + 1);

    if (startDate > yesterday) {
      console.log('データは既に最新です。');
      return;
    }
    startDateString = Utilities.formatDate(startDate, accountTimezone, "yyyyMMdd");
  }

  const query = `
    SELECT
      segments.date,
      segments.device,
      campaign.name,
      segments.conversion_action_name,
      metrics.conversions,
      campaign.advertising_channel_type
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

      // ★★★【変更点】デバイス名とチャネル名の表記を統一する ★★★
      let device = row['segments.device'];
      if (device === 'Mobile devices with full browsers') device = 'MOBILE';
      if (device === 'Computers') device = 'DESKTOP';
      if (device === 'Tablets with full browsers') device = 'TABLET';

      const channel = row['campaign.advertising_channel_type'].toUpperCase();

      dataToWrite.push([
        row['segments.date'],
        device, // 統一したデバイス名
        row['campaign.name'],
        row['segments.conversion_action_name'],
        row['metrics.conversions'],
        channel // 統一したチャネル名
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