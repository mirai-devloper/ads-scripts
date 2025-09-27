/**
 * 【日次更新専用版】
 * 常に、シートの最終行の翌日から昨日までの未取得データを追記します。
 */
 function main() {

  // ▼▼【要設定】▼▼ 記録したいスプレッドシートのURLを貼り付けてください
  const SPREADSHEET_URL = 'スプレッドシートのURLをここに貼り付けてください';

  // ▼設定▼ 記録先のシート名を指定してください
  const SHEET_NAME = 'CV内訳データ';

  // --- スプレッドシートの準備 ---
  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  // --- ヘッダー行の準備 ---
  const japaneseHeaders = [
    '日付', 'デバイス', 'キャンペーン名', 'キャンペーンID', '広告グループ名', '広告グループID',
    '広告グループステータス', '広告グループタイプ', 'コンバージョンアクション名', 'コンバージョン数', '広告チャネルタイプ'
  ];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(japaneseHeaders);
  }

  // ★★★【変更点】日次更新専用のシンプルな期間設定ロジック ★★★
  const accountTimezone = AdsApp.currentAccount().getTimeZone();
  let startDateString, endDateString;

  const today = new Date();
  const yesterday = new Date();
  yesterday.setDate(today.getDate() - 1);
  endDateString = Utilities.formatDate(yesterday, accountTimezone, "yyyyMMdd");

  // シートにヘッダー行しかない場合、開始日も昨日とする
  if (sheet.getLastRow() <= 1) {
    console.log('データがないため、昨日1日分のデータを取得します。');
    startDateString = endDateString;
  } else { // データがある場合は、最終日の翌日から取得
    console.log('通常実行：未取得の期間のデータを取得します。');
    const lastDate = new Date(sheet.getRange(sheet.getLastRow(), 1).getValue());
    const startDate = new Date(lastDate);
    startDate.setDate(lastDate.getDate() + 1);

    // 既にデータが最新の場合は処理を終了
    if (startDate > yesterday) {
      console.log('データは既に最新です。');
      return;
    }
    startDateString = Utilities.formatDate(startDate, accountTimezone, "yyyyMMdd");
  }

  console.log(`取得期間: ${startDateString} から ${endDateString}`);

  // --- GAQLクエリを作成 ---
  const query = `
    SELECT
      segments.date,
      segments.device,
      campaign.name,
      campaign.id,
      ad_group.name,
      ad_group.id,
      ad_group.status,
      ad_group.type,
      segments.conversion_action_name,
      metrics.conversions,
      campaign.advertising_channel_type
    FROM ad_group
    WHERE
      segments.date >= '${startDateString}' AND segments.date <= '${endDateString}'
      AND metrics.conversions > 0
    ORDER BY
      segments.date ASC,
      campaign.name ASC,
      ad_group.name ASC
  `;

  try {
    const report = AdsApp.report(query);
    const rows = report.rows();
    const dataToWrite = [];

    while (rows.hasNext()) {
      const row = rows.next();

      let device = row['segments.device'];
      if (device === 'Mobile devices with full browsers') device = 'MOBILE';
      if (device === 'Computers') device = 'DESKTOP';
      if (device === 'Tablets with full browsers') device = 'TABLET';
      if (device === 'Other') device = 'OTHER';
      if (device === 'Connected TV') device = 'STREAMING_TV';

      const channel = row['campaign.advertising_channel_type'].toUpperCase();

      dataToWrite.push([
        row['segments.date'],
        device,
        row['campaign.name'],
        row['campaign.id'],
        row['ad_group.name'],
        row['ad_group.id'],
        row['ad_group.status'],
        row['ad_group.type'],
        row['segments.conversion_action_name'],
        row['metrics.conversions'],
        channel
      ]);
    }

    if (dataToWrite.length > 0) {
      // データをシートの末尾に一括で追記
      sheet.getRange(sheet.getLastRow() + 1, 1, dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);
      console.log(`${dataToWrite.length}件のデータを追記しました。`);
    } else {
      console.log('期間内に記録対象のデータはありませんでした。');
    }
  } catch (e) {
    console.error('コンバージョン内訳レポートのエラー:', e.toString());
  }
}