/**
 * 【年指定・追記・自動ソート版】
 * 指定した1年分のコンバージョン内訳データを広告グループ単位で取得します。
 */
 function main() {

  // ▼▼【要設定】▼▼ 取得したい年を西暦で指定してください
  const TARGET_YEAR = 2025; // 例: 2024年のデータを取得する場合は 2024 と入力

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

  // ★★★【変更点】ご要望の項目リストに合わせてヘッダーを修正 ★★★
  const japaneseHeaders = [
    '日付', 'デバイス', 'キャンペーン名', 'キャンペーンID', '広告グループ名', '広告グループID',
    '広告グループステータス', '広告グループタイプ', 'コンバージョンアクション名', 'コンバージョン数', '広告チャネルタイプ'
  ];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(japaneseHeaders);
  }

  // --- 取得期間を決定するロジック ---
  const accountTimezone = AdsApp.currentAccount().getTimeZone();
  const startDate = new Date(TARGET_YEAR, 0, 1); // 1月1日
  const endDate = new Date(TARGET_YEAR, 11, 31); // 12月31日
  const startDateString = Utilities.formatDate(startDate, accountTimezone, "yyyyMMdd");
  const endDateString = Utilities.formatDate(endDate, accountTimezone, "yyyyMMdd");
  console.log(`取得期間: ${TARGET_YEAR}年1月1日 から ${TARGET_YEAR}年12月31日`);

  // ★★★【変更点】ご要望の項目リストに合わせてクエリを修正 ★★★
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

      // ★★★【変更点】書き込むデータを修正 ★★★
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

      // シート全体を日付で並べ替え
      if (sheet.getLastRow() > 1) {
        const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
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