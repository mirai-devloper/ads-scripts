/**
 * 【日次更新・追記・自動ソート版】
 * 未取得のコンバージョン内訳データを取得します。
 * ★P-MAXのアセットグループと、通常の広告グループを両方取得します。
 * ★データの取得は最大で「前々日」までとします。
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

  // 「広告グループ名」を「グループ名」に変更し、アセットグループ名も含むようにします
  const japaneseHeaders = [
    '日付', 'デバイス', 'キャンペーン名', 'キャンペーンID', 'グループ名', 'グループID',
    'グループステータス', 'グループタイプ', 'コンバージョンアクション名', 'コンバージョン数', '広告チャネルタイプ'
  ];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(japaneseHeaders);
  }

  // --- ここから修正箇所：日次更新用の取得期間を決定 ---
  const accountTimezone = AdsApp.currentAccount().getTimeZone();
  let startDateString, endDateString;

  const today = new Date();
  const dayBeforeYesterday = new Date(today);
  dayBeforeYesterday.setDate(today.getDate() - 2);
  endDateString = Utilities.formatDate(dayBeforeYesterday, accountTimezone, "yyyyMMdd");

  if (sheet.getLastRow() <= 1) { // ヘッダー行のみ、またはデータがない場合
    console.log('データがないため、昨日1日分のデータを取得します。(本スクリプトでは前々日までのデータが対象です)');
    // 最初の取得日は手動で設定するか、特定の期間を指定することを推奨します。
    // ここでは、仮に30日前から取得する例を示しますが、環境に合わせて調整してください。
    const tempStartDate = new Date();
    tempStartDate.setDate(tempStartDate.getDate() - 30);
    startDateString = Utilities.formatDate(tempStartDate, accountTimezone, "yyyyMMdd");
  } else { // データがある場合は、最終日の翌日から取得
    console.log('通常実行：未取得の期間のデータを取得します。');
    const lastDateValue = sheet.getRange(sheet.getLastRow(), 1).getValue();
    const lastDate = new Date(lastDateValue);
    const startDate = new Date(lastDate);
    startDate.setDate(lastDate.getDate() + 1);

    // 既にデータが最新の場合は処理を終了
    if (startDate > dayBeforeYesterday) {
      console.log('データは既に最新です。');
      return;
    }
    startDateString = Utilities.formatDate(startDate, accountTimezone, "yyyyMMdd");
  }
  // --- ここまで修正箇所 ---

  console.log(`取得期間: ${startDateString} から ${endDateString}`);

  const dataToWrite = [];

  try {
    // --- Query 1: P-MAX以外の広告グループデータを取得 ---
    const adGroupQuery = `
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
        AND campaign.advertising_channel_type != 'PERFORMANCE_MAX'
    `;
    console.log('P-MAX以外のキャンペーンの広告グループデータを取得しています...');
    const adGroupReport = AdsApp.report(adGroupQuery);
    const adGroupRows = adGroupReport.rows();
    while (adGroupRows.hasNext()) {
      const row = adGroupRows.next();
      let device = row['segments.device'];
      if (device === 'Mobile devices with full browsers') device = 'MOBILE';
      if (device === 'Computers') device = 'DESKTOP';
      if (device === 'Tablets with full browsers') device = 'TABLET';
      if (device === 'Other') device = 'OTHER';
      if (device === 'Connected TV') device = 'STREAMING_TV';

      dataToWrite.push([
        row['segments.date'],
        device,
        row['campaign.name'],
        row['campaign.id'],
        row['ad_group.name'], // 広告グループ名
        row['ad_group.id'],
        row['ad_group.status'],
        row['ad_group.type'],
        row['segments.conversion_action_name'],
        row['metrics.conversions'],
        row['campaign.advertising_channel_type'].toUpperCase()
      ]);
    }
    console.log(`${dataToWrite.length}件の広告グループデータを処理しました。`);


    // --- Query 2: P-MAXのアセットグループデータを取得 ---
    const pmaxQuery = `
      SELECT
        segments.date,
        segments.device,
        campaign.name,
        campaign.id,
        asset_group.name,
        segments.conversion_action_name,
        metrics.conversions,
        campaign.advertising_channel_type
      FROM asset_group
      WHERE
        segments.date >= '${startDateString}' AND segments.date <= '${endDateString}'
        AND metrics.conversions > 0
        AND campaign.advertising_channel_type = 'PERFORMANCE_MAX'
    `;
    console.log('P-MAXキャンペーンのアセットグループデータを取得しています...');
    const pmaxReport = AdsApp.report(pmaxQuery);
    const pmaxRows = pmaxReport.rows();
    const pmaxDataCount = dataToWrite.length;

    while (pmaxRows.hasNext()) {
      const row = pmaxRows.next();
      let device = row['segments.device'];
      if (device === 'Mobile devices with full browsers') device = 'MOBILE';
      if (device === 'Computers') device = 'DESKTOP';
      if (device === 'Tablets with full browsers') device = 'TABLET';
      if (device === 'Other') device = 'OTHER';
      if (device === 'Connected TV') device = 'STREAMING_TV';

      dataToWrite.push([
        row['segments.date'],
        device,
        row['campaign.name'],
        row['campaign.id'],
        row['asset_group.name'], // アセットグループ名
        '(P-MAX)', // グループID
        '(P-MAX)', // グループステータス
        '(P-MAX)', // グループタイプ
        row['segments.conversion_action_name'],
        row['metrics.conversions'],
        row['campaign.advertising_channel_type'].toUpperCase()
      ]);
    }
    console.log(`${dataToWrite.length - pmaxDataCount}件のアセットグループデータを処理しました。`);


    // --- データをスプレッドシートに書き込み ---
    if (dataToWrite.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);
      console.log(`合計 ${dataToWrite.length}件のデータを追記しました。`);

      if (sheet.getLastRow() > 1) {
        const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
        dataRange.sort({column: 1, ascending: true});
        console.log('シート全体を日付順に並べ替えました。');
      }
    } else {
        console.log('期間内に記録対象のデータはありませんでした。');
    }
  } catch (e) {
    console.error('レポートのエラー:', e.toString());
  }
}