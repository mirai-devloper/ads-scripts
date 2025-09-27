/**
 * 【性別データ取得・日次更新版】
 * 初回は過去90日分、以降は毎日、日別の性別データを追記する。
 */
 function main() {

  // ▼▼【要設定】▼▼ 記録したいスプレッドシートのURLを貼り付けてください
  const SPREADSHEET_URL = 'スプレッドシートのURLをここに貼り付けてください';

  // ▼設定▼ 記録先のシート名を指定してください
  const SHEET_NAME = '性別データ';

  // --- スプレッドシートの準備 ---
  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  try {
    const japaneseHeaders = [
      '日付', 'キャンペーン名', '広告グループ名', '性別',
      '表示回数', 'クリック数', '費用', 'コンバージョン数'
    ];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(japaneseHeaders);
      console.log('ヘッダー行を新規設定しました。');
    }

    // ★★★【変更点】日次更新用の期間設定ロジック ★★★
    const accountTimezone = AdsApp.currentAccount().getTimeZone();
    let startDateString, endDateString;

    const today = new Date();
    const yesterday = new Date();
    yesterday.setDate(today.getDate() - 1);
    endDateString = Utilities.formatDate(yesterday, accountTimezone, "yyyyMMdd");

    if (sheet.getLastRow() <= 1) { // 初回実行
      console.log('初回実行：過去90日分のデータを取得します。');
      const startDate = new Date();
      startDate.setDate(today.getDate() - 90);
      startDateString = Utilities.formatDate(startDate, accountTimezone, "yyyyMMdd");
    } else { // 通常実行
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

    console.log('取得期間: ' + startDateString + ' - ' + endDateString);

    // --- GAQLクエリを作成 ---
    const query = `
      SELECT
        segments.date,
        campaign.name,
        ad_group.name,
        ad_group_criterion.gender.type,
        metrics.impressions,
        metrics.clicks,
        metrics.cost_micros,
        metrics.conversions
      FROM gender_view
      WHERE
        segments.date >= '${startDateString}'
        AND segments.date <= '${endDateString}'
      ORDER BY
        segments.date ASC,
        campaign.name ASC,
        ad_group.name ASC
    `;

    const report = AdsApp.report(query);
    const rows = report.rows();
    const dataToWrite = [];

    while (rows.hasNext()) {
      const row = rows.next();

      let gender = row['ad_group_criterion.gender.type'];
      if (gender === 'MALE') gender = '男性';
      if (gender === 'FEMALE') gender = '女性';
      if (gender === 'UNDETERMINED') gender = '不明';

      dataToWrite.push([
        row['segments.date'],
        row['campaign.name'],
        row['ad_group.name'],
        gender,
        row['metrics.impressions'],
        row['metrics.clicks'],
        row['metrics.cost_micros'] / 1000000,
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
    console.error('スクリプトの実行中にエラーが発生しました: ' + e.toString());
  }
}