/**
 * 【性別データ取得・年指定版】
 * 指定した1年分の性別データを追記し、シート全体を日付順に並べ替える。
 * 項目名の誤りを修正。
 */
 function main() {

  // ▼▼【要設定】▼▼ 取得したい年を西暦で指定してください
  const TARGET_YEAR = 2025; // 例: 2024年

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

    // --- 取得期間を決定するロジック ---
    const accountTimezone = AdsApp.currentAccount().getTimeZone();
    const startDate = new Date(TARGET_YEAR, 0, 1); // 1月1日
    const endDate = new Date(TARGET_YEAR, 11, 31); // 12月31日
    const startDateString = Utilities.formatDate(startDate, accountTimezone, "yyyyMMdd");
    const endDateString = Utilities.formatDate(endDate, accountTimezone, "yyyyMMdd");

    console.log(`取得期間: ${TARGET_YEAR}年1月1日 から ${TARGET_YEAR}年12月31日`);

    // ★★★【変更点】性別の項目名を「ad_group_criterion.gender.type」に修正 ★★★
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
    `;

    const report = AdsApp.report(query);
    const rows = report.rows();
    const dataToWrite = [];

    while (rows.hasNext()) {
      const row = rows.next();

      // ★★★【変更点】性別の項目名を「ad_group_criterion.gender.type」に修正 ★★★
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
      console.log(`${dataToWrite.length}件のデータを追記しました。`);

      if (sheet.getLastRow() > 1) {
        const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
        dataRange.sort({column: 1, ascending: true});
        console.log('シート全体を日付順に並べ替えました。');
      }
    } else {
      console.log('期間内に記録対象のデータはありませんでした。');
    }

  } catch (e) {
    console.error('スクリプトの実行中にエラーが発生しました: ' + e.toString());
  }
}