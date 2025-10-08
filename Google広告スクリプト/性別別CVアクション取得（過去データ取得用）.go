/**
 * 【性別・CVアクション別データ取得・年指定版】
 * 指定した1年分の性別・コンバージョンアクション別のデータを追記し、シート全体を日付順に並べ替える。
 * ★広告チャネルタイプを追加（大文字）
 * ★データの取得は最大で「前々日」までとします。
 */
 function main() {

  // ▼▼【要設定】▼▼ 取得したい年を西暦で指定してください
  const TARGET_YEAR = 2025; // 例: 2024年

  // ▼▼【要設定】▼▼ 記録したいスプレッドシートのURLを貼り付けてください
  const SPREADSHEET_URL = 'スプレッドシートのURLをここに貼り付けてください';

  // ▼設定▼ 記録先のシート名を指定してください
  const SHEET_NAME = '性別CVアクションデータ'; // シート名を変更

  // --- スプレッドシートの準備 ---
  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  try {
    // ヘッダーに「広告チャネルタイプ」を追加
    const japaneseHeaders = [
      '日付', 'キャンペーン名', '広告チャネルタイプ', '広告グループ名', '性別',
      'コンバージョンアクション名', 'コンバージョン数'
    ];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(japaneseHeaders);
      console.log('ヘッダー行を新規設定しました。');
    }

    // --- 取得期間を決定するロジック ---
    const accountTimezone = AdsApp.currentAccount().getTimeZone();
    const startDate = new Date(TARGET_YEAR, 0, 1); // 指定年の1月1日

    // スクリプト実行日の前々日を計算
    const today = new Date();
    const dayBeforeYesterday = new Date(today);
    dayBeforeYesterday.setDate(today.getDate() - 2);

    // 取得終了日を、指定年の12月31日と前々日のうち、どちらか早い方に設定
    let endDate = new Date(TARGET_YEAR, 11, 31); // 指定年の12月31日
    if (endDate > dayBeforeYesterday) {
      endDate = dayBeforeYesterday;
    }

    // もしstartDateがendDateより後の日付になってしまう場合は、処理をスキップ
    if (startDate > endDate) {
      console.log(`指定年(${TARGET_YEAR})のデータは、まだ前々日までの範囲に達していません。`);
      return;
    }

    const startDateString = Utilities.formatDate(startDate, accountTimezone, "yyyyMMdd");
    const endDateString = Utilities.formatDate(endDate, accountTimezone, "yyyyMMdd");

    console.log(`取得期間: ${startDateString} から ${endDateString}`);

    // クエリに「campaign.advertising_channel_type」を追加
    const query = `
      SELECT
        segments.date,
        campaign.name,
        campaign.advertising_channel_type,
        ad_group.name,
        ad_group_criterion.gender.type,
        segments.conversion_action_name,
        metrics.conversions
      FROM gender_view
      WHERE
        segments.date >= '${startDateString}'
        AND segments.date <= '${endDateString}'
        AND metrics.conversions > 0
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

      // 書き込むデータに「campaign.advertising_channel_type」を追加
      dataToWrite.push([
        row['segments.date'],
        row['campaign.name'],
        row['campaign.advertising_channel_type'], // 広告チャネルタイプを追加
        row['ad_group.name'],
        gender,
        row['segments.conversion_action_name'],
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