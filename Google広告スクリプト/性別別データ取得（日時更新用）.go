/**
 * 【性別データ取得・日次更新版】
 * 未取得の性別データを追記し、シート全体を日付順に並べ替える。
 * 項目名の誤りを修正。
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

    // --- ここから修正箇所：日次更新用の取得期間を決定 ---
    const accountTimezone = AdsApp.currentAccount().getTimeZone();
    let startDateString, endDateString;

    const today = new Date();
    const yesterday = new Date();
    yesterday.setDate(today.getDate() - 1);
    endDateString = Utilities.formatDate(yesterday, accountTimezone, "yyyyMMdd");

    if (sheet.getLastRow() <= 1) { // ヘッダー行のみ、またはデータがない場合
      console.log('データがないため、昨日1日分のデータを取得します。');
      startDateString = endDateString;
    } else { // データがある場合は、最終日の翌日から取得
      const lastDateValue = sheet.getRange(sheet.getLastRow(), 1).getValue();
      const lastDate = new Date(lastDateValue);
      const startDate = new Date(lastDate);
      startDate.setDate(lastDate.getDate() + 1);

      // 既にデータが最新の場合は処理を終了
      if (startDate > yesterday) {
        console.log('データは既に最新です。処理を終了します。');
        return;
      }
      startDateString = Utilities.formatDate(startDate, accountTimezone, "yyyyMMdd");
    }

    console.log(`取得期間: ${startDateString} から ${endDateString}`);
    // --- ここまで修正箇所 ---


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