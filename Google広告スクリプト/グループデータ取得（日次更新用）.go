/**
 * 【広告グループ版・日次更新専用】
 * 常に、シートの最終行の翌日から昨日までの未取得データを追記します。
 */
 function main() {

  // ▼▼【要設定】▼▼ 記録したいスプレッドシートのURLを貼り付けてください
  const SPREADSHEET_URL = 'スプレッドシートのURLをここに貼り付けてください';

  // ▼設定▼ 記録先のシート名を指定してください
  const SHEET_NAME = 'グループデータ';

  // --- スプレッドシートの準備 ---
  if (SPREADSHEET_URL.indexOf('https://docs.google.com/spreadsheets/d/') === -1) {
    throw new Error('スプレッドシートのURLを正しく設定してください。');
  }
  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  try {
    const apiFields = [
      'Date', 'CampaignId', 'CampaignName', 'AdGroupId', 'AdGroupName',
      'AdGroupStatus', 'AdGroupType', 'Device', 'Conversions',
      'Impressions', 'Clicks', 'Cost'
    ];
    const japaneseHeaders = [
      '日付', 'キャンペーンID', 'キャンペーン名', '広告グループID', '広告グループ名',
      '広告グループステータス', '広告グループタイプ', 'デバイス', 'コンバージョン',
      '表示回数', 'クリック数', '費用'
    ];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(japaneseHeaders);
      console.log('ヘッダー行を新規設定しました。');
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

    console.log('取得期間: ' + startDateString + ' - ' + endDateString);

    // --- レポートを取得 ---
    const query =
      'SELECT ' + apiFields.join(', ') + ' ' +
      'FROM ADGROUP_PERFORMANCE_REPORT ' +
      'DURING ' + startDateString + ',' + endDateString + ' ' +
      'ORDER BY Date ASC';

    const report = AdsApp.report(query);
    const rows = report.rows();

    const dataToWrite = [];
    while (rows.hasNext()) {
      const row = rows.next();
      const newRow = [];

      for (let i = 0; i < apiFields.length; i++) {
        const fieldName = apiFields[i];
        let value = row[fieldName];

        // デバイス名の表記を統一
        if (fieldName === 'Device') {
          if (value === 'Mobile devices with full browsers') value = 'MOBILE';
          if (value === 'Computers') value = 'DESKTOP';
          if (value === 'Tablets with full browsers') value = 'TABLET';
          if (value === 'Other') value = 'OTHER';
          if (value === 'Devices streaming video content to TV screens') value = 'STREAMING_TV';
        }

        // 費用のカンマ区切りを数値に変換（単位変換はしない）
        if (fieldName === 'Cost') {
            if (typeof value === 'string' && value.includes(',')) {
               value = parseFloat(value.replace(/,/g, ''));
            }
        }

        newRow.push(value);
      }
      dataToWrite.push(newRow);
    }

    if (dataToWrite.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);
      console.log(dataToWrite.length + '件のデータを追記しました。');
    } else {
      console.log('期間内に記録対象のデータはありませんでした。');
    }

  } catch (e) {
    console.error('スクリプトの実行中にエラーが発生しました: ' + e.toString());
    console.error('エラー詳細: ' + e.stack);
  }
}