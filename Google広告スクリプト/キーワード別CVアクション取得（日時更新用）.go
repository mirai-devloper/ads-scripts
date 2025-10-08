/**
 * 【キーワード別CVアクションレポート版・日次更新版】
 * 未取得のキーワード・コンバージョンアクション別データを追記し、シート全体を日付順に並べ替えます。
 * マッチタイプを日本語に変換。
 * ★コンバージョン数を整数に丸める処理を追加。
 */
 function main() {

  // ▼▼【要設定】▼▼ 記録したいスプレッドシートのURLを貼り付けてください
  const SPREADSHEET_URL = 'スプレッドシートのURLをここに貼り付けてください';

  // ▼設定▼ 記録先のシート名を指定してください
  const SHEET_NAME = 'キーワードCVアクションデータ'; // シート名を変更

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
    // APIフィールドから表示回数・クリック数・費用を削除し、コンバージョンアクション名を追加
    const apiFields = [
      'Date', 'Device', 'CampaignName', 'AdGroupName', 'Criteria', 'KeywordMatchType',
      'ConversionTypeName', 'Conversions', 'ConversionValue'
    ];
    // ヘッダーも上記に合わせて変更
    const japaneseHeaders = [
      '日付', 'デバイス', 'キャンペーン名', '広告グループ名', 'キーワード', 'マッチタイプ',
      'コンバージョンアクション名', 'コンバージョン数', 'コンバージョン価値'
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

    const reportPeriod = startDateString + ',' + endDateString;
    console.log('取得期間: ' + reportPeriod);
    // --- ここまで修正箇所 ---

    // --- レポートを取得 ---
    // 条件に「Conversions > 0」を追加し、CVが発生したデータのみ取得
    const query =
      'SELECT ' + apiFields.join(', ') + ' ' +
      'FROM KEYWORDS_PERFORMANCE_REPORT ' +
      'WHERE Conversions > 0 ' +
      'DURING ' + reportPeriod + ' ' +
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

        // マッチタイプを日本語に変換
        if (fieldName === 'KeywordMatchType') {
          if (value === 'Broad') value = '部分一致';
          else if (value === 'Exact') value = '完全一致';
          else if (value === 'Phrase') value = 'フレーズ一致';
        }

        // ★★★ 修正箇所: コンバージョン数に小数点がある場合、整数に丸める ★★★
        if (fieldName === 'Conversions' && typeof value === 'number') {
          value = Math.round(value);
        }

        newRow.push(value);
      }
      dataToWrite.push(newRow);
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