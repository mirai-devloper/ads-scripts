/**
 * 【広告グループ版・年指定・追記・自動ソート】
 * 指定した1年分の広告グループデータを取得し、シートに追記後、全体を日付順に並べ替えます。
 */
 function main() {

  // ▼▼【要設定】▼▼ 取得したい年を西暦で指定してください
  const TARGET_YEAR = 2025; // 例: 2025年

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
    // APIフィールド名（英語）- ConversionActionName を削除
    const apiFields = [
      'Date',
      'CampaignId',
      'CampaignName',
      'AdGroupId',
      'AdGroupName',
      'AdGroupStatus',
      'AdGroupType',
      'Device',
      'Conversions',
      'Impressions',
      'Clicks',
      'Cost'
    ];

    // スプレッドシートのヘッダー行（日本語）- コンバージョンアクション名 を削除
    const japaneseHeaders = [
      '日付',
      'キャンペーンID',
      'キャンペーン名',
      '広告グループID',
      '広告グループ名',
      '広告グループステータス',
      '広告グループタイプ',
      'デバイス',
      'コンバージョン',
      '表示回数',
      'クリック数',
      '費用'
    ];

    // ヘッダー行がなければ書き込む
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(japaneseHeaders);
      console.log('ヘッダー行を新規設定しました。');
    }

    // --- 取得期間を決定するロジック ---
    const accountTimezone = AdsApp.currentAccount().getTimeZone();
    const startDate = new Date(TARGET_YEAR, 0, 1);
    const endDate = new Date(TARGET_YEAR, 11, 31);
    const startDateString = Utilities.formatDate(startDate, accountTimezone, "yyyyMMdd");
    const endDateString = Utilities.formatDate(endDate, accountTimezone, "yyyyMMdd");
    const reportPeriod = startDateString + ',' + endDateString;

    console.log('取得期間: ' + reportPeriod);

    // --- レポートを取得 ---
    // 広告グループレポート（ADGROUP_PERFORMANCE_REPORT）から取得
    const query = 'SELECT ' + apiFields.join(', ') +
                  ' FROM ADGROUP_PERFORMANCE_REPORT' +
                  ' DURING ' + reportPeriod;

    const report = AdsApp.report(query);
    const rows = report.rows();

    const dataToWrite = [];
    while (rows.hasNext()) {
      const row = rows.next();
      const newRow = [];

      // データを1つずつ処理し、必要に応じて表記を統一
      for (let i = 0; i < apiFields.length; i++) {
        const fieldName = apiFields[i];
        let value = row[fieldName];

        // デバイス名の表記を統一
        if (fieldName === 'Device') {
          if (value === 'Mobile devices with full browsers') value = 'MOBILE';
          if (value === 'Computers') value = 'DESKTOP';
          if (value === 'Tablets with full browsers') value = 'TABLET';
          if (value === 'Other') value = 'OTHER';
          if (value === 'Devices streaming video content to TV screens') value = 'STREAMING_TV'; // APIのバージョンによって名称が異なる場合がある
        }

        // ★★★【変更点】費用の単位変換ロジックを削除 ★★★
        // 現在はマイクロ円で返されないため、単位変換は不要。
        // ただし、カンマ区切りの文字列が返される可能性を考慮し、数値への変換は残す。
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
      // データをシートの末尾に一括で追記
      sheet.getRange(sheet.getLastRow() + 1, 1, dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);
      console.log(dataToWrite.length + '件のデータを追記しました。');

      // シート全体を日付で並べ替え（ヘッダー行を除く）
      if (sheet.getLastRow() > 1) {
        // 日付列は8番目 (H列)
        const dateColumnIndex = 1;
        const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
        dataRange.sort({column: dateColumnIndex, ascending: true});
        console.log('シート全体を日付順に並べ替えました。');
      }
    } else {
      console.log('期間内に記録対象のデータはありませんでした。');
    }

  } catch (e) {
    console.error('スクリプトの実行中にエラーが発生しました: ' + e.toString());
    // エラー詳細をログに出力
    console.error('エラー詳細: ' + e.stack);
  }
}
