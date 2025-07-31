/**
 * 【年指定・追記・自動ソート版】
 * 指定した1年分のデータを追記し、シート全体を日付順に並べ替えます。
 * デバイス名・チャネル名の表記を統一し、金額の単位を修正。
 */
 function main() {

  // ▼▼【要設定】▼▼ 取得したい年を西暦で指定してください
  const TARGET_YEAR = 2025; // 年数を指定してください

  // ▼▼【要設定】▼▼ 記録したいスプレッドシートのURLを貼り付けてください
  const SPREADSHEET_URL = '（スプレッドシートのURLを入力）';

  // ▼設定▼ 記録先のシート名を指定してください
  const SHEET_NAME = '基本データ';

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
      'Date', 'Device', 'AccountDescriptiveName', 'CampaignId', 'CampaignName', 'CampaignStatus', 'AdvertisingChannelType', 'BiddingStrategyType', 'Impressions', 'Clicks', 'Cost', 'Ctr', 'AverageCpc', 'Conversions', 'ConversionRate', 'CostPerConversion', 'AllConversions', 'AllConversionRate', 'CostPerAllConversion', 'ViewThroughConversions', 'Interactions', 'InteractionRate', 'AverageCost', 'AverageCpm', 'AverageCpv', 'SearchImpressionShare', 'SearchTopImpressionShare', 'SearchAbsoluteTopImpressionShare', 'SearchBudgetLostImpressionShare', 'SearchRankLostImpressionShare', 'ContentImpressionShare', 'ContentBudgetLostImpressionShare', 'ContentRankLostImpressionShare', 'VideoViews', 'VideoViewRate', 'VideoQuartile25Rate', 'VideoQuartile50Rate', 'VideoQuartile75Rate', 'VideoQuartile100Rate'
    ];
    const japaneseHeaders = [
      '日付', 'デバイス', 'アカウント名', 'キャンペーンID', 'キャンペーン名', 'キャンペーンステータス', '広告チャネルタイプ', '入札戦略タイプ', '表示回数', 'クリック数', 'ご利用額', 'クリック率', '平均クリック単価', 'コンバージョン', 'コンバージョン率', 'コンバージョン単価', 'すべてのコンバージョン', 'すべてのコンバージョン率', 'すべてのコンバージョン単価', 'ビュースルーコンバージョン', 'インタラクション', 'インタラクション率', '平均費用', '平均CPM', '平均CPV', '検索IS', '検索TOP IS', '検索Abs.TOP IS', '検索IS損失率(予算)', '検索IS損失率(ランク)', 'コンテンツIS', 'コンテンツIS損失率(予算)', 'コンテンツIS損失率(ランク)', '動画再生回数', '動画再生率', '動画再生25%', '動画再生50%', '動画再生75%', '動画再生100%'
    ];

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
    const query = 'SELECT ' + apiFields.join(', ') + ' FROM CAMPAIGN_PERFORMANCE_REPORT DURING ' + reportPeriod + ' ORDER BY Date ASC';
    const report = AdsApp.report(query);
    const rows = report.rows();

    const dataToWrite = [];
    while (rows.hasNext()) {
      const row = rows.next();
      const newRow = [];

      // ★★★【変更点】データを1つずつ処理し、表記と単位を統一 ★★★
      for (let i = 0; i < apiFields.length; i++) {
        const fieldName = apiFields[i];
        let value = row[fieldName];

        // デバイス名の表記を統一
        if (fieldName === 'Device') {
          if (value === 'Mobile devices with full browsers') value = 'MOBILE';
          if (value === 'Computers') value = 'DESKTOP';
          if (value === 'Tablets with full browsers') value = 'TABLET';
        }

        // チャネル名を大文字に統一
        if (fieldName === 'AdvertisingChannelType') {
          if (value) value = value.toUpperCase();
        }

        newRow.push(value);
      }
      dataToWrite.push(newRow);
    }

    if (dataToWrite.length > 0) {
      // データをシートの末尾に一括で追記
      sheet.getRange(sheet.getLastRow() + 1, 1, dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);
      console.log(dataToWrite.length + '件のデータを追記しました。');

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
    console.error('スクリプトの実行中にエラーが発生しました: ' + e.toString());
  }
}