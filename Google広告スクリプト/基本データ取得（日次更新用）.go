/**
 * 【最終完成版・金額修正済み】
 * 毎日、未取得のデータを自動で補完します。
 * 金額はAPIから取得した円単位の値をそのまま利用します。
 */
 function main() {

  // ▼▼【要設定】▼▼ 記録したいスプレッドシートのURLを貼り付けてください
  const SPREADSHEET_URL = 'スプレッドシートのURLをここに貼り付けてください';

  // ▼設定▼ 記録先のシート名を指定してください
  const SHEET_NAME = '基本データ'; // シート名は変更OK

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
      'Date', 'Device', 'AccountDescriptiveName', 'CampaignId', 'CampaignName', 'CampaignStatus',
      'AdvertisingChannelType', 'BiddingStrategyType', 'Impressions', 'Clicks', 'Cost', 'Ctr', 'AverageCpc',
      'Conversions', 'ConversionRate', 'CostPerConversion', 'AllConversions', 'AllConversionRate', 'CostPerAllConversion',
      'ViewThroughConversions', 'Interactions', 'InteractionRate', 'AverageCost', 'AverageCpm', 'AverageCpv',
      'SearchImpressionShare', 'SearchTopImpressionShare', 'SearchAbsoluteTopImpressionShare', 'SearchBudgetLostImpressionShare', 'SearchRankLostImpressionShare',
      'ContentImpressionShare', 'ContentBudgetLostImpressionShare', 'ContentRankLostImpressionShare',
      'VideoViews', 'VideoViewRate', 'VideoQuartile25Rate', 'VideoQuartile50Rate', 'VideoQuartile75Rate', 'VideoQuartile100Rate'
    ];
    const japaneseHeaders = [
      '日付', 'デバイス', 'アカウント名', 'キャンペーンID', 'キャンペーン名', 'キャンペーンステータス',
      '広告チャネルタイプ', '入札戦略タイプ', '表示回数', 'クリック数', 'ご利用額', 'クリック率', '平均クリック単価',
      'コンバージョン', 'コンバージョン率', 'コンバージョン単価', 'すべてのコンバージョン', 'すべてのコンバージョン率', 'すべてのコンバージョン単価',
      'ビュースルーコンバージョン', 'インタラクション', 'インタラクション率', '平均費用', '平均CPM', '平均CPV',
      '検索IS', '検索TOP IS', '検索Abs.TOP IS', '検索IS損失率(予算)', '検索IS損失率(ランク)',
      'コンテンツIS', 'コンテンツIS損失率(予算)', 'コンテンツIS損失率(ランク)',
      '動画再生回数', '動画再生率', '動画再生25%', '動画再生50%', '動画再生75%', '動画再生100%'
    ];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(japaneseHeaders);
      console.log('ヘッダー行を新規設定しました。');
    }

    // --- 取得期間を決定するロジック ---
    const accountTimezone = AdsApp.currentAccount().getTimeZone();
    let reportPeriod = '';

    const today = new Date();
    const yesterday = new Date();
    yesterday.setDate(today.getDate() - 1);
    const endDateString = Utilities.formatDate(yesterday, accountTimezone, "yyyyMMdd");

    if (sheet.getLastRow() <= 1) {
      console.log('データがないため、昨日1日分のデータを取得します。');
      const startDateString = endDateString;
      reportPeriod = startDateString + ',' + endDateString;
    } else {
      console.log('通常実行（デイリー更新）を開始します。');
      const lastDate = new Date(sheet.getRange(sheet.getLastRow(), 1).getValue());
      const startDate = new Date(lastDate);
      startDate.setDate(lastDate.getDate() + 1);

      if (startDate > yesterday) {
        console.log('データは既に最新です。処理を終了します。');
        return;
      }

      const startDateString = Utilities.formatDate(startDate, accountTimezone, "yyyyMMdd");
      reportPeriod = startDateString + ',' + endDateString;
    }

    console.log('取得期間: ' + reportPeriod);

    // --- レポートを取得 ---
    const query =
      'SELECT ' + apiFields.join(', ') + ' ' +
      'FROM CAMPAIGN_PERFORMANCE_REPORT ' +
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

        if (typeof value === 'number' && !isFinite(value)) {
          value = 0;
        }

        // --- 表記統一（デバイス名・チャネル名） ---
        if (fieldName === 'Device') {
          if (value === 'Mobile devices with full browsers') value = 'MOBILE';
          if (value === 'Computers') value = 'DESKTOP';
          if (value === 'Tablets with full browsers') value = 'TABLET';
          if (value === 'Other') value = 'OTHER';
          if (value === 'Devices streaming video content to TV screens') value = 'STREAMING_TV';
        }
        if (fieldName === 'AdvertisingChannelType') {
          if (value) value = value.toUpperCase();
        }

        newRow.push(value);
      }
      dataToWrite.push(newRow);
    }

    if (dataToWrite.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);
      console.log(dataToWrite.length + '件のデータを記録しました。');
    } else {
      console.log('期間内に記録対象のデータはありませんでした。');
    }

  } catch (e) {
    console.error('スクリプトの実行中にエラーが発生しました: ' + e.toString());
  }
}