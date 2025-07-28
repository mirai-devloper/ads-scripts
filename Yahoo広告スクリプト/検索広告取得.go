/************************************
 * 設定項目
 ************************************/
// データを出力したいGoogleスプレッドシートのURL
const SPREADSHEET_URL = 'YOUR_SPREADSHEET_URL';

// データを取得したい年（西暦4桁）
const TARGET_YEAR = 2024;

// 出力先のシート名
const SHEET_NAME = '検索広告（YSA）';


/************************************
 * メイン処理
 ************************************/
function main() {
  // --- 1. レポート項目の定義 ---
  // APIにリクエストするフィールド名（英語）
  const reportFields = [
    'ACCOUNT_ID', 'ACCOUNT_NAME',
    'CAMPAIGN_ID', 'CAMPAIGN_NAME',
    'ADGROUP_ID', 'ADGROUP_NAME',
    'AD_ID', 'AD_NAME', 'AD_TYPE',
    'FINAL_URL',
    'DEVICE',
    'DAY',
    'IMPS', 'CLICKS', 'COST', 'AVG_CPC',
    'CONVERSIONS', 'CONV_RATE', 'COST_PER_CONV', 'VALUE_PER_CONV', 'CONV_VALUE'
  ];

  // 日本語ヘッダーとの対応表
  const headerMapping = {
    'ACCOUNT_ID': 'アカウントID', 'ACCOUNT_NAME': 'アカウント名',
    'CAMPAIGN_ID': 'キャンペーンID', 'CAMPAIGN_NAME': 'キャンペーン名',
    'ADGROUP_ID': '広告グループID', 'ADGROUP_NAME': '広告グループ名',
    'AD_ID': '広告ID', 'AD_NAME': '広告名', 'AD_TYPE': '広告タイプ',
    'FINAL_URL': '最終リンク先URL',
    'DEVICE': 'デバイス',
    'DAY': '日',
    'IMPS': 'インプレッション数', 'CLICKS': 'クリック数', 'COST': 'コスト', 'AVG_CPC': '平均CPC',
    'CONVERSIONS': 'コンバージョン数', 'CONV_RATE': 'コンバージョン率', 'COST_PER_CONV': 'コンバージョン単価', 'VALUE_PER_CONV': 'コンバージョンあたりの価値', 'CONV_VALUE': 'コンバージョンの価値'
  };
  // --------------------------------

  // --- 2. レポート取得の準備 ---
  const startDate = TARGET_YEAR + '0101';
  const endDate = TARGET_YEAR + '1231';

  // AdsUtilities を使用したセレクターを作成
  const selector = {
    accountId: AdsUtilities.getCurrentAccountId(), // 実行中のアカウントを取得
    reportType: 'AD', // 広告レポート
    fields: reportFields,
    reportDateRangeType: 'CUSTOM_DATE',
    dateRange: {
      startDate: startDate,
      endDate: endDate
    },
    reportSkipColumnHeader: 'TRUE' // ヘッダーは自前で付けるので、レポートからは除外
  };

  // --- 3. レポートの取得と処理 ---
  try {
    Logger.log('レポート取得を開始します...');
    // AdsUtilities を使ってレポートを取得
    const report = AdsUtilities.getSearchReport(selector);

    if (report && report.reports && report.reports[0].rows) {
      const reportData = report.reports[0].rows;
      Logger.log(reportData.length + '行のデータを取得しました。');

      // ▼▼▼ 追加: 日付でデータを並び替え ▼▼▼
      const dayIndex = reportFields.indexOf('DAY');
      if (dayIndex !== -1) {
        Logger.log('日付でデータを並び替えます...');
        // reportDataは二次元配列なので、日付カラムを基準にソート
        reportData.sort((a, b) => {
          // a[dayIndex] と b[dayIndex] は 'YYYY/MM/DD' 形式の文字列
          // そのまま文字列として比較することで日付順にソートできる
          if (a[dayIndex] < b[dayIndex]) return -1;
          if (a[dayIndex] > b[dayIndex]) return 1;
          return 0;
        });
      }
      // ▲▲▲ 並び替えここまで ▲▲▲

      writeDataToSheet(reportData, reportFields, headerMapping);
      Logger.log('スプレッドシートへの書き込みが完了しました。');
    } else {
      Logger.log('レポートデータを取得できませんでした。');
      if(report.errors){
        // エラー内容を分かりやすく整形してログに出力
        const errorMessages = report.errors.map(e => `[${e.errorCode}] ${e.message} (${JSON.stringify(e.details)})`).join('\n');
        Logger.log('エラー詳細:\n' + errorMessages);
      }
    }
  } catch (e) {
    Logger.log('レポート取得中にエラーが発生しました: ' + e);
    throw e;
  }
}

/************************************
 * スプレッドシートにデータを書き込む関数
 ************************************/
function writeDataToSheet(reportData, reportFields, headerMapping) {
  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }
  sheet.clear();

  // 日本語ヘッダーを作成
  const japaneseHeaders = reportFields.map(field => headerMapping[field] || field);
  sheet.getRange(1, 1, 1, japaneseHeaders.length).setValues([japaneseHeaders]);

  // データ部を書き込み
  if (reportData.length > 0) {
    sheet.getRange(2, 1, reportData.length, reportData[0].length).setValues(reportData);
  }
}
