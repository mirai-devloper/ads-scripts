/************************************
 * 設定項目
 ************************************/
// データを出力したいGoogleスプレッドシートのURL
const SPREADSHEET_URL = 'YOUR_SPREADSHEET_URL';

// データ出力先のシート名
const DATA_SHEET_NAME = '検索広告（YSA）';

// 実行履歴を記録するシート名（自動作成されます）
const HISTORY_SHEET_NAME = '実行履歴';


/************************************
 * メイン処理
 ************************************/
function main() {
  // --- 1. レポート項目の定義 ---
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

  // --- 2. レポート取得期間の決定 ---
  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const lastDate = getLastExecutionDate(spreadsheet);
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);

  // 最後に取得した日の翌日を開始日とする
  const startDate = new Date(lastDate);
  startDate.setDate(startDate.getDate() + 1);

  // 最終取得日がなく、初回実行の場合
  if (!lastDate) {
    startDate.setTime(yesterday.getTime());
    Logger.log('初回実行のため、昨日のデータを取得します。');
  }

  // 開始日が昨日より後の場合は、処理の必要なし
  if (startDate > yesterday) {
    Logger.log('取得対象の新しいデータはありません。処理を終了します。');
    return;
  }

  const startDateStr = formatDate(startDate, 'YYYYMMDD');
  const endDateStr = formatDate(yesterday, 'YYYYMMDD');
  Logger.log('レポート取得期間: ' + startDateStr + ' - ' + endDateStr);

  // --- 3. レポート取得の準備 ---
  const selector = {
    accountId: AdsUtilities.getCurrentAccountId(),
    reportType: 'AD',
    fields: reportFields,
    reportDateRangeType: 'CUSTOM_DATE',
    dateRange: {
      startDate: startDateStr,
      endDate: endDateStr
    },
    reportSkipColumnHeader: 'TRUE'
  };

  // --- 4. レポートの取得と処理 ---
  try {
    Logger.log('レポート取得を開始します...');
    const report = AdsUtilities.getSearchReport(selector);

    if (report && report.reports && report.reports[0].rows) {
      const reportData = report.reports[0].rows;
      Logger.log(reportData.length + '行のデータを取得しました。');

      const dayIndex = reportFields.indexOf('DAY');
      if (dayIndex !== -1) {
        Logger.log('日付でデータを並び替えます...');
        reportData.sort((a, b) => (a[dayIndex] < b[dayIndex] ? -1 : 1));
      }

      appendDataToSheet(spreadsheet, reportData, reportFields, headerMapping);
      updateExecutionDate(spreadsheet, yesterday); // 成功したので最終実行日を更新
      Logger.log('スプレッドシートへの書き込みと実行日の更新が完了しました。');

    } else {
      Logger.log('レポート期間内にデータが見つかりませんでした。');
      if (report.errors) {
        const errorMessages = report.errors.map(e => `[${e.errorCode}] ${e.message} (${JSON.stringify(e.details)})`).join('\n');
        Logger.log('エラー詳細:\n' + errorMessages);
      } else {
        // データが0件でも、期間の更新は行う
        updateExecutionDate(spreadsheet, yesterday);
        Logger.log('データは0件でしたが、実行日は更新しました。');
      }
    }
  } catch (e) {
    Logger.log('レポート取得中にエラーが発生しました: ' + e);
    throw e;
  }
}

/************************************
 * 実行履歴シートから最終実行日を取得する関数
 ************************************/
function getLastExecutionDate(spreadsheet) {
  let historySheet = spreadsheet.getSheetByName(HISTORY_SHEET_NAME);
  if (!historySheet) {
    Logger.log('実行履歴シートが見つからないため、作成します。');
    historySheet = spreadsheet.insertSheet(HISTORY_SHEET_NAME);
    historySheet.getRange('A1').setValue('最終データ取得日');
  }
  const lastDate = historySheet.getRange('A2').getValue();
  return lastDate instanceof Date ? lastDate : null;
}

/************************************
 * 実行履歴シートに最終実行日を記録する関数
 ************************************/
function updateExecutionDate(spreadsheet, date) {
  let historySheet = spreadsheet.getSheetByName(HISTORY_SHEET_NAME);
  if (!historySheet) {
    historySheet = spreadsheet.insertSheet(HISTORY_SHEET_NAME);
    historySheet.getRange('A1').setValue('最終データ取得日');
  }
  historySheet.getRange('A2').setValue(date);
  Logger.log('最終データ取得日を ' + formatDate(date, 'YYYY/MM/DD') + ' に更新しました。');
}

/************************************
 * スプレッドシートにデータを追記する関数
 ************************************/
function appendDataToSheet(spreadsheet, reportData, reportFields, headerMapping) {
  let dataSheet = spreadsheet.getSheetByName(DATA_SHEET_NAME);
  if (!dataSheet) {
    dataSheet = spreadsheet.insertSheet(DATA_SHEET_NAME);
  }

  // ヘッダー行がなければ書き込む
  if (dataSheet.getLastRow() === 0) {
    const japaneseHeaders = reportFields.map(field => headerMapping[field] || field);
    dataSheet.getRange(1, 1, 1, japaneseHeaders.length).setValues([japaneseHeaders]);
  }

  // データ部を最終行に追記
  if (reportData.length > 0) {
    dataSheet.getRange(dataSheet.getLastRow() + 1, 1, reportData.length, reportData[0].length).setValues(reportData);
  }
}

/************************************
 * 日付オブジェクトを文字列にフォーマットする関数
 ************************************/
function formatDate(dateObj, format) {
  const year = dateObj.getFullYear();
  const month = ('0' + (dateObj.getMonth() + 1)).slice(-2);
  const day = ('0' + dateObj.getDate()).slice(-2);
  if (format === 'YYYYMMDD') {
    return `${year}${month}${day}`;
  }
  return `${year}/${month}/${day}`;
}
