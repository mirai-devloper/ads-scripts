/**
 * メインの処理を実行する関数
 * コンバージョンレポートを取得します。
 */
 function fetchMetaAdsConversions() {
  try {
    // 設定ファイル(Config.gs)から設定値を参照します
    const targetYear = CV_REPORT_TARGET_YEAR;
    const sheetName = CV_REPORT_SHEET_NAME;

    const conversionData = getConversionInsights(targetYear);
    if (!conversionData || conversionData.length === 0) {
      Logger.log(`'${targetYear}'年に取得できるコンバージョンデータがありませんでした。`);
      return;
    }
    writeConversionsToSheet(conversionData, sheetName);
    Logger.log(`コンバージョンレポートの書き込みが完了しました。合計 ${conversionData.length} 件のデータを取得しました。`);
  } catch (e) {
    Logger.log('エラーが発生しました: ' + e.toString());
    SpreadsheetApp.getUi().alert('エラー: ' + e.message);
  }
}

/**
 * Meta Marketing APIからコンバージョンデータを取得する
 * @param {number} targetYear - 取得対象の西暦年
 * @return {Array} APIから取得したデータの配列
 */
function getConversionInsights(targetYear) {
  const startDate = `${targetYear}-01-01`;
  const endDate = `${targetYear}-12-31`;
  Logger.log(`データ取得期間: ${startDate} 〜 ${endDate}`);

  // 【修正】APIバージョンを最新版(v23.0)に更新
  const apiVersion = 'v23.0';
  let url = `https://graph.facebook.com/${apiVersion}/${AD_ACCOUNT_ID}/insights`;

  // ★★★ コンバージョン分析に必要な項目 ★★★
  const fields = [
    'campaign_name',
    'adset_name',
    'ad_name',
    'spend',
    'actions', // 全てのアクションを取得
    'action_values' // 全てのアクションの価値を取得
  ].join(',');

  const params = {
    'access_token': ACCESS_TOKEN,
    'level': 'ad',
    'fields': fields,
    // 【修正】エラーの原因となっていたbreakdownsを削除
    'time_range': JSON.stringify({'since': startDate, 'until': endDate}),
    'time_increment': 1, // 日別のデータを取得
    'limit': 500
  };

  let allData = [];
  let requestUrl = url + '?' + Object.keys(params).map(key => `${encodeURIComponent(key)}=${encodeURIComponent(params[key])}`).join('&');

  while (requestUrl) {
    Logger.log(`データを取得中... URL: ${requestUrl.substring(0, 150)}...`);
    const response = UrlFetchApp.fetch(requestUrl, { 'muteHttpExceptions': true, 'headers': { 'Authorization': 'Bearer ' + ACCESS_TOKEN } });
    const result = JSON.parse(response.getContentText());

    if (result.error) {
      throw new Error(`APIエラー: ${result.error.message}`);
    }

    if (result.data && result.data.length > 0) {
      allData = allData.concat(result.data);
    }

    requestUrl = (result.paging && result.paging.next) ? result.paging.next : null;
  }

  return allData;
}


/**
 * 取得したコンバージョンデータをスプレッドシートに書き込む
 * @param {Array} data - 書き込むデータ
 * @param {string} sheetName - 書き込み先のシート名
 */
function writeConversionsToSheet(data, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  sheet.clear();
  const headers = [
    '日付', 'キャンペーン名', '広告セット名', '広告名',
    'アクションタイプ', 'アクション数', 'アクションの価値(売上など)', '消化金額'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

  // 【修正】1つの広告データから複数のコンバージョン行を生成する
  const rows = data.flatMap(item => {
    if (!item.actions || item.actions.length === 0) {
      return []; // アクションがなければ行を生成しない
    }

    // 各アクションを行に変換
    return item.actions.map(action => {
      // 対応するアクションの価値を探す
      const actionValueData = item.action_values ? item.action_values.find(v => v.action_type === action.action_type) : null;
      const actionValue = actionValueData ? Number(actionValueData.value) : 0;

      return [
        item.date_start,
        item.campaign_name,
        item.adset_name,
        item.ad_name,
        action.action_type,
        Number(action.value || 0),
        actionValue,
        Number(item.spend || 0)
      ];
    });
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    sheet.autoResizeColumns(1, headers.length);
  }
}
