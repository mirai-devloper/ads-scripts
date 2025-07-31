/**
 * メインの処理を実行する関数
 * 指定された1年分の総合レポートを取得します。
 */
 function fetchYearlyReport() {
  try {
    // 設定ファイル(Config.gs)から設定値を参照します
    const targetYear = YEARLY_REPORT_TARGET_YEAR;
    const sheetName = YEARLY_REPORT_SHEET_NAME;

    Logger.log(`${targetYear}年の総合レポートを取得します...`);

    const reportData = getYearlyInsights(targetYear);
    if (!reportData || reportData.length === 0) {
      Logger.log(`'${targetYear}'年に取得できるデータがありませんでした。`);
      return;
    }

    writeYearlyReportToSheet(reportData, sheetName);
    Logger.log(`年次レポートの書き込みが完了しました。合計 ${reportData.length} 件のデータを取得しました。`);

  } catch (e) {
    Logger.log('エラーが発生しました: ' + e.toString());
    SpreadsheetApp.getUi().alert('エラー: ' + e.message);
  }
}

/**
 * Meta Marketing APIから指定した1年分のインサイトデータを取得する
 * @param {number} targetYear - 取得対象の西暦年
 * @returns {Array} - 取得したデータ配列
 */
function getYearlyInsights(targetYear) {
  const startDate = `${targetYear}-01-01`;
  const endDate = `${targetYear}-12-31`;

  // 【修正】APIバージョンを最新版(v23.0)に更新
  const apiVersion = 'v23.0';
  let url = `https://graph.facebook.com/${apiVersion}/${AD_ACCOUNT_ID}/insights`;

  const fields = [
    'campaign_name','adset_name','ad_name','spend','impressions','reach','frequency','clicks','ctr','cpc','cpm',
    'inline_link_clicks','inline_link_click_ctr','cost_per_inline_link_click','inline_post_engagement','cost_per_inline_post_engagement',
    'video_p25_watched_actions','video_p50_watched_actions','video_p75_watched_actions','video_p100_watched_actions','video_avg_time_watched_actions',
    'actions','action_values','cost_per_action_type','conversions','cost_per_conversion'
  ].join(',');

  const breakdowns = ['publisher_platform', 'device_platform'].join(',');

  const params = {
    'access_token': ACCESS_TOKEN,
    'level': 'ad',
    'fields': fields,
    'breakdowns': breakdowns,
    'time_range': JSON.stringify({'since': startDate, 'until': endDate}),
    'time_increment': 1,
    'limit': 500
  };

  let allData = [];
  let requestUrl = url + '?' + Object.keys(params).map(key => `${encodeURIComponent(key)}=${encodeURIComponent(params[key])}`).join('&');

  while (requestUrl) {
    const response = UrlFetchApp.fetch(requestUrl, { 'muteHttpExceptions': true, 'headers': { 'Authorization': 'Bearer ' + ACCESS_TOKEN } });
    const result = JSON.parse(response.getContentText());

    if (result.error) throw new Error(`APIエラー: ${result.error.message}`);
    if (result.data && result.data.length > 0) allData = allData.concat(result.data);
    requestUrl = (result.paging && result.paging.next) ? result.paging.next : null;
  }
  return allData;
}

/**
 * 取得したデータをスプレッドシートに書き込む（全クリア方式）
 * @param {Array} data - 書き込むデータ
 * @param {string} sheetName - 書き込み先のシート名
 */
function writeYearlyReportToSheet(data, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  sheet.clear(); // シートをクリア

  const headers = [
    '日付', 'キャンペーン名', '広告セット名', '広告名', '配信プラットフォーム', 'デバイス',
    '消化金額', 'インプレッション数', 'リーチ数', 'フリークエンシー', 'クリック数', 'CTR(%)', 'CPC', 'CPM',
    'リンククリック数', 'リンクCTR(%)', 'リンクCPC', '投稿エンゲージメント', 'エンゲージメント単価',
    '動画再生数', '動画25%再生', '動画50%再生', '動画75%再生', '動画100%再生', '平均再生時間',
    'カート追加数', 'チェックアウト開始数', '登録完了数', 'リード獲得数', '購入数', '購入金額'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

  const rows = data.map(item => {
    const videoPlays = parseAction(item, 'video_view');
    return [
      item.date_start, item.campaign_name, item.adset_name, item.ad_name, item.publisher_platform, item.device_platform,
      Number(item.spend || 0), Number(item.impressions || 0), Number(item.reach || 0), Number(item.frequency || 0), Number(item.clicks || 0), Number(item.ctr || 0), Number(item.cpc || 0), Number(item.cpm || 0),
      Number(item.inline_link_clicks || 0), Number(item.inline_link_click_ctr || 0), Number(item.cost_per_inline_link_click || 0), Number(item.inline_post_engagement || 0), Number(item.cost_per_inline_post_engagement || 0),
      videoPlays,
      item.video_p25_watched_actions ? Number(item.video_p25_watched_actions[0].value) : 0,
      item.video_p50_watched_actions ? Number(item.video_p50_watched_actions[0].value) : 0,
      item.video_p75_watched_actions ? Number(item.video_p75_watched_actions[0].value) : 0,
      item.video_p100_watched_actions ? Number(item.video_p100_watched_actions[0].value) : 0,
      item.video_avg_time_watched_actions ? Number(item.video_avg_time_watched_actions[0].value) : 0,
      parseAction(item, 'add_to_cart'), parseAction(item, 'initiate_checkout'), parseAction(item, 'complete_registration'), parseAction(item, 'lead'),
      parseAction(item, 'omni_purchase') || parseAction(item, 'purchase'),
      parseActionValue(item, 'omni_purchase') || parseActionValue(item, 'purchase')
    ];
  });

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
}

// --- 以下は補助的な関数 ---
function parseAction(item, actionType) {
  if (!item.actions) return 0;
  const action = item.actions.find(a => a.action_type === actionType);
  return action ? Number(action.value) : 0;
}
function parseActionValue(item, actionType) {
  if (!item.action_values) return 0;
  const action = item.action_values.find(a => a.action_type === actionType);
  return action ? Number(action.value) : 0;
}
