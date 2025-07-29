// ================================================================
// ▼▼▼ お客様が設定する箇所 ▼▼▼
// ================================================================

// 1. Meta広告のアクセストークン
const ACCESS_TOKEN = '（アクセストークン）';

// 2. Meta広告のアカウントID（"act_"から始まるもの）
const AD_ACCOUNT_ID = 'act_（アカウントID）';

// 3. 取得したい年を西暦4桁で指定
const TARGET_YEAR = 2024;

// 4. データを書き込むシート名
const SHEET_NAME = 'Meta広告レポート';

// ================================================================
// ▲▲▲ 設定はここまで ▲▲▲
// ================================================================


/**
 * メインの処理を実行する関数
 */
function fetchMetaAdsReport() {
  try {
    const reportData = getInsights();
    if (!reportData || reportData.length === 0) {
      Logger.log(`'${TARGET_YEAR}'年に取得できるデータがありませんでした。`);
      return;
    }
    writeToSheet(reportData);
    Logger.log(`レポートの書き込みが完了しました。合計 ${reportData.length} 件のデータを取得しました。`);
  } catch (e) {
    Logger.log('エラーが発生しました: ' + e.toString());
    SpreadsheetApp.getUi().alert('エラー: ' + e.message);
  }
}

/**
 * Meta Marketing APIからインサイトデータを取得する（ページネーション対応）
 */
function getInsights() {
  const startDate = `${TARGET_YEAR}-01-01`;
  const endDate = `${TARGET_YEAR}-12-31`;
  Logger.log(`データ取得期間: ${startDate} 〜 ${endDate}`);

  const apiVersion = 'v20.0';
  let url = `https://graph.facebook.com/${apiVersion}/${AD_ACCOUNT_ID}/insights`;

  // ★★★ 取得したい項目（フィールド）★★★
  // 不要な項目は行頭に // を付けてコメントアウトすると、処理が速くなります。
  const fields = [
    'campaign_name',
    'adset_name',
    'ad_name',
    // --- 主要指標 ---
    'spend',              // 消化金額
    'impressions',        // インプレッション数
    'reach',              // リーチ数
    'frequency',          // フリークエンシー
    'clicks',             // 全てのクリック
    'ctr',                // クリック率（全体）
    'cpc',                // クリック単価（全体）
    'cpm',                // 1,000インプレッションあたりのコスト
    // --- エンゲージメント指標 ---
    'inline_link_clicks', // リンクのクリック数
    'inline_link_click_ctr', // リンクのクリック率
    'cost_per_inline_link_click', // リンククリック単価
    'inline_post_engagement',   // 投稿エンゲージメント
    'cost_per_inline_post_engagement', // 投稿エンゲージメント単価
    // --- 動画指標 ---
    'video_p25_watched_actions',
    'video_p50_watched_actions',
    'video_p75_watched_actions',
    'video_p100_watched_actions',
    'video_avg_time_watched_actions',
    // --- コンバージョン指標 ---
    'actions',            // 全てのアクション（購入、カート追加など）
    'action_values',      // 全てのアクションの価値（売上金額など）
    'cost_per_action_type',
    'conversions',
    'cost_per_conversion'
  ].join(',');

  // 【修正点】内訳（breakdowns）を正しく設定
  // 'publisher_platform'と'device_platform'は指標(fields)ではなく内訳(breakdowns)で指定
  const breakdowns = [
    'publisher_platform',
    'device_platform'
  ].join(',');

  const params = {
    'access_token': ACCESS_TOKEN,
    'level': 'ad',
    'fields': fields,
    'breakdowns': breakdowns, // ★ 修正
    'time_range': JSON.stringify({'since': startDate, 'until': endDate}),
    'time_increment': 1,
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
 * actions配列を解析して、特定のコンバージョン指標を抽出するヘルパー関数
 */
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

/**
 * 取得したデータをスプレッドシートに書き込む
 */
function writeToSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  sheet.clear();
  const headers = [
    '日付', 'キャンペーン名', '広告セット名', '広告名', '配信プラットフォーム', 'デバイス',
    '消化金額', 'インプレッション数', 'リーチ数', 'フリークエンシー', 'クリック数', 'CTR(%)', 'CPC', 'CPM',
    'リンククリック数', 'リンクCTR(%)', 'リンクCPC', '投稿エンゲージメント', 'エンゲージメント単価',
    '動画再生数', '動画25%再生', '動画50%再生', '動画75%再生', '動画100%再生', '平均再生時間',
    'カート追加数', 'チェックアウト開始数', '登録完了数', 'リード獲得数', '購入数', '購入金額'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

  const rows = data.map(item => {
    // コンバージョンアクションを解析
    const addToCart = parseAction(item, 'add_to_cart');
    const initiateCheckout = parseAction(item, 'initiate_checkout');
    const completeRegistration = parseAction(item, 'complete_registration');
    const lead = parseAction(item, 'lead');
    const purchase = parseAction(item, 'omni_purchase') || parseAction(item, 'purchase');
    const purchaseValue = parseActionValue(item, 'omni_purchase') || parseActionValue(item, 'purchase');
    // 【修正点】'video_plays'の代わりに'video_view'アクションタイプから動画再生数を取得
    const videoPlays = parseAction(item, 'video_view');

    return [
      item.date_start, item.campaign_name, item.adset_name, item.ad_name, item.publisher_platform, item.device_platform,
      Number(item.spend || 0), Number(item.impressions || 0), Number(item.reach || 0), Number(item.frequency || 0), Number(item.clicks || 0), Number(item.ctr || 0), Number(item.cpc || 0), Number(item.cpm || 0),
      Number(item.inline_link_clicks || 0), Number(item.inline_link_click_ctr || 0), Number(item.cost_per_inline_link_click || 0), Number(item.inline_post_engagement || 0), Number(item.cost_per_inline_post_engagement || 0),
      videoPlays, // ★ 修正
      item.video_p25_watched_actions ? Number(item.video_p25_watched_actions[0].value) : 0,
      item.video_p50_watched_actions ? Number(item.video_p50_watched_actions[0].value) : 0,
      item.video_p75_watched_actions ? Number(item.video_p75_watched_actions[0].value) : 0,
      item.video_p100_watched_actions ? Number(item.video_p100_watched_actions[0].value) : 0,
      item.video_avg_time_watched_actions ? Number(item.video_avg_time_watched_actions[0].value) : 0,
      addToCart, initiateCheckout, completeRegistration, lead, purchase, purchaseValue
    ];
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    sheet.autoResizeColumns(1, headers.length);
  }
}