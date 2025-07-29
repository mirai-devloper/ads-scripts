// ================================================================
// ▼▼▼ お客様が設定する箇所 ▼▼▼
// ================================================================

// 1. Meta広告のアクセストークン
const ACCESS_TOKEN = '（アクセストークン）';

// 2. Meta広告のアカウントID（"act_"から始まるもの）
const AD_ACCOUNT_ID = 'act_（アカウントID）';

// 3. データを書き込むシート名
const SHEET_NAME = 'Meta広告レポート';

// ================================================================
// ▲▲▲ 設定はここまで ▲▲▲
// ================================================================


/**
 * 日次トリガーで実行するメイン関数
 */
function runDailyUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

  try {
    const { startDate, endDate } = getTargetDateRange(sheet);

    // 取得対象期間がなければ（＝昨日分まで取得済みなら）処理を終了
    if (!startDate) {
      Logger.log('データは最新の状態です。処理を終了します。');
      return;
    }

    Logger.log(`データ取得期間: ${startDate} 〜 ${endDate}`);

    // APIからレポートデータを取得
    const reportData = getInsights(startDate, endDate);
    if (!reportData || reportData.length === 0) {
      Logger.log('期間内に取得できるデータがありませんでした。');
      return;
    }

    // スプレッドシートに追記
    appendToSheet(sheet, reportData);
    Logger.log(`レポートの書き込みが完了しました。合計 ${reportData.length} 件のデータを追記しました。`);

  } catch (e) {
    Logger.log('エラーが発生しました: ' + e.toString());
    // エラー発生時にメールで通知したい場合は以下のコメントを外す
    // MailApp.sendEmail(Session.getEffectiveUser().getEmail(), '[エラー] Meta広告レポート取得', e.toString());
  }
}

/**
 * スプレッドシートの最終記録日から、取得すべき日付の範囲を決定する
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @returns {{startDate: string|null, endDate: string|null}} - 取得開始日と終了日
 */
function getTargetDateRange(sheet) {
  const lastRow = sheet.getLastRow();
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const endDate = formatDate(yesterday);

  // シートが空かヘッダーのみの場合、昨日1日分を取得
  if (lastRow < 2) {
    return { startDate: endDate, endDate: endDate };
  }

  // A列の最終行から最後の日付を取得
  const lastRecordedDateStr = sheet.getRange(lastRow, 1).getValue();
  const lastRecordedDate = new Date(lastRecordedDateStr);

  // 取得開始日を計算（最終記録日の翌日）
  const startDate = new Date(lastRecordedDate.getTime());
  startDate.setDate(startDate.getDate() + 1);

  // 最終記録日が昨日以降の場合、取得対象はない
  if (startDate > yesterday) {
    return { startDate: null, endDate: null };
  }

  return { startDate: formatDate(startDate), endDate: endDate };
}

/**
 * Meta Marketing APIから指定期間のインサイトデータを取得する
 * @param {string} startDate - 取得開始日 (YYYY-MM-DD)
 * @param {string} endDate - 取得終了日 (YYYY-MM-DD)
 * @returns {Array} - 取得したデータ配列
 */
function getInsights(startDate, endDate) {
  const apiVersion = 'v20.0';
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
 * 取得したデータをスプレッドシートに追記する
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {Array} data - 書き込むデータ
 */
function appendToSheet(sheet, data) {
  const lastRow = sheet.getLastRow();

  // ヘッダーがなければ書き込む
  if (lastRow < 1) {
    const headers = [
      '日付', 'キャンペーン名', '広告セット名', '広告名', '配信プラットフォーム', 'デバイス',
      '消化金額', 'インプレッション数', 'リーチ数', 'フリークエンシー', 'クリック数', 'CTR(%)', 'CPC', 'CPM',
      'リンククリック数', 'リンクCTR(%)', 'リンクCPC', '投稿エンゲージメント', 'エンゲージメント単価',
      '動画再生数', '動画25%再生', '動画50%再生', '動画75%再生', '動画100%再生', '平均再生時間',
      'カート追加数', 'チェックアウト開始数', '登録完了数', 'リード獲得数', '購入数', '購入金額'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }

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

// --- 以下は補助的な関数（変更なし） ---

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

function formatDate(date) {
  const y = date.getFullYear();
  const m = ('0' + (date.getMonth() + 1)).slice(-2);
  const d = ('0' + date.getDate()).slice(-2);
  return `${y}-${m}-${d}`;
}