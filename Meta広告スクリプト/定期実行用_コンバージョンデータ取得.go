/**
 * 日次トリガーで実行するメイン関数
 * コンバージョンレポートを更新します。
 */
 function runDailyConversionUpdate() {
  try {
    // 設定ファイル(Config.gs)からシート名を取得
    const sheetName = CV_REPORT_SHEET_NAME;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

    // スプレッドシートの記録から、取得すべき日付の範囲を決定
    const { startDate, endDate } = getTargetDateRange(sheet);

    // 取得対象期間がなければ（＝昨日分まで取得済みなら）処理を終了
    if (!startDate) {
      Logger.log('コンバージョンデータは最新の状態です。処理を終了します。');
      return;
    }

    Logger.log(`コンバージョンデータ取得期間: ${startDate} 〜 ${endDate}`);

    // APIからレポートデータを取得
    const conversionData = getConversionInsights(startDate, endDate);
    if (!conversionData || conversionData.length === 0) {
      Logger.log('期間内に取得できるコンバージョンデータがありませんでした。');
      return;
    }

    // スプレッドシートに追記
    appendConversionsToSheet(sheet, conversionData);
    Logger.log(`コンバージョンレポートの書き込みが完了しました。合計 ${conversionData.length} 件のデータを追記しました。`);

  } catch (e) {
    Logger.log('エラーが発生しました: ' + e.toString());
    SpreadsheetApp.getUi().alert('エラー: ' + e.message);
  }
}

/**
 * Meta Marketing APIから指定期間のコンバージョンデータを取得する
 * @param {string} startDate - 取得開始日 (YYYY-MM-DD)
 * @param {string} endDate - 取得終了日 (YYYY-MM-DD)
 * @return {Array} APIから取得したデータの配列
 */
function getConversionInsights(startDate, endDate) {
  // APIバージョンを最新版に設定
  const apiVersion = 'v23.0';
  let url = `https://graph.facebook.com/${apiVersion}/${AD_ACCOUNT_ID}/insights`;

  const fields = [
    'campaign_name',
    'adset_name',
    'ad_name',
    'spend',
    'actions',
    'action_values'
  ].join(',');

  const params = {
    'access_token': ACCESS_TOKEN,
    'level': 'ad',
    'fields': fields,
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
 * 取得したコンバージョンデータをスプレッドシートに追記する
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 書き込み先のシート
 * @param {Array} data - 書き込むデータ
 */
function appendConversionsToSheet(sheet, data) {
  // ヘッダーがなければ書き込む
  if (sheet.getLastRow() < 1) {
    const headers = [
      '日付', 'キャンペーン名', '広告セット名', '広告名',
      'アクションタイプ', 'アクション数', 'アクションの価値(売上など)', '消化金額'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }

  // 1つの広告データから複数のコンバージョン行を生成する
  const rows = data.flatMap(item => {
    if (!item.actions || item.actions.length === 0) {
      return [];
    }

    return item.actions.map(action => {
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
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    sheet.autoResizeColumns(1, 8);
  }
}

// --- 以下は補助的な関数 ---

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

  if (lastRow < 2) {
    return { startDate: endDate, endDate: endDate };
  }

  const lastRecordedDateStr = sheet.getRange(lastRow, 1).getValue();
  const lastRecordedDate = new Date(lastRecordedDateStr);

  const startDate = new Date(lastRecordedDate.getTime());
  startDate.setDate(startDate.getDate() + 1);

  if (startDate > yesterday) {
    return { startDate: null, endDate: null };
  }

  return { startDate: formatDate(startDate), endDate: endDate };
}

/**
 * Dateオブジェクトを 'YYYY-MM-DD' 形式の文字列に変換する
 * @param {Date} date - 変換するDateオブジェクト
 * @returns {string} - フォーマットされた日付文字列
 */
function formatDate(date) {
  const y = date.getFullYear();
  const m = ('0' + (date.getMonth() + 1)).slice(-2);
  const d = ('0' + date.getDate()).slice(-2);
  return `${y}-${m}-${d}`;
}
