/**
 * 【パフォーマンスデータ】
 * 指定した1年分の広告グループレポートを取得し、スプレッドシートに追記します。
 * 日付順に並べ替え、各項目を日本語に変換して出力します。
 * ★データの取得は最大で「前々日」までとします。
 */
 function main() {

  // ▼▼【要設定】▼▼ 取得したい年を西暦で指定してください
  const TARGET_YEAR = 2025; // 年数を指定してください

  // ▼▼【要設定】▼▼ 記録したいスプレッドシートのURLを貼り付けてください
  const SPREADSHEET_URL = 'スプレッドシートのURLをここに貼り付けてください';

  // ▼設定▼ 記録先のシート名を指定してください
  const SHEET_NAME = 'パフォーマンスデータ';

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
    // --- 翻訳用マッピング（APIのEnum値とユーザー指定リストを完全に対応） ---
    const campaignTypeMap = {
      'SEARCH': '検索',
      'DISPLAY': 'ディスプレイ',
      'SHOPPING': 'ショッピング',
      'VIDEO': '動画',
      'MULTI_CHANNEL': 'アプリ',
      'SMART': 'スマート',
      'HOTEL': 'ホテル',
      'LOCAL': 'ローカル',
      'DEMAND_GEN': 'デマンド ジェネレーション',
      'PERFORMANCE_MAX': 'P-MAX',
      'UNKNOWN': '（不明）',
      'UNSPECIFIED': '（未指定）'
    };

    const adGroupTypeMap = {
      'SEARCH_STANDARD': '標準',
      'SEARCH_DYNAMIC_ADS': '動的広告',
      'DISPLAY_STANDARD': 'ディスプレイ',
      'DISPLAY_ENGAGEMENT_AD': 'ディスプレイ エンゲージメント',
      'SHOPPING_PRODUCT_ADS': 'ショッピング - 商品',
      'SHOPPING_SHOWCASE_ADS': 'ショッピング - ショーケース',
      'SHOPPING_SMART_ADS': 'ショッピング - スマート',
      'SHOPPING_COMPARISON_LISTING_ADS': 'ショッピング - コレクション',
      'VIDEO_TRUE_VIEW_IN_STREAM': 'インストリーム',
      'VIDEO_RESPONSIVE': 'インストリーム',
      'VIDEO_ACTION': 'インストリーム',
      'VIDEO_NON_SKIPPABLE_IN_STREAM': 'インストリーム',
      'VIDEO_BUMPER': 'インストリーム',
      'VIDEO_OUTSTREAM': 'インストリーム',
      'VIDEO_DISCOVERY': 'インフィード動画',
      'VIDEO_TRUE_VIEW_IN_DISPLAY': 'インフィード動画',
      'HOTEL_ADS': 'ホテル広告',
      'UNKNOWN': '（不明）',
      'UNSPECIFIED': '（未指定）'
    };

    const deviceMap = {
      'DESKTOP': 'コンピュータ',
      'TABLET': 'タブレット',
      'MOBILE': 'スマートフォン',
      'CONNECTED_TV': 'テレビ画面',
      'OTHER': 'その他',
      'UNKNOWN': '（不明）',
      'UNSPECIFIED': '（未指定）'
    };

    const dayOfWeekMap = ['日', '月', '火', '水', '木', '金', '土'];

    // --- ヘッダーの設定 ---
    const japaneseHeaders = [
      '日付', '曜日', 'キャンペーン名', 'キャンペーンタイプ', '広告グループ名', '広告グループの種類', 'デバイス', '費用', '表示回数', 'クリック数', 'コンバージョン数'
    ];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(japaneseHeaders);
      console.log('ヘッダー行を新規設定しました。');
    }

    // --- 取得期間を決定するロジック ---
    const accountTimezone = AdsApp.currentAccount().getTimeZone();
    const startDate = new Date(TARGET_YEAR, 0, 1);
    const today = new Date();
    const dayBeforeYesterday = new Date(today);
    dayBeforeYesterday.setDate(today.getDate() - 2);
    let endDate = new Date(TARGET_YEAR, 11, 31);
    if (endDate > dayBeforeYesterday) {
      endDate = dayBeforeYesterday;
    }

    if (startDate > endDate) {
      console.log(`指定年(${TARGET_YEAR})のデータは、まだ前々日までの範囲に達していません。`);
      return;
    }

    const startDateString = Utilities.formatDate(startDate, accountTimezone, "yyyy-MM-dd");
    const endDateString = Utilities.formatDate(endDate, accountTimezone, "yyyy-MM-dd");

    console.log('取得期間: ' + startDateString + ' - ' + endDateString);

    const dataToWrite = [];

    // --- レポートを取得（ご指摘に基づき、impressions > 0 のフィルタを削除） ---
    const query = `
      SELECT
        segments.date,
        campaign.name,
        campaign.advertising_channel_type,
        ad_group.name,
        ad_group.type,
        segments.device,
        metrics.cost_micros,
        metrics.impressions,
        metrics.clicks,
        metrics.conversions
      FROM ad_group
      WHERE segments.date BETWEEN '${startDateString}' AND '${endDateString}'`;

    const report = AdsApp.search(query);

    for (const row of report) {
      const dateStr = row.segments.date;
      if (!dateStr) continue;

      const date = new Date(dateStr);
      const dayOfWeek = dayOfWeekMap[date.getDay()];

      // 【根本修正】オブジェクトが存在するかを必ず確認し、安全にデータを取得
      const campaignName = row.campaign ? row.campaign.name : '（キャンペーン情報なし）';
      const campaignTypeRaw = row.campaign ? row.campaign.advertising_channel_type : 'UNKNOWN';
      const campaignType = campaignTypeMap[campaignTypeRaw] || campaignTypeRaw;

      const adGroupName = row.ad_group ? row.ad_group.name : '（該当なし）'; // P-MAXなど広告グループがない場合
      const adGroupTypeRaw = row.ad_group ? row.ad_group.type : 'UNKNOWN';
      const adGroupType = adGroupTypeMap[adGroupTypeRaw] || adGroupTypeRaw;

      const deviceRaw = row.segments.device;
      const device = deviceMap[deviceRaw] || deviceRaw;

      // 【APIの仕様】費用は「マイクロ円」で返されるため、100万で割ることで「円」に変換
      const cost = row.metrics.cost_micros / 1000000;

      const impressions = row.metrics.impressions;
      const clicks = row.metrics.clicks;
      const conversions = row.metrics.conversions;

      dataToWrite.push([
        Utilities.formatDate(date, accountTimezone, 'yyyy-MM-dd'),
        dayOfWeek,
        campaignName,
        campaignType,
        adGroupName,
        adGroupType,
        device,
        cost,
        impressions,
        clicks,
        conversions
      ]);
    }

    // --- データの書き込みと並べ替え ---
    if (dataToWrite.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);
      console.log(dataToWrite.length + '件のデータを追記しました。');

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
    console.error('Stack: ' + e.stack);
  }
}

