// ▼▼▼ このスクリプト専用の設定 ▼▼▼
const FunnelReportConfig = {
  SHEET_NAME: 'グループ集計',
  HEADER_ROWS: 1,
  // B列にキャンペーンが追加されたため、列番号を更新
  COL: {
    DATE: 1,          // A列: 日付
    CAMPAIGN: 2,      // B列: キャンペーン名
    AD_GROUP: 3,      // C列: 広告グループ名
    COST: 4,          // D列: 費用
    IMPRESSIONS: 5,   // E列: 表示回数
    CLICKS: 6,        // F列: クリック数
    CONVERSIONS: 8    // H列: コンバージョン数
  }
};
// ▲▲▲ 設定はここまで ▲▲▲

function doGet(e) {
  const allData = getAllAdGroupData_();
  const availableCampaigns = getAvailableCampaigns_(allData);
  const template = HtmlService.createTemplateFromFile('index'); // HTMLファイル名を指定

  template.allDataJson = JSON.stringify(allData);
  template.availableCampaignsJson = JSON.stringify(availableCampaigns);

  return template.evaluate()
    .setTitle('広告パフォーマンス インフォグラフィック')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function getAllAdGroupData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(FunnelReportConfig.SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= FunnelReportConfig.HEADER_ROWS) { return []; }

  const data = sheet.getRange(
    FunnelReportConfig.HEADER_ROWS + 1, 1,
    sheet.getLastRow() - FunnelReportConfig.HEADER_ROWS, sheet.getLastColumn()
  ).getValues();

  // スプレッドシートのタイムゾーンを取得
  const timezone = ss.getSpreadsheetTimeZone();

  return data.map(function(row) {
    const date = row[FunnelReportConfig.COL.DATE - 1];
    if (!(date instanceof Date) || !row[FunnelReportConfig.COL.AD_GROUP - 1]) { return null; }
    return {
      // ★変更点: 日付フィルタリング用に yyyy-MM-dd 形式の日付データを追加
      date: Utilities.formatDate(date, timezone, "yyyy-MM-dd"),
      campaignName: row[FunnelReportConfig.COL.CAMPAIGN - 1],
      groupName: row[FunnelReportConfig.COL.AD_GROUP - 1],
      cost: parseNumber_(row[FunnelReportConfig.COL.COST - 1]),
      impressions: parseNumber_(row[FunnelReportConfig.COL.IMPRESSIONS - 1]),
      clicks: parseNumber_(row[FunnelReportConfig.COL.CLICKS - 1]),
      conversions: parseNumber_(row[FunnelReportConfig.COL.CONVERSIONS - 1])
    };
  }).filter(function(item) { return item !== null; });
}

function getAvailableCampaigns_(allData) {
    const uniqueCampaigns = new Set();
    allData.forEach(function(row) {
        if (row.campaignName) {
            uniqueCampaigns.add(row.campaignName);
        }
    });
    const campaignList = Array.from(uniqueCampaigns).sort();
    campaignList.unshift("すべてのキャンペーン");
    return campaignList;
}

function parseNumber_(value) {
  if (typeof value === 'number') { return value; }
  if (typeof value === 'string') {
    const num = parseFloat(value.replace(/[^0-9.-]+/g, ""));
    return isNaN(num) ? 0 : num;
  }
  return 0;
}