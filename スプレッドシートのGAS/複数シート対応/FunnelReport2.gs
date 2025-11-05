// FunnelReport2.gs
const FunnelReport2 = { // 1. オブジェクト名を変更
  // --- ① レポート2専用の設定 ---
  config: {
    SHEET_NAME: 'もう一つの広告データ', // 2. 2つ目のシート名に変更
    HEADER_ROWS: 1,
    // ※もし2つ目のシートの列構成が違う場合は、ここも修正してください
    COL: {
      DATE: 1, CAMPAIGN: 2, AD_GROUP: 3, COST: 4,
      IMPRESSIONS: 5, CLICKS: 6, CONVERSIONS: 8
    }
  },

  // --- ② doGetをrenderに改名し、この中に入れる ---
  render: function() {
    const allData = this.getAllAdGroupData_();
    const availableMonths = this.getAvailableMonths_(allData);
    const availableCampaigns = this.getAvailableCampaigns_(allData);

    const template = HtmlService.createTemplateFromFile('FunnelPage2.html'); // 3. 対応するHTMLファイルを変更

    template.allDataJson = JSON.stringify(allData);
    template.availableMonthsJson = JSON.stringify(availableMonths);
    template.availableCampaignsJson = JSON.stringify(availableCampaigns);

    if (availableMonths.length > 0) {
      template.defaultPeriod = availableMonths[0].period;
    } else {
      template.defaultPeriod = null;
    }

    return template.evaluate().setTitle('ファネルレポート2');
  },

  // --- ③ これまでのヘルパー関数を全てこの中に入れる ---
  // (getAllAdGroupData_ や getAvailableMonths_ など、FunnelReport1からコピーした残りの関数はそのまま)
  getAllAdGroupData_: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(this.config.SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= this.config.HEADER_ROWS) { return []; }
    const data = sheet.getRange(this.config.HEADER_ROWS + 1, 1, sheet.getLastRow() - this.config.HEADER_ROWS, sheet.getLastColumn()).getValues();
    const self = this;
    return data.map(function(row) {
      const date = row[self.config.COL.DATE - 1];
      if (!(date instanceof Date) || !row[self.config.COL.AD_GROUP - 1]) { return null; }
      return {
        period: date.getFullYear() + '-' + date.getMonth(), campaignName: row[self.config.COL.CAMPAIGN - 1],
        groupName: row[self.config.COL.AD_GROUP - 1], cost: self.parseNumber_(row[self.config.COL.COST - 1]),
        impressions: self.parseNumber_(row[self.config.COL.IMPRESSIONS - 1]), clicks: self.parseNumber_(row[self.config.COL.CLICKS - 1]),
        conversions: self.parseNumber_(row[self.config.COL.CONVERSIONS - 1])
      };
    }).filter(function(item) { return item !== null; });
  },
  getAvailableMonths_: function(allData) {
    const uniquePeriods = {};
    allData.forEach(function(row) { uniquePeriods[row.period] = true; });
    return Object.keys(uniquePeriods).map(function(period) {
      const parts = period.split('-');
      const year = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10);
      return { period: period, display: `${year}年${month + 1}月` };
    }).sort(function(a, b) {
      const partsA = a.period.split('-'), partsB = b.period.split('-');
      return (parseInt(partsB[0], 10) - parseInt(partsA[0], 10)) || (parseInt(partsB[1], 10) - parseInt(partsA[1], 10));
    });
  },
  getAvailableCampaigns_: function(allData) {
    const uniqueCampaigns = new Set();
    allData.forEach(function(row) { if (row.campaignName) { uniqueCampaigns.add(row.campaignName); } });
    const campaignList = Array.from(uniqueCampaigns).sort();
    campaignList.unshift("すべてのキャンペーン");
    return campaignList;
  },
  parseNumber_: function(value) {
    if (typeof value === 'number') { return value; }
    if (typeof value === 'string') {
      const num = parseFloat(value.replace(/[^0-9.-]+/g, ""));
      return isNaN(num) ? 0 : num;
    }
    return 0;
  }
};