// FunnelReport1.gs
const FunnelReport1 = {
  // --- ① レポート1専用の設定 ---
  config: {
    SHEET_NAME: '広告データ', // 1つ目のシート名
    HEADER_ROWS: 1,
    COL: {
      DATE: 1, CAMPAIGN: 2, AD_GROUP: 3, COST: 4,
      IMPRESSIONS: 5, CLICKS: 6, CONVERSIONS: 8
    }
  },

  // --- ② doGetをrenderに改名し、この中に入れる ---
  render: function() {
    const allData = this.getAllAdGroupData_(); // 内部の関数を呼び出す際は「this.」を付ける
    const availableMonths = this.getAvailableMonths_(allData);
    const availableCampaigns = this.getAvailableCampaigns_(allData);

    const template = HtmlService.createTemplateFromFile('FunnelPage1.html'); // 対応するHTMLファイルを指定

    template.allDataJson = JSON.stringify(allData);
    template.availableMonthsJson = JSON.stringify(availableMonths);
    template.availableCampaignsJson = JSON.stringify(availableCampaigns);

    if (availableMonths.length > 0) {
      template.defaultPeriod = availableMonths[0].period;
    } else {
      template.defaultPeriod = null;
    }

    return template.evaluate().setTitle('ファネルレポート1');
  },

  // --- ③ これまでのヘルパー関数を全てこの中に入れる ---
  getAllAdGroupData_: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(this.config.SHEET_NAME); // 設定は「this.config」で参照
    if (!sheet || sheet.getLastRow() <= this.config.HEADER_ROWS) { return []; }

    const data = sheet.getRange(
      this.config.HEADER_ROWS + 1, 1,
      sheet.getLastRow() - this.config.HEADER_ROWS, sheet.getLastColumn()
    ).getValues();

    const self = this; // mapの中でthis.parseNumber_を呼ぶため
    return data.map(function(row) {
      const date = row[self.config.COL.DATE - 1];
      if (!(date instanceof Date) || !row[self.config.COL.AD_GROUP - 1]) { return null; }
      return {
        period: date.getFullYear() + '-' + date.getMonth(),
        campaignName: row[self.config.COL.CAMPAIGN - 1],
        groupName: row[self.config.COL.AD_GROUP - 1],
        cost: self.parseNumber_(row[self.config.COL.COST - 1]),
        impressions: self.parseNumber_(row[self.config.COL.IMPRESSIONS - 1]),
        clicks: self.parseNumber_(row[self.config.COL.CLICKS - 1]),
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