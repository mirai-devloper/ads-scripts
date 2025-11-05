function main() {
  // ① 書き込みたいGoogleスプレッドシートの情報を設定
  const SPREADSHEET_URL = "スプレッドシートのURLをここに貼り付けてください";
  const SHEET_NAME = "費用"; // 対象のシート名

  // ---ここから自動処理---
  try {
    const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
    }

    let startDate;
    const today = new Date();

    // ② シートのデータ状況を確認
    if (sheet.getLastRow() < 2) {
      startDate = "20180101"; // アカウントの開始年に合わせて調整
      sheet.clear();
    } else {
      const lastDate = new Date(sheet.getRange(sheet.getLastRow(), 1).getValue());
      const nextMonthDate = new Date(lastDate.getFullYear(), lastDate.getMonth() + 1, 1);
      const y = nextMonthDate.getFullYear();
      const m = ("0" + (nextMonthDate.getMonth() + 1)).slice(-2);
      startDate = y + m + "01";
    }

    // ③ 取得期間の終了日を「先月の末日」に設定
    const firstDayOfThisMonth = new Date(today.getFullYear(), today.getMonth(), 1);
    const lastDayOfPreviousMonth = new Date(firstDayOfThisMonth.getTime() - 1);
    const y = lastDayOfPreviousMonth.getFullYear();
    const m = ("0" + (lastDayOfPreviousMonth.getMonth() + 1)).slice(-2);
    const d = ("0" + lastDayOfPreviousMonth.getDate()).slice(-2);
    const endDate = y + m + d;

    if (startDate > endDate) {
      Logger.log("更新する新しいデータはありませんでした。処理を終了します。");
      return;
    }

    // ④-1 アカウント全体の合計費用を取得
    const totalCostQuery = `SELECT Month, Cost FROM ACCOUNT_PERFORMANCE_REPORT DURING ${startDate},${endDate}`;
    const totalReport = AdsApp.report(totalCostQuery);
    const totalRows = totalReport.rows();
    const totalCosts = {};
    while (totalRows.hasNext()) {
      const row = totalRows.next();
      totalCosts[row["Month"]] = parseFloat(row["Cost"].replace(/,/g, ''));
    }

    // ④-2 キャンペーン別の費用を取得
    const campaignCostQuery = `SELECT CampaignName, Month, Cost FROM CAMPAIGN_PERFORMANCE_REPORT DURING ${startDate},${endDate}`;
    const campaignReport = AdsApp.report(campaignCostQuery);
    const campaignRows = campaignReport.rows();

    // ⑤ データを整形
    const monthlyCampaignData = {};
    const campaignSet = new Set();

    while (campaignRows.hasNext()) {
      const row = campaignRows.next();
      const campaignName = row["CampaignName"];
      const month = row["Month"];
      const cost = parseFloat(row["Cost"].replace(/,/g, ''));

      if (!monthlyCampaignData[month]) {
        monthlyCampaignData[month] = {};
      }
      monthlyCampaignData[month][campaignName] = cost;
      campaignSet.add(campaignName);
    }

    if (Object.keys(totalCosts).length === 0) {
      Logger.log("期間内に広告費用データが見つかりませんでした。");
      return;
    }

    // ⑥ スプレッドシートのヘッダーを更新
    // ★★★★★ 修正点 ★★★★★
    // ヘッダー名を「費用」に戻しました
    let headers = ["月", "費用"];
    if (sheet.getLastRow() > 0) {
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
       // 万が一 "合計費用" になっていた場合も "費用" に戻すための処理
      if (headers[1] === "合計費用") {
          headers[1] = "費用";
      }
    }

    let newCampaignsAdded = false;
    campaignSet.forEach(campaign => {
      if (headers.indexOf(campaign) === -1) {
        headers.push(campaign);
        newCampaignsAdded = true;
      }
    });

    if (sheet.getLastRow() < 1 || newCampaignsAdded) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    // ⑦ 整形したデータを追記
    const months = Object.keys(totalCosts).sort();
    const dataToAppend = [];

    for (const month of months) {
      const newRow = [new Date(month), totalCosts[month]]; // A列: 月, B列: 費用
      for (let i = 2; i < headers.length; i++) { // C列以降のキャンペーン
        const campaignName = headers[i];
        const cost = (monthlyCampaignData[month] && monthlyCampaignData[month][campaignName]) ? monthlyCampaignData[month][campaignName] : 0;
        newRow.push(cost);
      }
      dataToAppend.push(newRow);
    }

    if (dataToAppend.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, dataToAppend.length, headers.length).setValues(dataToAppend);
    }

    // ⑧ ソートとフォーマット設定
    if (sheet.getLastRow() > 1) {
      const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
      range.sort({column: 1, ascending: true});

      sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).setNumberFormat("yyyy/mm/dd"); // A列
      sheet.getRange(2, 2, sheet.getLastRow() - 1, sheet.getLastColumn() - 1).setNumberFormat("#,##0"); // B列以降
    }

    Logger.log(`${dataToAppend.length}ヶ月分のデータをシートに追記し、更新しました。`);

  } catch (e) {
    Logger.log(`エラーが発生しました: ${e.message} (Line: ${e.lineNumber})`);
  }
}