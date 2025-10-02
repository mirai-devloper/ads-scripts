// --------------------------------------------------------------------------------
// 設定項目
// --------------------------------------------------------------------------------

// ▼▼▼【要設定】出力先のGoogleスプレッドシートのURLを指定してください ▼▼▼
const SPREADSHEET_URL = 'スプレッドシートのURLをここに貼り付けてください';

// ▼▼▼【任意設定】出力先のシート名を指定してください ▼▼▼
const SHEET_NAME = '地域別データ';

// --------------------------------------------------------------------------------
// メイン処理
// --------------------------------------------------------------------------------
function main() {
  // --- スプレッドシートの準備 ---
  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  // --- ここから修正箇所：日次更新用の取得期間を決定 ---
  const accountTimezone = AdsApp.currentAccount().getTimeZone();
  let startDateString, endDateString;

  const today = new Date();
  const yesterday = new Date();
  yesterday.setDate(today.getDate() - 1);
  endDateString = Utilities.formatDate(yesterday, accountTimezone, "yyyy-MM-dd");

  if (sheet.getLastRow() <= 1) { // ヘッダー行のみ、またはデータがない場合
    Logger.log('データがないため、昨日1日分のデータを取得します。');
    startDateString = endDateString;
  } else { // データがある場合は、最終日の翌日から取得
    const lastDateValue = sheet.getRange(sheet.getLastRow(), 1).getValue();
    const lastDate = new Date(lastDateValue);
    const startDate = new Date(lastDate);
    startDate.setDate(lastDate.getDate() + 1);

    // 既にデータが最新の場合は処理を終了
    if (startDate > yesterday) {
      Logger.log('データは既に最新です。処理を終了します。');
      return;
    }
    startDateString = Utilities.formatDate(startDate, accountTimezone, "yyyy-MM-dd");
  }

  Logger.log(`データ取得範囲: ${startDateString} から ${endDateString}`);
  // --- ここまで修正箇所 ---


  // --- Step 1: パフォーマンス指標と地域IDを日別に取得 ---
  Logger.log('Step 1: パフォーマンスデータを取得しています...');
  const performanceQuery = `
    SELECT
      segments.date,
      campaign_criterion.criterion_id,
      metrics.clicks,
      metrics.impressions,
      metrics.cost_micros,
      metrics.conversions
    FROM
      location_view
    WHERE
      segments.date BETWEEN '${startDateString}' AND '${endDateString}'
      AND campaign.status = 'ENABLED'
  `;
  const performanceReport = AdsApp.report(performanceQuery);
  const performanceRows = performanceReport.rows();

  const performanceData = {};
  const allCriterionIds = new Set();
  for (const row of performanceRows) {
    const criterionId = row['campaign_criterion.criterion_id'];
    const date = row['segments.date'];
    if (!criterionId || !date) continue;
    allCriterionIds.add(criterionId);
    const key = `${date}_${criterionId}`;
    if (!performanceData[key]) {
      performanceData[key] = { date: date, criterionId: criterionId, clicks: 0, impressions: 0, cost: 0, conversions: 0 };
    }
    performanceData[key].clicks += parseFloat(row['metrics.clicks']);
    performanceData[key].impressions += parseFloat(row['metrics.impressions']);
    performanceData[key].cost += parseFloat(row['metrics.cost_micros']) / 1000000;
    performanceData[key].conversions += parseFloat(row['metrics.conversions']);
  }

  if (allCriterionIds.size === 0) {
    Logger.log('期間内にパフォーマンスデータが見つかりませんでした。');
    return;
  }

  // --- Step 2: 全ての地域IDの詳細情報を、種類を判別しながら取得 ---
  Logger.log(`Step 2: ${allCriterionIds.size} 件の地域IDから名前を特定しています...`);
  const locationInfoMap = new Map();

  const criteriaQuery = `
    SELECT
      campaign_criterion.criterion_id,
      campaign_criterion.type,
      campaign_criterion.location.geo_target_constant,
      campaign_criterion.proximity.radius,
      campaign_criterion.proximity.radius_units,
      campaign_criterion.proximity.address.street_address,
      campaign_criterion.proximity.address.city_name
    FROM
      campaign_criterion
    WHERE
      campaign_criterion.criterion_id IN (${Array.from(allCriterionIds).join(',')})
  `;
  const criteriaReport = AdsApp.report(criteriaQuery);
  const criteriaRows = criteriaReport.rows();

  const geoTargetIdsToLookup = new Set();
  const tempCriterionInfo = new Map();

  for (const row of criteriaRows) {
    const id = row['campaign_criterion.criterion_id'];
    const type = row['campaign_criterion.type'];

    if (type === 'PROXIMITY') {
      const addressParts = [
        row['campaign_criterion.proximity.address.city_name'],
        row['campaign_criterion.proximity.address.street_address']
      ].filter(Boolean).join(' '); // 住所を結合
      const radius = row['campaign_criterion.proximity.radius'];
      const units = row['campaign_criterion.proximity.radius_units'];
      locationInfoMap.set(id, `[半径] ${addressParts} (${radius} ${units})`);
    } else if (type === 'LOCATION') {
      const geoTarget = row['campaign_criterion.location.geo_target_constant'];
      if (geoTarget && geoTarget.startsWith('geoTargetConstants/')) {
        geoTargetIdsToLookup.add(`'${geoTarget}'`);
        tempCriterionInfo.set(id, { geoTarget: geoTarget });
      }
    }
  }

  if (geoTargetIdsToLookup.size > 0) {
    const geoQuery = `
      SELECT geo_target_constant.name, geo_target_constant.resource_name
      FROM geo_target_constant
      WHERE geo_target_constant.resource_name IN (${Array.from(geoTargetIdsToLookup).join(',')})
    `;
    const geoReport = AdsApp.report(geoQuery);
    const geoNameMap = new Map();
    for (const row of geoReport.rows()) {
      geoNameMap.set(row['geo_target_constant.resource_name'], row['geo_target_constant.name']);
    }

    for (const [id, info] of tempCriterionInfo.entries()) {
      if (geoNameMap.has(info.geoTarget)) {
        locationInfoMap.set(id, geoNameMap.get(info.geoTarget));
      }
    }
  }

  // --- Step 3: データを結合して出力 ---
  Logger.log('Step 3: データを結合して出力します...');

  const headers = ['日付', 'ターゲット地域', 'クリック数', '表示回数', '費用', 'コンバージョン数'];

  // ヘッダー行がまだない場合（シートが空の場合）のみ、ヘッダーを追加します
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.getRange("A:A").setNumberFormat('yyyy-mm-dd');
    sheet.getRange("B:B").setNumberFormat('@');
  }

  const outputData = [];
  for (const key in performanceData) {
    const data = performanceData[key];
    const name = locationInfoMap.get(data.criterionId) || data.criterionId;

    outputData.push([
        data.date, name,
        data.clicks, data.impressions,
        Math.round(data.cost), data.conversions
    ]);
  }

  if (outputData.length > 0) {
    // データをシートの末尾に追記します
    sheet.getRange(sheet.getLastRow() + 1, 1, outputData.length, headers.length).setValues(outputData);
    Logger.log(`${outputData.length} 行のデータをスプレッドシートに追記しました。`);

    // 追記後、ヘッダー行を除いたシート全体を日付でソートします
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    dataRange.sort({column: 1, ascending: true}); // 1列目（日付）を昇順でソート
    Logger.log('シート全体を日付順にソートしました。');

  } else {
    Logger.log('期間内に記録対象のデータはありませんでした。');
  }
}