// --------------------------------------------------------------------------------
// 設定項目
// --------------------------------------------------------------------------------

// ▼▼▼【要設定】データ取得対象の年（西暦）を指定してください ▼▼▼
const YEAR = 2025;

// ▼▼▼【要設定】出力先のGoogleスプレッドシートのURLを指定してください ▼▼▼
const SPREADSHEET_URL = 'スプレッドシートのURLをここに貼り付けてください';

// ▼▼▼【任意設定】出力先のシート名を指定してください ▼▼▼
const SHEET_NAME = '地域別CVアクションデータ'; // シート名を変更

// --------------------------------------------------------------------------------
// メイン処理
// --------------------------------------------------------------------------------
function main() {
  // --- 取得期間を決定するロジック ---
  const accountTimezone = AdsApp.currentAccount().getTimeZone();
  const startDate = new Date(YEAR, 0, 1); // 指定年の1月1日

  // スクリプト実行日の前々日を計算
  const today = new Date();
  const dayBeforeYesterday = new Date(today);
  dayBeforeYesterday.setDate(today.getDate() - 2);

  // 取得終了日を、指定年の12月31日と前々日のうち、どちらか早い方に設定
  let endDate = new Date(YEAR, 11, 31); // 指定年の12月31日
  if (endDate > dayBeforeYesterday) {
    endDate = dayBeforeYesterday;
  }

  // もしstartDateがendDateより後の日付になってしまう場合は、処理をスキップ
  if (startDate > endDate) {
    Logger.log(`指定年(${YEAR})のデータは、まだ前々日までの範囲に達していません。`);
    return;
  }

  const START_DATE_STR = Utilities.formatDate(startDate, accountTimezone, "yyyy-MM-dd");
  const END_DATE_STR = Utilities.formatDate(endDate, accountTimezone, "yyyy-MM-dd");
  Logger.log(`データ取得範囲: ${START_DATE_STR} から ${END_DATE_STR}`);


  // --- スプレッドシートの準備 ---
  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  // ヘッダーを定義
  const headers = ['日付', 'ターゲット地域', '広告チャネルタイプ', 'コンバージョンアクション名', 'コンバージョン数'];

  // ヘッダー行がまだない場合（シートが空の場合）のみ、ヘッダーを追加します
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.getRange("A:A").setNumberFormat('yyyy-mm-dd');
    sheet.getRange("B:B").setNumberFormat('@'); // ターゲット地域もテキスト形式に
    Logger.log('ヘッダー行を新規設定しました。');
  }

  try {
    // --- Step 1: 地域別のコンバージョンデータを取得 ---
    Logger.log('Step 1: コンバージョンデータを取得しています...');
    const convQuery = `
      SELECT
        segments.date,
        campaign_criterion.criterion_id,
        campaign.advertising_channel_type,  -- ★追加: 広告チャネルタイプ
        segments.conversion_action_name,
        metrics.conversions
      FROM
        location_view
      WHERE
        segments.date BETWEEN '${START_DATE_STR}' AND '${END_DATE_STR}'
        AND campaign.status = 'ENABLED'
        AND metrics.conversions > 0
    `;
    const convReport = AdsApp.report(convQuery);
    const convRows = Array.from(convReport.rows()); // 後で使うため、一度配列に格納

    if (convRows.length === 0) {
      Logger.log('期間内にコンバージョンデータが見つかりませんでした。');
      return;
    }

    // --- Step 2: 全ての地域IDの詳細情報を取得 ---
    // レポートから重複を除いた地域IDのリストを作成
    const allCriterionIds = new Set(convRows.map(row => row['campaign_criterion.criterion_id']));

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

    const outputData = [];
    for (const row of convRows) {
      const criterionId = row['campaign_criterion.criterion_id'];
      const name = locationInfoMap.get(criterionId) || criterionId;

      outputData.push([
        row['segments.date'],
        name,
        row['campaign.advertising_channel_type'], // ★追加: 広告チャネルタイプ
        row['segments.conversion_action_name'],
        row['metrics.conversions']
      ]);
    }

    if (outputData.length > 0) {
      // データをシートの末尾に追記します
      sheet.getRange(sheet.getLastRow() + 1, 1, outputData.length, headers.length).setValues(outputData);
      Logger.log(`${outputData.length} 行のデータをスプレッドシートに追記しました。`);

      // 追記後、ヘッダー行を除いたシート全体を日付でソートします
      if (sheet.getLastRow() > 1) { // データ行が1行以上ある場合のみソート
        const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
        dataRange.sort({column: 1, ascending: true}); // 1列目（日付）を昇順でソート
        Logger.log('シート全体を日付順にソートしました。');
      }

    } else {
      Logger.log('最終的な出力データが見つかりませんでした。');
    }

  } catch (e) {
    Logger.log('スクリプトの実行中にエラーが発生しました: ' + e.toString());
  }
}