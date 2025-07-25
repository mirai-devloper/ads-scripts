// --- 設定項目 ---

// 1. 休日リストが記載されているスプレッドシートのURL
const SPREADSHEET_URL = '（スプレッドシートのURLを入力）';

// 2. シート名
const SHEET_NAME = '合算';

// 3. 休日の日付が入力されている列番号 (A列なら1, B列なら2)
const DATE_COLUMN = 1;

// 4. 操作したいキャンペーン名のリスト（完全一致）
//    例: ['キャンペーンA', 'キャンペーンB']
//    空のまま [] にすると、アカウントの全有効キャンペーンが対象になります。
const TARGET_CAMPAIGN_NAMES = [];

// --- 設定はここまで ---


/**
 * メイン関数
 */
function main() {
  const holidays = getHolidaysFromSheet();
  if (holidays.size === 0) {
    Logger.log('休日リストが空か、取得できませんでした。処理を終了します。');
    return;
  }
  Logger.log(`${holidays.size}件の休日を読み込みました。`);

  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowString = Utilities.formatDate(tomorrow, AdsApp.currentAccount().getTimeZone(), 'yyyy/MM/dd');
  Logger.log(`判定対象日（明日）: ${tomorrowString}`);

  const shouldBePaused = holidays.has(tomorrowString);
  if (shouldBePaused) {
    Logger.log('明日は休日のため、広告をオフにします。');
  } else {
    Logger.log('明日は平日（休みではない）ため、広告をオンにします。');
  }

  // ▼▼▼ ここからロジックを全面的に変更 ▼▼▼

  // まず全ての有効なキャンペーンを取得
  const campaigns = AdsApp.campaigns().withCondition("campaign.status != 'REMOVED'").get();

  if (!campaigns.hasNext()) {
    Logger.log('対象アカウントに有効なキャンペーンが存在しません。');
    return;
  }

  // 1つずつキャンペーンをチェックし、条件に合致すれば操作する
  while (campaigns.hasNext()) {
    const campaign = campaigns.next();
    const campaignName = campaign.getName();

    // TARGET_CAMPAIGN_NAMESが空、またはリストに現在のキャンペーン名が含まれているか
    if (TARGET_CAMPAIGN_NAMES.length === 0 || TARGET_CAMPAIGN_NAMES.includes(campaignName)) {
      if (shouldBePaused) {
        campaign.pause();
        Logger.log(`キャンペーン「${campaignName}」を一時停止しました。`);
      } else {
        campaign.enable();
        Logger.log(`キャンペーン「${campaignName}」を有効にしました。`);
      }
    }
  }
  // ▲▲▲ ここまで変更 ▲▲▲

  Logger.log('処理が完了しました。');
}

/**
 * スプレッドシートから休日リストを取得して、Setとして返す
 * @return {Set<string>} 'yyyy/MM/dd' 形式の日付文字列のSet
 */
function getHolidaysFromSheet() {
  try {
    const sheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL).getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return new Set();
    }
    const range = sheet.getRange(2, DATE_COLUMN, lastRow - 1, 1);
    const timezone = AdsApp.currentAccount().getTimeZone();

    const holidays = range.getValues()
      .flat()
      .filter(cell => cell instanceof Date)
      .map(date => Utilities.formatDate(date, timezone, 'yyyy/MM/dd'));

    return new Set(holidays);
  } catch (e) {
    Logger.log(`エラー: スプレッドシートの読み込みに失敗しました。URLやシート名が正しいか確認してください。 - ${e.toString()}`);
    return new Set();
  }
}