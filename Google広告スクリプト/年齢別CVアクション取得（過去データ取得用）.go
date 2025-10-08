/**
 * 【年齢別・CVアクション別データ取得・年指定版】
 * 指定した1年分の年齢別・コンバージョンアクション別のデータを追記し、シート全体を日付順に並べ替える。
 * ★広告チャネルタイプを追加（大文字）
 * ★年齢の値を日本語に変換（AGE_RANGE_UNDETERMINED と UNDETERMINED に対応）
 * ★データの取得は最大で「前々日」までとします。
 */
 function main() {

  // ▼▼【要設定】▼▼ 取得したい年を西暦で指定してください
  const TARGET_YEAR = 2025; // 例: 2024年

  // ▼▼【要設定】▼▼ 記録したいスプレッドシートのURLを貼り付けてください
  const SPREADSHEET_URL = 'スプレッドシートのURLをここに貼り付けてください';

  // ▼設定▼ 記録先のシート名を指定してください
  const SHEET_NAME = '年齢別CVアクションデータ'; // シート名を変更

  // --- スプレッドシートの準備 ---
  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  try {
    // ヘッダーの「性別」を「年齢」に変更
    const japaneseHeaders = [
      '日付', 'キャンペーン名', '広告チャネルタイプ', '広告グループ名', '年齢',
      'コンバージョンアクション名', 'コンバージョン数'
    ];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(japaneseHeaders);
      console.log('ヘッダー行を新規設定しました。');
    }

    // --- 取得期間を決定するロジック ---
    const accountTimezone = AdsApp.currentAccount().getTimeZone();
    const startDate = new Date(TARGET_YEAR, 0, 1); // 指定年の1月1日

    // スクリプト実行日の前々日を計算
    const today = new Date();
    const dayBeforeYesterday = new Date(today);
    dayBeforeYesterday.setDate(today.getDate() - 2);

    // 取得終了日を、指定年の12月31日と前々日のうち、どちらか早い方に設定
    let endDate = new Date(TARGET_YEAR, 11, 31); // 指定年の12月31日
    if (endDate > dayBeforeYesterday) {
      endDate = dayBeforeYesterday;
    }

    // もしstartDateがendDateより後の日付になってしまう場合は、処理をスキップ
    if (startDate > endDate) {
      console.log(`指定年(${TARGET_YEAR})のデータは、まだ前々日までの範囲に達していません。`);
      return;
    }

    const startDateString = Utilities.formatDate(startDate, accountTimezone, "yyyyMMdd");
    const endDateString = Utilities.formatDate(endDate, accountTimezone, "yyyyMMdd");

    console.log(`取得期間: ${startDateString} から ${endDateString}`);

    // クエリを性別用から年齢用に変更
    const query = `
      SELECT
        segments.date,
        campaign.name,
        campaign.advertising_channel_type,
        ad_group.name,
        ad_group_criterion.age_range.type,
        segments.conversion_action_name,
        metrics.conversions
      FROM age_range_view
      WHERE
        segments.date >= '${startDateString}'
        AND segments.date <= '${endDateString}'
        AND metrics.conversions > 0
    `;

    const report = AdsApp.report(query);
    const rows = report.rows();
    const dataToWrite = [];

    while (rows.hasNext()) {
      const row = rows.next();

      let ageRange = row['ad_group_criterion.age_range.type'];
      // ★年齢の値を日本語に変換する処理を修正 (AGE_RANGE_UNDETERMINED と UNDETERMINED に対応)
      switch (ageRange) {
        case 'AGE_RANGE_18_24':
          ageRange = '18歳～24歳';
          break;
        case 'AGE_RANGE_25_34':
          ageRange = '25歳～34歳';
          break;
        case 'AGE_RANGE_35_44':
          ageRange = '35歳～44歳';
          break;
        case 'AGE_RANGE_45_54':
          ageRange = '45歳～54歳';
          break;
        case 'AGE_RANGE_55_64':
          ageRange = '55歳～64歳';
          break;
        case 'AGE_RANGE_65_UP':
          ageRange = '65歳～';
          break;
        case 'UNDETERMINED':
        case 'AGE_RANGE_UNDETERMINED': // ★AGE_RANGE_UNDETERMINED も「不明」に変換
          ageRange = '不明';
          break;
        // その他の値があればここに追加
      }

      // 書き込むデータを年齢用に変更
      dataToWrite.push([
        row['segments.date'],
        row['campaign.name'],
        row['campaign.advertising_channel_type'],
        row['ad_group.name'],
        ageRange, // 変換後の年齢データを追加
        row['segments.conversion_action_name'],
        row['metrics.conversions']
      ]);
    }

    if (dataToWrite.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);
      console.log(`${dataToWrite.length}件のデータを追記しました。`);

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
  }
}