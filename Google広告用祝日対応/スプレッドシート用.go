/**
 * 日本の祝日データを取得し、シートに書き込む最終版関数
 * 今年のデータがない場合は、クリアしてから今年のデータを最優先で書き込みます。
 */
function populateHolidaySheet_Final() {
  // --- 設定項目 ---
  const SHEET_NAME = '祝日データ'; // データを書き込むシート名
  const CALENDAR_ID = 'ja.japanese#holiday@group.v.calendar.google.com'; // 日本の祝日カレンダーID
  // --- 設定はここまで ---

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = ['日付', '祝日名'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }

  try {
    let targetYear;
    const today = new Date();
    const currentYear = today.getFullYear();
    const currentMonth = today.getMonth(); // 1月は0, 12月は11

    // --- 取得対象年を決定する最終ロジック ---

    // シート内の全日付データを取得
    const lastRow = sheet.getLastRow();
    let existingYears = [];
    if (lastRow > 1) {
      existingYears = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
                           .flat()
                           .map(d => new Date(d).getFullYear());
    }

    // Setに変換してユニークな年のリストを取得
    const uniqueExistingYears = new Set(existingYears);

    if (!uniqueExistingYears.has(currentYear)) {
      // ケース1: シートに「今年」のデータが存在しない場合 (最優先)
      Logger.log(`シートに${currentYear}年のデータが存在しないため、取得します。`);
      targetYear = currentYear;

      // ヘッダー行(1行目)以外をすべてクリア
      if (sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
        Logger.log('既存のデータをクリアしました。');
      }

    } else {
      // ケース2: シートに「今年」のデータが既に存在する場合
      if (currentMonth === 11) { // 12月なら来年データを取得
        const nextYear = currentYear + 1;
        if (uniqueExistingYears.has(nextYear)) {
          Logger.log(`${nextYear}年のデータは既に存在するため、処理をスキップします。`);
          return;
        }
        targetYear = nextYear;
        Logger.log(`12月のため、来年(${targetYear}年)のデータを追記します。`);
      } else {
        Logger.log(`シートは最新(${currentYear}年)です。処理をスキップします。`);
        return; // 処理を終了
      }
    }

    // --- データ取得と書き込み処理 ---
    const startTime = new Date(targetYear, 0, 1);
    const endTime = new Date(targetYear, 11, 31);

    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    const events = calendar.getEvents(startTime, endTime);

    if (events.length === 0) {
      Logger.log(`${targetYear}年の祝日データが見つかりませんでした。`);
      return;
    }

    const holidayData = events.map(event => [event.getAllDayStartDate(), event.getTitle()]);

    // データを追記
    const appendRow = sheet.getLastRow() + 1;
    const dataRange = sheet.getRange(appendRow, 1, holidayData.length, holidayData[0].length);
    dataRange.setValues(holidayData);

    // 書式設定
    dataRange.setNumberFormat('yyyy/MM/dd');

    // 列幅調整
    sheet.autoResizeColumn(1);
    sheet.autoResizeColumn(2);

    Logger.log(`✅ ${targetYear}年の祝日${holidayData.length}件をシートに書き込みました。`);

  } catch (e) {
    Logger.log('⚠️ エラーが発生しました: ' + e.toString());
  }
}