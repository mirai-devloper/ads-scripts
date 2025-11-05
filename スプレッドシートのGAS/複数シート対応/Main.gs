// Main.gs
/**
 * Webアプリのメインの入り口。
 * URLのパラメータを見て、表示するページを振り分ける。
 */
function doGet(e) {
  const page = e.parameter.page;

  switch (page) {
    case 'report1':
      return FunnelReport1.render(); // FunnelReport1.gs の render関数を呼び出す

    case 'report2':
      return FunnelReport2.render(); // FunnelReport2.gs の render関数を呼び出す

    default:
      return HtmlService.createHtmlOutput(
        'ページが指定されていません。URLの末尾に ?page=report1 または ?page=report2 を追加してください。'
      );
  }
}