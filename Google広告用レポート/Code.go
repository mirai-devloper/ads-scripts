/**
 * 月次広告レポートをWebアプリとして表示するスクリプト (デバッグログ付き)
 * * 実行時の処理状況をログに出力し、問題の原因を特定します。
 */


/**
 * Webアプリとしてアクセスされたときに実行されるメイン関数
 */
 function doGet(e) {
  try {
    console.log("doGet: 開始");
    const htmlOutput = HtmlService.createTemplateFromFile('index').evaluate();
    htmlOutput.setTitle("広告運用詳細レポート (検索広告)");
    htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    console.log("doGet: 正常終了");
    return htmlOutput;
  } catch (error) {
    console.error("doGet Error: " + error.toString());
    return ContentService.createTextOutput("レポートの表示中にエラーが発生しました。");
  }
}

/**
 * HTML側から呼び出され、レポートに必要なすべてのデータを返す関数
 */
function getReportData(refresh) {
  try {
    console.log("getReportData: 開始");
    const cache = CacheService.getScriptCache();
    const today = new Date();
    const cacheKey = `report_data_main_${Utilities.formatDate(today, 'JST', 'yyyy-MM')}`;

    if (refresh) {
      const summaryCacheKey = `summary_text_${Utilities.formatDate(today, 'JST', 'yyyy-MM')}`;
      cache.removeAll([cacheKey, summaryCacheKey]);
      console.log('キャッシュをクリアしました。');
    }

    const cachedData = cache.get(cacheKey);
    if (cachedData) {
      console.log('キャッシュからレポートデータを返します。');
      return JSON.parse(cachedData);
    }

    console.log('キャッシュが見つからないため、新しいレポートデータを生成します。');

    const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    const baseSheet = ss.getSheetByName(SHEET_NAME_BASE);
    const cvSheet = ss.getSheetByName(SHEET_NAME_CV);
    const keywordSheet = ss.getSheetByName(SHEET_NAME_KEYWORD);

    if (!baseSheet || !cvSheet || !keywordSheet) {
      throw new Error(`必要なシートが見つかりません。`);
    }
    console.log("シートの取得完了");

    const baseData = baseSheet.getDataRange().getValues();
    const cvData = cvSheet.getDataRange().getValues();
    const keywordData = keywordSheet.getDataRange().getValues();
    console.log("シートからのデータ読み込み完了");

    const baseHeaders = baseData.shift();
    const cvHeaders = cvData.shift();
    const keywordHeaders = keywordData.shift();

    const todayForPeriods = new Date();
    const lastMonthEndDate = new Date(todayForPeriods.getFullYear(), todayForPeriods.getMonth(), 0);
    const lastMonthStartDate = new Date(todayForPeriods.getFullYear(), todayForPeriods.getMonth() - 1, 1);
    const prevMonthEndDate = new Date(todayForPeriods.getFullYear(), todayForPeriods.getMonth() - 1, 0);
    const prevMonthStartDate = new Date(todayForPeriods.getFullYear(), todayForPeriods.getMonth() - 2, 1);

    console.log("データ集計処理を開始します...");
    const { lastMonthData, prevMonthData, monthlyData } = processAllData(
      baseData, cvData, keywordData,
      baseHeaders, cvHeaders, keywordHeaders,
      lastMonthStartDate, lastMonthEndDate,
      prevMonthStartDate, prevMonthEndDate
    );
    console.log("データ集計処理が完了しました。");

    const reportData = { lastMonthData, prevMonthData, monthlyData };

    cache.put(cacheKey, JSON.stringify(reportData), 21600); // 6時間キャッシュ
    console.log('新しいレポートデータを生成し、キャッシュに保存しました。');

    return reportData;

  } catch (e) {
    console.error("getReportData Error: " + e.toString());
    throw new Error("レポートデータの生成中にエラーが発生しました: " + e.message);
  }
}

/**
 * 全データを1回のループで効率的に集計する関数
 */
function processAllData(baseData, cvData, keywordData, baseHeaders, cvHeaders, keywordHeaders, lastMonthStartDate, lastMonthEndDate, prevMonthStartDate, prevMonthEndDate) {
  const getIndex = (headers, name) => headers.indexOf(name);
  const col = {
    base: { date: getIndex(baseHeaders, '日付'), device: getIndex(baseHeaders, 'デバイス'), campaign: getIndex(baseHeaders, 'キャンペーン名'), channel: getIndex(baseHeaders, '広告チャネルタイプ'), cost: getIndex(baseHeaders, 'ご利用額'), clicks: getIndex(baseHeaders, 'クリック数'), imp: getIndex(baseHeaders, '表示回数'), },
    cv: { date: getIndex(cvHeaders, '日付'), device: getIndex(cvHeaders, 'デバイス'), campaign: getIndex(cvHeaders, 'キャンペーン名'), action: getIndex(cvHeaders, 'コンバージョンアクション名'), cvs: getIndex(cvHeaders, 'コンバージョン数'), channel: getIndex(cvHeaders, '広告チャネルタイプ') },
    kw: { date: getIndex(keywordHeaders, '日付'), keyword: getIndex(keywordHeaders, 'キーワード'), match: getIndex(keywordHeaders, 'マッチタイプ'), cost: getIndex(keywordHeaders, 'ご利用額'), clicks: getIndex(keywordHeaders, 'クリック数'), cvs: getIndex(keywordHeaders, 'コンバージョン数'), }
  };

  const monthlyAgg = {};

  cvData.forEach(row => {
    try {
      const rowDate = new Date(row[col.cv.date]);
      if (isNaN(rowDate.getTime())) return;
      const monthKey = Utilities.formatDate(rowDate, 'JST', 'yyyy-MM');
      const actionName = row[col.cv.action] || '';
      const channel = row[col.cv.channel];

      if (channel === 'SEARCH' && !actionName.includes('中間')) {
        if (!monthlyAgg[monthKey]) monthlyAgg[monthKey] = { imp: 0, clicks: 0, cost: 0, cv: 0 };
        monthlyAgg[monthKey].cv += parseFloat(row[col.cv.cvs]) || 0;
      }
    } catch (e) { /* 無視 */ }
  });

  baseData.forEach(row => {
    try {
      const rowDate = new Date(row[col.base.date]);
      if (isNaN(rowDate.getTime())) return;
      if (row[col.base.channel] === 'SEARCH') {
        const monthKey = Utilities.formatDate(rowDate, 'JST', 'yyyy-MM');
        if (!monthlyAgg[monthKey]) monthlyAgg[monthKey] = { imp: 0, clicks: 0, cost: 0, cv: 0 };
        monthlyAgg[monthKey].imp += parseInt(row[col.base.imp]) || 0;
        monthlyAgg[monthKey].clicks += parseInt(row[col.base.clicks]) || 0;
        monthlyAgg[monthKey].cost += parseFloat(String(row[col.base.cost]).replace(/,/g, '')) || 0;
      }
    } catch (e) { /* 無視 */ }
  });

  Object.keys(monthlyAgg).forEach(monthKey => {
    const data = monthlyAgg[monthKey];
    data.ctr = data.imp > 0 ? (data.clicks / data.imp) : 0;
    data.cpc = data.clicks > 0 ? (data.cost / data.clicks) : 0;
    data.cvr = data.clicks > 0 ? (data.cv / data.clicks) : 0;
    data.cpa = data.cv > 0 ? (data.cost / data.cv) : 0;
  });

  const { lastMonthBreakdowns, prevMonthBreakdowns } = getPeriodBreakdowns(baseData, cvData, keywordData, col, lastMonthStartDate, lastMonthEndDate, prevMonthStartDate, prevMonthEndDate);

  const lastMonthKey = Utilities.formatDate(lastMonthStartDate, 'JST', 'yyyy-MM');
  const prevMonthKey = Utilities.formatDate(prevMonthStartDate, 'JST', 'yyyy-MM');

  const lastMonthTotals = monthlyAgg[lastMonthKey] || { imp: 0, clicks: 0, cost: 0, cv: 0, ctr: 0, cpc: 0, cvr: 0, cpa: 0 };
  const prevMonthTotals = monthlyAgg[prevMonthKey] || { imp: 0, clicks: 0, cost: 0, cv: 0, ctr: 0, cpc: 0, cvr: 0, cpa: 0 };

  const lastMonthResult = {
    period: `${Utilities.formatDate(lastMonthStartDate, 'JST', 'yyyy/MM/dd')} - ${Utilities.formatDate(lastMonthEndDate, 'JST', 'yyyy/MM/dd')}`,
    totalCost: lastMonthTotals.cost, totalClicks: lastMonthTotals.clicks, totalImpressions: lastMonthTotals.imp, totalConversions: lastMonthTotals.cv,
    ctr: lastMonthTotals.ctr, cvr: lastMonthTotals.cvr, cpa: lastMonthTotals.cpa,
    ...lastMonthBreakdowns
  };

  const prevMonthResult = {
    period: `${Utilities.formatDate(prevMonthStartDate, 'JST', 'yyyy/MM/dd')} - ${Utilities.formatDate(prevMonthEndDate, 'JST', 'yyyy/MM/dd')}`,
    totalCost: prevMonthTotals.cost, totalClicks: prevMonthTotals.clicks, totalImpressions: prevMonthTotals.imp, totalConversions: prevMonthTotals.cv,
    ...prevMonthBreakdowns
  };

  return { lastMonthData: lastMonthResult, prevMonthData: prevMonthResult, monthlyData: monthlyAgg };
}

function getPeriodBreakdowns(baseData, cvData, keywordData, col, lastMonthStartDate, lastMonthEndDate, prevMonthStartDate, prevMonthEndDate) {
    const lastMonthBreakdowns = { campaignData: {}, deviceData: {}, keywordData: {} };
    const prevMonthBreakdowns = { campaignData: {}, deviceData: {}, keywordData: {} };

    const cvMap = {};
    cvData.forEach(row => {
        try {
            const rowDate = new Date(row[col.cv.date]);
            if (isNaN(rowDate.getTime())) return;
            const actionName = row[col.cv.action] || '';
            if (!actionName.includes('中間')) {
                const key = `${Utilities.formatDate(rowDate, 'JST', 'yyyy-MM-dd')}|${row[col.cv.campaign]}|${row[col.cv.device]}`;
                const cvs = parseFloat(row[col.cv.cvs]) || 0;
                cvMap[key] = (cvMap[key] || 0) + cvs;
            }
        } catch (e) {}
    });

    baseData.forEach(row => {
        try {
            const rowDate = new Date(row[col.base.date]);
            if (isNaN(rowDate.getTime())) return;
            if (row[col.base.channel] !== 'SEARCH') return;

            let targetBreakdown = null;
            if (rowDate >= lastMonthStartDate && rowDate <= lastMonthEndDate) {
                targetBreakdown = lastMonthBreakdowns;
            } else if (rowDate >= prevMonthStartDate && rowDate <= prevMonthEndDate) {
                targetBreakdown = prevMonthBreakdowns;
            }

            if (targetBreakdown) {
                const key = `${Utilities.formatDate(rowDate, 'JST', 'yyyy-MM-dd')}|${row[col.base.campaign]}|${row[col.base.device]}`;
                const conversions = cvMap[key] || 0;
                const cost = parseFloat(String(row[col.base.cost]).replace(/,/g, '')) || 0;
                const clicks = parseInt(row[col.base.clicks]) || 0;

                const campaignName = row[col.base.campaign];
                if (!targetBreakdown.campaignData[campaignName]) targetBreakdown.campaignData[campaignName] = { cost: 0, clicks: 0, conversions: 0 };
                targetBreakdown.campaignData[campaignName].cost += cost;
                targetBreakdown.campaignData[campaignName].clicks += clicks;
                targetBreakdown.campaignData[campaignName].conversions += conversions;

                const deviceName = row[col.base.device];
                if (!targetBreakdown.deviceData[deviceName]) targetBreakdown.deviceData[deviceName] = { conversions: 0 };
                targetBreakdown.deviceData[deviceName].conversions += conversions;
            }
        } catch (e) {}
    });

    keywordData.forEach(row => {
        try {
            const rowDate = new Date(row[col.kw.date]);
            if (isNaN(rowDate.getTime())) return;
            if (rowDate >= lastMonthStartDate && rowDate <= lastMonthEndDate) {
                const kw = row[col.kw.keyword];
                if (!lastMonthBreakdowns.keywordData[kw]) lastMonthBreakdowns.keywordData[kw] = { clicks: 0, cost: 0, cvs: 0, match: row[col.kw.match] };
                lastMonthBreakdowns.keywordData[kw].clicks += parseInt(row[col.kw.clicks]) || 0;
                lastMonthBreakdowns.keywordData[kw].cost += parseFloat(String(row[col.kw.cost]).replace(/,/g, '')) || 0;
                lastMonthBreakdowns.keywordData[kw].cvs += parseInt(row[col.kw.cvs]) || 0;
            }
        } catch (e) {}
    });

    return { lastMonthBreakdowns, prevMonthBreakdowns };
}


/**
 * HTML側から呼び出され、Gemini APIで総括を生成する関数
 */
function getGeminiSummary(lastMonth, prevMonth) {
  const cache = CacheService.getScriptCache();
  const today = new Date();
  const cacheKey = `summary_text_${Utilities.formatDate(today, 'JST', 'yyyy-MM')}`;

  const cachedSummary = cache.get(cacheKey);
  if (cachedSummary) {
    console.log('キャッシュから総括を返します。');
    return cachedSummary;
  }

  console.log('キャッシュが見つからないため、新しい総括を生成します。');

  try {
    const getChange = (current, previous) => previous > 0 ? ((current / previous) - 1) * 100 : 0;
    const costChange = getChange(lastMonth.totalCost, prevMonth.totalCost);
    const clicksChange = getChange(lastMonth.totalClicks, prevMonth.totalClicks);
    const cvChange = getChange(lastMonth.totalConversions, prevMonth.totalConversions);

    const prompt = `
あなたはプロの広告運用コンサルタントです。以下のデータに基づいて、クライアント（不動産会社）向けの検索広告運用レポートの「総括」を記述してください。

# データ概要
- レポート対象: 検索広告
- 期間: ${lastMonth.period}
- 比較対象期間: ${prevMonth.period}

# 主要KPI (先月実績と前月比)
- ご利用額: ${Math.round(lastMonth.totalCost).toLocaleString()}円 (${costChange >= 0 ? '+' : ''}${costChange.toFixed(1)}%)
- クリック数: ${lastMonth.totalClicks.toLocaleString()}回 (${clicksChange >= 0 ? '+' : ''}${clicksChange.toFixed(1)}%)
- コンバージョン数: ${lastMonth.totalConversions.toLocaleString()}件 (${cvChange >= 0 ? '+' : ''}${cvChange.toFixed(1)}%)
- コンバージョン単価 (CPA): ${Math.round(lastMonth.cpa).toLocaleString()}円
- コンバージョン率 (CVR): ${(lastMonth.cvr * 100).toFixed(2)}%

# 指示
- 上記の数値を分析し、良かった点、考えられる課題、そして来月に向けた具体的な改善提案（ネクストアクション）をまとめてください。
- 箇条書きを用いて、簡潔で分かりやすく記述してください。
- 必ず以下のHTML形式で出力してください。Markdownなどの他の形式は使用しないでください。
<p><strong>【総括】</strong></p>
<p><strong>良かった点：</strong></p>
<ul>
  <li>（ここに良かった点を記述）</li>
</ul>
<p><strong>課題点：</strong></p>
<ul>
  <li>（ここに課題点を記述）</li>
</ul>
<p><strong>ネクストアクション：</strong></p>
<ul>
  <li>（ここに具体的な提案を記述）</li>
</ul>
`;


    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-05-20:generateContent?key=${geminiApiKey}`;

    if (!geminiApiKey || geminiApiKey === "ここにGenerative Language APIキーを貼り付けてください") {
      return "<p>（総括を自動生成するには、スクリプトにAPIキーを設定してください）</p>";
    }

    const payload = {
      contents: [{
        parts: [{ text: prompt }]
      }]
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
      const result = JSON.parse(response.getContentText());
      if (result.candidates && result.candidates.length > 0 &&
          result.candidates[0].content && result.candidates[0].content.parts &&
          result.candidates[0].content.parts.length > 0) {
        let text = result.candidates[0].content.parts[0].text;
        const formattedHtml = formatSummaryAsHtml(text);
        cache.put(cacheKey, formattedHtml, 21600); // 6時間キャッシュ
        console.log('新しい総括を生成し、キャッシュに保存しました。');
        return formattedHtml;
      } else {
        return "<p>総括を生成できませんでした。APIからの応答が予期しない形式です。</p>";
      }
    } else {
      console.error("Gemini API Error: " + response.getContentText());
      return "<p>総括の自動生成中にエラーが発生しました。</p>";
    }
  } catch(e) {
    console.error("Gemini API 呼び出し中に例外エラー: " + e.toString());
    return "<p>総括の自動生成中にエラーが発生しました。</p>";
  }
}

/**
 * GeminiからのテキストをHTMLに整形する関数
 */
function formatSummaryAsHtml(text) {
  if (text.trim().startsWith('<p>')) {
    return text; // すでにHTML形式の場合はそのまま返す
  }

  // Markdownの **太字** を <strong> タグに変換
  text = text.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');

  const lines = text.split('\n');
  let html = '';
  let inList = false;

  lines.forEach(line => {
    line = line.trim();
    if (line.startsWith('* ') || line.startsWith('- ')) {
      if (!inList) {
        html += '<ul>';
        inList = true;
      }
      html += '<li>' + line.substring(2) + '</li>';
    } else {
      if (inList) {
        html += '</ul>';
        inList = false;
      }
      if (line.length > 0) {
        if (line.includes('【') && line.includes('】')) {
            html += '<p><strong>' + line + '</strong></p>';
        } else {
            html += '<p>' + line + '</p>';
        }
      }
    }
  });

  if (inList) {
    html += '</ul>';
  }

  return html;
}
