/**
 * 月次広告レポートをHTML形式で自動生成するスクリプト (月別データ強化・3シート対応版)
 * * 実行すると、月別の実績やサマリーを含むレポートをHTMLファイルとして
 * Googleドライブのルートフォルダに保存します。
 * * 「基本データ」「コンバージョンデータ」「キーワード別データ」の3シートからデータを取得します。
 * * 検索広告のデータのみを対象とします。
 * * Gemini APIを使用して総括を自動生成します。
 */


/**
 * メインの実行関数
 */
function createMonthlyReportFrom3Sheets() {
  try {
    // --- 1. スプレッドシートとデータの準備 ---
    const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);

    const baseSheet = ss.getSheetByName(SHEET_NAME_BASE);
    const cvSheet = ss.getSheetByName(SHEET_NAME_CV);
    const keywordSheet = ss.getSheetByName(SHEET_NAME_KEYWORD);

    if (!baseSheet || !cvSheet || !keywordSheet) {
      throw new Error(`必要なシート（${SHEET_NAME_BASE}, ${SHEET_NAME_CV}, ${SHEET_NAME_KEYWORD}）のいずれかが見つかりません。`);
    }

    const baseData = baseSheet.getDataRange().getValues();
    const cvData = cvSheet.getDataRange().getValues();
    const keywordData = keywordSheet.getDataRange().getValues();

    const baseHeaders = baseData.shift();
    const cvHeaders = cvData.shift();
    const keywordHeaders = keywordData.shift();

    // --- 2. 期間の定義 (先月・先々月) ---
    const today = new Date();
    const lastMonthEndDate = new Date(today.getFullYear(), today.getMonth(), 0);
    const lastMonthStartDate = new Date(today.getFullYear(), today.getMonth() - 1, 1);
    const prevMonthEndDate = new Date(today.getFullYear(), today.getMonth() - 1, 0);
    const prevMonthStartDate = new Date(today.getFullYear(), today.getMonth() - 2, 1);

    // --- 3. 各期間のデータを集計 ---
    const lastMonthData = aggregateData(baseData, cvData, keywordData, baseHeaders, cvHeaders, keywordHeaders, lastMonthStartDate, lastMonthEndDate);
    const prevMonthData = aggregateData(baseData, cvData, keywordData, baseHeaders, cvHeaders, keywordHeaders, prevMonthStartDate, prevMonthEndDate);
    const monthlyData = aggregateDataForMonthlyView(baseData, cvData, baseHeaders, cvHeaders);


    // --- 4. HTMLレポートを生成 ---
    const reportHtml = generateHtmlReport(lastMonthData, prevMonthData, monthlyData);

    // --- 5. HTMLファイルをドライブに保存 ---
    const reportTitle = `【広告レポート_検索】${Utilities.formatDate(lastMonthStartDate, 'JST', 'yyyy-MM')}.html`;
    DriveApp.createFile(reportTitle, reportHtml, MimeType.HTML);

    console.log(`レポート「${reportTitle}」が正常に作成されました。`);
    SpreadsheetApp.getUi().alert(`レポート「${reportTitle}」がGoogleドライブに作成されました。`);

  } catch (e) {
    console.error('レポート作成中にエラーが発生しました: ' + e.toString());
    SpreadsheetApp.getUi().alert('エラー: ' + e.message);
  }
}

/**
 * 3つのシートからデータを集計する関数
 */
function aggregateData(baseData, cvData, keywordData, baseHeaders, cvHeaders, keywordHeaders, startDate, endDate) {
  const getIndex = (headers, name) => headers.indexOf(name);

  const col = {
    base: { date: getIndex(baseHeaders, '日付'), device: getIndex(baseHeaders, 'デバイス'), campaign: getIndex(baseHeaders, 'キャンペーン名'), channel: getIndex(baseHeaders, '広告チャネルタイプ'), cost: getIndex(baseHeaders, 'ご利用額'), clicks: getIndex(baseHeaders, 'クリック数'), imp: getIndex(baseHeaders, '表示回数'), },
    cv: { date: getIndex(cvHeaders, '日付'), device: getIndex(cvHeaders, 'デバイス'), campaign: getIndex(cvHeaders, 'キャンペーン名'), action: getIndex(cvHeaders, 'コンバージョンアクション名'), cvs: getIndex(cvHeaders, 'コンバージョン数'), },
    kw: { date: getIndex(keywordHeaders, '日付'), keyword: getIndex(keywordHeaders, 'キーワード'), match: getIndex(keywordHeaders, 'マッチタイプ'), cost: getIndex(keywordHeaders, 'ご利用額'), clicks: getIndex(keywordHeaders, 'クリック数'), cvs: getIndex(keywordHeaders, 'コンバージョン数'), }
  };

  const cvMap = {};
  cvData.forEach(row => {
    try {
      const rowDate = new Date(row[col.cv.date]);
      const actionName = row[col.cv.action] || '';
      if (rowDate >= startDate && rowDate <= endDate && !actionName.includes('中間')) {
        const key = `${Utilities.formatDate(rowDate, 'JST', 'yyyy-MM-dd')}|${row[col.cv.campaign]}|${row[col.cv.device]}`;
        const cvs = parseFloat(row[col.cv.cvs]) || 0;
        cvMap[key] = (cvMap[key] || 0) + cvs;
      }
    } catch(e) {}
  });

  let totalCost = 0, totalClicks = 0, totalImpressions = 0, totalConversions = 0;
  const campaignAgg = {}, deviceAgg = {};

  baseData.forEach(row => {
    try {
      const rowDate = new Date(row[col.base.date]);
      if (rowDate >= startDate && rowDate <= endDate && row[col.base.channel] === 'SEARCH') {
        const key = `${Utilities.formatDate(rowDate, 'JST', 'yyyy-MM-dd')}|${row[col.base.campaign]}|${row[col.base.device]}`;
        const conversions = cvMap[key] || 0;
        const cost = parseFloat(String(row[col.base.cost]).replace(/,/g, '')) || 0;
        const clicks = parseInt(row[col.base.clicks]) || 0;
        const impressions = parseInt(row[col.base.imp]) || 0;
        totalCost += cost; totalClicks += clicks; totalImpressions += impressions; totalConversions += conversions;
        const campaignName = row[col.base.campaign];
        if (!campaignAgg[campaignName]) campaignAgg[campaignName] = { cost: 0, clicks: 0, conversions: 0 };
        campaignAgg[campaignName].cost += cost; campaignAgg[campaignName].clicks += clicks; campaignAgg[campaignName].conversions += conversions;
        const deviceName = row[col.base.device];
        if (!deviceAgg[deviceName]) deviceAgg[deviceName] = { conversions: 0 };
        deviceAgg[deviceName].conversions += conversions;
      }
    } catch(e) {}
  });

  const keywordAgg = {};
  keywordData.forEach(row => {
    try {
      const rowDate = new Date(row[col.kw.date]);
      if (rowDate >= startDate && rowDate <= endDate) {
        const kw = row[col.kw.keyword];
        if (!keywordAgg[kw]) keywordAgg[kw] = { clicks: 0, cost: 0, cvs: 0, match: row[col.kw.match] };
        keywordAgg[kw].clicks += parseInt(row[col.kw.clicks]) || 0;
        keywordAgg[kw].cost += parseFloat(String(row[col.kw.cost]).replace(/,/g, '')) || 0;
        keywordAgg[kw].cvs += parseInt(row[col.kw.cvs]) || 0;
      }
    } catch(e) {}
  });

  return {
    period: `${Utilities.formatDate(startDate, 'JST', 'yyyy/MM/dd')} - ${Utilities.formatDate(endDate, 'JST', 'yyyy/MM/dd')}`,
    totalCost, totalClicks, totalImpressions, totalConversions,
    ctr: totalImpressions > 0 ? (totalClicks / totalImpressions) : 0,
    cvr: totalClicks > 0 ? (totalConversions / totalClicks) : 0,
    cpa: totalConversions > 0 ? (totalCost / totalConversions) : 0,
    campaignData: campaignAgg, deviceData: deviceAgg, keywordData: keywordAgg
  };
}


/**
 * 月別実績データを集計する関数 (修正版)
 */
function aggregateDataForMonthlyView(baseData, cvData, baseHeaders, cvHeaders) {
    const getIndex = (headers, name) => headers.indexOf(name);
    const col = {
        base: { date: getIndex(baseHeaders, '日付'), channel: getIndex(baseHeaders, '広告チャネルタイプ'), imp: getIndex(baseHeaders, '表示回数'), clicks: getIndex(baseHeaders, 'クリック数'), cost: getIndex(baseHeaders, 'ご利用額'), },
        cv: { date: getIndex(cvHeaders, '日付'), action: getIndex(cvHeaders, 'コンバージョンアクション名'), cvs: getIndex(cvHeaders, 'コンバージョン数'), channel: getIndex(cvHeaders, '広告チャネルタイプ') }
    };

    const cvMapMonthly = {};
    cvData.forEach(row => {
      try {
        const rowDate = new Date(row[col.cv.date]);
        const monthKey = Utilities.formatDate(rowDate, 'JST', 'yyyy-MM');
        const actionName = row[col.cv.action] || '';
        const channel = row[col.cv.channel];

        if (channel === 'SEARCH' && !actionName.includes('中間')) {
            const cvs = parseFloat(row[col.cv.cvs]) || 0;
            cvMapMonthly[monthKey] = (cvMapMonthly[monthKey] || 0) + cvs;
        }
      } catch(e) {}
    });

    const monthlyAgg = {};
    baseData.forEach(row => {
      try {
        if (row[col.base.channel] === 'SEARCH') {
            const rowDate = new Date(row[col.base.date]);
            const monthKey = Utilities.formatDate(rowDate, 'JST', 'yyyy-MM');
            if (!monthlyAgg[monthKey]) {
                monthlyAgg[monthKey] = { imp: 0, clicks: 0, cost: 0, cv: 0 };
            }
            monthlyAgg[monthKey].imp += parseInt(row[col.base.imp]) || 0;
            monthlyAgg[monthKey].clicks += parseInt(row[col.base.clicks]) || 0;
            monthlyAgg[monthKey].cost += parseFloat(String(row[col.base.cost]).replace(/,/g, '')) || 0;
        }
      } catch(e) {}
    });

    Object.keys(monthlyAgg).forEach(monthKey => {
        monthlyAgg[monthKey].cv = cvMapMonthly[monthKey] || 0;
        const data = monthlyAgg[monthKey];
        data.ctr = data.imp > 0 ? (data.clicks / data.imp) : 0;
        data.cpc = data.clicks > 0 ? (data.cost / data.clicks) : 0;
        data.cvr = data.clicks > 0 ? (data.cv / data.clicks) : 0;
        data.cpa = data.cv > 0 ? (data.cost / data.cv) : 0;
    });

    return monthlyAgg;
}

/**
 * Gemini APIを呼び出して総括を生成する関数
 */
function generateSummaryWithGemini(lastMonth, prevMonth, costChange, clicksChange, cvChange) {
  try {
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
        const text = result.candidates[0].content.parts[0].text;
        return text;
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


/**
 * 集計データからHTMLレポートを生成する関数
 */
function generateHtmlReport(lastMonth, prevMonth, monthlyData) {
  const getChange = (current, previous) => previous > 0 ? ((current / previous) - 1) * 100 : 0;
  const costChange = getChange(lastMonth.totalCost, prevMonth.totalCost);
  const clicksChange = getChange(lastMonth.totalClicks, prevMonth.totalClicks);
  const cvChange = getChange(lastMonth.totalConversions, prevMonth.totalConversions);

  // Gemini APIで総括を生成し、HTMLに整形
  let summaryText = generateSummaryWithGemini(lastMonth, prevMonth, costChange, clicksChange, cvChange);
  summaryText = formatSummaryAsHtml(summaryText);

  const deviceLabels = Object.keys(lastMonth.deviceData);
  const deviceCvData = deviceLabels.map(label => lastMonth.deviceData[label].conversions);

  const sortedMonths = Object.keys(monthlyData).sort();
  const monthlyHeaders = sortedMonths.map(m => {
      const [year, month] = m.split('-');
      return `${year}年${parseInt(month, 10)}月`;
  });

  const monthlyChartLabels = sortedMonths.map(m => m.replace('-', '/'));
  const monthlyCpcData = sortedMonths.map(m => Math.round(monthlyData[m].cpc));
  const monthlyCvrData = sortedMonths.map(m => (monthlyData[m].cvr * 100).toFixed(2));

  const simulationData = [
    "出稿費|¥70,000|¥70,000|¥70,000|¥70,000|¥70,000",
    "クリック単価|¥600|¥550|¥500|¥450|¥400",
    "クリック数|117|127|140|156|175",
    "実CVR 0.37%|0.43|0.47|0.52|0.58|0.65",
    "実CVR 0.43%|0.50|0.55|0.60|0.67|0.75",
    "実CVR 1.00%|1.17|1.27|1.40|1.56|1.75",
    "実CVR 1.41%|1.65|1.79|1.97|2.19|2.47"
  ];

  const simulationTableRows = simulationData.map(rowStr => {
      const cells = rowStr.split('|');
      const header = cells.shift();
      const dataCells = cells.map(cell => `<td class="px-6 py-4 text-right">${cell}</td>`).join('');
      return `<tr><td class="px-6 py-4 font-medium">${header}</td>${dataCells}</tr>`;
  }).join('');


  const getMonthlyRow = (label, key, format) => {
      let cells = '';
      sortedMonths.forEach(m => {
          let value = monthlyData[m][key];
          switch (format) {
              case 'yen': value = `¥${Math.round(value).toLocaleString()}`; break;
              case 'percent': value = `${(value * 100).toFixed(2)}%`; break;
              case 'number': value = value.toLocaleString(); break;
          }
          cells += `<td class="px-3 py-3 text-right whitespace-nowrap border-l border-gray-300">${value}</td>`;
      });
      return `<tr><td class="px-3 py-3 font-medium whitespace-nowrap sticky left-0 bg-white z-10 border-l-4 border-white border-r border-gray-300">${label}</td>${cells}</tr>`;
  };

  const htmlTemplate = `
    <!DOCTYPE html>
    <html lang="ja">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>広告運用詳細レポート (検索広告)</title>
        <script src="https://cdn.tailwindcss.com"></script>
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Noto+Sans+JP:wght@400;500;700&display=swap" rel="stylesheet">
        <style>
          body { font-family: 'Inter', 'Noto Sans JP', sans-serif; background-color: #f3f4f6; }
          .tab-active { border-color: #3b82f6; color: #3b82f6; background-color: #ffffff; }
        </style>
    </head>
    <body class="p-4 sm:p-6 md:p-8">
        <div class="max-w-7xl mx-auto bg-gray-100">
            <div class="mb-6">
                <h1 class="text-3xl font-bold text-gray-800">広告運用詳細レポート (検索広告)</h1>
                <p class="text-gray-500">期間: ${lastMonth.period}</p>
                <p class="text-sm text-gray-500">比較対象期間: ${prevMonth.period}</p>
            </div>

            <div class="mb-6"><div class="border-b border-gray-200"><nav class="-mb-px flex space-x-6" aria-label="Tabs">
                <button onclick="changeTab('summary')" id="tab-summary" class="tab-active whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm">サマリー</button>
                <button onclick="changeTab('monthly')" id="tab-monthly" class="text-gray-500 hover:text-gray-700 hover:border-gray-300 whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm">月別データ</button>
                <button onclick="changeTab('keyword')" id="tab-keyword" class="text-gray-500 hover:text-gray-700 hover:border-gray-300 whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm">キーワード別実績</button>
            </nav></div></div>

            <div id="content-summary" class="tab-content">
                <div class="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-4 mb-6">
                    <div class="bg-white p-4 rounded-lg shadow-sm text-center"><h3 class="text-sm font-medium text-gray-500">費用</h3><p class="mt-1 text-2xl font-bold text-gray-900">¥${Math.round(lastMonth.totalCost).toLocaleString()}</p><p class="mt-1 text-xs ${costChange >= 0 ? 'text-red-600' : 'text-green-600'}">(${costChange.toFixed(1)}% vs 前月)</p></div>
                    <div class="bg-white p-4 rounded-lg shadow-sm text-center"><h3 class="text-sm font-medium text-gray-500">クリック数</h3><p class="mt-1 text-2xl font-bold text-gray-900">${lastMonth.totalClicks.toLocaleString()}</p><p class="mt-1 text-xs ${clicksChange >= 0 ? 'text-green-600' : 'text-red-600'}">(${clicksChange.toFixed(1)}% vs 前月)</p></div>
                    <div class="bg-white p-4 rounded-lg shadow-sm text-center"><h3 class="text-sm font-medium text-gray-500">CTR</h3><p class="mt-1 text-2xl font-bold text-gray-900">${(lastMonth.ctr * 100).toFixed(2)}%</p></div>
                    <div class="bg-white p-4 rounded-lg shadow-sm text-center"><h3 class="text-sm font-medium text-gray-500">CV</h3><p class="mt-1 text-2xl font-bold text-gray-900">${lastMonth.totalConversions.toLocaleString()}</p><p class="mt-1 text-xs ${cvChange >= 0 ? 'text-green-600' : 'text-red-600'}">(${cvChange.toFixed(1)}% vs 前月)</p></div>
                    <div class="bg-white p-4 rounded-lg shadow-sm text-center"><h3 class="text-sm font-medium text-gray-500">CVR</h3><p class="mt-1 text-2xl font-bold text-gray-900">${(lastMonth.cvr * 100).toFixed(2)}%</p></div>
                    <div class="bg-white p-4 rounded-lg shadow-sm text-center"><h3 class="text-sm font-medium text-gray-500">CPA</h3><p class="mt-1 text-2xl font-bold text-gray-900">¥${Math.round(lastMonth.cpa).toLocaleString()}</p></div>
                </div>
                <div class="grid grid-cols-1 lg:grid-cols-5 gap-6 mb-6"><div class="lg:col-span-3 bg-white p-6 rounded-lg shadow-sm"><h3 class="font-semibold text-gray-800 mb-4">デバイス別CV比率</h3><div class="relative h-80"><canvas id="deviceChart"></canvas></div></div><div class="lg:col-span-2 bg-white p-6 rounded-lg shadow-sm"><h3 class="font-semibold text-gray-800 mb-4">総括</h3><div class="space-y-4 text-sm text-gray-700">${summaryText}</div></div></div>
                <div class="bg-white p-4 sm:p-6 rounded-lg shadow-sm overflow-x-auto"><h3 class="font-semibold text-gray-800 mb-4">キャンペーン別実績</h3><table class="w-full text-sm text-left text-gray-500"><thead class="text-xs text-gray-700 uppercase bg-gray-50"><tr><th scope="col" class="px-3 py-3">キャンペーン</th><th scope="col" class="px-3 py-3 text-right">費用</th><th scope="col" class="px-3 py-3 text-right">クリック数</th><th scope="col" class="px-3 py-3 text-right">CV</th><th scope="col" class="px-3 py-3 text-right">CPA</th></tr></thead><tbody>${Object.keys(lastMonth.campaignData).sort((a,b) => lastMonth.campaignData[b].cost - lastMonth.campaignData[a].cost).map(name => { const c = lastMonth.campaignData[name]; const cpa = c.conversions > 0 ? Math.round(c.cost / c.conversions) : 0; return `<tr class="bg-white border-b hover:bg-gray-50"><th scope="row" class="px-3 py-3 font-medium text-gray-900 whitespace-nowrap">${name}</th><td class="px-3 py-3 text-right">¥${Math.round(c.cost).toLocaleString()}</td><td class="px-3 py-3 text-right">${c.clicks.toLocaleString()}</td><td class="px-3 py-3 text-right font-bold">${c.conversions.toLocaleString()}</td><td class="px-3 py-3 text-right">¥${cpa.toLocaleString()}</td></tr>`; }).join('')}</tbody></table></div>
            </div>

            <div id="content-monthly" class="tab-content hidden">
                 <div class="bg-white p-4 sm:p-6 rounded-lg shadow-sm mb-6">
                    <h3 class="font-semibold text-gray-800 mb-4">月別実績データ</h3>
                    <div id="monthly-table-container" class="overflow-x-auto">
                        <table class="w-full text-sm text-left text-gray-500">
                            <thead class="text-xs text-gray-700 uppercase bg-gray-50 sticky top-0 z-20">
                                <tr>
                                    <th class="px-3 py-3 sticky left-0 bg-gray-50 z-30 border-l-4 border-gray-50 border-r border-gray-300">指標</th>
                                    ${monthlyHeaders.map(h => `<th class="px-3 py-3 text-right min-w-[150px] border-l border-gray-300">${h}</th>`).join('')}
                                </tr>
                            </thead>
                            <tbody class="divide-y divide-gray-200">
                                ${getMonthlyRow('表示回数', 'imp', 'number')}
                                ${getMonthlyRow('クリック数', 'clicks', 'number')}
                                ${getMonthlyRow('クリック率 (CTR)', 'ctr', 'percent')}
                                ${getMonthlyRow('平均クリック単価 (CPC)', 'cpc', 'yen')}
                                ${getMonthlyRow('コンバージョン数 (CV)', 'cv', 'number')}
                                ${getMonthlyRow('コンバージョン率 (CVR)', 'cvr', 'percent')}
                                ${getMonthlyRow('コンバージョン単価 (CPA)', 'cpa', 'yen')}
                                ${getMonthlyRow('ご利用額', 'cost', 'yen')}
                            </tbody>
                        </table>
                    </div>
                </div>
                <div class="bg-white p-6 rounded-lg shadow-sm mb-6"><h3 class="font-semibold text-gray-800 mb-4">月別 CPC・CVR 推移</h3><div class="relative h-80"><canvas id="monthlyTrendChart"></canvas></div></div>
                <div class="bg-white p-4 sm:p-6 rounded-lg shadow-sm overflow-x-auto">
                    <h3 class="font-semibold text-gray-800 mb-4">シミュレーション</h3>
                    <table class="w-full text-sm text-left text-gray-500"><thead class="text-xs text-gray-700 uppercase bg-gray-50"><tr><th class="px-6 py-3">指標</th><th class="px-6 py-3 text-right">1ヶ月目</th><th class="px-6 py-3 text-right">2ヶ月目</th><th class="px-6 py-3 text-right">3ヶ月目</th><th class="px-6 py-3 text-right">4ヶ月目</th><th class="px-6 py-3 text-right">5ヶ月目</th></tr></thead><tbody class="divide-y divide-gray-200">${simulationTableRows}</tbody></table>
                </div>
            </div>

            <div id="content-keyword" class="tab-content hidden">
                <div class="bg-white p-4 sm:p-6 rounded-lg shadow-sm overflow-x-auto"><h3 class="font-semibold text-gray-800 mb-4">キーワード別実績</h3><p class="text-xs text-gray-500 mb-4">※この表のコンバージョン数はキーワード別データの数値を参照しています。</p><table class="w-full text-sm text-left text-gray-500"><thead class="text-xs text-gray-700 uppercase bg-gray-50"><tr><th scope="col" class="px-6 py-3">キーワード</th><th scope="col" class="px-6 py-3">マッチタイプ</th><th scope="col" class="px-6 py-3 text-right">費用</th><th scope="col" class="px-6 py-3 text-right">クリック数</th><th scope="col" class="px-6 py-3 text-right">CV数</th></tr></thead><tbody>${Object.keys(lastMonth.keywordData).sort((a,b) => lastMonth.keywordData[b].cost - lastMonth.keywordData[a].cost).slice(0, 50).map(kw => { const k = lastMonth.keywordData[kw]; return `<tr class="bg-white border-b hover:bg-gray-50"><th scope="row" class="px-6 py-4 font-medium text-gray-900 whitespace-nowrap">${kw}</th><td class="px-6 py-4">${k.match}</td><td class="px-6 py-4 text-right">¥${Math.round(k.cost).toLocaleString()}</td><td class="px-6 py-4 text-right">${k.clicks.toLocaleString()}</td><td class="px-6 py-4 text-right font-bold">${k.cvs.toLocaleString()}</td></tr>`; }).join('')}</tbody></table></div>
            </div>
        </div>
        <script>
            function changeTab(selectedTab) {
                ['summary', 'monthly', 'keyword'].forEach(tab => {
                    document.getElementById(\`tab-\${tab}\`).classList.toggle('tab-active', tab === selectedTab);
                    document.getElementById(\`tab-\${tab}\`).classList.toggle('text-gray-500', tab !== selectedTab);
                    document.getElementById(\`content-\${tab}\`).classList.toggle('hidden', tab !== selectedTab);
                });
                if (selectedTab === 'monthly') {
                    const tableContainer = document.getElementById('monthly-table-container');
                    if(tableContainer) tableContainer.scrollLeft = tableContainer.scrollWidth;
                }
            }
            // Summary Chart
            const deviceCtx = document.getElementById('deviceChart').getContext('2d');
            new Chart(deviceCtx, { type: 'doughnut', data: { labels: ${JSON.stringify(deviceLabels)}, datasets: [{ data: ${JSON.stringify(deviceCvData)}, backgroundColor: ['#3b82f6', '#60a5fa', '#93c5fd', '#bfdbfe'] }] }, options: { responsive: true, maintainAspectRatio: false } });

            // Monthly Trend Chart
            const monthlyTrendCtx = document.getElementById('monthlyTrendChart').getContext('2d');
            new Chart(monthlyTrendCtx, {
                type: 'line',
                data: {
                    labels: ${JSON.stringify(monthlyChartLabels)},
                    datasets: [
                        { label: '平均クリック単価 (CPC)', data: ${JSON.stringify(monthlyCpcData)}, borderColor: '#3b82f6', backgroundColor: '#3b82f6', yAxisID: 'yCpc', tension: 0.1 },
                        { label: 'コンバージョン率 (CVR)', data: ${JSON.stringify(monthlyCvrData)}, borderColor: '#f97616', backgroundColor: '#f97616', yAxisID: 'yCvr', tension: 0.1 }
                    ]
                },
                options: { responsive: true, maintainAspectRatio: false, scales: { yCpc: { type: 'linear', display: true, position: 'left', title: { display: true, text: 'CPC (円)' } }, yCvr: { type: 'linear', display: true, position: 'right', title: { display: true, text: 'CVR (%)' }, grid: { drawOnChartArea: false } } } }
            });

            // Initial scroll for monthly table if it's the default view (it's not, but good practice)
            document.addEventListener('DOMContentLoaded', (event) => {
                const tableContainer = document.getElementById('monthly-table-container');
                if(tableContainer && !tableContainer.parentElement.classList.contains('hidden')) {
                    tableContainer.scrollLeft = tableContainer.scrollWidth;
                }
            });
        </script>
    </body>
    </html>
  `;
  return htmlTemplate;
}
