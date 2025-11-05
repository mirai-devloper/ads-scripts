<?php
require __DIR__ . '/vendor/autoload.php';

// --- 設定 ---
$spreadsheetId = 'YOUR_SPREADSHEET_ID'; // あなたのスプレッドシートID
$credentialsPath = __DIR__ . '/credentials.json';

// 読み込むシート名と範囲を定義
const SHEET_CONFIG = [
    'group' => [
        'name' => '全_グループ集計',
        'range' => 'A:J',
        'parser' => 'parseGroupData'
    ],
    'performance' => [
        'name' => 'パフォーマンス',
        'range' => 'A:H',
        'parser' => 'parsePerformanceData'
    ],
    'gender' => [
        'name' => '性別',
        'range' => 'A:I',
        'parser' => 'parseGenderData'
    ]
];

// --- データ取得・整形 ---

function getSheetData(string $spreadsheetId, string $range, string $credentialsPath): array {
    try {
        $client = new Google_Client();
        $client->setAuthConfig($credentialsPath);
        $client->addScope(Google_Service_Sheets::SPREADSHEETS_READONLY);
        $service = new Google_Service_Sheets($client);
        $response = $service->spreadsheets_values->get($spreadsheetId, $range);
        $values = $response->getValues() ?: [];
        if ($values) array_shift($values); // ヘッダーを削除
        return $values;
    } catch (Exception $e) {
        error_log('Error fetching Google Sheets data for range ' . $range . ': ' . $e->getMessage());
        return [];
    }
}

function parseNumber($value) {
    return is_numeric($v = preg_replace('/[^0-9.-]+/', '', $value)) ? floatval($v) : 0;
}

function parseDate($dateString) {
    if (empty($dateString)) return null;
    try {
        return (new DateTime($dateString))->format('Y-m-d');
    } catch (Exception $e) {
        return null;
    }
}

function parseGroupData(array $row): ?array {
    if (empty($row[2]) || empty($row[0])) return null; // グループ名と日付は必須
    return [
        'date' => parseDate($row[0]), 'campaignName' => $row[1] ?? '', 'groupName' => $row[2] ?? '',
        'cost' => parseNumber($row[3] ?? 0), 'impressions' => parseNumber($row[4] ?? 0),
        'clicks' => parseNumber($row[5] ?? 0), 'conversions' => parseNumber($row[7] ?? 0)
    ];
}

function parsePerformanceData(array $row): ?array {
    if (empty($row[1]) || empty($row[0])) return null; // キャンペーン名と日付は必須
    return [
        'date' => parseDate($row[0]), 'campaignName' => $row[1] ?? '',
        'cost' => parseNumber($row[3] ?? 0), 'impressions' => parseNumber($row[4] ?? 0),
        'clicks' => parseNumber($row[5] ?? 0), 'conversions' => parseNumber($row[7] ?? 0)
    ];
}

function parseGenderData(array $row): ?array {
    if (empty($row[8]) || empty($row[0])) return null; // 性別と日付は必須
    return [
        'date' => parseDate($row[0]), 'campaignName' => $row[1] ?? '', 'gender' => $row[8] ?? '',
        'cost' => parseNumber($row[3] ?? 0), 'impressions' => parseNumber($row[4] ?? 0),
        'clicks' => parseNumber($row[5] ?? 0), 'conversions' => parseNumber($row[7] ?? 0)
    ];
}

// --- メイン処理 ---
$allDataSets = [];
$allCampaigns = [];

foreach (SHEET_CONFIG as $key => $config) {
    $fullRange = $config['name'] . '!' . $config['range'];
    $rawData = getSheetData($spreadsheetId, $fullRange, $credentialsPath);
    $parsedData = [];
    foreach ($rawData as $row) {
        $parsedRow = $config['parser']($row);
        if ($parsedRow && $parsedRow['date']) {
            $parsedData[] = $parsedRow;
            if (!empty($parsedRow['campaignName'])) {
                $allCampaigns[] = $parsedRow['campaignName'];
            }
        }
    }
    $allDataSets[$key] = $parsedData;
}

// 全データからユニークなキャンペーンリストを作成
$uniqueCampaigns = array_values(array_unique($allCampaigns));
sort($uniqueCampaigns);
array_unshift($uniqueCampaigns, 'すべてのキャンペーン');

// JSONとしてフロントエンドに渡す
$allDataSetsJson = json_encode($allDataSets);
$availableCampaignsJson = json_encode($uniqueCampaigns);

require 'template.html';
?>