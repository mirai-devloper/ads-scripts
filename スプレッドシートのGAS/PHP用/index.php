<?php
// エラー表示を有効化
ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);
error_reporting(E_ALL);


// Google APIクライアントライブラリを読み込む
require __DIR__ . '/vendor/autoload.php';

// --- 設定 ---
// スプレッドシートのID（URLから取得）
// 例: https://docs.google.com/spreadsheets/d/【スプレッドシートID】/edit
$spreadsheetId = '1ViOwzqt-uqbEG-MEQJKrAQPnuw-iYDLiSsDId7bPGCE';

// 取得するシート名と範囲
// A列からH列の全ての行を取得
$range = '全_グループ集計!A:J';

// 認証情報ファイルのパス
$credentialsPath = __DIR__ . '/credentials.json';

// --- 列番号の定義 (0始まり) ---
const COL = [
    'DATE'        => 0, 'CAMPAIGN'    => 1, 'AD_GROUP'    => 2,
    'COST'        => 3, 'IMPRESSIONS' => 4, 'CLICKS'      => 5,
    'CONVERSIONS' => 7,
];

/**
 * 数値に変換するヘルパー関数
 */
function parseNumber($value) {
    $cleanedValue = preg_replace('/[^0-9.-]+/', '', $value);
    return is_numeric($cleanedValue) ? floatval($cleanedValue) : 0;
}

/**
 * Google Sheets APIからデータを取得し、配列に変換する関数
 */
function getAllAdGroupData(string $spreadsheetId, string $range, string $credentialsPath): array {
  try {
      $client = new Google_Client();
      $client->setAuthConfig($credentialsPath);
      $client->addScope(Google_Service_Sheets::SPREADSHEETS_READONLY); // 読み取り専用スコープ

      $service = new Google_Service_Sheets($client);
      $response = $service->spreadsheets_values->get($spreadsheetId, $range);
      $values = $response->getValues();

      // ヘッダー行を削除
      if ($values) {
          array_shift($values);
      }

      $allData = [];
      foreach ($values as $row) {
          // C列（広告グループ）かA列（日付）が空の行はスキップ
          if (empty($row[COL['AD_GROUP']]) || empty($row[COL['DATE']])) {
              continue;
          }

          // ▼▼▼ ここが修正部分 ▼▼▼
          // 日付文字列を直接DateTimeオブジェクトに変換
          // '2025/10/09' や '2025-10-09' のような一般的な形式に対応
          try {
              $date = new DateTime($row[COL['DATE']]);
          } catch (Exception $e) {
              // もし日付の形式が特殊でパースに失敗した場合は、エラーを防ぐためその行をスキップ
              continue;
          }
          // ▲▲▲ 修正はここまで ▲▲▲

          $allData[] = [
              'date'         => $date->format('Y-m-d'),
              'campaignName' => $row[COL['CAMPAIGN']] ?? '',
              'groupName'    => $row[COL['AD_GROUP']] ?? '',
              'cost'         => parseNumber($row[COL['COST']] ?? '0'),
              'impressions'  => parseNumber($row[COL['IMPRESSIONS']] ?? '0'),
              'clicks'       => parseNumber($row[COL['CLICKS']] ?? '0'),
              'conversions'  => parseNumber($row[COL['CONVERSIONS']] ?? '0'),
          ];
      }
      return $allData;

  } catch (Exception $e) {
      // エラーが発生した場合は空の配列を返し、エラー内容をログに出力
      error_log('Error fetching Google Sheets data: ' . $e->getMessage());
      return [];
  }
}

/**
 * データから利用可能なキャンペーン名のリストを取得する関数
 */
function getAvailableCampaigns(array $allData): array {
    $campaignNames = array_column($allData, 'campaignName');
    $uniqueCampaigns = array_filter(array_unique($campaignNames));
    sort($uniqueCampaigns);
    array_unshift($uniqueCampaigns, 'すべてのキャンペーン');
    return $uniqueCampaigns;
}

// --- メイン処理 ---
$allData = getAllAdGroupData($spreadsheetId, $range, $credentialsPath);
$availableCampaigns = getAvailableCampaigns($allData);

// データをJavaScriptで使えるようにJSON形式に変換
$allDataJson = json_encode($allData);
$availableCampaignsJson = json_encode($availableCampaigns);

// HTMLテンプレートを読み込んで表示
require 'template.html';
