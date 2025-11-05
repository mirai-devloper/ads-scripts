<?php
// エラーを強制的に表示する設定
ini_set('display_errors', 1);
error_reporting(E_ALL);

// --- ここから診断スクリプト ---
header('Content-Type: text/html; charset=utf-8');
echo '<!DOCTYPE html><html><head><title>サーバー環境診断</title><style>body{font-family: sans-serif; padding: 20px;} .ok{color: green; font-weight: bold;} .ng{color: red; font-weight: bold;}</style></head><body>';
echo '<h1>サーバー環境 診断レポート</h1>';

// --- 診断1: PHPバージョンとパスの確認 ---
echo '<h2>ステップ1: 基本情報の確認</h2>';
echo '<p>PHPバージョン: ' . phpversion() . '</p>';
$autoloadPath = __DIR__ . '/vendor/autoload.php';
echo '<p>想定している autoload.php のフルパス: ' . $autoloadPath . '</p>';
echo '<hr>';

// --- 診断2: ファイルの存在と読み取り権限の確認 ---
echo '<h2>ステップ2: ファイルの存在と権限の確認</h2>';
if (file_exists($autoloadPath)) {
    echo '<p class="ok">✅ OK: autoload.php は指定されたパスに存在します。</p>';

    if (is_readable($autoloadPath)) {
        echo '<p class="ok">✅ OK: autoload.php はPHPスクリプトから読み取り可能です。</p>';
    } else {
        echo '<p class="ng">❌ NG: autoload.php は存在しますが、読み取り権限がありません！(パーミッションの問題)</p>';
    }
} else {
    echo '<p class="ng">❌ NG: autoload.php が指定されたパスに存在しません！(パスまたはアップロードの問題)</p>';
}
echo '<hr>';

// --- 診断3: autoload.php の読み込みテスト ---
echo '<h2>ステップ3: ライブラリ読み込みテスト</h2>';
try {
    // require_once を試行
    require_once $autoloadPath;
    echo '<p class="ok">✅ OK: require_once は成功しました。ファイルは正常に読み込まれました。</p>';

    // --- 診断4: クラスの存在確認 ---
    echo '<h2>ステップ4: クラスの存在確認</h2>';
    $targetClass = 'Google\Auth\Cache\MemoryCacheItemPool';
    if (class_exists($targetClass)) {
        echo '<p class="ok">✅ OK: 目的のクラス (' . $targetClass . ') は正常に認識されています。</p>';
    } else {
        echo '<p class="ng">❌ NG: autoload.php は読み込めましたが、目的のクラス (' . $targetClass . ') が見つかりません。vendorフォルダが破損しているか、不完全な可能性があります。</p>';
    }

} catch (Throwable $t) {
    // require_once で致命的なエラーが発生した場合
    echo '<p class="ng">❌ NG: require_once の実行中に致命的なエラーが発生しました。</p>';
    echo '<p><strong>エラーメッセージ:</strong></p>';
    echo '<pre style="background: #eee; padding: 10px; border: 1px solid #ccc;">' . htmlspecialchars($t->getMessage()) . '</pre>';
}

echo '</body></html>';
?>