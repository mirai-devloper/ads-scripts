<?php
// PHPでcomposer updateを実行するスクリプト例
// 注意: 実行ユーザーの権限やcomposerのパスに注意してください。

// composer.pharが同じディレクトリにある場合
$composerPath = __DIR__ . '/composer.phar';

// composerコマンドがグローバルインストールされている場合は下記のように
// $composerPath = 'composer';

$cmd = escapeshellcmd("php $composerPath update 2>&1");
$output = [];
$return_var = 0;
exec($cmd, $output, $return_var);

// 結果を表示
header('Content-Type: text/plain; charset=UTF-8');
echo "コマンド: $cmd\n";
echo "リターンコード: $return_var\n";
echo "---- 出力 ----\n";
echo implode("\n", $output);
