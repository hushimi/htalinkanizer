# https://qiita.com/asterisk9101/items/71b1b3dae68916a98271
# cmdプロンプト dirコマンドの結果をPowerShellで取得

Param([string]$watchPath)

# 前回の記録を消去
"---" | Out-File -Encoding default ($PSScriptRoot + "\result.txt");

# 空文字が渡された場合処理終了
if ([string]::IsNullOrEmpty($watchPath)) {
    "000" | Out-File -Encoding default ($PSScriptRoot + "\result.txt");
    exit;
}

[long]$freeSpace = cmd /c dir $watchPath `
    | Select-Object -last 1 `
    | ForEach-Object { $_ -replace ',' -split '\s+' } `
    | Select-Object -index 3
;

# 1024B = 1KB, 1024KB = 1MB, 1024MB = 1GB
# 取得した値をGB単位で表示
$convertedSize = [Math]::Ceiling(($freeSpace / 1GB) * [Math]::Pow(10, 2)) / [Math]::Pow(10, 2);
$convertedSize | Out-File -Encoding default ($PSScriptRoot + "\result.txt");