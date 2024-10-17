# https://qiita.com/asterisk9101/items/71b1b3dae68916a98271
# cmd�v�����v�g dir�R�}���h�̌��ʂ�PowerShell�Ŏ擾

Param([string]$watchPath)

# �O��̋L�^������
"---" | Out-File -Encoding default ($PSScriptRoot + "\result.txt");

# �󕶎����n���ꂽ�ꍇ�����I��
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
# �擾�����l��GB�P�ʂŕ\��
$convertedSize = [Math]::Ceiling(($freeSpace / 1GB) * [Math]::Pow(10, 2)) / [Math]::Pow(10, 2);
$convertedSize | Out-File -Encoding default ($PSScriptRoot + "\result.txt");