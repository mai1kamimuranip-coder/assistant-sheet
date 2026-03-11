# Gitブランチを develop に切り替える
Write-Host "🌿 Gitブランチを develop に切り替え中..." -ForegroundColor Cyan
git checkout develop

# 開発用スクリプトID
$devScriptId = "1cJ5cakSu9w5Lvvw5eYWMJeAi6jQNUSg1zGc6XY3VS0RdW7tl02v3Egvt"

Write-Host "🔄 開発環境（Development）に切り替え中..." -ForegroundColor Cyan

# .clasp.json の内容を読み込んで scriptId を書き換える
if (Test-Path ".clasp.json") {
    $config = Get-Content ".clasp.json" | ConvertFrom-Json
    $config.scriptId = $devScriptId
    $config | ConvertTo-Json | Set-Content ".clasp.json"
    Write-Host "✅ 開発環境への切り替えが完了しました。" -ForegroundColor Green
} else {
    Write-Host "❌ .clasp.json が見つかりません。" -ForegroundColor Red
    exit 1
}

Write-Host "`n確認: clasp setting" -ForegroundColor Yellow
clasp setting
