# Gitブランチを main に切り替える
Write-Host "🚢 Gitブランチを main に切り替え中..." -ForegroundColor Magenta
git checkout main

# 本番用スクリプトID
$prodScriptId = "122olPcWgL95V-rthNGb3QBgp_qFcmSYJ-MosAvVjc_9o97BCbp3xLcAC"

Write-Host "🚀 本番環境（Production）に切り替え中..." -ForegroundColor Magenta

# .clasp.json の内容を読み込んで scriptId を書き換える
if (Test-Path ".clasp.json") {
    $config = Get-Content ".clasp.json" | ConvertFrom-Json
    $config.scriptId = $prodScriptId
    $config | ConvertTo-Json | Set-Content ".clasp.json"
    Write-Host "✅ 本番環境への切り替えが完了しました。" -ForegroundColor Green
} else {
    Write-Host "❌ .clasp.json が見つかりません。" -ForegroundColor Red
    exit 1
}

Write-Host "`n確認: clasp setting" -ForegroundColor Yellow
clasp setting
