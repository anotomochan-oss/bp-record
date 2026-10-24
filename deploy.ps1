# こけログ（血圧記録システム）：変更を GitHub（画面）と Apps Script（データ担当）に反映する
# 使い方：powershell -ExecutionPolicy Bypass -File deploy.ps1 "変更内容のメモ"
param([string]$Message = '更新')

Set-Location $PSScriptRoot

# ウェブアプリのURL（https://script.google.com/macros/s/○○○/exec）の ○○○ の部分
$DeploymentId = 'AKfycbxnD0JHhgeQi3GyYNb0QUVMk6sT1edyCpgNiFH8TCQIAu2yhjSn5iC_J35rKylvqZT2'
$GasHashFile = '.last-gas-deploy'

function Invoke-Step($label, [scriptblock]$block) {
  Write-Host "== $label" -ForegroundColor Cyan
  & $block
  if ($LASTEXITCODE -ne 0) {
    Write-Host "失敗しました：$label" -ForegroundColor Red
    exit 1
  }
}

# 1. GitHub（push すると GitHub Actions が web/ を Pages に公開する）
git add -A
git diff --cached --quiet
if ($LASTEXITCODE -ne 0) {
  Invoke-Step 'コミット' { git commit -m $Message }
}
Invoke-Step 'GitHub へ送信' { git push }

# 2. Apps Script（gas/ が前回から変わったときだけ）
#    デプロイのたびに「バージョン」が1つ増え、上限（200）があるので、変更がなければ何もしない
$hash = (Get-FileHash -Path gas\Code.gs, gas\appsscript.json | ForEach-Object { $_.Hash }) -join ''
$last = if (Test-Path $GasHashFile) { (Get-Content $GasHashFile -Raw).Trim() } else { '' }
if ($hash -eq $last) {
  Write-Host '== Apps Script は変更なし（スキップ）'
} else {
  Invoke-Step 'Apps Script へアップロード' { clasp push --force }
  Invoke-Step 'Apps Script を新バージョンでデプロイ（URLは変わりません）' { clasp update-deployment $DeploymentId -d $Message }
  Set-Content -Path $GasHashFile -Value $hash
}

Write-Host '完了しました。画面への反映は1〜2分後です。' -ForegroundColor Green
