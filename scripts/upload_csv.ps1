# マネーフォワード収支CSV → Google Sheets 自動アップロードスクリプト (Windows版)
#
# 使い方:
#   .\scripts\upload_csv.ps1 ~\Downloads\収入・支出詳細_*.csv
#   .\scripts\upload_csv.ps1 file1.csv file2.csv
#
# CSVファイルを data/ にコピーして git push するだけで、
# GitHub Actions が自動で集計 → Google Sheets に書き出します。

param(
    [Parameter(Mandatory=$true, Position=0, ValueFromRemainingArguments=$true)]
    [string[]]$CsvFiles
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoDir = Split-Path -Parent $ScriptDir
$DataDir = Join-Path $RepoDir "data"

Write-Host "=== マネーフォワード → スプレッドシート自動連携 ===" -ForegroundColor Cyan
Write-Host ""

# ワイルドカード展開
$ExpandedFiles = @()
foreach ($pattern in $CsvFiles) {
    $resolved = Resolve-Path -Path $pattern -ErrorAction SilentlyContinue
    if ($resolved) {
        $ExpandedFiles += $resolved
    } else {
        Write-Host "  ⚠ ファイルが見つかりません: $pattern" -ForegroundColor Yellow
    }
}

if ($ExpandedFiles.Count -eq 0) {
    Write-Host "エラー: CSVファイルが見つかりませんでした" -ForegroundColor Red
    Write-Host ""
    Write-Host "使い方:" -ForegroundColor Yellow
    Write-Host '  .\scripts\upload_csv.ps1 ~\Downloads\収入・支出詳細_*.csv'
    Write-Host '  .\scripts\upload_csv.ps1 C:\Users\pama6\Downloads\*.csv'
    exit 1
}

# リポジトリの最新状態を取得
Write-Host "[1/4] リポジトリを最新に更新中..."
Push-Location $RepoDir
try {
    git pull origin main --quiet 2>$null
} catch {
    Write-Host "  (リモートからの取得をスキップ)" -ForegroundColor Yellow
}

# data/ フォルダがなければ作成
if (-not (Test-Path $DataDir)) {
    New-Item -ItemType Directory -Path $DataDir | Out-Null
}

# CSVファイルをdata/にコピー
Write-Host "[2/4] CSVファイルをdata/にコピー中..."
$Copied = 0
foreach ($file in $ExpandedFiles) {
    $fileName = Split-Path -Leaf $file
    Copy-Item -Path $file -Destination (Join-Path $DataDir $fileName) -Force
    Write-Host "  ✅ $fileName" -ForegroundColor Green
    $Copied++
}

# コミット＆プッシュ
Write-Host "[3/4] GitHubにpush中..."
git add data/*.csv
$Month = Get-Date -Format "yyyy年MM月"
git commit -m "${Month}の収支データを追加" --quiet
git push origin main --quiet

Write-Host "[4/4] 完了!" -ForegroundColor Green
Write-Host ""
Write-Host "GitHub Actions が自動でレポートを生成します。"
Write-Host "進捗はこちらで確認: https://github.com/pama6663516/atsuyama-github/actions" -ForegroundColor Cyan
Write-Host ""
Write-Host "数十秒後にGoogle Sheetsに反映されます。"

Pop-Location
