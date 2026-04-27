# マネーフォワード収支CSV → Google Sheets 自動アップロードスクリプト (Windows版)
#
# 使い方:
#   .\scripts\upload_csv.ps1 ~\Downloads\収入・支出詳細_*.csv
#   .\scripts\upload_csv.ps1 file1.csv file2.csv
#
# CSVファイルを data/ にコピーして git push するだけで、
# GitHub Actions が自動で集計 → Google Sheets に書き出します。
# 完了後、Excelレポートを指定フォルダに自動コピーします。

param(
    [Parameter(Mandatory=$true, Position=0, ValueFromRemainingArguments=$true)]
    [string[]]$CsvFiles
)

$ErrorActionPreference = "Stop"

# レポート保存先フォルダ
$OUTPUT_DIR = "C:\Users\pama6\AI_AGENT\agents\CFO_Amon"

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
Write-Host "[1/6] リポジトリを最新に更新中..."
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
Write-Host "[2/6] CSVファイルをdata/にコピー中..."
$Copied = 0
foreach ($file in $ExpandedFiles) {
    $fileName = Split-Path -Leaf $file
    Copy-Item -Path $file -Destination (Join-Path $DataDir $fileName) -Force
    Write-Host "  ✅ $fileName" -ForegroundColor Green
    $Copied++
}

# コミット＆プッシュ
Write-Host "[3/6] GitHubにpush中..."
git add data/*.csv
$Month = Get-Date -Format "yyyy年MM月"
git commit -m "${Month}の収支データを追加" --quiet
git push origin main --quiet

Write-Host "[4/6] GitHub Actionsの完了を待っています..."
Write-Host "  (Google Sheets + Excelレポート生成中... 約1-2分)" -ForegroundColor Gray

# ワークフロー完了を待つ（最大3分）
$MaxWait = 180
$Elapsed = 0
$Interval = 15
Start-Sleep -Seconds 10
$Elapsed += 10

while ($Elapsed -lt $MaxWait) {
    try {
        git fetch origin main --quiet 2>$null
        $RemoteLog = git log origin/main --oneline -1
        if ($RemoteLog -match "\[skip ci\]") {
            Write-Host "  ✅ ワークフロー完了!" -ForegroundColor Green
            break
        }
    } catch {}
    Write-Host "  ... 待機中 (${Elapsed}秒)" -ForegroundColor Gray
    Start-Sleep -Seconds $Interval
    $Elapsed += $Interval
}

# 最新のレポートを取得
Write-Host "[5/6] レポートを取得中..."
git pull origin main --quiet 2>$null

# 出力フォルダにコピー
Write-Host "[6/6] レポートを保存中..."
$OutputRepoDir = Join-Path $RepoDir "output"

if (-not (Test-Path $OUTPUT_DIR)) {
    New-Item -ItemType Directory -Path $OUTPUT_DIR -Force | Out-Null
}

if (Test-Path $OutputRepoDir) {
    $XlsxFiles = Get-ChildItem -Path $OutputRepoDir -Filter "*.xlsx" -ErrorAction SilentlyContinue
    if ($XlsxFiles) {
        foreach ($xlsx in $XlsxFiles) {
            Copy-Item -Path $xlsx.FullName -Destination $OUTPUT_DIR -Force
            Write-Host "  ✅ $($xlsx.Name) → $OUTPUT_DIR" -ForegroundColor Green
        }
    } else {
        Write-Host "  ⚠ Excelファイルが見つかりませんでした" -ForegroundColor Yellow
    }
} else {
    Write-Host "  ⚠ output/フォルダがまだ作成されていません" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=== 完了! ===" -ForegroundColor Cyan
Write-Host "  Google Sheets: 自動更新済み"
Write-Host "  Excelレポート: $OUTPUT_DIR"
Write-Host "  Actions確認: https://github.com/pama6663516/atsuyama-github/actions" -ForegroundColor Gray

Pop-Location
