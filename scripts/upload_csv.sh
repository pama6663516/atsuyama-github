#!/bin/bash
# マネーフォワード収支CSV → Google Sheets 自動アップロードスクリプト
#
# 使い方:
#   ./scripts/upload_csv.sh ~/Downloads/収入・支出詳細_*.csv
#   ./scripts/upload_csv.sh file1.csv file2.csv
#
# CSVファイルを data/ にコピーして git push するだけで、
# GitHub Actions が自動で集計 → Google Sheets に書き出します。

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_DIR="$(dirname "$SCRIPT_DIR")"
DATA_DIR="$REPO_DIR/data"

# 引数チェック
if [ $# -eq 0 ]; then
    echo "使い方: $0 <CSVファイル> [CSVファイル2] ..."
    echo ""
    echo "例:"
    echo "  $0 ~/Downloads/収入・支出詳細_2026-03-01_2026-03-31.csv"
    echo "  $0 ~/Downloads/*.csv"
    exit 1
fi

# リポジトリの最新状態を取得
echo "=== マネーフォワード → スプレッドシート自動連携 ==="
echo ""
echo "[1/4] リポジトリを最新に更新中..."
cd "$REPO_DIR"
git pull origin main --quiet

# CSVファイルをdata/にコピー
echo "[2/4] CSVファイルをdata/にコピー中..."
COPIED=0
for CSV_FILE in "$@"; do
    if [ ! -f "$CSV_FILE" ]; then
        echo "  ⚠ ファイルが見つかりません: $CSV_FILE"
        continue
    fi
    FILENAME="$(basename "$CSV_FILE")"
    cp "$CSV_FILE" "$DATA_DIR/$FILENAME"
    echo "  ✅ $FILENAME"
    COPIED=$((COPIED + 1))
done

if [ "$COPIED" -eq 0 ]; then
    echo "エラー: コピーできるCSVファイルがありませんでした"
    exit 1
fi

# コミット＆プッシュ
echo "[3/4] GitHubにpush中..."
cd "$REPO_DIR"
git add data/*.csv
MONTH=$(date +"%Y年%m月")
git commit -m "${MONTH}の収支データを追加" --quiet
git push origin main --quiet

echo "[4/4] 完了!"
echo ""
echo "GitHub Actions が自動でレポートを生成します。"
echo "進捗はこちらで確認: https://github.com/pama6663516/atsuyama-github/actions"
echo ""
echo "数十秒後にGoogle Sheetsに反映されます。"
