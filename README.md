# マネーフォワード → スプレッドシート連携ツール

マネーフォワード ME からエクスポートしたCSVファイルを読み込み、収支データを整理してスプレッドシートに書き出すツールです。

## 機能

- マネーフォワードCSVの自動解析（Shift_JIS / UTF-8 対応）
- 複数CSVファイルの結合・重複除去
- **ダッシュボード**: 収支概要、月平均、支出最多カテゴリ
- **月次サマリー**: 月ごとの収入・支出・収支・累計収支
- **カテゴリ別集計**: 支出の大項目・中項目別の合計と構成比
- **月別カテゴリ**: 月×カテゴリのクロス集計表（グラフ作成用）
- **取引一覧**: 全取引の詳細リスト
- Google Sheets / Excel 両方に対応
- **毎月末自動実行**: GitHub Actions で自動取得・集計・出力
- マネーフォワードからのCSV自動ダウンロード（Playwright使用）

## セットアップ

```bash
pip install -r requirements.txt
```

## 使い方

### 1. マネーフォワードからCSVをエクスポート

1. [マネーフォワード ME](https://moneyforward.com/) にログイン
2. 「家計簿」→「家計簿（月ごと）」を開く
3. 「CSVダウンロード」ボタンからエクスポート

### 2. Excel出力（すぐ使える）

```bash
python src/main.py sample_data/sample.csv --format excel
# → output/収支レポート.xlsx に出力されます

# 出力先を指定
python src/main.py data.csv --format excel --output 家計簿2025.xlsx

# 複数ファイルを結合
python src/main.py 1月.csv 2月.csv 3月.csv --format excel

# 期間を指定
python src/main.py data.csv --format excel --start 2025-01 --end 2025-06

# 振替取引を含める
python src/main.py data.csv --format excel --include-transfers
```

### 3. Google Sheets出力

#### 事前準備

1. [Google Cloud Console](https://console.cloud.google.com/) でプロジェクトを作成
2. Google Sheets API と Google Drive API を有効化
3. サービスアカウントを作成し、JSONキーをダウンロード
4. JSONキーを `config/credentials.json` に配置
5. `config/settings.yaml.example` を `config/settings.yaml` にコピーして編集
6. 出力先スプレッドシートをサービスアカウントのメールアドレスと共有

```bash
# 設定ファイルを準備
cp config/settings.yaml.example config/settings.yaml
# settings.yaml を編集して spreadsheet_id を設定

# 実行
python src/main.py data.csv --format google_sheets --config config/settings.yaml
```

## 出力シート構成

| シート名 | 内容 |
|---|---|
| ダッシュボード | 総収入・総支出・純収支・月平均などの概要 |
| 月次サマリー | 月ごとの収入/支出/収支/累計収支 |
| カテゴリ別集計 | 支出の大項目・中項目別の合計金額・件数・構成比 |
| 月別カテゴリ | 月×大項目カテゴリのクロス集計 |
| 取引一覧 | 全取引の詳細データ |

## 毎月末の自動実行（GitHub Actions）

毎月末日の23:00 (JST) に自動で「データ取得 → 集計 → スプレッドシート出力」を行います。

### セットアップ手順

1. **GitHub リポジトリの Settings → Secrets and variables → Actions** で以下を登録:

| Secret 名 | 内容 |
|---|---|
| `MF_EMAIL` | マネーフォワードのログインメールアドレス |
| `MF_PASSWORD` | マネーフォワードのログインパスワード |
| `GOOGLE_CREDENTIALS_JSON` | サービスアカウントJSONファイルの中身をそのまま貼り付け |
| `SPREADSHEET_ID` | 出力先Google SheetsのスプレッドシートID |

2. **出力先スプレッドシートの共有設定**で、サービスアカウントのメールアドレス（`xxx@xxx.iam.gserviceaccount.com`）に編集権限を付与

3. 設定完了! 毎月末に自動実行されます

### 手動実行

GitHub の Actions タブから「月次収支レポート自動生成」ワークフローを選択し、「Run workflow」で手動実行も可能です。年月を指定して過去のデータも取得できます。

### ローカルでの自動実行

```bash
# Playwright のインストール
pip install playwright
playwright install chromium

# 環境変数を設定
export MF_EMAIL="your@email.com"
export MF_PASSWORD="yourpassword"
export SPREADSHEET_ID="your-spreadsheet-id"
export GOOGLE_CREDENTIALS_PATH="config/credentials.json"

# 当月のデータを自動取得・集計・出力
python src/auto_run.py

# 特定の月を指定
python src/auto_run.py --year 2025 --month 3

# Excel出力のみ（Google Sheets設定不要）
python src/auto_run.py --format excel
```

## プロジェクト構成

```
├── config/
│   └── settings.yaml.example  # 設定ファイルのテンプレート
├── sample_data/
│   └── sample.csv             # サンプルデータ
├── .github/workflows/
│   └── monthly_report.yml     # 毎月末自動実行ワークフロー
├── src/
│   ├── main.py                # メインスクリプト（CLI・手動用）
│   ├── auto_run.py            # 自動実行スクリプト（取得→集計→出力）
│   ├── moneyforward/
│   │   ├── csv_parser.py      # CSVファイル解析
│   │   └── scraper.py         # MF自動ログイン＆CSVダウンロード
│   ├── processor/
│   │   └── data_processor.py  # データ集計・分析
│   └── spreadsheet/
│       ├── sheets_writer.py   # Google Sheets出力
│       └── excel_writer.py    # Excel出力
├── requirements.txt
└── README.md
```
