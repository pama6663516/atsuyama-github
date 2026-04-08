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
- **自動実行**: CSVをpushするだけでGoogle Sheetsに自動反映（GitHub Actions）

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
python src/main.py data/sample.csv --format excel
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

## 自動実行（GitHub Actions）

CSVファイルをリポジトリの `data/` フォルダにpushするだけで、自動的にGoogle Sheetsに収支データが書き出されます。

### 毎月の流れ

```
1. マネーフォワードからCSVをダウンロード（30秒）
2. data/ フォルダに入れてgit push
3. → GitHub Actionsが自動で集計 → Google Sheetsに反映!
```

### セットアップ手順

#### 1. Google Cloud の設定

1. [Google Cloud Console](https://console.cloud.google.com/) → 新規プロジェクト作成
2. 「APIとサービス」→「ライブラリ」→ **Google Sheets API** と **Google Drive API** を有効化
3. 「認証情報」→「認証情報を作成」→「サービスアカウント」→ 作成
4. 作成したサービスアカウント →「キー」タブ →「鍵を追加」→「JSON」→ ダウンロード
5. 出力先のGoogle Sheetsを新規作成
6. スプレッドシートの「共有」で、サービスアカウントのメールアドレス（JSON内の `client_email`）を**編集者として追加**

#### 2. GitHub Secrets の登録

リポジトリの **Settings** → **Secrets and variables** → **Actions** で登録:

| Secret 名 | 内容 |
|---|---|
| `GOOGLE_CREDENTIALS_JSON` | ダウンロードしたJSONファイルの中身を全文貼り付け |
| `SPREADSHEET_ID` | スプレッドシートURLの `/d/` と `/edit` の間の文字列 |

#### 3. CSVをpushして自動実行

```bash
# マネーフォワードからダウンロードしたCSVを data/ に配置
cp ~/Downloads/収入・支出詳細_2025.csv data/

# pushするだけで自動実行!
git add data/
git commit -m "2025年4月の収支データ"
git push
```

Actions タブで実行状況を確認できます。手動実行も「Run workflow」から可能です。

## 出力シート構成

| シート名 | 内容 |
|---|---|
| ダッシュボード | 総収入・総支出・純収支・月平均などの概要 |
| 月次サマリー | 月ごとの収入/支出/収支/累計収支 |
| カテゴリ別集計 | 支出の大項目・中項目別の合計金額・件数・構成比 |
| 月別カテゴリ | 月×大項目カテゴリのクロス集計 |
| 取引一覧 | 全取引の詳細データ |

## プロジェクト構成

```
├── .github/workflows/
│   └── monthly_report.yml     # 自動実行ワークフロー（CSV push時 + 毎月末）
├── data/                      # ← ここにCSVを入れてpush
├── src/
│   ├── main.py                # メインスクリプト（CLI）
│   ├── moneyforward/
│   │   ├── csv_parser.py      # CSVファイル解析
│   │   └── scraper.py         # ローカル用: MF自動ログイン＆CSVダウンロード
│   ├── processor/
│   │   └── data_processor.py  # データ集計・分析
│   └── spreadsheet/
│       ├── sheets_writer.py   # Google Sheets出力
│       └── excel_writer.py    # Excel出力
├── config/
│   └── settings.yaml.example  # 設定ファイルのテンプレート
├── sample_data/
│   └── sample.csv             # サンプルデータ
├── requirements.txt
└── README.md
```
