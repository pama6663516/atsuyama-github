# マネーフォワード → スプレッドシート自動連携

## プロジェクト概要

マネーフォワード ME の収支CSVデータを Google Sheets に自動で整理・書き出すツール。

## アーキテクチャ

```
[ユーザー] CSVダウンロード → data/ にpush → [GitHub Actions] 集計 → [Google Sheets] 出力
```

- `src/main.py`: CLI メインスクリプト（CSV読み込み→集計→出力）
- `src/moneyforward/csv_parser.py`: マネーフォワードCSV解析（Shift_JIS/UTF-8自動判定）
- `src/processor/data_processor.py`: データ集計（月次サマリー、カテゴリ別、クロス集計）
- `src/spreadsheet/sheets_writer.py`: Google Sheets書き出し（gspread）
- `src/spreadsheet/excel_writer.py`: Excel書き出し（openpyxl）
- `.github/workflows/monthly_report.yml`: 自動実行ワークフロー

## よく使うコマンド

### CSVをアップロードして自動実行

```bash
./scripts/upload_csv.sh ~/Downloads/収入・支出詳細_*.csv
```

### ローカルでExcel出力

```bash
python src/main.py data/*.csv --format excel
```

### ローカルでGoogle Sheets出力

```bash
python src/main.py data/*.csv --format google_sheets --config config/settings.yaml
```

## 出力シート構成

| シート | 内容 |
|---|---|
| ダッシュボード | 総収入・支出・月平均・支出最多カテゴリ |
| 月次サマリー | 月ごとの収入/支出/収支/累計収支 |
| カテゴリ別集計 | 大項目・中項目別の合計金額・構成比 |
| 月別カテゴリ | 月×カテゴリのクロス集計表 |
| 取引一覧 | 全取引の詳細データ |

## GitHub Secrets

| Secret | 用途 |
|---|---|
| `GOOGLE_CREDENTIALS_JSON` | サービスアカウントJSONの全文 |
| `SPREADSHEET_ID` | 出力先スプレッドシートのID |

## 注意事項

- `config/credentials.json` と `config/settings.yaml` は .gitignore で除外済み（認証情報保護）
- CSVは必ず `data/` フォルダに入れること（ワークフローのトリガー条件）
- マネーフォワードCSVのカラム: 計算対象, 日付, 内容, 金額（円）, 保有金融機関, 大項目, 中項目, メモ, 振替, ID
