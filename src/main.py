"""マネーフォワード → スプレッドシート連携ツール

マネーフォワード ME からエクスポートしたCSVファイルを読み込み、
収支データを整理してGoogle SheetsまたはExcelに書き出します。

使い方:
    # Excel出力（Google設定不要ですぐ使える）
    python src/main.py sample_data/sample.csv --format excel

    # Google Sheets出力
    python src/main.py data.csv --format google_sheets --config config/settings.yaml

    # 複数CSVファイルを指定
    python src/main.py data1.csv data2.csv --format excel

    # 振替を含めて集計
    python src/main.py data.csv --format excel --include-transfers

    # 期間を指定して集計
    python src/main.py data.csv --format excel --start 2025-01 --end 2025-12
"""

import argparse
import os
import sys
from pathlib import Path

import yaml

from moneyforward.csv_parser import parse_csv, parse_multiple_csv
from processor.data_processor import process_data
from spreadsheet.excel_writer import write_to_excel


def main():
    args = parse_args()

    # 設定ファイルの読み込み
    config = load_config(args.config) if args.config else {}

    # CSVファイルの読み込み
    print(f"CSVファイルを読み込み中...")
    if len(args.csv_files) == 1:
        df = parse_csv(args.csv_files[0])
    else:
        df = parse_multiple_csv(args.csv_files)
    print(f"  {len(df)} 件の取引を読み込みました")

    # データ処理
    print("データを集計中...")
    exclude_transfers = not args.include_transfers
    result = process_data(
        df,
        exclude_transfers=exclude_transfers,
        start_month=args.start or config.get("processing", {}).get("start_month", ""),
        end_month=args.end or config.get("processing", {}).get("end_month", ""),
    )

    stats = result["stats"]
    print(f"  期間: {stats['num_months']}ヶ月")
    print(f"  総収入: ¥{stats['total_income']:,}")
    print(f"  総支出: ¥{stats['total_expense']:,}")
    print(f"  純収支: ¥{stats['net']:,}")

    # 出力
    output_format = args.format or config.get("output", {}).get("format", "excel")

    if output_format == "google_sheets":
        _output_google_sheets(result, config, args)
    else:
        _output_excel(result, config, args)


def _output_google_sheets(result: dict, config: dict, args) -> None:
    """Google Sheetsに出力する。"""
    from spreadsheet.sheets_writer import write_to_sheets

    gs_config = config.get("google_sheets", {})

    # 認証ファイルのパス: 環境変数 > 設定ファイル > デフォルト
    credentials_path = (
        os.environ.get("GOOGLE_CREDENTIALS_PATH", "")
        or gs_config.get("credentials_path", "config/credentials.json")
    )

    spreadsheet_id = (
        os.environ.get("SPREADSHEET_ID", "")
        or gs_config.get("spreadsheet_id", "")
    )

    if not spreadsheet_id:
        print("エラー: spreadsheet_id が設定されていません", file=sys.stderr)
        print("config/settings.yaml または環境変数 SPREADSHEET_ID を設定してください", file=sys.stderr)
        sys.exit(1)

    if not Path(credentials_path).exists():
        print(f"エラー: 認証ファイルが見つかりません: {credentials_path}", file=sys.stderr)
        sys.exit(1)

    print("Google Sheetsに書き出し中...")
    url = write_to_sheets(result, credentials_path, spreadsheet_id)
    print(f"完了! スプレッドシート: {url}")


def _output_excel(result: dict, config: dict, args) -> None:
    """Excelファイルに出力する。"""
    output_path = args.output or config.get("output", {}).get(
        "excel_path", "output/収支レポート.xlsx"
    )
    print(f"Excelファイルに書き出し中...")
    path = write_to_excel(result, output_path)
    print(f"完了! 出力先: {path}")


def load_config(config_path: str) -> dict:
    """設定ファイルを読み込む。"""
    path = Path(config_path)
    if not path.exists():
        print(f"警告: 設定ファイルが見つかりません: {config_path}", file=sys.stderr)
        return {}
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def parse_args():
    parser = argparse.ArgumentParser(
        description="マネーフォワード収支データをスプレッドシートに書き出すツール",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  python src/main.py export.csv --format excel
  python src/main.py export.csv --format excel --output 家計簿.xlsx
  python src/main.py export.csv --format google_sheets --config config/settings.yaml
  python src/main.py jan.csv feb.csv mar.csv --format excel --start 2025-01 --end 2025-03
        """,
    )
    parser.add_argument(
        "csv_files",
        nargs="+",
        help="マネーフォワードからエクスポートしたCSVファイル",
    )
    parser.add_argument(
        "--format", "-f",
        choices=["excel", "google_sheets"],
        default="excel",
        help="出力形式 (デフォルト: excel)",
    )
    parser.add_argument(
        "--config", "-c",
        help="設定ファイルのパス (Google Sheets出力時に必要)",
    )
    parser.add_argument(
        "--output", "-o",
        help="出力ファイルのパス (Excel出力時)",
    )
    parser.add_argument(
        "--include-transfers",
        action="store_true",
        help="振替取引を集計に含める",
    )
    parser.add_argument(
        "--start",
        help="集計開始月 (YYYY-MM形式)",
    )
    parser.add_argument(
        "--end",
        help="集計終了月 (YYYY-MM形式)",
    )
    return parser.parse_args()


if __name__ == "__main__":
    main()
