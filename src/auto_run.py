"""毎月末の自動実行スクリプト

マネーフォワードからCSVを自動取得 → 集計 → スプレッドシートに書き出し
を一括で行います。

環境変数:
    MF_EMAIL: マネーフォワードのログインメールアドレス
    MF_PASSWORD: マネーフォワードのログインパスワード
    GOOGLE_CREDENTIALS_JSON: サービスアカウントJSONの内容（GitHub Actions用）
    SPREADSHEET_ID: 出力先スプレッドシートのID

使い方:
    # 当月のデータを自動取得・集計・出力
    python src/auto_run.py

    # 特定の月を指定
    python src/auto_run.py --year 2025 --month 3

    # Excel出力（Google Sheets不要）
    python src/auto_run.py --format excel
"""

import argparse
import json
import os
import sys
import tempfile
from datetime import datetime
from pathlib import Path

from moneyforward.scraper import download_csv
from moneyforward.csv_parser import parse_csv
from processor.data_processor import process_data
from spreadsheet.excel_writer import write_to_excel


def main():
    args = parse_args()

    now = datetime.now()
    year = args.year or now.year
    month = args.month or now.month

    print(f"=== マネーフォワード 月次自動レポート ({year}/{month:02d}) ===")
    print()

    # Step 1: マネーフォワードからCSVダウンロード
    print("[1/3] マネーフォワードからデータを取得中...")
    csv_path = download_csv(
        year=year,
        month=month,
        download_dir="downloads",
        headless=True,
    )
    print(f"  CSV取得完了: {csv_path}")
    print()

    # Step 2: データ集計
    print("[2/3] データを集計中...")
    df = parse_csv(csv_path)
    print(f"  {len(df)} 件の取引を読み込みました")

    result = process_data(df, exclude_transfers=True)
    stats = result["stats"]
    print(f"  収入: ¥{stats['total_income']:,}")
    print(f"  支出: ¥{stats['total_expense']:,}")
    print(f"  収支: ¥{stats['net']:,}")
    print()

    # Step 3: スプレッドシートに出力
    output_format = args.format or os.environ.get("OUTPUT_FORMAT", "google_sheets")

    if output_format == "google_sheets":
        _output_to_sheets(result, year, month)
    else:
        _output_to_excel(result, year, month, args.output)

    print()
    print("=== 完了 ===")


def _output_to_sheets(result: dict, year: int, month: int) -> None:
    """Google Sheetsに出力する。"""
    from spreadsheet.sheets_writer import write_to_sheets

    spreadsheet_id = os.environ.get("SPREADSHEET_ID", "")
    if not spreadsheet_id:
        print("エラー: SPREADSHEET_ID 環境変数が設定されていません", file=sys.stderr)
        sys.exit(1)

    # GitHub Actions 用: 環境変数からJSON認証情報を取得
    credentials_json = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")
    if credentials_json:
        # 一時ファイルに認証情報を書き出し
        creds_path = Path(tempfile.mktemp(suffix=".json"))
        creds_path.write_text(credentials_json)
        credentials_path = str(creds_path)
    else:
        credentials_path = os.environ.get(
            "GOOGLE_CREDENTIALS_PATH", "config/credentials.json"
        )

    if not Path(credentials_path).exists():
        print(f"エラー: 認証ファイルが見つかりません: {credentials_path}", file=sys.stderr)
        sys.exit(1)

    print(f"[3/3] Google Sheetsに書き出し中...")
    url = write_to_sheets(result, credentials_path, spreadsheet_id)
    print(f"  完了! {url}")

    # 一時ファイルのクリーンアップ
    if credentials_json:
        creds_path.unlink(missing_ok=True)


def _output_to_excel(result: dict, year: int, month: int, output: str | None) -> None:
    """Excelファイルに出力する。"""
    output_path = output or f"output/収支レポート_{year}{month:02d}.xlsx"
    print(f"[3/3] Excelに書き出し中...")
    path = write_to_excel(result, output_path)
    print(f"  完了! {path}")


def parse_args():
    parser = argparse.ArgumentParser(
        description="マネーフォワード月次自動レポート生成",
    )
    parser.add_argument("--year", type=int, help="対象年 (デフォルト: 当年)")
    parser.add_argument("--month", type=int, help="対象月 (デフォルト: 当月)")
    parser.add_argument(
        "--format", choices=["excel", "google_sheets"],
        help="出力形式 (デフォルト: google_sheets)",
    )
    parser.add_argument("--output", help="Excel出力時のファイルパス")
    return parser.parse_args()


if __name__ == "__main__":
    main()
