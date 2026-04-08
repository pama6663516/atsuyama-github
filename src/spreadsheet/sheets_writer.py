"""Google Sheets への書き出しモジュール"""

import gspread
from google.oauth2.service_account import Credentials
import pandas as pd


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# シート名の定義
SHEET_DASHBOARD = "ダッシュボード"
SHEET_MONTHLY = "月次サマリー"
SHEET_CATEGORY = "カテゴリ別集計"
SHEET_MATRIX = "月別カテゴリ"
SHEET_TRANSACTIONS = "取引一覧"


def write_to_sheets(
    data: dict,
    credentials_path: str,
    spreadsheet_id: str,
) -> str:
    """集計データをGoogle Sheetsに書き出す。

    Args:
        data: data_processor.process_data() の出力
        credentials_path: サービスアカウントJSON
        spreadsheet_id: スプレッドシートのID

    Returns:
        スプレッドシートのURL
    """
    gc = _authorize(credentials_path)
    spreadsheet = gc.open_by_key(spreadsheet_id)

    # 各シートにデータを書き出し
    _write_dashboard(spreadsheet, data["stats"], data["monthly_summary"])
    _write_dataframe_sheet(spreadsheet, SHEET_MONTHLY, data["monthly_summary"])
    _write_dataframe_sheet(spreadsheet, SHEET_CATEGORY, data["category_breakdown"])
    if not data["monthly_category"].empty:
        _write_dataframe_sheet(spreadsheet, SHEET_MATRIX, data["monthly_category"])
    _write_dataframe_sheet(spreadsheet, SHEET_TRANSACTIONS, data["transactions"])

    # 書式を適用
    _apply_formatting(spreadsheet)

    return spreadsheet.url


def _authorize(credentials_path: str) -> gspread.Client:
    """Google APIの認証を行う。"""
    creds = Credentials.from_service_account_file(credentials_path, scopes=SCOPES)
    return gspread.authorize(creds)


def _get_or_create_sheet(
    spreadsheet: gspread.Spreadsheet, title: str
) -> gspread.Worksheet:
    """シートを取得、なければ作成する。"""
    try:
        worksheet = spreadsheet.worksheet(title)
        worksheet.clear()
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=title, rows=1000, cols=26)
    return worksheet


def _write_dashboard(
    spreadsheet: gspread.Spreadsheet,
    stats: dict,
    monthly_summary: pd.DataFrame,
) -> None:
    """ダッシュボードシートを作成する。"""
    ws = _get_or_create_sheet(spreadsheet, SHEET_DASHBOARD)

    rows = [
        ["マネーフォワード 収支レポート"],
        [],
        ["--- 全体サマリー ---"],
        ["集計期間（月数）", stats["num_months"]],
        ["取引件数", stats["num_transactions"]],
        [],
        ["--- 収支概要 ---"],
        ["総収入", f"¥{stats['total_income']:,}"],
        ["総支出", f"¥{stats['total_expense']:,}"],
        ["純収支", f"¥{stats['net']:,}"],
        [],
        ["--- 月平均 ---"],
        ["月平均収入", f"¥{stats['avg_monthly_income']:,}"],
        ["月平均支出", f"¥{stats['avg_monthly_expense']:,}"],
        [],
        ["--- その他 ---"],
        ["支出最多カテゴリ", stats["top_expense_category"]],
    ]

    # 月次推移データも追加（グラフ作成用）
    if not monthly_summary.empty:
        rows.append([])
        rows.append(["--- 月次推移（グラフ用データ） ---"])
        header = monthly_summary.columns.tolist()
        rows.append(header)
        for _, row in monthly_summary.iterrows():
            rows.append(row.tolist())

    ws.update(range_name="A1", values=rows)


def _write_dataframe_sheet(
    spreadsheet: gspread.Spreadsheet,
    title: str,
    df: pd.DataFrame,
) -> None:
    """DataFrameをシートに書き出す。"""
    if df.empty:
        return

    ws = _get_or_create_sheet(spreadsheet, title)

    # ヘッダー + データ
    header = df.columns.tolist()
    values = df.fillna("").values.tolist()

    # 数値型以外を文字列に変換
    converted = []
    for row in values:
        converted_row = []
        for val in row:
            if isinstance(val, (int, float)):
                converted_row.append(val)
            else:
                converted_row.append(str(val))
        converted.append(converted_row)

    all_rows = [header] + converted
    ws.update(range_name="A1", values=all_rows)


def _apply_formatting(spreadsheet: gspread.Spreadsheet) -> None:
    """各シートに書式を適用する。"""
    # ダッシュボードのタイトル行を太字に
    try:
        ws = spreadsheet.worksheet(SHEET_DASHBOARD)
        ws.format("A1", {
            "textFormat": {"bold": True, "fontSize": 14},
        })
        ws.format("A3", {"textFormat": {"bold": True, "fontSize": 11}})
        ws.format("A7", {"textFormat": {"bold": True, "fontSize": 11}})
        ws.format("A12", {"textFormat": {"bold": True, "fontSize": 11}})
        ws.format("A16", {"textFormat": {"bold": True, "fontSize": 11}})
    except Exception:
        pass  # 書式設定の失敗は無視

    # 各データシートのヘッダーを太字に
    for sheet_name in [SHEET_MONTHLY, SHEET_CATEGORY, SHEET_MATRIX, SHEET_TRANSACTIONS]:
        try:
            ws = spreadsheet.worksheet(sheet_name)
            ws.format("1", {
                "textFormat": {"bold": True},
                "backgroundColor": {"red": 0.9, "green": 0.93, "blue": 0.98},
            })
            ws.freeze(rows=1)
        except Exception:
            pass
