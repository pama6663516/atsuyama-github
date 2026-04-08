"""マネーフォワード ME のCSVエクスポートファイルを解析するモジュール"""

import pandas as pd
from pathlib import Path


# マネーフォワードCSVの標準カラム名
MF_COLUMNS = {
    "計算対象": "included",
    "日付": "date",
    "内容": "description",
    "金額（円）": "amount",
    "保有金融機関": "institution",
    "大項目": "major_category",
    "中項目": "sub_category",
    "メモ": "memo",
    "振替": "is_transfer",
    "ID": "id",
}


def parse_csv(file_path: str | Path) -> pd.DataFrame:
    """マネーフォワードのCSVファイルを読み込み、DataFrameに変換する。

    Args:
        file_path: CSVファイルのパス

    Returns:
        整形済みのDataFrame
    """
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"CSVファイルが見つかりません: {file_path}")

    # マネーフォワードCSVはShift_JISエンコーディング
    df = _try_read_csv(file_path)

    # カラム名を正規化
    df = _normalize_columns(df)

    # データ型を変換
    df = _convert_types(df)

    return df


def parse_multiple_csv(file_paths: list[str | Path]) -> pd.DataFrame:
    """複数のCSVファイルを読み込み、結合する。

    Args:
        file_paths: CSVファイルのパスのリスト

    Returns:
        結合・重複除去済みのDataFrame
    """
    dfs = [parse_csv(p) for p in file_paths]
    combined = pd.concat(dfs, ignore_index=True)

    # ID列がある場合、重複を除去
    if "id" in combined.columns:
        combined = combined.drop_duplicates(subset=["id"], keep="last")

    combined = combined.sort_values("date", ascending=True).reset_index(drop=True)
    return combined


def _try_read_csv(file_path: Path) -> pd.DataFrame:
    """複数のエンコーディングを試してCSVを読み込む。"""
    encodings = ["shift_jis", "cp932", "utf-8", "utf-8-sig"]
    last_error = None

    for encoding in encodings:
        try:
            return pd.read_csv(file_path, encoding=encoding)
        except (UnicodeDecodeError, UnicodeError) as e:
            last_error = e
            continue

    raise ValueError(
        f"CSVファイルの読み込みに失敗しました（対応エンコーディング: {encodings}）: {last_error}"
    )


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """カラム名を英語の内部名に変換する。"""
    # 前後の空白を除去
    df.columns = df.columns.str.strip()

    rename_map = {}
    for jp_name, en_name in MF_COLUMNS.items():
        if jp_name in df.columns:
            rename_map[jp_name] = en_name

    if rename_map:
        df = df.rename(columns=rename_map)
    return df


def _convert_types(df: pd.DataFrame) -> pd.DataFrame:
    """各カラムのデータ型を適切に変換する。"""
    # 日付
    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        df["year_month"] = df["date"].dt.to_period("M")

    # 金額: カンマ区切りの文字列を数値に
    if "amount" in df.columns:
        if df["amount"].dtype == object:
            df["amount"] = (
                df["amount"]
                .astype(str)
                .str.replace(",", "", regex=False)
                .str.replace("¥", "", regex=False)
                .str.replace("円", "", regex=False)
                .str.strip()
            )
        df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0).astype(int)

    # 振替フラグ
    if "is_transfer" in df.columns:
        df["is_transfer"] = df["is_transfer"].astype(str).str.strip().isin(["1", "振替"])

    # 計算対象フラグ
    if "included" in df.columns:
        df["included"] = df["included"].astype(str).str.strip().isin(["1", "○"])

    # 収入/支出の判定
    if "amount" in df.columns:
        df["type"] = df["amount"].apply(lambda x: "income" if x > 0 else "expense")

    return df
