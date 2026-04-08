"""収支データの集計・分析を行うモジュール"""

import pandas as pd


def process_data(
    df: pd.DataFrame,
    exclude_transfers: bool = True,
    start_month: str = "",
    end_month: str = "",
) -> dict:
    """取引データを集計し、各種サマリーを生成する。

    Args:
        df: CSVパーサーで読み込んだDataFrame
        exclude_transfers: 振替取引を除外するか
        start_month: 集計開始月 (YYYY-MM)
        end_month: 集計終了月 (YYYY-MM)

    Returns:
        各シートに書き出すデータを格納した辞書
    """
    # フィルタリング
    filtered = _filter_data(df, exclude_transfers, start_month, end_month)

    return {
        "monthly_summary": _monthly_summary(filtered),
        "category_breakdown": _category_breakdown(filtered),
        "monthly_category": _monthly_category_matrix(filtered),
        "transactions": _format_transactions(filtered),
        "stats": _calculate_stats(filtered),
    }


def _filter_data(
    df: pd.DataFrame,
    exclude_transfers: bool,
    start_month: str,
    end_month: str,
) -> pd.DataFrame:
    """条件に応じてデータをフィルタリングする。"""
    filtered = df.copy()

    if exclude_transfers and "is_transfer" in filtered.columns:
        filtered = filtered[~filtered["is_transfer"]]

    if start_month and "year_month" in filtered.columns:
        start = pd.Period(start_month, freq="M")
        filtered = filtered[filtered["year_month"] >= start]

    if end_month and "year_month" in filtered.columns:
        end = pd.Period(end_month, freq="M")
        filtered = filtered[filtered["year_month"] <= end]

    return filtered.reset_index(drop=True)


def _monthly_summary(df: pd.DataFrame) -> pd.DataFrame:
    """月次の収支サマリーを作成する。

    出力カラム: 年月, 収入, 支出, 収支, 累計収支
    """
    if df.empty or "year_month" not in df.columns:
        return pd.DataFrame(columns=["年月", "収入", "支出", "収支", "累計収支"])

    income = (
        df[df["amount"] > 0]
        .groupby("year_month")["amount"]
        .sum()
        .rename("収入")
    )
    expense = (
        df[df["amount"] < 0]
        .groupby("year_month")["amount"]
        .sum()
        .abs()
        .rename("支出")
    )

    # 全ての月を含むインデックスを作成
    all_months = df["year_month"].sort_values().unique()
    summary = pd.DataFrame(index=all_months)
    summary["収入"] = income.reindex(all_months, fill_value=0)
    summary["支出"] = expense.reindex(all_months, fill_value=0)
    summary["収支"] = summary["収入"] - summary["支出"]
    summary["累計収支"] = summary["収支"].cumsum()

    summary = summary.reset_index()
    summary = summary.rename(columns={"index": "年月", "year_month": "年月"})
    summary["年月"] = summary["年月"].astype(str)

    return summary


def _category_breakdown(df: pd.DataFrame) -> pd.DataFrame:
    """カテゴリ別の支出集計を作成する。

    出力カラム: 大項目, 中項目, 合計金額, 取引件数, 構成比(%)
    """
    expenses = df[df["amount"] < 0].copy()
    if expenses.empty:
        return pd.DataFrame(
            columns=["大項目", "中項目", "合計金額", "取引件数", "構成比(%)"]
        )

    expenses["abs_amount"] = expenses["amount"].abs()

    group_cols = []
    if "major_category" in expenses.columns:
        group_cols.append("major_category")
    if "sub_category" in expenses.columns:
        group_cols.append("sub_category")

    if not group_cols:
        return pd.DataFrame(
            columns=["大項目", "中項目", "合計金額", "取引件数", "構成比(%)"]
        )

    grouped = (
        expenses.groupby(group_cols)
        .agg(合計金額=("abs_amount", "sum"), 取引件数=("abs_amount", "count"))
        .reset_index()
        .sort_values("合計金額", ascending=False)
    )

    total = grouped["合計金額"].sum()
    grouped["構成比(%)"] = (grouped["合計金額"] / total * 100).round(1)

    rename_map = {}
    if "major_category" in grouped.columns:
        rename_map["major_category"] = "大項目"
    if "sub_category" in grouped.columns:
        rename_map["sub_category"] = "中項目"
    grouped = grouped.rename(columns=rename_map)

    return grouped.reset_index(drop=True)


def _monthly_category_matrix(df: pd.DataFrame) -> pd.DataFrame:
    """月×カテゴリのクロス集計表を作成する。

    行: 大項目, 列: 年月
    """
    expenses = df[df["amount"] < 0].copy()
    if expenses.empty or "major_category" not in expenses.columns:
        return pd.DataFrame()

    expenses["abs_amount"] = expenses["amount"].abs()

    pivot = expenses.pivot_table(
        index="major_category",
        columns="year_month",
        values="abs_amount",
        aggfunc="sum",
        fill_value=0,
    )

    # カラム名を文字列に
    pivot.columns = [str(c) for c in pivot.columns]
    pivot = pivot.reset_index().rename(columns={"major_category": "カテゴリ"})

    # 合計行を追加
    total_row = pivot.select_dtypes(include="number").sum()
    total_row_df = pd.DataFrame([["合計"] + total_row.tolist()], columns=pivot.columns)
    pivot = pd.concat([pivot, total_row_df], ignore_index=True)

    return pivot


def _format_transactions(df: pd.DataFrame) -> pd.DataFrame:
    """取引一覧を見やすい形式に整形する。"""
    if df.empty:
        return pd.DataFrame(columns=["日付", "内容", "金額", "大項目", "中項目", "金融機関", "メモ"])

    col_map = {
        "date": "日付",
        "description": "内容",
        "amount": "金額",
        "major_category": "大項目",
        "sub_category": "中項目",
        "institution": "金融機関",
        "memo": "メモ",
    }

    available = {k: v for k, v in col_map.items() if k in df.columns}
    result = df[list(available.keys())].rename(columns=available).copy()

    if "日付" in result.columns:
        result["日付"] = result["日付"].dt.strftime("%Y-%m-%d")

    result = result.sort_values("日付", ascending=False).reset_index(drop=True)
    return result


def _calculate_stats(df: pd.DataFrame) -> dict:
    """全体の統計情報を算出する。"""
    if df.empty:
        return {
            "total_income": 0,
            "total_expense": 0,
            "net": 0,
            "avg_monthly_income": 0,
            "avg_monthly_expense": 0,
            "num_months": 0,
            "num_transactions": 0,
            "top_expense_category": "N/A",
        }

    income = df[df["amount"] > 0]["amount"].sum()
    expense = df[df["amount"] < 0]["amount"].sum()
    num_months = df["year_month"].nunique() if "year_month" in df.columns else 1

    top_category = "N/A"
    if "major_category" in df.columns:
        expenses_by_cat = (
            df[df["amount"] < 0]
            .groupby("major_category")["amount"]
            .sum()
            .abs()
        )
        if not expenses_by_cat.empty:
            top_category = expenses_by_cat.idxmax()

    return {
        "total_income": int(income),
        "total_expense": int(abs(expense)),
        "net": int(income + expense),
        "avg_monthly_income": int(income / num_months) if num_months > 0 else 0,
        "avg_monthly_expense": int(abs(expense) / num_months) if num_months > 0 else 0,
        "num_months": num_months,
        "num_transactions": len(df),
        "top_expense_category": top_category,
    }
