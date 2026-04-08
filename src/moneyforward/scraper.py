"""マネーフォワード ME から自動でCSVをダウンロードするスクレイパー

Playwright を使用してブラウザ操作を自動化し、
家計簿データのCSVエクスポートをダウンロードします。

必要な環境変数:
    MF_EMAIL: マネーフォワードのログインメールアドレス
    MF_PASSWORD: マネーフォワードのログインパスワード
"""

import os
import time
from datetime import datetime, timedelta
from pathlib import Path

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout


MF_BASE_URL = "https://moneyforward.com"
MF_LOGIN_URL = f"{MF_BASE_URL}/sign_in"
MF_CF_URL = f"{MF_BASE_URL}/cf"

DEFAULT_DOWNLOAD_DIR = "downloads"


def download_csv(
    email: str | None = None,
    password: str | None = None,
    year: int | None = None,
    month: int | None = None,
    download_dir: str = DEFAULT_DOWNLOAD_DIR,
    headless: bool = True,
) -> str:
    """マネーフォワードにログインし、指定月のCSVをダウンロードする。

    Args:
        email: ログインメールアドレス（未指定時は環境変数 MF_EMAIL）
        password: ログインパスワード（未指定時は環境変数 MF_PASSWORD）
        year: 対象年（未指定時は当月）
        month: 対象月（未指定時は当月）
        download_dir: ダウンロード先ディレクトリ
        headless: ヘッドレスモードで実行するか

    Returns:
        ダウンロードしたCSVファイルのパス
    """
    email = email or os.environ.get("MF_EMAIL", "")
    password = password or os.environ.get("MF_PASSWORD", "")

    if not email or not password:
        raise ValueError(
            "ログイン情報が必要です。環境変数 MF_EMAIL, MF_PASSWORD を設定するか、"
            "引数で email, password を指定してください。"
        )

    # 対象月を決定（未指定時は前月）
    if year is None or month is None:
        today = datetime.now()
        # 月末実行を想定し、当月のデータを取得
        target = today.replace(day=1)
        year = target.year
        month = target.month

    download_path = Path(download_dir).resolve()
    download_path.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(
            accept_downloads=True,
            locale="ja-JP",
        )
        page = context.new_page()

        try:
            # ログイン
            _login(page, email, password)

            # 家計簿ページに遷移
            csv_path = _download_monthly_csv(page, year, month, download_path)

            return csv_path
        finally:
            browser.close()


def download_csv_range(
    start_year: int,
    start_month: int,
    end_year: int,
    end_month: int,
    email: str | None = None,
    password: str | None = None,
    download_dir: str = DEFAULT_DOWNLOAD_DIR,
    headless: bool = True,
) -> list[str]:
    """複数月分のCSVを一括ダウンロードする。

    Returns:
        ダウンロードしたCSVファイルパスのリスト
    """
    email = email or os.environ.get("MF_EMAIL", "")
    password = password or os.environ.get("MF_PASSWORD", "")

    if not email or not password:
        raise ValueError(
            "ログイン情報が必要です。環境変数 MF_EMAIL, MF_PASSWORD を設定してください。"
        )

    download_path = Path(download_dir).resolve()
    download_path.mkdir(parents=True, exist_ok=True)

    csv_files = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(
            accept_downloads=True,
            locale="ja-JP",
        )
        page = context.new_page()

        try:
            _login(page, email, password)

            # 各月のCSVをダウンロード
            current = datetime(start_year, start_month, 1)
            end = datetime(end_year, end_month, 1)

            while current <= end:
                print(f"  {current.year}/{current.month:02d} のデータをダウンロード中...")
                csv_path = _download_monthly_csv(
                    page, current.year, current.month, download_path
                )
                csv_files.append(csv_path)

                # 次の月へ
                if current.month == 12:
                    current = current.replace(year=current.year + 1, month=1)
                else:
                    current = current.replace(month=current.month + 1)

                time.sleep(2)  # サーバーへの負荷軽減

        finally:
            browser.close()

    return csv_files


def _login(page, email: str, password: str) -> None:
    """マネーフォワードにログインする。"""
    print("  マネーフォワードにログイン中...")
    page.goto(MF_LOGIN_URL, wait_until="domcontentloaded")
    time.sleep(2)

    # メールアドレスでログインを選択
    email_login_btn = page.locator("a:has-text('メールアドレスでログイン')")
    if email_login_btn.count() > 0:
        email_login_btn.first.click()
        time.sleep(1)

    # メールアドレス入力
    page.fill('input[name="mfid_user[email]"]', email)
    page.click('input[type="submit"], button[type="submit"]')
    time.sleep(2)

    # パスワード入力
    page.fill('input[name="mfid_user[password]"]', password)
    page.click('input[type="submit"], button[type="submit"]')
    time.sleep(3)

    # ログイン成功の確認
    try:
        page.wait_for_url("**/", timeout=10000)
    except PlaywrightTimeout:
        # 2段階認証やCAPTCHAが出た場合
        if "sign_in" in page.url:
            raise RuntimeError(
                "ログインに失敗しました。メールアドレス/パスワードを確認するか、"
                "2段階認証を一時的に無効にしてください。"
            )

    print("  ログイン成功")


def _download_monthly_csv(
    page, year: int, month: int, download_path: Path
) -> str:
    """指定月の家計簿CSVをダウンロードする。"""
    # 家計簿ページ（月指定）
    cf_url = f"{MF_CF_URL}?month={year}-{month:02d}"
    page.goto(cf_url, wait_until="domcontentloaded")
    time.sleep(3)

    # CSVダウンロードボタンをクリック
    with page.expect_download(timeout=30000) as download_info:
        # 「CSV」ダウンロードリンクを探す
        csv_link = page.locator('a:has-text("CSV"), a[href*="csv"]')
        if csv_link.count() == 0:
            raise RuntimeError(
                f"CSVダウンロードリンクが見つかりません ({year}/{month:02d})"
            )
        csv_link.first.click()

    download = download_info.value
    filename = f"moneyforward_{year}{month:02d}.csv"
    save_path = str(download_path / filename)
    download.save_as(save_path)

    print(f"  ダウンロード完了: {save_path}")
    return save_path
