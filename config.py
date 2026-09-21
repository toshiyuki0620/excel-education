import os
import sys
from dotenv import load_dotenv
from supabase import Client, create_client

# .env ファイルから環境変数を読み込む
load_dotenv()

def _get_required_env(var_name: str) -> str:
    """指定された環境変数を取得し、存在しないまたは空の場合はエラーを出力して終了する"""
    value = os.environ.get(var_name, "").strip()
    if not value:
        print(
            f"\n[CRITICAL ERROR] 必須の環境変数 '{var_name}' が設定されていません。"
        )
        print(
            "プロジェクトルートに `.env` ファイルが存在するか、値が空になっていないか確認してください。\n"
        )
        sys.exit(1)
    return value


# --- 環境変数の検証と取得 ---
raw_url = _get_required_env("SUPABASE_URL")
SUPABASE_KEY = _get_required_env("SUPABASE_KEY")

# URLの末尾や /rest/v1 パスの自動補正
if "/rest/v1" in raw_url:
    raw_url = raw_url.split("/rest/v1")[0]
SUPABASE_URL = raw_url.rstrip("/")

# Flask等の動作環境設定（デフォルト: development）
FLASK_ENV = os.environ.get("FLASK_ENV", "development").strip()
PORT = int(os.environ.get("PORT", "5000"))

# --- Supabase クライアントの初期化 ---
try:
    print("--- 【起動時】Supabase 接続設定確認 ---")
    print(f"URL: [{SUPABASE_URL}]")
    print("---------------------------------------")

    supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
except Exception as e:
    print(
        f"\n[CRITICAL ERROR] Supabase クライアントの初期化に失敗しました: {e}"
    )
    sys.exit(1)