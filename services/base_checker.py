from abc import ABC, abstractmethod
from typing import List, Optional
from openpyxl.worksheet.worksheet import Worksheet
from supabase import Client
from config import supabase as default_supabase


class BaseStageChecker(ABC):
    """各ステージの判定ロジック用抽象基底クラス"""

    def __init__(
        self,
        ws: Worksheet,
        file_stream,
        supabase_client: Optional[Client] = None,
    ):
        self.ws = ws
        self.file_stream = file_stream
        # 引数で渡されなければ config.py のクライアントを使用（依存性の注入）
        self.supabase = supabase_client or default_supabase

    @abstractmethod
    def check_cell(self, cell, val: str) -> List[str]:
        """セル単位の判定処理。
        エラーメッセージまたは検出ログのリストを返す。
        （サブクラスで実装必須）
        """
        pass

    def check_sheet(self) -> List[str]:
        """シート/ファイル全体の判定処理。
        デフォルトでは何も実行しない（必要なステージのみオーバーライドする）。
        """
        return []