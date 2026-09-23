from abc import ABC, abstractmethod
from typing import List
from openpyxl.cell.cell import MergedCell


class BaseStageChecker(ABC):
    """全ステージチェッカーの基底クラス"""

    def __init__(self, ws, file_stream=None):
        self.ws = ws
        self.file_stream = file_stream

    @abstractmethod
    def check_cell(self, cell, val: str) -> List[str]:
        """各セル単位のチェック（子クラスで実装）"""
        pass

    @abstractmethod
    def check_sheet(self) -> List[str]:
        """シート・ファイル全体単位のチェック（子クラスで実装）"""
        pass

    def run_check(self) -> List[str]:
        """全セルのチェックを実行する共通エントリーポイント"""
        errors = []

        # 1. シート全体のセル走査
        for row in self.ws.iter_rows():
            for cell in row:
                # ★共通ガード: 結合セル（MergedCell）はチェック対象外としてスキップ
                if isinstance(cell, MergedCell):
                    continue

                # 値が存在する場合のみ文字列化して渡す（None対策）
                val = str(cell.value) if cell.value is not None else ""
                
                # 安全な cell のみを各チェッカーの check_cell に渡す
                cell_errors = self.check_cell(cell, val)
                if cell_errors:
                    errors.extend(cell_errors)

        # 2. シート単位のチェック実行
        sheet_errors = self.check_sheet()
        if sheet_errors:
            errors.extend(sheet_errors)

        return errors