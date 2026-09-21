import re
from typing import List
from services.base_checker import BaseStageChecker


class Stage7Checker(BaseStageChecker):
    """ステージ7: 分析基礎（ピボットテーブル・統計関数・複数シート参照） 判定チェッカー"""

    REQUIRED_STATS_CELLS = {
        "C5": ("RANK", "順位計算（RANK.EQ関数など）"),
        "C6": ("RANK", "順位計算（RANK.EQ関数など）"),
        "F10": ("LARGE", "上位データの取得（LARGE関数）"),
    }

    def check_cell(self, cell, val: str) -> List[str]:
        """セル単位の判定: 他シート参照時の構文エラーや絶対参照漏れを検証"""
        errors = []
        upper_val = val.upper().replace(" ", "")

        if not upper_val.startswith("="):
            return errors

        # 1. 他シート参照（例: =データ!A1 ）の構文チェック
        if "!" in upper_val:
            if upper_val.endswith("!"):
                errors.append(
                    f"セル {cell.coordinate}: ⚠️シート参照（!）の後ろにセル範囲が指定されていません。"
                )

        # 2. RANK関数の範囲固定チェック
        if "RANK" in upper_val and "$" not in upper_val:
            errors.append(
                f"セル {cell.coordinate}: ⚠️RANK関数の参照範囲に絶対参照（$）が使われていません。"
                "順位比較の基準範囲がズレないように $ で固定しましょう。"
            )

        return errors

    def check_sheet(self) -> List[str]:
        """シート単位の判定: ピボットテーブルの挿入有無と必須の統計関数の存在を検証"""
        errors = []

        # 1. ピボットテーブルの存在確認（openpyxl の _pivots を参照）
        has_pivot = hasattr(self.ws, "_pivots") and len(self.ws._pivots) > 0
        if not has_pivot:
            errors.append(
                "⚠️ピボットテーブルが作成されていません。"
                "「挿入」→「ピボットテーブル」からクロス集計表を作成しましょう。"
            )

        # 2. 統計関数の必須セルチェック
        for cell_ref, (func_name, label) in self.REQUIRED_STATS_CELLS.items():
            try:
                cell_val = (
                    str(self.ws[cell_ref].value or "").upper().replace(" ", "")
                )
            except Exception:
                continue

            if not cell_val.startswith("="):
                errors.append(
                    f"セル {cell_ref}（{label}）: ⚠️数式が未入力です。"
                )
            elif func_name not in cell_val:
                errors.append(
                    f"セル {cell_ref}（{label}）: ⚠️{func_name}関数が使われていません。"
                )

        return errors