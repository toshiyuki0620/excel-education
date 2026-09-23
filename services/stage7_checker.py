import re
from typing import List
from openpyxl.cell.cell import MergedCell
from services.base_checker import BaseStageChecker


class Stage7Checker(BaseStageChecker):
    """ステージ7: 分析基礎（ピボットテーブル・RANK関数・他シート参照） 判定チェッカー

    - ピボットテーブルオブジェクト（ws._pivots）の有無チェック
    - RANK / RANK.EQ 関数の使用および参照範囲の絶対参照（$）チェック
    - 別シート参照（!）時の記述妥当性チェック
    """

    def check_cell(self, cell, val: str) -> List[str]:
        """セル単位の判定ロジック"""
        errors = []

        if isinstance(cell, MergedCell):
            return errors

        if cell.value is None or not str(cell.value).strip().startswith("="):
            return errors

        raw_val = str(cell.value).strip()
        upper_val = raw_val.upper().replace(" ", "").replace(" ", "")
        coord = cell.coordinate.upper()

        # --- RANK / RANK.EQ 関数のチェック ---
        if "RANK(" in upper_val or "RANK.EQ(" in upper_val:
            # 第2引数（参照範囲）を取得して絶対参照 ($) が付いているか確認
            match = re.search(r"RANK(?:\.EQ)?\(([^,]+),([^,]+)", upper_val)
            if match:
                ref_range = match.group(2)
                if ":" in ref_range and "$" not in ref_range:
                    errors.append(
                        f"セル {coord}: ⚠️RANK関数の参照範囲（{ref_range}）に絶対参照記号（$）が付いていません。"
                        "下方向にオートフィル（コピー）した際に範囲がズレないよう $ マーク（例: $C$5:$C$20）で固定しましょう。"
                    )

        # --- 別シート参照（!）のチェック ---
        if "!" in upper_val:
            # 例: Sheet2!A1 のような形式
            if not re.search(r"['\"]?[\w\s]+['\"]?!", upper_val):
                errors.append(
                    f"セル {coord}: ⚠️他シートのセル参照形式が不正確です（例: シート名!セル番地）。"
                )

        return errors

    def check_sheet(self) -> List[str]:
        """シート単位の判定ロジック"""
        errors = []

        # ピボットテーブルの存在確認（openpyxl では _pivots に格納されます）
        pivots = getattr(self.ws, "_pivots", [])
        if not pivots or len(pivots) == 0:
            # 必須課題がピボットテーブル作成の場合のアドバイス
            errors.append(
                "💡ヒント: ピボットテーブルが作成されていないか、別シートに作成されている可能性があります。[挿入] ＞ [ピボットテーブル] を確認してください。"
            )

        return errors