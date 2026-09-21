import os
import re
from typing import List
from services.base_checker import BaseStageChecker


class Stage0Checker(BaseStageChecker):
    """ステージ0: 基本操作（ファイル名・シート名・基本フォーマット） 判定チェッカー"""

    # 提出時に入力必須の判定対象セル範囲（例: A1:D10の範囲に空欄がないか確認）
    REQUIRED_CELLS = ["A1", "A2", "B2", "C2", "D2"]

    def check_cell(self, cell, val: str) -> List[str]:
        """セル単位の判定: 文字列前後の不要な空白や全角スペース混入を検証"""
        errors = []

        if not val:
            return errors

        # 1. 前後に不要なスペースが混入しているかチェック
        if val.startswith(" ") or val.endswith(" ") or val.startswith(" ") or val.endswith(" "):
            errors.append(
                f"セル {cell.coordinate}: ⚠️値の前後に不要な空白（スペース）が入っています。"
            )

        # 2. 数値データに全角スペースが含まれていないかチェック
        if re.search(r"\d", val) and " " in val:
            errors.append(
                f"セル {cell.coordinate}: ⚠️数値の中に全角スペースが含まれています。"
            )

        return errors

    def check_sheet(self) -> List[str]:
        """シート・ファイル単位の判定: シート名、ファイル名、空欄セルの有無を検証"""
        errors = []

        # 1. シート名のチェック（デフォルトの「Sheet1」のままになっていないか）
        current_sheet_name = self.ws.title.strip()
        if current_sheet_name.lower() in ["sheet1", "シート1"]:
            errors.append(
                f"⚠️シート名が「{current_sheet_name}」のままです。"
                "「基本操作」など、指定された分かりやすいシート名に変更しましょう。"
            )

        # 2. 必須入力セルの空欄チェック
        for cell_ref in self.REQUIRED_CELLS:
            try:
                cell_val = self.ws[cell_ref].value
                if cell_val is None or str(cell_val).strip() == "":
                    errors.append(
                        f"セル {cell_ref}: ⚠️値が未入力です。指定されたデータを入力してください。"
                    )
            except Exception:
                continue

        # 3. 非表示の行列・シートが残っていないかの確認
        if self.ws.sheet_state != "visible":
            errors.append("⚠️ワークシートが非表示になっています。")

        return errors