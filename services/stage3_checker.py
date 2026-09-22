import re
from typing import List
from openpyxl.cell.cell import MergedCell
from services.base_checker import BaseStageChecker


class Stage3Checker(BaseStageChecker):
    """ステージ3: 参照の理解（相対参照・絶対参照・SUM関数） 判定チェッカー

    対象セル範囲: E10:F15
    - E10〜E14: 小計（例: =C10*D10）
    - F10〜F14: 税込合計（例: =E10*(1+$D$4) ※D4への絶対参照 $D$4 が必要）
    - E15: 小計合計（例: =SUM(E10:E14)）
    - F15: 税込合計合計（例: =SUM(F10:F14)）
    """

    # チェック対象のセル範囲（E10:F15）
    TARGET_ROWS = range(10, 16)  # 10〜15行目
    TARGET_COLS = {"E", "F"}     # E列・F列

    def check_cell(self, cell, val: str) -> List[str]:
        """セル単位の判定: 指定範囲（E10:F15）のみをチェック"""
        errors = []

        # 1. 結合セル（MergedCell）の場合は属性アクセスエラーを防ぐため即座にスキップ
        if isinstance(cell, MergedCell):
            return errors

        # 2. 列名（"E", "F" など）の安全な取得
        col_letter = (
            cell.column_letter.upper()
            if hasattr(cell, "column_letter")
            else ""
        )
        row_num = cell.row
        coord = cell.coordinate.upper()

        # 3. E10:F15 以外のセルは対象外としてスキップ
        if col_letter not in self.TARGET_COLS or row_num not in self.TARGET_ROWS:
            return errors

        # 4. 未入力チェック
        if cell.value is None or str(cell.value).strip() == "":
            errors.append(f"セル {coord}: ⚠️数式が未入力です。")
            return errors

        upper_val = str(val).upper().replace(" ", "").replace(" ", "")

        # 5. 数式（=）開始チェック
        if not upper_val.startswith("=") and not upper_val.startswith("＝"):
            errors.append(
                f"セル {coord}: ⚠️数式が「=」から始まっていません。計算式を入力してください。"
            )
            return errors

        # --- 10〜14行目（商品別計算） ---
        if 10 <= row_num <= 14:
            # E列（小計）: 相対参照による掛け算チェック
            if col_letter == "E":
                if "*" not in upper_val:
                    errors.append(
                        f"セル {coord}: ⚠️「単価×数量」の計算が行われていません（「*」記号を使用してください）。"
                    )

            # F列（税込合計）: D4 に対する絶対参照（$D$4）チェック
            elif col_letter == "F":
                # 消費税率（0.1）を直接ベタ打ちしている場合（例: =E10*1.1 や =E10*(1+0.1)）
                if re.search(r"1\.1|0\.1", upper_val):
                    errors.append(
                        f"セル {coord}: ⚠️消費税率を数式内に直接数値（0.1 や 1.1）で入力しています。"
                        "税率セル「D4」を絶対参照（$D$4）で参照しましょう。"
                    )
                # D4 への参照はあるが $ マークが付いていない場合（例: =E10*(1+D4)）
                elif "D4" in upper_val and "$D$4" not in upper_val and "D$4" not in upper_val:
                    errors.append(
                        f"セル {coord}: ⚠️消費税率セル「D4」に絶対参照記号（$）が付いていません（例: $D$4）。"
                        "下方向にオートフィル（コピー）した際に参照先がズレてしまわないか確認しましょう。"
                    )

        # --- 15行目（合計欄） ---
        elif row_num == 15:
            if "SUM" not in upper_val:
                errors.append(
                    f"セル {coord}: ⚠️合計の計算には SUM 関数を使用してください。"
                )

        return errors

    def check_sheet(self) -> List[str]:
        """シート単位の判定: 必須領域（E10:F15）の未入力が一括でないか確認"""
        errors = []

        for row in self.TARGET_ROWS:
            for col in ["E", "F"]:
                coord = f"{col}{row}"
                cell = self.ws[coord]

                # 結合セルならチェック不要
                if isinstance(cell, MergedCell):
                    continue

                cell_val = cell.value
                if cell_val is None or str(cell_val).strip() == "":
                    errors.append(f"セル {coord}: ⚠️課題入力欄が空欄になっています。")

        return errors