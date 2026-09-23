import re
from typing import List
from openpyxl.cell.cell import MergedCell
from services.base_checker import BaseStageChecker


class Stage4Checker(BaseStageChecker):
    """ステージ4: 応用関数（VLOOKUP / XLOOKUP） 判定チェッカー

    - VLOOKUP または XLOOKUP 関数の記述確認
    - VLOOKUP 第4引数（検索型: FALSE / 0）の完全一致設定漏れチェック
    - 参照範囲（範囲マスタ）における絶対参照（$）の指定漏れチェック
    - 数式エラー（#N/A 等）が発生していないかの確認
    """

    def check_cell(self, cell, val: str) -> List[str]:
        """セル単位の判定ロジック"""
        errors = []

        # 1. 結合セル（MergedCell）は属性アクセスエラー防止のため即座にスキップ
        if isinstance(cell, MergedCell):
            return errors

        # 2. 値が空または数式でない（=で始まらない）場合はスキップ
        if cell.value is None:
            return errors

        raw_val = str(cell.value).strip()
        if not raw_val.startswith("=") and not raw_val.startswith("＝"):
            return errors

        coord = cell.coordinate.upper()
        upper_val = raw_val.upper().replace(" ", "").replace(" ", "")

        # 3. VLOOKUP 関数のチェック
        if "VLOOKUP(" in upper_val:
            # (a) 第4引数（検索型: FALSE または 0）の完全一致チェック
            # VLOOKUP(検索値, 範囲, 列番号, [検索型])
            # 第4引数が省かれている、または TRUE になっている場合を検出
            if not (",FALSE)" in upper_val or ",0)" in upper_val or ",FALSE," in upper_val or ",0," in upper_val):
                errors.append(
                    f"セル {coord}: ⚠️VLOOKUP 関数の第4引数（検索型）に「FALSE」または「0」（完全一致）が設定されていません。"
                    "検索型の指定を省略すると誤ったデータを取得する原因になります。"
                )

            # (b) 範囲参照の絶対参照（$）チェック
            # 例: VLOOKUP(A2, A10:B20, 2, FALSE) のように範囲に $ が無いケースを抽出
            # カンマで区切られた第2引数（範囲指定）に ':' が含まれ、$ が含まれていないか検証
            vlookup_match = re.search(r"VLOOKUP\(([^,]+),([^,]+),", upper_val)
            if vlookup_match:
                table_array = vlookup_match.group(2)
                if ":" in table_array and "$" not in table_array:
                    errors.append(
                        f"セル {coord}: ⚠️VLOOKUP 関数の参照範囲「{table_array}」に絶対参照記号（$）が付いていません。"
                        "数式を下にコピーした際に範囲がズレないよう「$」で固定しましょう。"
                    )

        # 4. XLOOKUP 関数のチェック（推奨関数の利用確認）
        elif "XLOOKUP(" in upper_val:
            # XLOOKUP の参照範囲における絶対参照指定チェック
            xlookup_match = re.search(r"XLOOKUP\(([^,]+),([^,]+),([^,]+)", upper_val)
            if xlookup_match:
                lookup_array = xlookup_match.group(2)
                return_array = xlookup_match.group(3)
                if (":" in lookup_array and "$" not in lookup_array) or (":" in return_array and "$" not in return_array):
                    errors.append(
                        f"セル {coord}: ⚠️XLOOKUP 関数の検索範囲・戻り配列に絶対参照記号（$）が付いていません。"
                        "オートフィル時の参照ズレを防ぐため「$」で固定しましょう。"
                    )

        return errors

    def check_sheet(self) -> List[str]:
        """シート単位の判定ロジック"""
        errors = []
        has_lookup = False
        has_na_error = False

        for row in self.ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue

                val = str(cell.value) if cell.value is not None else ""
                upper_val = val.upper()

                # VLOOKUP または XLOOKUP の存在チェック
                if "VLOOKUP(" in upper_val or "XLOOKUP(" in upper_val:
                    has_lookup = True

                # #N/A エラーの発生チェック
                if val == "#N/A":
                    has_na_error = True

        if not has_lookup:
            errors.append("⚠️シート内に VLOOKUP 関数または XLOOKUP 関数が見つかりませんでした。")

        if has_na_error:
            errors.append(
                "⚠️シート内に参照エラー（#N/A）が発生しているセルがあります。"
                "検索値や参照範囲、完全一致（FALSE）の設定を見直してください。"
            )

        return errors