import re
from typing import List
from services.base_checker import BaseStageChecker


class Stage4Checker(BaseStageChecker):
    """ステージ4: 応用関数（VLOOKUP / XLOOKUP） 判定チェッカー"""

    # VLOOKUPの入力必須セルとラベルのマッピング
    VLOOKUP_REQUIRED = {
        "C9": "商品名（C列）",
        "C10": "商品名（C列）",
        "C11": "商品名（C列）",
        "C12": "商品名（C列）",
        "C13": "商品名（C列）",
        "C14": "商品名（C列）",
        "D9": "単価（D列）",
        "D10": "単価（D列）",
        "D11": "単価（D列）",
        "D12": "単価（D列）",
        "D13": "単価（D列）",
        "D14": "単価（D列）",
    }

    def check_cell(self, cell, val: str) -> List[str]:
        """セル単位の判定: VLOOKUP/XLOOKUP の引数不備や絶対参照（$）漏れを検証"""
        errors = []
        upper_val = val.upper().replace(" ", "")

        if not upper_val.startswith("="):
            return errors

        # --- VLOOKUP の構文・参照チェック ---
        if "VLOOKUP(" in upper_val:
            vlookup_args = upper_val.count(",")
            has_false = upper_val.endswith(",FALSE)") or upper_val.endswith(",0)")

            # 第4引数（検索方法: FALSE/0）の確認
            if vlookup_args < 3 or not has_false:
                errors.append(
                    f"セル {cell.coordinate}: ⚠️VLOOKUP関数の第4引数（検索方法）に 'FALSE' または '0' が指定されていません。"
                )

            # マスタ範囲の絶対参照（$）チェック
            has_full_abs = bool(re.search(r"\$[A-Z]+\$\d+", upper_val))
            has_col_only = (
                bool(re.search(r"\$[A-Z]+\d+", upper_val)) and not has_full_abs
            )

            if has_col_only:
                errors.append(
                    f"セル {cell.coordinate}: ⚠️マスタ範囲の行番号に$がついていません（例：$A3）。"
                )
            elif not has_full_abs:
                errors.append(
                    f"セル {cell.coordinate}: ⚠️VLOOKUPのマスタ範囲に絶対参照（$）がついていません。"
                )

        # --- XLOOKUP の構文チェック ---
        elif "XLOOKUP(" in upper_val:
            if upper_val.count(",") < 2:
                errors.append(
                    f"セル {cell.coordinate}: ⚠️XLOOKUP関数の引数が足りません。"
                )
            if "$" not in upper_val:
                errors.append(
                    f"セル {cell.coordinate}: ⚠️XLOOKUP関数の検索範囲・戻り範囲に絶対参照（$）が使われていない可能性があります。"
                )

        return errors

    def check_sheet(self) -> List[str]:
        """シート単位の判定: 「マスタ」以外の作業シートで必須関数が記述されているか検証"""
        errors = []

        # 2シート構成に対応: 「マスタ」以外のシートを対象とする
        target_ws = next(
            (wb_sheet for wb_sheet in self.ws.parent.worksheets if wb_sheet.title != "マスタ"),
            self.ws,
        )

        for cell_ref, label in self.VLOOKUP_REQUIRED.items():
            try:
                cell_val = (
                    str(target_ws[cell_ref].value or "").upper().replace(" ", "")
                )
            except Exception:
                continue

            if not cell_val.startswith("="):
                errors.append(
                    f"セル {cell_ref}（{label}）: ⚠️数式が入力されていません。"
                    f"VLOOKUP関数を使って商品マスタから自動取得しましょう。"
                )
            elif "VLOOKUP(" not in cell_val and "XLOOKUP(" not in cell_val:
                errors.append(
                    f"セル {cell_ref}（{label}）: ⚠️VLOOKUP関数（またはXLOOKUP関数）が使われていません。"
                    f"=VLOOKUP(検索値, マスタ!$A$3:$C$10, 列番号, FALSE) の形式で入力しましょう。"
                )

        return errors