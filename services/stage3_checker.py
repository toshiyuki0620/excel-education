import re
from typing import List
from services.base_checker import BaseStageChecker


class Stage3Checker(BaseStageChecker):
    """ステージ3: 参照の理解（相対参照・絶対参照・複合参照） 判定チェッカー"""

    # 必須入力セルと期待される参照形式（表記・説明用）
    REQUIRED_REFERENCES = {
        "D5": ("絶対参照", "単価・税率セルへの参照（$指定）"),
        "D6": ("絶対参照", "単価・税率セルへの参照（$指定）"),
        "D7": ("絶対参照", "単価・税率セルへの参照（$指定）"),
        "E5": ("複合参照/絶対参照", "行または列の固定参照"),
    }

    def check_cell(self, cell, val: str) -> List[str]:
        """セル単位の判定: 絶対参照（$）の漏れや不要な参照形式を検証"""
        errors = []
        upper_val = val.upper().replace(" ", "")

        if not upper_val.startswith("="):
            return errors

        # 1. 共通の数式チェック: セル直接の数値入力（ベタ打ち）チェック
        # 例: =B5*1.1 や =B5*100 のようにセル参照を使わずに数値を直打ちしているケース
        if re.search(r"=\s*[A-Z]+\d+\s*[\*\+\-\/]\s*\d+", upper_val):
            errors.append(
                f"セル {cell.coordinate}: ⚠️数式内に数値が直接入力（ベタ打ち）されています。"
                "条件や税率が変更された際に対応できるよう、該当数値が入ったセルを参照しましょう。"
            )

        # 2. オートフィル（縦展開・横展開）想定セルでの絶対参照漏れ確認
        # D5〜D12 などの計算列で固定セル（例: C2）を参照しているのに $ がついていない場合
        if cell.row >= 5 and cell.column_letter in ["D", "E"]:
            # 数式内にセル参照（例: C2, B3 など）が含まれているか抽出
            cell_refs = re.findall(r"\b[A-Z]+\d+\b", upper_val)
            if cell_refs:
                errors.append(
                    f"セル {cell.coordinate}: ⚠️数式に絶対参照記号（$）が付いていません（例: {cell_refs[0]} -> ${cell_refs[0][0]}${cell_refs[0][1:]}）。"
                    "下方向にオートフィル（コピー）した際に参照先がズレてしまわないか確認しましょう。"
                )

        return errors

    def check_sheet(self) -> List[str]:
        """シート単位の判定: 必須計算セルで正しい参照設定・計算が行われているか検証"""
        errors = []

        for cell_ref, (ref_type, label) in self.REQUIRED_REFERENCES.items():
            try:
                cell_val = (
                    str(self.ws[cell_ref].value or "").upper().replace(" ", "")
                )
            except Exception:
                continue

            # 数式未入力チェック
            if not cell_val.startswith("="):
                errors.append(
                    f"セル {cell_ref}（{label}）: ⚠️数式が入力されていません。"
                    f"{ref_type}を活用して計算式を作成しましょう。"
                )
            # $ 記号が含まれていないチェック（絶対参照/複合参照が求められるセル）
            elif "$" not in cell_val:
                errors.append(
                    f"セル {cell_ref}（{label}）: ⚠️{ref_type}のドル記号（$）が設定されていません。"
                    "F4キーを押して参照形式を切り替えましょう。"
                )

        return errors