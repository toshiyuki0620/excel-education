import datetime
from typing import List
from openpyxl.cell.cell import MergedCell
from services.base_checker import BaseStageChecker
from utils import is_date_cell, normalize_for_compare


class Stage1Checker(BaseStageChecker):
    """ステージ1: データ型（日付・文字列・数値） 判定チェッカー

    対象セル範囲: B9 〜 D9
    - B9: 日付（例: 2026/09/01 など日付形式であること）
    - C9: 品目（文字列）
    - D9: 金額（数値、「円」などの文字列や全角数字の混入をチェック）
    """

    # チェック対象のセル番地（課題1の入力欄）
    TARGET_CELLS = {"B9", "C9", "D9"}

    def check_cell(self, cell, val: str) -> List[str]:
        """セル単位の判定: 指定範囲（B9〜D9）のみをチェック"""
        errors = []

        # 1. 結合セル（MergedCell）は属性アクセスエラー防止のため即座にスキップ
        if isinstance(cell, MergedCell):
            return errors

        coord = cell.coordinate.upper()

        # 2. B9, C9, D9 以外のセルはスキップ
        if coord not in self.TARGET_CELLS:
            return errors

        # 3. 未入力（空欄）チェック
        if cell.value is None or str(cell.value).strip() == "":
            errors.append(f"セル {coord}: ⚠️値が未入力です。")
            return errors

        norm_val = normalize_for_compare(val)

        # --- B9 セル: 日付型チェック ---
        if coord == "B9":
            if not is_date_cell(cell):
                errors.append(
                    f"セル {coord}: ⚠️「日付」として認識されていません。"
                    "「2026/9/1」のようにスラッシュ区切りで入力してください。"
                )

        # --- C9 セル: 品目（文字列）チェック ---
        elif coord == "C9":
            # 数値のみで入力されている場合は注意
            if isinstance(cell.value, (int, float)) or (val.isdigit() and norm_val.isdigit()):
                errors.append(
                    f"セル {coord}: ⚠️品目（文字列）欄に数値のみが入力されています。"
                    "「食費」や「日用品」などのテキストを入力してください。"
                )

        # --- D9 セル: 金額（数値型）チェック ---
        elif coord == "D9":
            # (a) 「1000円」のように「円」が文字列として直接入力されている場合
            if "円" in val:
                errors.append(
                    f"セル {coord}: ⚠️金額に「円」が直接入力されています。"
                    "数値のみを入力し、単位は表示形式で設定しましょう。"
                )

            # (b) 文字列型として入力されている場合（例: `'1500` や全角数字）
            elif isinstance(cell.value, str):
                errors.append(
                    f"セル {coord}: ⚠️金額が「文字列型」として入力されています。"
                    "半角数値のみを入力してください。"
                )

        return errors

    def check_sheet(self) -> List[str]:
        """シート単位の判定: 必須セル（B9〜D9）の未入力がないか一括確認"""
        errors = []

        for coord in ["B9", "C9", "D9"]:
            cell = self.ws[coord]

            # 結合セルならチェック不要
            if isinstance(cell, MergedCell):
                continue

            if cell.value is None or str(cell.value).strip() == "":
                errors.append(f"セル {coord}: ⚠️課題入力欄が空欄になっています。")

        return errors