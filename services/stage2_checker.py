import re
from typing import List
from openpyxl.cell.cell import MergedCell
from services.base_checker import BaseStageChecker


class Stage2Checker(BaseStageChecker):
    """ステージ2: 基本関数（SUM, AVERAGE, COUNT, IF等） 判定チェッカー

    - 数式入力の有無チェック（「=」で始まるか）
    - 全角記号（＝, （, ）など）の混入チェック
    - 基本関数（SUM, AVERAGE, COUNT, IF 等）の使用チェック
    - 指定セルに対する個別チェック
    """

    # チェック対象の基本関数リスト
    REQUIRED_FUNCTIONS = ["SUM", "AVERAGE", "COUNT", "COUNTA", "IF"]

    def check_cell(self, cell, val: str) -> List[str]:
        """セル単位の判定ロジック"""
        errors = []

        # 1. 結合セル（MergedCell）は属性アクセスエラー防止のため即座にスキップ
        if isinstance(cell, MergedCell):
            return errors

        # 2. 値が空の場合はスキップ（未入力チェックは check_sheet や特定セルで実施）
        if cell.value is None or str(cell.value).strip() == "":
            return errors

        raw_val = str(cell.value).strip()
        upper_val = raw_val.upper()
        coord = cell.coordinate.upper()

        # 3. 全角の等号（＝）で始まっている場合の警告
        if raw_val.startswith("＝"):
            errors.append(
                f"セル {coord}: ⚠️数式の先頭が全角の「＝」になっています。半角の「=」を使用してください。"
            )

        # 4. 数式セルの場合の検証（半角または全角の「=」で始まるセル）
        if raw_val.startswith("=") or raw_val.startswith("＝"):
            # (a) 全角の括弧やコンマなどの混入チェック
            if re.search(r"[（），＋−＊／]", raw_val):
                errors.append(
                    f"セル {coord}: ⚠️数式内に全角文字（括弧や演算子）が含まれています。すべて半角で入力してください。"
                )

            # (b) 関数名の大文字・小文字チェック（推奨表記指導）
            # 例: =sum(a1:a5) のような小文字入力を検知してアドバイス
            for func_name in self.REQUIRED_FUNCTIONS:
                pattern = rf"\b{func_name.lower()}\("
                if re.search(pattern, raw_val):
                    errors.append(
                        f"セル {coord}: 💡関数名「{func_name.lower()}」が小文字で入力されています。"
                        f"大文字の「{func_name}」で入力すると読みやすくなります。"
                    )

        return errors

    def check_sheet(self) -> List[str]:
        """シート単位の判定ロジック（シート全体の数式利用状況等の確認）"""
        errors = []
        has_formula = False

        # シート内に一つでも数式が存在するか確認
        for row in self.ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                if cell.value and str(cell.value).startswith("="):
                    has_formula = True
                    break
            if has_formula:
                break

        if not has_formula:
            errors.append("⚠️シート内に計算式（「=」で始まるセル）が見つかりませんでした。")

        return errors