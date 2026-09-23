import re
from typing import List
from openpyxl.cell.cell import MergedCell
from services.base_checker import BaseStageChecker


class Stage5Checker(BaseStageChecker):
    """ステージ5: データ整形・テーブル・フィルター連動 判定チェッカー

    - Excelテーブル機能（ws.tables）の作成有無チェック
    - SUBTOTAL 関数（または AGGREGATE 関数）の使用チェック
    - SUBTOTAL 関数の第1引数（集計機能番号: 9/109 など）の検証
    """

    def check_cell(self, cell, val: str) -> List[str]:
        """セル単位の判定ロジック"""
        errors = []

        # 1. 結合セル（MergedCell）は属性アクセスエラー防止のためスキップ
        if isinstance(cell, MergedCell):
            return errors

        # 2. 空セルまたは数式でないセルはスキップ
        if cell.value is None or not str(cell.value).strip().startswith("="):
            return errors

        raw_val = str(cell.value).strip()
        upper_val = raw_val.upper().replace(" ", "").replace(" ", "")
        coord = cell.coordinate.upper()

        # SUBTOTAL 関数が使われている場合の詳細検証
        if "SUBTOTAL(" in upper_val:
            # SUBTOTAL(機能番号, 参照範囲) の第1引数をチェック
            match = re.search(r"SUBTOTAL\(([^,]+),", upper_val)
            if match:
                func_num = match.group(1).strip()
                # 9 (SUM: 非表示行含む) または 109 (SUM: 手動非表示行無視)
                if func_num not in ["9", "109", "1", "101"]:
                    errors.append(
                        f"セル {coord}: 💡SUBTOTAL関数の第1引数（集計方法）に「{func_num}」が指定されています。"
                        "合計の計算には「9」または「109」を使用するのが一般的です。"
                    )

        return errors

    def check_sheet(self) -> List[str]:
        """シート単位の判定ロジック（テーブル機能や特定関数の設定状況）"""
        errors = []

        # 1. シート内に Excel テーブルオブジェクトが定義されているかチェック
        # openpyxl では ws.tables に定義済みのテーブルが格納されます
        if not self.ws.tables or len(self.ws.tables) == 0:
            errors.append(
                "⚠️ワークシート内に「Excelテーブル」が設定されていません。[挿入] ＞ [テーブル] からテーブル化しましょう。"
            )

        # 2. SUBTOTAL 関数がどこかに使用されているかチェック
        has_subtotal = False
        for row in self.ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                if cell.value and "SUBTOTAL(" in str(cell.value).upper():
                    has_subtotal = True
                    break
            if has_subtotal:
                break

        if not has_subtotal:
            errors.append(
                "⚠️フィルター連動集計用の「SUBTOTAL 関数」が見つかりませんでした。"
            )

        return errors