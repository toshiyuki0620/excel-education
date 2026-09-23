from typing import List
from openpyxl.cell.cell import MergedCell
from services.base_checker import BaseStageChecker


class Stage0Checker(BaseStageChecker):
    """ステージ0: 基本操作・入力マナー 判定チェッカー

    - 値の前後の不要な空白（半角・全角スペース）チェック
    - 数値入力内の全角スペース混入チェック
    - デフォルトシート名（Sheet1）の放置チェック
    - ワークシート全体の空欄チェック
    """

    def check_cell(self, cell, val: str) -> List[str]:
        """セル単位の判定ロジック"""
        errors = []

        # 1. 結合セル（MergedCell）は属性アクセスエラー防止のためスキップ
        if isinstance(cell, MergedCell):
            return errors

        # 2. 空セルはスキップ
        if cell.value is None:
            return errors

        raw_str = str(cell.value)
        coord = cell.coordinate.upper()

        # (a) 値の前後に余分な空白（全角・半角）が含まれているか
        if raw_str != raw_str.strip() or raw_str != raw_str.strip(" "):
            errors.append(
                f"セル {coord}: ⚠️値の前後に不要な空白（スペース）が入っています。トリムして削除しましょう。"
            )

        # (b) 数値型または数式の中に全角スペースが混入していないか
        if (" " in raw_str) and (isinstance(cell.value, (int, float)) or raw_str.startswith("=")):
            errors.append(
                f"セル {coord}: ⚠️入力値の中に全角スペースが含まれています。"
            )

        return errors

    def check_sheet(self) -> List[str]:
        """シート単位の判定ロジック"""
        errors = []

        # 1. デフォルトシート名のままになっていないか
        sheet_name = self.ws.title.strip()
        if sheet_name.lower() in ["sheet1", "シート1"]:
            errors.append(
                f"⚠️ワークシート名が「{sheet_name}」のままです。分かりやすいシート名に変更しましょう。"
            )

        # 2. シート全体のデータ存在チェック
        has_data = False
        for row in self.ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                if cell.value is not None and str(cell.value).strip() != "":
                    has_data = True
                    break
            if has_data:
                break

        if not has_data:
            errors.append("⚠️ワークシートにデータが入力されていません。")

        return errors