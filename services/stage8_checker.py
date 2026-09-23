from typing import List
from openpyxl.cell.cell import MergedCell
from services.base_checker import BaseStageChecker


class Stage8Checker(BaseStageChecker):
    """ステージ8: 総合演習（エラー制御・ダッシュボード構築） 判定チェッカー

    - シート内の数式エラー（#REF!, #VALUE!, #N/A, #DIV/0! 等）の残存検知
    - IFERROR 関数によるエラー制御（マスク処理）の利用有無チェック
    """

    # Excel の一般的なエラー値のリスト
    EXCEL_ERRORS = ["#REF!", "#VALUE!", "#N/A", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!"]

    def check_cell(self, cell, val: str) -> List[str]:
        """セル単位の判定ロジック"""
        errors = []

        if isinstance(cell, MergedCell):
            return errors

        if cell.value is None:
            return errors

        str_val = str(cell.value).strip()
        upper_val = str_val.upper()
        coord = cell.coordinate.upper()

        # 1. 数式結果またはセル値に数式エラーが残っていないか検知
        for err in self.EXCEL_ERRORS:
            if err in upper_val:
                errors.append(
                    f"セル {coord}: ⚠️計算エラー（{err}）が発生しています。"
                    "数式や参照先のセルを見直すか、IFERROR関数でエラー処理を行いましょう。"
                )
                break

        # 2. VLOOKUP などでエラーが出やすい箇所に IFERROR が組み込まれているかのアドバイス
        if upper_val.startswith("=") and "VLOOKUP(" in upper_val and "IFERROR(" not in upper_val:
            errors.append(
                f"セル {coord}: 💡VLOOKUP関数が単体で使用されています。"
                "データが見つからない場合のエラー表示を防ぐため、IFERROR関数（例: =IFERROR(VLOOKUP(...), \"\")）を組み合わせるとより実践的です。"
            )

        return errors

    def check_sheet(self) -> List[str]:
        """シート単位の判定ロジック（エラーセルの総数カウント）"""
        errors = []
        error_count = 0

        for row in self.ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                if cell.value:
                    u_val = str(cell.value).upper()
                    if any(err in u_val for err in self.EXCEL_ERRORS):
                        error_count += 1

        if error_count > 0:
            errors.append(
                f"⚠️シート全体で {error_count} 個のエラーセル（#N/A や #REF! など）が残っています。"
            )

        return errors