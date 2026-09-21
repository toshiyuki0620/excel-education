import re
from typing import List
from services.base_checker import BaseStageChecker


class Stage8Checker(BaseStageChecker):
    """ステージ8: 総合演習（実務レベルのデータ処理・ダッシュボード作成） 判定チェッカー"""

    def check_cell(self, cell, val: str) -> List[str]:
        """セル単位の判定: 数式エラーの残存や複雑な複合関数の入力不備を検証"""
        errors = []
        upper_val = val.upper().replace(" ", "")

        # 1. シート全体の エラー値 残存チェック
        for err_code in ["#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NAME?"]:
            if err_code in val:
                errors.append(
                    f"セル {cell.coordinate}: ⚠️計算エラー（{err_code}）が発生しています。"
                    "IFERROR関数で囲むか、参照先の入力データを見直してください。"
                )
                return errors

        if not upper_val.startswith("="):
            return errors

        # 2. 複合関数のネストチェック (例: IFERROR(VLOOKUP(...), "未登録") )
        if "VLOOKUP(" in upper_val and "IFERROR(" not in upper_val:
            # 総合演習ではエラーハンドリング（IFERROR）の併用を推奨する例
            errors.append(
                f"セル {cell.coordinate}: 💡ヒント: VLOOKUP関数を IFERROR関数 と組み合わせると、該当データがない場合のエラー表示をきれいに防げます。"
            )

        return errors

    def check_sheet(self) -> List[str]:
        """シート単位の判定: 総合課題としての完成度（複数シート連携、最終集計セルの確認）を検証"""
        errors = []

        # 1. ワークブック全体のシート数チェック（データシート・集計シートの分離）
        wb = self.ws.parent
        if len(wb.worksheets) < 2:
            errors.append(
                "⚠️ワークシートが1枚しかありません。"
                "「売上データ」と「分析ダッシュボード」のように目的別にシートを分けて構成しましょう。"
            )

        # 2. 最終集計セル（例: B3セル）の確認
        try:
            summary_val = str(self.ws["B3"].value or "").strip()
            if not summary_val.startswith("="):
                errors.append(
                    "セル B3（総合売上集計）: ⚠️最終集計値が数式で計算されていません。"
                )
        except Exception:
            pass

        return errors