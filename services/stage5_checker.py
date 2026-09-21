import re
from typing import List
from services.base_checker import BaseStageChecker


class Stage5Checker(BaseStageChecker):
    """ステージ5: データ整形（テーブル・フィルタ・条件付き書式） 判定チェッカー"""

    def check_cell(self, cell, val: str) -> List[str]:
        """セル単位の判定: ステージ5はシート全体の設定確認が主のため、セル個別チェックはなし"""
        return []

    def check_sheet(self) -> List[str]:
        """シート単位の判定: テーブル化、オートフィルタ、条件付き書式の設定有無を検証"""
        errors = []

        # 1. テーブル化（openpyxl.worksheet.worksheet.Worksheet.tables）のチェック
        if not self.ws.tables:
            errors.append(
                "⚠️テーブル化（Ctrl+T）が行われていません。"
                "表全体（B10:F25）を選択して「挿入」→「テーブル」で設定しましょう。"
            )
        else:
            # テーブルの指定範囲が適切かどうかの確認
            for tbl in self.ws.tables.values():
                ref = tbl.ref  # 例: "B10:F25"
                match_tbl = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", ref)
                if match_tbl:
                    _, start_row, _, end_row = match_tbl.groups()
                    if int(start_row) > 10 or int(end_row) < 20:
                        errors.append(
                            f"⚠️テーブルの範囲（{ref}）が在庫データ全体を含んでいない可能性があります。"
                            f"ヘッダー行（10行目）から全データ行（25行目）を含めて設定しましょう。"
                        )

        # 2. オートフィルタのチェック
        # （シート直下の auto_filter または テーブル内の autoFilter を確認）
        has_filter = bool(self.ws.auto_filter.ref) or any(
            table.autoFilter and table.autoFilter.ref
            for table in self.ws.tables.values()
        )
        if not has_filter:
            errors.append(
                "⚠️フィルタが設定されていません。"
                "テーブル化するとフィルタは自動で付きます。テーブル化を先に確認しましょう。"
            )

        # 3. 条件付き書式のチェック
        if not any(
            cf_range.rules for cf_range in self.ws.conditional_formatting
        ):
            errors.append(
                "⚠️条件付き書式が設定されていません。"
                "「ホーム」→「条件付き書式」→「新しいルール」で在庫切れ（在庫数≦発注点）の"
                "行をオレンジ色にハイライトしましょう。"
            )

        return errors