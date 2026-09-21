from typing import List
from services.base_checker import BaseStageChecker
from utils import is_date_cell, normalize_for_compare

class Stage1Checker(BaseStageChecker):
    """ステージ1: データ型 判定チェッカー"""

    def check_cell(self, cell, val: str) -> List[str]:
        """セル単位の判定: 数値セルに「円」が直接入力されていないかを検証"""
        errors = []

        # 数値セル（数式以外）に直接「円」が文字として入力されているかチェック
        if "円" in val and not val.startswith("="):
            errors.append(
                f"セル {cell.coordinate}: ⚠️値に「円」が直接入力されています。"
                f"数値の後ろに単位をつけたい場合は『セルの書式設定』を使いましょう。"
            )

        return errors

    def check_sheet(self) -> List[str]:
        """シート単位の判定: answer_keys テーブルから正解のデータ型を取得して検証"""
        errors = []
        type_labels = {
            "n": "数値",
            "s": "文字列",
            "d": "日付",
            "b": "真偽値",
            "f": "数式",
        }

        # 1. Supabaseからステージ1の型チェックキーを取得
        try:
            response = (
                self.supabase.table("answer_keys")
                .select("*")
                .eq("stage_id", 1)
                .execute()
            )
            type_keys = [
                k for k in (response.data or []) if k.get("expected_type")
            ]
        except Exception as e:
            return [
                f"【注意】型チェック用データの取得に失敗しました（{e}）。管理者に確認してください。"
            ]

        # 2. 各対象セルのデータ型チェック
        for key in type_keys:
            cell_ref = key.get("cell")
            expected_type = key.get("expected_type")

            try:
                cell = self.ws[cell_ref]
            except Exception:
                errors.append(
                    f"⚠️セル指定「{cell_ref}」が不正なため確認できませんでした。"
                )
                continue

            actual_type = cell.data_type
            cell_val_str = (
                str(cell.value) if cell.value is not None else ""
            )

            # 日付型の判定（utils.py の is_date_cell を利用）
            if expected_type == "d" and is_date_cell(cell):
                continue

            # 型が異なる場合の処理
            if actual_type != expected_type:
                # 数値型期待で文字列扱いに「円」が含まれている場合は check_cell 側で警告を出すためスキップ
                if (
                    expected_type == "n"
                    and actual_type == "s"
                    and "円" in cell_val_str
                ):
                    continue

                expected_label = type_labels.get(
                    expected_type, expected_type
                )

                # 初心者向けのアドバイスメッセージ生成
                guidance = ""
                if expected_type == "n" and actual_type == "s":
                    guidance = "数字の前にアポストロフィ（'）が付いていないか確認しましょう。"
                elif expected_type == "d" and actual_type == "s":
                    guidance = "「2026/9/8」のようにスラッシュ区切りで入力すると自動的に日付として認識されます。"

                errors.append(
                    f"セル {cell_ref}: ⚠️{expected_label}を入力してほしいところですが、"
                    f"セルが左寄せ（文字列扱い）になっています。{guidance}"
                )

        return errors