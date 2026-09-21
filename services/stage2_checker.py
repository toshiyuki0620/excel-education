import re
from typing import List
from services.base_checker import BaseStageChecker


class Stage2Checker(BaseStageChecker):
    """ステージ2: 基本関数 判定チェッカー"""

    # 必須関数チェック用のテーブルマッピング
    REQUIRED_FUNCTIONS = {
        "F8": ("SUM", "田中 一郎の合計"),
        "F9": ("SUM", "鈴木 花子の合計"),
        "F10": ("SUM", "佐藤 次郎の合計"),
        "F11": ("SUM", "山田 三枝の合計"),
        "F12": ("SUM", "伊藤 四朗の合計"),
        "F15": ("COUNT", "受験者数"),
        "F16": ("AVERAGE", "平均点"),
        "F17": ("MAX", "最高点"),
        "F18": ("MIN", "最低点"),
    }

    def check_cell(self, cell, val: str) -> List[str]:
        """セル単位の判定: 関数のスペルミスや引数・範囲区切りの間違いを検証"""
        errors = []
        upper_val = val.upper().replace(" ", "")

        # 1. スペルミスの検出 (#NAME? または Typpo 疑い)
        if "#NAME?" in val or "AVARAGE" in upper_val or "SAM(" in upper_val:
            if "AVARAGE" in upper_val or "AVE(" in upper_val:
                errors.append(
                    f"セル {cell.coordinate}: ⚠️AVERAGE関数のスペルミス（AVARAGEやAVEなど）の疑いがあります。"
                )
            elif "SAM(" in upper_val:
                errors.append(
                    f"セル {cell.coordinate}: ⚠️SUM関数のスペルミス（SAM）の疑いがあります。"
                )
            else:
                errors.append(
                    f"セル {cell.coordinate}: ⚠️関数名が間違っているため、#NAME? エラーが発生しています。"
                )
            return errors

        # 2. 範囲指定のミス ( SUM(A1,A5) のようにコロンではなくカンマで区切っている )
        if (
            upper_val.startswith("=SUM(")
            or upper_val.startswith("=AVERAGE(")
        ) and ("," in upper_val and ":" not in upper_val):
            errors.append(
                f"セル {cell.coordinate}: ⚠️関数の範囲がコロン（:）ではなくカンマ（,）で区切られているため、2つのセルしか計算されていません。"
            )

        # 3. IF関数の引数不足 ( =IF(A1>=80, "合格") のように偽の場合が未設定 )
        if upper_val.startswith("=IF(") and upper_val.count(",") == 1:
            errors.append(
                f"セル {cell.coordinate}: ⚠️IF関数の引数が足りません。条件に合わない（偽の）場合の表示内容も設定しましょう。"
            )

        return errors

    def check_sheet(self) -> List[str]:
        """シート単位の判定: 指定された集計セルに正しい関数が入力されているか検証"""
        errors = []

        for cell_ref, (func_name, label) in self.REQUIRED_FUNCTIONS.items():
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
                    f"{func_name}関数を使って求めましょう。"
                )
            # 指定関数が使われていないチェック
            elif f"{func_name}(" not in cell_val:
                errors.append(
                    f"セル {cell_ref}（{label}）: ⚠️{func_name}関数が使われていません。"
                    f"=で始まる数式の中に {func_name}( ) を使って求めましょう。"
                )

        return errors