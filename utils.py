import datetime
import re
from typing import Optional
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet


def normalize_for_compare(val: Optional[str]) -> str:
    """比較用に文字列を正規化する関数。
    - 全角英数字・記号を半角に変換
    - 余分な空白（半角/全角スペース）の削除
    - 英字の大文字化
    """
    if val is None:
        return ""

    s = str(val).strip()

    # 全角英数字・記号を半角に変換
    zenkaku = "０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ＝＋－＊／（），％"
    hankaku = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZABCDEFGHIJKLMNOPQRSTUVWXYZ=-*/(),%"
    trans_table = str.maketrans(zenkaku, hankaku)
    s = s.translate(trans_table)

    # 空白スペース（全角・半角）を削除し、英字を大文字化
    s = s.replace(" ", "").replace(" ", "").upper()

    return s


def is_date_cell(cell: openpyxl.cell.cell.Cell) -> bool:
    """セルが日付型かどうかを判定する関数。
    Pythonの datetime オブジェクトか、Excelのナンバーフォーマットから判定。
    """
    # 1. 値自体が datetime / date オブジェクトの場合
    if isinstance(cell.value, (datetime.datetime, datetime.date)):
        return True

    # 2. openpyxl の is_date プロパティ判定
    if cell.is_date:
        return True

    # 3. 表示形式（number_format）による判定
    fmt = str(cell.number_format).lower() if cell.number_format else ""
    date_patterns = ["yyyy", "yy", "mm", "dd", "m/d", "yyyy/m/d", "yyyy-mm-dd"]

    return any(p in fmt for p in date_patterns)


def extract_target_sheet(
    wb: openpyxl.Workbook, stage_id: int
) -> Worksheet:
    """stage_id やシート構成に応じて、チェック対象の適切なワークシートを自動判定して抽出する。
    - 2枚構成（例: Stage4 で「マスタ」と「作業シート」がある場合）は「マスタ」以外のシートを取得
    - それ以外はアクティブシート（または先頭シート）を返却
    """
    sheets = wb.worksheets

    if not sheets:
        raise ValueError("Excelファイル内にワークシートが存在しません。")

    # ステージ4などで「マスタ」シートと「データ入力」シートに分かれている場合
    if stage_id == 4 and len(sheets) > 1:
        for sheet in sheets:
            if sheet.title != "マスタ":
                return sheet

    # 特定の指定がない場合はアクティブシートを返す
    return wb.active or sheets[0]


def has_pivot_table(ws: Worksheet) -> bool:
    """ワークシート内にピボットテーブルが存在するか判定する関数（Stage7 / 8等で使用）"""
    # openpyxlの _pivots プロパティを確認
    if hasattr(ws, "_pivots") and len(ws._pivots) > 0:
        return True
    return False


def is_formula_cell(cell: openpyxl.cell.cell.Cell) -> bool:
    """セルが数式セルかどうかを判定する関数（全角の『＝』から始まる誤入力も検出）"""
    val = str(cell.value) if cell.value is not None else ""
    val_str = val.strip()
    return val_str.startswith("=") or val_str.startswith("＝")