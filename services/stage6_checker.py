from typing import List
from openpyxl.cell.cell import MergedCell
from services.base_checker import BaseStageChecker


class Stage6Checker(BaseStageChecker):
    """ステージ6: データの可視化（グラフ作成） 判定チェッカー

    - ワークシート内へのグラフオブジェクト（ws._charts）の挿入チェック
    - グラフタイトルの設定有無チェック
    """

    def check_cell(self, cell, val: str) -> List[str]:
        """セル単位の判定ロジック（グラフ作成用セル入力等の個別チェック）"""
        errors = []

        # 結合セルはスキップ
        if isinstance(cell, MergedCell):
            return errors

        # 必要に応じて特定のセル入力をチェックする際はこちらに記述
        return errors

    def check_sheet(self) -> List[str]:
        """シート単位の判定ロジック（グラフ描画オブジェクトの確認）"""
        errors = []

        # openpyxl では _charts 属性内にシート内のグラフが格納されます
        charts = getattr(self.ws, "_charts", [])

        if not charts or len(charts) == 0:
            errors.append(
                "⚠️ワークシート内にグラフが見つかりませんでした。[挿入] ＞ [グラフ] からグラフを作成しましょう。"
            )
        else:
            # グラフが存在する場合の詳細チェック
            for idx, chart in enumerate(charts, 1):
                # グラフタイトルの有無を確認
                if not hasattr(chart, "title") or chart.title is None:
                    errors.append(
                        f"⚠️グラフ {idx}: グラフタイトルが設定されていません。グラフの目的が伝わるタイトルを設定しましょう。"
                    )

        return errors