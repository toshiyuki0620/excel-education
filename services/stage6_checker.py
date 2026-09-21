from typing import List
from services.base_checker import BaseStageChecker


class Stage6Checker(BaseStageChecker):
    """ステージ6: 可視化（グラフ・チャート作成） 判定チェッカー"""

    def check_cell(self, cell, val: str) -> List[str]:
        """セル単位の判定: グラフステージのためセル個別チェックはなし"""
        return []

    def check_sheet(self) -> List[str]:
        """シート単位の判定: グラフオブジェクトの有無やタイトル・要素の設定を検証"""
        errors = []

        # 1. ワークシート内にグラフが存在するかチェック
        charts = self.ws._charts
        if not charts:
            return [
                "⚠️シート内にグラフが挿入されていません。"
                "対象の表を選択し、「挿入」タブから適切なグラフ（棒グラフ・折れ線グラフなど）を作成しましょう。"
            ]

        # 2. 挿入されたグラフの構成要素チェック
        for idx, chart in enumerate(charts, 1):
            # グラフタイトルの存在確認
            if not chart.title:
                errors.append(
                    f"⚠️グラフ #{idx}: グラフタイトルが設定されていないか、空になっています。"
                    "何を表しているグラフなのか分かるタイトルを入力しましょう。"
                )

            # 軸ラベルの確認（レーダーチャートや円グラフ以外の2D/3Dグラフ）
            if hasattr(chart, "x_axis") and hasattr(chart, "y_axis"):
                if not chart.y_axis.title:
                    errors.append(
                        f"⚠️グラフ #{idx}: 縦軸（Y軸）のラベル（単位や項目名）が設定されていません。"
                    )

        return errors