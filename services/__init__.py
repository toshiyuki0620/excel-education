from typing import Optional, Type
from openpyxl.worksheet.worksheet import Worksheet
from services.base_checker import BaseStageChecker

# 各ステージチェッカーをインポート
from services.stage0_checker import Stage0Checker
from services.stage1_checker import Stage1Checker
from services.stage2_checker import Stage2Checker
from services.stage3_checker import Stage3Checker
from services.stage4_checker import Stage4Checker
from services.stage5_checker import Stage5Checker
from services.stage6_checker import Stage6Checker
from services.stage7_checker import Stage7Checker
from services.stage8_checker import Stage8Checker

# ステージIDとチェッカークラスのマッピング（レジストリ）
CHECKER_MAP: dict[int, Type[BaseStageChecker]] = {
    0: Stage0Checker,
    1: Stage1Checker,
    2: Stage2Checker,
    3: Stage3Checker,
    4: Stage4Checker,
    5: Stage5Checker,
    6: Stage6Checker,
    7: Stage7Checker,
    8: Stage8Checker,
}


def get_checker(
    stage_id: int, ws: Worksheet, file_stream
) -> Optional[BaseStageChecker]:
    """stage_id に対応するチェッカーのインスタンスを生成して返す。
    未定義のステージIDの場合は None を返す。
    """
    checker_cls = CHECKER_MAP.get(stage_id)
    if checker_cls:
        return checker_cls(ws, file_stream)
    return None