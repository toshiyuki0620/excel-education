import os
import traceback
from flask import Flask, jsonify, redirect, render_template, request, url_for
import openpyxl

# 分離した設定・Supabaseクライアントの読み込み
import config
from config import supabase

# 分離したサービス・共通ユーティリティの読み込み
from services import get_checker
from utils import extract_target_sheet

app = Flask(__name__)

# 各ステージのタイトル定義
STAGE_TITLES = {
    0: "基本操作",
    1: "データ型",
    2: "基本関数",
    3: "参照の理解",
    4: "応用関数",
    5: "データ整形",
    6: "可視化",
    7: "分析基礎",
    8: "総合演習",
}


def should_check_cell(cell, target_cells: set) -> bool:
    """チェック対象セルかどうかを判定する。
    answer_keys に登録がある場合はそのセルのみ、未登録の場合は全セル対象。
    """
    if not target_cells:
        return True
    return cell.coordinate.upper() in target_cells


def run_common_cell_checks(cell, val: str) -> list[str]:
    """全ステージ共通のセル単位エラーチェック（エラー値や全角混入の判定）"""
    errors = []

    # 共通のエラー値チェック（#REF! など）
    if any(err in val for err in ["#REF!", "#VALUE!", "#NAME?", "#DIV/0!"]):
        if "#NAME?" not in val:
            errors.append(
                f"セル {cell.coordinate}: 数式エラー（{val}）が発生しています。"
            )
            return errors

    # 全角文字チェック（数式セルのみ対象）
    looks_like_formula = val.startswith("=") or val.startswith("＝")
    if looks_like_formula and ("＝" in val or "（" in val or "）" in val):
        errors.append(
            f"セル {cell.coordinate}: ⚠️数式に全角文字（＝やかっこ）が混入しています。"
        )

    return errors


def analyze_excel_file(file_stream, stage_id: int) -> list[str]:
    """Excelファイルを解析し、共通チェックおよび各ステージ固有チェッカーを実行する"""
    wb = openpyxl.load_workbook(file_stream, data_only=False)
    ws = extract_target_sheet(wb, stage_id)

    all_logs = []

    # --- Supabase からチェック対象セルの取得 ---
    try:
        ak_response = (
            supabase.table("answer_keys")
            .select("cell")
            .eq("stage_id", stage_id)
            .execute()
        )
        target_cells = {
            row["cell"].upper()
            for row in (ak_response.data or [])
            if row.get("cell")
        }
    except Exception:
        target_cells = set()

    # --- stage_id に応じたチェッカーインスタンスの取得 ---
    checker = get_checker(stage_id, ws, file_stream)

    # --- 1. セル単位のループ処理 ---
    for row in ws.iter_rows(values_only=False):
        for cell in row:
            val = str(cell.value) if cell.value else ""

            if not should_check_cell(cell, target_cells):
                continue

            # 共通チェックの実行
            all_logs.extend(run_common_cell_checks(cell, val))

            # ステージ個別チェッカーの実行（check_cell）
            if checker:
                all_logs.extend(checker.check_cell(cell, val))

    # --- 2. シート単位の判定処理（check_sheet） ---
    if checker:
        all_logs.extend(checker.check_sheet())

    return all_logs


# --- 受講生用：ファイル提出画面 ---
@app.route("/")
def upload_page():
    return render_template("upload.html")


# --- 受講生用：ファイル受け取り・判定処理 ---
@app.route("/upload_progress", methods=["POST"])
def upload_progress():
    student_id = request.form.get("student_id")
    raw_stage_id = request.form.get("stage_id")
    file = request.files.get("excel_file")

    if not file or not file.filename.endswith(".xlsx"):
        return (
            render_template(
                "result.html",
                success=False,
                student_id=None,
                stage_name=None,
                status=None,
                error_count=0,
                errors=[],
                info_logs=[],
                message="有効な.xlsxファイルをアップロードしてください。",
            ),
            400,
        )

    if not raw_stage_id or not raw_stage_id.isdigit():
        return (
            render_template(
                "result.html",
                success=False,
                message="stage_id が正しく指定されていません。",
            ),
            400,
        )

    stage_id = int(raw_stage_id)

    try:
        # 未登録受講生ならダミー作成
        try:
            supabase.table("students").insert(
                {"student_id": student_id, "name": "テスト受講生"}
            ).execute()
        except Exception:
            pass

        # 解析処理呼び出し
        all_logs = analyze_excel_file(file, stage_id)

        real_errors = [
            msg for msg in all_logs if "⚠️" in msg or "エラー" in msg
        ]
        info_logs = [msg for msg in all_logs if msg not in real_errors]
        error_count = len(real_errors)

        # ステータス決定規則
        if stage_id in (3, 6, 8):
            status = "目視レビュー待ち"
        else:
            status = "要確認" if error_count > 0 else "提出済"

        progress_data = {
            "student_id": student_id,
            "stage_id": stage_id,
            "status": status,
            "error_logs": {
                "detected_errors": all_logs,
                "error_count": error_count,
            },
        }
        supabase.table("progress_results").insert(progress_data).execute()

        return render_template(
            "result.html",
            success=True,
            student_id=student_id,
            stage_name=STAGE_TITLES.get(stage_id, f"ステージ{stage_id}"),
            status=status,
            error_count=error_count,
            errors=real_errors,
            info_logs=info_logs,
            message=None,
        )
    except Exception as e:
        print(f"================ ERROR: {type(e).__name__} ================")
        traceback.print_exc()
        print("========================================================")

        return (
            render_template(
                "result.html",
                success=False,
                student_id=student_id,
                stage_name=None,
                status=None,
                error_count=0,
                errors=[],
                info_logs=[],
                message=f"サーバーエラーが発生しました: {str(e)}",
            ),
            500,
        )


# --- 教員用：管理画面 ---
@app.route("/admin")
def admin_dashboard():
    response = (
        supabase.table("progress_results")
        .select("*")
        .order("updated_at", desc=True)
        .execute()
    )
    results = response.data or []
    for item in results:
        item["stage_name"] = STAGE_TITLES.get(
            item["stage_id"], f"ステージ{item['stage_id']}"
        )
        error_logs = item.get("error_logs", {}) or {}
        item["errors"] = error_logs.get("detected_errors", [])
        item["error_count"] = error_logs.get("error_count", 0)
    return render_template("admin.html", results=results)


# --- 教員用：評価更新 ---
@app.route("/admin/review/<record_id>", methods=["POST"])
def review_stage(record_id):
    new_status = request.form.get("status")
    teacher_comment = request.form.get("human_review")
    supabase.table("progress_results").update(
        {"status": new_status, "human_review": teacher_comment}
    ).eq("id", record_id).execute()
    return redirect(url_for("admin_dashboard"))


# --- 教員用：ステージ別 正解データ一覧 ---
@app.route("/admin/answer_keys")
def answer_keys_page():
    stage_id = int(request.args.get("stage_id", 0))
    response = (
        supabase.table("answer_keys")
        .select("*")
        .eq("stage_id", stage_id)
        .order("cell")
        .execute()
    )
    keys = response.data or []
    return render_template(
        "answer_keys.html",
        keys=keys,
        stage_id=stage_id,
        stage_titles=STAGE_TITLES,
    )


# --- 教員用：正解データの追加 ---
@app.route("/admin/answer_keys/add", methods=["POST"])
def answer_keys_add():
    stage_id = int(request.form.get("stage_id"))
    cell = request.form.get("cell", "").strip().upper()
    expected_value = request.form.get("expected_value", "").strip()
    expected_type = request.form.get("expected_type", "").strip()
    hint = request.form.get("hint", "").strip()

    if not cell or (not expected_value and not expected_type):
        return "セルに加えて、期待値または期待する型のどちらかは必須です", 400

    try:
        supabase.table("answer_keys").insert(
            {
                "stage_id": stage_id,
                "cell": cell,
                "expected_value": expected_value or None,
                "expected_type": expected_type or None,
                "hint": hint or None,
            }
        ).execute()
    except Exception as e:
        return (
            f"登録に失敗しました（同じセルが既に登録されている可能性があります）: {str(e)}",
            400,
        )

    return redirect(url_for("answer_keys_page", stage_id=stage_id))


# --- 教員用：正解データの削除 ---
@app.route("/admin/answer_keys/delete/<record_id>", methods=["POST"])
def answer_keys_delete(record_id):
    stage_id = request.form.get("stage_id", 0)
    supabase.table("answer_keys").delete().eq("id", record_id).execute()
    return redirect(url_for("answer_keys_page", stage_id=stage_id))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=config.PORT, debug=(config.FLASK_ENV == "development"))