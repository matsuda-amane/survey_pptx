#!/usr/bin/env python3
"""
アンケート結果（CSV / XLSX）をアップロードして PPTX を生成する Web アプリ（Streamlit）
社内利用を想定。ログインは環境変数 STREAMLIT_PASSWORD で設定したパスワードで行う。

起動: STREAMLIT_PASSWORD=your_password streamlit run app.py
"""

import io
import os
import sys
import tempfile
from pathlib import Path

import streamlit as st

# 同フォルダの survey_to_pptx を利用
sys.path.insert(0, str(Path(__file__).resolve().parent))
from survey_to_pptx import convert, load_data, classify_columns

# テンプレートはプロジェクト内の template/design.pptx
DEFAULT_TEMPLATE = Path(__file__).resolve().parent / "template" / "design.pptx"

st.set_page_config(
    page_title="アンケート → PPTX 変換",
    page_icon="📊",
    layout="centered",
)

# ── ログイン ─────────────────────────────────────────────────────────────────
def _get_expected_password() -> str:
    """環境変数で設定されたログインパスワード（未設定の場合は空＝ログイン不可）"""
    return (os.environ.get("STREAMLIT_PASSWORD") or "").strip()


def _check_login():
    if st.session_state.get("logged_in"):
        return
    expected = _get_expected_password()
    if not expected:
        st.error("ログイン機能が有効になっていません。管理者に連絡し、環境変数 **STREAMLIT_PASSWORD** を設定してもらってください。")
        st.stop()
    st.subheader("ログイン")
    pw = st.text_input("パスワード", type="password", key="login_password")
    if st.button("ログイン"):
        if pw == expected:
            st.session_state.logged_in = True
            st.rerun()
        else:
            st.error("パスワードが正しくありません。")
    st.stop()


_check_login()

# ログアウト（サイドバー）
with st.sidebar:
    if st.button("🚪 ログアウト"):
        st.session_state.logged_in = False
        st.rerun()

# ── メイン ───────────────────────────────────────────────────────────────────
st.title("📊 アンケート結果 → PPTX 変換")
st.caption("CSV または XLSX をアップロードし、出力に含める列を選んでから PowerPoint を生成できます。（社内利用）")

uploaded = st.file_uploader(
    "アンケートファイルを選択",
    type=["csv", "xlsx"],
    help="CSV または Excel（.xlsx）を選択してください",
)

if uploaded is not None:
    suffix = ".xlsx" if uploaded.name.lower().endswith(".xlsx") else ".csv"
    is_xlsx = suffix == ".xlsx"

    # アップロード内容を一時ファイルに保存して読み込み
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(uploaded.getvalue())
        tmp_path = tmp.name
    try:
        encoding = st.selectbox(
            "CSV の文字コード（CSV の場合）",
            ["utf-8", "utf-8-sig", "shift_jis", "cp932"],
            index=0,
        )
        df = load_data(tmp_path, encoding=encoding)
        col_types = classify_columns(df)
    finally:
        Path(tmp_path).unlink(missing_ok=True)

    # ファイル内容のプレビュー（ビューワー）
    st.subheader("ファイルの中身")
    preview_rows = min(50, len(df))
    st.dataframe(
        df.head(preview_rows),
        use_container_width=True,
        height=min(400, 35 * preview_rows + 38),
    )
    if len(df) > preview_rows:
        st.caption(f"先頭 {preview_rows} 件を表示（全 {len(df)} 件）")

    # 列タイプのラベル（表示用）
    type_labels = {
        "personal": "⛔ 個人情報",
        "metadata": "ℹ️ メタデータ",
        "categorical": "📊 選択式",
        "high_cardinality": "📊 選択式（多め）",
        "text": "📝 自由記述",
        "empty": "— 空",
        "appendix_company": "📋 Appendix（会社名）",
    }

    st.subheader("出力に含める列を選択")
    st.caption("PPTX に含めたい列だけを選んでください。氏名・連絡先などは含めないこともできます。")

    exclude_personal = st.checkbox(
        "個人情報として検出された列を初期選択から除外する",
        value=True,
        help="氏名・メール・電話番号・ID などは通常 PPTX に含めません",
    )
    if exclude_personal:
        default_cols = [c for c in df.columns if col_types.get(c) != "personal"]
    else:
        default_cols = list(df.columns)

    selected_columns = st.multiselect(
        "列を選択（選択した列だけが PPTX に出力されます）",
        options=list(df.columns),
        default=default_cols if default_cols else list(df.columns),
        format_func=lambda c: f"{c}  —  {type_labels.get(col_types.get(c, ''), col_types.get(c, '?'))}",
    )

    if not selected_columns:
        st.warning("少なくとも1列を選択してください。")

    col1, col2 = st.columns(2)
    with col1:
        title = st.text_input("タイトル（任意）", placeholder="省略時はファイル名または Survey Name 列を使用")
    with col2:
        subtitle = st.text_input("サブタイトル（任意）", placeholder="例: 2026年Q1")

    run = st.button("PPTX を生成", type="primary", disabled=not selected_columns)

    if run and selected_columns:
        with st.spinner("PPTX を生成しています…"):
            try:
                with tempfile.TemporaryDirectory() as tmp:
                    # 選択した列だけのデータで一時ファイルを作成
                    input_path = Path(tmp) / f"input{suffix}"
                    out_df = df[[c for c in selected_columns if c in df.columns]]
                    if is_xlsx:
                        out_df.to_excel(input_path, index=False)
                    else:
                        out_df.to_csv(input_path, index=False, encoding=encoding)

                    output_path = Path(tmp) / "アンケート結果.pptx"
                    template_path = str(DEFAULT_TEMPLATE) if DEFAULT_TEMPLATE.exists() else None

                    buf = io.StringIO()
                    old_stdout = sys.stdout
                    sys.stdout = buf
                    try:
                        convert(
                            str(input_path),
                            str(output_path),
                            title=title.strip(),
                            subtitle=subtitle.strip(),
                            encoding=encoding,
                            template_path=template_path,
                        )
                    finally:
                        sys.stdout = old_stdout
                    log = buf.getvalue()

                    if not output_path.exists():
                        st.error("PPTX の生成に失敗しました。")
                        st.code(log)
                    else:
                        st.success("PPTX を生成しました。下のボタンからダウンロードしてください。")
                        st.download_button(
                            label="📥 アンケート結果.pptx をダウンロード",
                            data=output_path.read_bytes(),
                            file_name="アンケート結果.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        )
                        with st.expander("変換ログ"):
                            st.text(log)
            except Exception as e:
                st.error(f"エラー: {e}")
                st.exception(e)
else:
    st.info("👆 CSV または XLSX ファイルをアップロードしてください。")
