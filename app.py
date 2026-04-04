#!/usr/bin/env python3
"""
アンケート結果（CSV / XLSX）をアップロードして PPTX を生成する Web アプリ（Streamlit）
社内利用を想定。ログインは環境変数 STREAMLIT_PASSWORD で設定したパスワードで行う。

起動: STREAMLIT_PASSWORD=your_password streamlit run app.py
"""

import io
import json
import os
import re
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

import streamlit as st

# 同フォルダの survey_to_pptx を利用
sys.path.insert(0, str(Path(__file__).resolve().parent))
from survey_to_pptx import convert, load_data, classify_columns, list_builtin_templates

_TEMPLATE_ROOT = Path(__file__).resolve().parent / "template"
_DEFAULT_TEMPLATE = _TEMPLATE_ROOT / "template_ligare.pptx"

st.set_page_config(
    page_title="アンケート → PPTX 変換",
    page_icon="📊",
    layout="centered",
)

# ── セッション初期化 ─────────────────────────────────────────────────────────
def _init_batch_session():
    if "column_presets" not in st.session_state:
        st.session_state.column_presets = {}
    if "audit_log" not in st.session_state:
        st.session_state.audit_log = []


_init_batch_session()


def _safe_pptx_filename(name: str, fallback: str) -> str:
    base = (name or "").strip() or fallback
    base = re.sub(r"[^\w\-_.\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FFF]+", "_", base)
    base = base.strip("._") or fallback
    if not base.lower().endswith(".pptx"):
        base = f"{base[:80]}.pptx"
    return base


def _merged_sponsor_columns(standard_cols: list[str], sponsor_extra: list[str]) -> list[str]:
    seen = set(standard_cols)
    out = list(standard_cols)
    for c in sponsor_extra:
        if c not in seen:
            out.append(c)
            seen.add(c)
    return out


def _columns_for_profile(
    profile: str, standard_cols: list[str], sponsor_extra: list[str]
) -> list[str]:
    if profile == "sponsor":
        return _merged_sponsor_columns(standard_cols, sponsor_extra)
    return list(standard_cols)


def _apply_row_filter(
    df,
    use_filter: bool,
    filter_col: str | None,
    filter_values: list,
):
    if not use_filter or not filter_col or not filter_values:
        return df
    if filter_col not in df.columns:
        return df
    chosen = {str(v).strip() for v in filter_values}
    s = df[filter_col].dropna().astype(str).str.strip()
    mask = df[filter_col].astype(str).str.strip().isin(chosen)
    return df.loc[mask]


def _run_convert_captured(
    out_path: Path,
    sub_df,
    title: str,
    subtitle: str,
    encoding: str,
    template_path: str | None,
) -> str:
    buf_tr = io.StringIO()
    old_stdout = sys.stdout
    sys.stdout = buf_tr
    try:
        convert(
            None,
            str(out_path),
            title=title.strip(),
            subtitle=subtitle.strip(),
            encoding=encoding,
            template_path=template_path,
            df=sub_df,
        )
    finally:
        sys.stdout = old_stdout
    return buf_tr.getvalue()


# ── ログイン ─────────────────────────────────────────────────────────────────
def _get_expected_password() -> str:
    return (os.environ.get("STREAMLIT_PASSWORD") or "").strip()


def _check_login():
    if st.session_state.get("logged_in"):
        return
    expected = _get_expected_password()
    if not expected:
        st.error(
            "ログイン機能が有効になっていません。管理者に連絡し、環境変数 **STREAMLIT_PASSWORD** を設定してもらってください。"
        )
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
st.caption(
    "CSV または XLSX をアップロードし、列を選んで PowerPoint を生成します。登壇者別レポートでは協賛向けに追加列を開示できます。（社内利用）"
)

uploaded = st.file_uploader(
    "アンケートファイルを選択",
    type=["csv", "xlsx"],
    help="CSV または Excel（.xlsx）を選択してください",
)

if uploaded is not None:
    suffix = ".xlsx" if uploaded.name.lower().endswith(".xlsx") else ".csv"
    is_xlsx = suffix == ".xlsx"

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

    st.subheader("ファイルの中身")
    preview_rows = min(50, len(df))
    st.dataframe(
        df.head(preview_rows),
        use_container_width=True,
        height=min(400, 35 * preview_rows + 38),
    )
    if len(df) > preview_rows:
        st.caption(f"先頭 {preview_rows} 件を表示（全 {len(df)} 件）")

    type_labels = {
        "personal": "⛔ 個人情報",
        "metadata": "ℹ️ メタデータ",
        "categorical": "📊 選択式",
        "high_cardinality": "📊 選択式（多め）",
        "text": "📝 自由記述",
        "empty": "— 空",
        "appendix_company": "📋 Appendix（会社名）",
    }

    exclude_personal = st.checkbox(
        "個人情報として検出された列を初期選択から除外する",
        value=True,
        help="氏名・メール・電話番号・ID などは通常 PPTX に含めません",
    )
    if exclude_personal:
        default_cols = [c for c in df.columns if col_types.get(c) != "personal"]
    else:
        default_cols = list(df.columns)

    output_mode = st.radio(
        "出力モード",
        ["単一レポート（1 本の PPTX）", "登壇者別レポート（ZIP・複数本）"],
        horizontal=True,
    )

    _builtins = list_builtin_templates()
    if _builtins:
        _labels = [x[0] for x in _builtins]
        _tmpl_choice = st.selectbox(
            "PPTX のフォーマット（テンプレート）",
            _labels,
            help="見た目・背景は template 内ので切り替えます。",
        )
        chosen_template_path = dict(_builtins)[_tmpl_choice]
    else:
        chosen_template_path = _DEFAULT_TEMPLATE if _DEFAULT_TEMPLATE.exists() else None

    template_path_str = (
        str(chosen_template_path)
        if chosen_template_path and chosen_template_path.exists()
        else None
    )

    if output_mode.startswith("単一"):
        st.subheader("出力に含める列")
        st.caption("PPTX に含めたい列だけを選んでください。")

        selected_columns = st.multiselect(
            "列を選択",
            options=list(df.columns),
            default=default_cols if default_cols else list(df.columns),
            format_func=lambda c: f"{c}  —  {type_labels.get(col_types.get(c, ''), col_types.get(c, '?'))}",
        )

        if not selected_columns:
            st.warning("少なくとも1列を選択してください。")

        col1, col2 = st.columns(2)
        with col1:
            title = st.text_input(
                "タイトル（任意）", placeholder="省略時は Survey Name 列またはファイル名"
            )
        with col2:
            subtitle = st.text_input("サブタイトル（任意）", placeholder="例: 2026年Q1")

        run = st.button("PPTX を生成", type="primary", disabled=not selected_columns)

        if run and selected_columns:
            with st.spinner("PPTX を生成しています…"):
                try:
                    with tempfile.TemporaryDirectory() as tmp:
                        input_path = Path(tmp) / f"input{suffix}"
                        out_df = df[[c for c in selected_columns if c in df.columns]]
                        if is_xlsx:
                            out_df.to_excel(input_path, index=False)
                        else:
                            out_df.to_csv(input_path, index=False, encoding=encoding)

                        output_path = Path(tmp) / "アンケート結果.pptx"

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
                                template_path=template_path_str,
                            )
                        finally:
                            sys.stdout = old_stdout
                        log = buf.getvalue()

                        if not output_path.exists():
                            st.error("PPTX の生成に失敗しました。")
                            st.code(log)
                        else:
                            st.success("PPTX を生成しました。")
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
        st.info(
            "**データの切り分け**: 1 ファイルに全会場・全登壇が入っているときは、各登壇者で「行フィルタ」を有効にし、"
            "セッション識別列と値（複数可）を選びます。**フィルタなし**は全行を対象にします（全員同一データに列だけ変える用途や、"
            "事前に別ツールで行を絞った CSV を載せる運用向け）。登壇者ごとに**別ファイル**だけを渡す場合は、この画面では単一モードを"
            "登壇者回数分ご利用ください（将来、複数アップロードに拡張可能です）。"
        )

        st.subheader("開示プロファイル別の列")
        cols_std = st.multiselect(
            "標準（一般登壇）向けに含める列",
            options=list(df.columns),
            default=default_cols if default_cols else list(df.columns),
            format_func=lambda c: f"{c}  —  {type_labels.get(col_types.get(c, ''), col_types.get(c, '?'))}",
            key="batch_cols_standard",
        )
        extra_default = [
            c
            for c in df.columns
            if col_types.get(c) in ("appendix_company", "text") and c not in cols_std
        ]
        cols_sponsor_extra = st.multiselect(
            "協賛向けに**追加で**含める列（標準の列に上乗せ）",
            options=[c for c in df.columns if c not in cols_std],
            default=[c for c in extra_default if c not in cols_std],
            format_func=lambda c: f"{c}  —  {type_labels.get(col_types.get(c, ''), col_types.get(c, '?'))}",
            key="batch_cols_sponsor_extra",
        )

        with st.expander("協賛で追加開示される列（差分プレビュー）"):
            if cols_sponsor_extra:
                st.write(sorted(cols_sponsor_extra))
            else:
                st.caption("協賛専用の追加列はありません（標準と同一の開示）。")

        with st.expander("列プリセット（名前付きで保存・ファイルで共有）"):
            preset_name = st.text_input("プリセット名", key="preset_name_input")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("現在の列セットを保存"):
                    if not preset_name.strip():
                        st.error("プリセット名を入力してください。")
                    else:
                        st.session_state.column_presets[preset_name.strip()] = {
                            "standard": list(cols_std),
                            "sponsor_extra": list(cols_sponsor_extra),
                        }
                        st.success(f"保存しました: {preset_name.strip()}")
            with c2:
                names = list(st.session_state.column_presets.keys())
                load_pick = st.selectbox("読み込むプリセット", [""] + names, key="preset_load_pick")
                if load_pick and st.button("このプリセットを列に適用"):
                    p = st.session_state.column_presets[load_pick]
                    valid = set(df.columns)
                    st.session_state["batch_cols_standard"] = [
                        c for c in p.get("standard", []) if c in valid
                    ]
                    st.session_state["batch_cols_sponsor_extra"] = [
                        c for c in p.get("sponsor_extra", []) if c in valid
                    ]
                    st.rerun()

            presets_blob = json.dumps(
                st.session_state.column_presets, ensure_ascii=False, indent=2
            )
            st.download_button(
                "プリセット JSON をダウンロード",
                data=presets_blob.encode("utf-8"),
                file_name="survey_pptx_column_presets.json",
                mime="application/json",
            )
            up = st.file_uploader("プリセット JSON を読み込み", type=["json"], key="preset_upload")
            if up is not None:
                try:
                    data = json.loads(up.getvalue().decode("utf-8"))
                    if isinstance(data, dict):
                        st.session_state.column_presets.update(data)
                        st.success("セッションにマージしました。読み込むプリセットから選べます。")
                    else:
                        st.error("JSON はオブジェクト（名前→列リスト）形式である必要があります。")
                except json.JSONDecodeError as e:
                    st.error(f"JSON の解析に失敗しました: {e}")

        if not cols_std:
            st.warning("標準向けには少なくとも 1 列選んでください。")

        n_speakers = st.number_input("登壇者数", min_value=1, max_value=8, value=3, step=1)

        speaker_defs = []
        for i in range(int(n_speakers)):
            with st.expander(f"登壇者 {i + 1}", expanded=(i == 0)):
                sp_name = st.text_input("表示名・ファイル名のもと", value=f"登壇者{i + 1}", key=f"sp_{i}_label")
                profile = st.selectbox(
                    "開示プロファイル",
                    ["standard", "sponsor"],
                    format_func=lambda x: "一般登壇（標準）" if x == "standard" else "協賛（標準＋追加列）",
                    key=f"sp_{i}_profile",
                )
                t1, t2 = st.columns(2)
                with t1:
                    sp_title = st.text_input("スライドタイトル（任意）", key=f"sp_{i}_title", value="")
                with t2:
                    sp_sub = st.text_input("スライドサブタイトル（任意）", key=f"sp_{i}_sub", value="")

                use_f = st.checkbox(
                    "行フィルタを使う（セッションで絞り込む）",
                    value=False,
                    key=f"sp_{i}_use_filter",
                )
                filter_col = None
                filter_vals = []
                if use_f:
                    filter_col = st.selectbox(
                        "フィルタ列",
                        [""] + list(df.columns),
                        key=f"sp_{i}_fcol",
                    )
                    if filter_col and filter_col in df.columns:
                        uniques = sorted(
                            df[filter_col].dropna().astype(str).str.strip().unique().tolist(),
                            key=str,
                        )
                        if len(uniques) > 200:
                            st.caption(f"選択肢が多いです（{len(uniques)} 件）。先頭 200 件のみ一覧します。")
                            uniques = uniques[:200]
                        filter_vals = st.multiselect(
                            "含める値（複数可）",
                            options=uniques,
                            key=f"sp_{i}_fval",
                        )

                out_fn = st.text_input(
                    "出力ファイル名",
                    value=_safe_pptx_filename(sp_name, f"speaker_{i + 1}"),
                    key=f"sp_{i}_file",
                )

                speaker_defs.append(
                    {
                        "index": i,
                        "label": sp_name,
                        "profile": profile,
                        "title": sp_title,
                        "subtitle": sp_sub,
                        "use_filter": use_f,
                        "filter_col": filter_col if filter_col else None,
                        "filter_values": filter_vals,
                        "out_name": out_fn,
                    }
                )

        batch_run = st.button(
            "登壇者分を一括生成（ZIP）",
            type="primary",
            disabled=not cols_std,
        )

        with st.expander("監査ログ（このブラウザセッション）"):
            if not st.session_state.audit_log:
                st.caption("まだ記録がありません。ZIP 一括生成すると追記されます。")
            else:
                for entry in reversed(st.session_state.audit_log[-30:]):
                    st.json(entry)
            if st.button("ログをクリア（セッション）"):
                st.session_state.audit_log = []
                st.rerun()

        if batch_run and cols_std:
            zip_buf = io.BytesIO()
            results = []
            combined_log = []

            with st.spinner("各登壇者分の PPTX を生成しています…"):
                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for sp in speaker_defs:
                        tag = sp["label"] or f"speaker{sp['index'] + 1}"
                        try:
                            sub = _apply_row_filter(
                                df,
                                sp["use_filter"],
                                sp["filter_col"],
                                sp["filter_values"],
                            )
                            if len(sub) == 0:
                                results.append((tag, False, "フィルタ後が 0 件です"))
                                combined_log.append(f"=== {tag} SKIP: 0 rows ===\n")
                                continue

                            use_cols = _columns_for_profile(
                                sp["profile"], cols_std, cols_sponsor_extra
                            )
                            use_cols = [c for c in use_cols if c in sub.columns]
                            if not use_cols:
                                results.append((tag, False, "選択列がデータにありません"))
                                continue

                            sub_out = sub[use_cols]

                            with tempfile.TemporaryDirectory() as tmpd:
                                out_p = Path(tmpd) / "out.pptx"
                                log_piece = _run_convert_captured(
                                    out_p,
                                    sub_out,
                                    sp["title"],
                                    sp["subtitle"],
                                    encoding,
                                    template_path_str,
                                )
                                combined_log.append(f"=== {tag} ===\n{log_piece}\n")

                                if not out_p.exists():
                                    results.append((tag, False, "PPTX が生成されませんでした"))
                                    continue

                                name_in_zip = _safe_pptx_filename(
                                    sp["out_name"].replace(".pptx", ""),
                                    f"report_{sp['index'] + 1}",
                                )
                                zf.writestr(name_in_zip, out_p.read_bytes())
                                results.append((tag, True, f"{len(sub)} 件 / {len(use_cols)} 列"))

                        except Exception as ex:
                            results.append((tag, False, str(ex)))
                            combined_log.append(f"=== {tag} ERROR: {ex} ===\n")

            any_ok = any(ok for _, ok, _ in results)
            for tag, ok, msg in results:
                if ok:
                    st.success(f"{tag}: OK（{msg}）")
                else:
                    st.error(f"{tag}: {msg}")

            if any_ok:
                st.download_button(
                    label="📦 ZIP をダウンロード（全員分まとめ）",
                    data=zip_buf.getvalue(),
                    file_name=f"seminar_reports_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M')}.zip",
                    mime="application/zip",
                )
                entry = {
                    "utc": datetime.now(timezone.utc).isoformat(),
                    "speakers": [
                        {
                            "label": sp["label"],
                            "profile": sp["profile"],
                            "rows_after_filter": len(
                                _apply_row_filter(
                                    df,
                                    sp["use_filter"],
                                    sp["filter_col"],
                                    sp["filter_values"],
                                )
                            ),
                            "ok": ok,
                            "message": msg,
                        }
                        for sp, (_, ok, msg) in zip(speaker_defs, results)
                    ],
                    "standard_columns_n": len(cols_std),
                    "sponsor_extra_columns_n": len(cols_sponsor_extra),
                }
                st.session_state.audit_log.append(entry)

            with st.expander("一括変換ログ"):
                st.text("".join(combined_log))

else:
    st.info("👆 CSV または XLSX ファイルをアップロードしてください。")
