#!/usr/bin/env python3
"""
アンケート結果（CSV / XLSX）をアップロードして PPTX を生成する Web アプリ（Streamlit）
社内利用を想定。ログインは環境変数 STREAMLIT_PASSWORD で設定したパスワードで行う。

起動: STREAMLIT_PASSWORD=your_password streamlit run app.py
"""

import hashlib
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

sys.path.insert(0, str(Path(__file__).resolve().parent))
from survey_to_pptx import (
    COLUMN_LAYOUT_AUTO,
    COLUMN_LAYOUT_CHOICES_UI,
    TEMPLATE_DIR,
    convert,
    load_data,
    classify_columns,
    list_builtin_templates,
)

_DEFAULT_TEMPLATE = TEMPLATE_DIR / "template_ligare.pptx"

st.set_page_config(
    page_title="アンケート → PPTX 変換",
    page_icon="📊",
    layout="centered",
)


def _init_session():
    if "column_presets" not in st.session_state:
        st.session_state.column_presets = {}
    if "audit_log" not in st.session_state:
        st.session_state.audit_log = []
    if "wizard_loaded" not in st.session_state:
        st.session_state.wizard_loaded = False
    if "wizard_upload_sig" not in st.session_state:
        st.session_state.wizard_upload_sig = None


_init_session()


def _col_widget_hash(col: str) -> str:
    return hashlib.sha256(str(col).encode("utf-8")).hexdigest()[:14]


def _safe_pptx_filename(name: str, fallback: str) -> str:
    base = (name or "").strip() or fallback
    base = re.sub(r"[^\w\-_.\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FFF]+", "_", base)
    base = base.strip("._") or fallback
    if not base.lower().endswith(".pptx"):
        base = f"{base[:80]}.pptx"
    return base


def _apply_row_filter(df, use_filter: bool, filter_col: str | None, filter_values: list):
    if not use_filter or not filter_col or not filter_values:
        return df
    if filter_col not in df.columns:
        return df
    chosen = {str(v).strip() for v in filter_values}
    mask = df[filter_col].astype(str).str.strip().isin(chosen)
    return df.loc[mask]


def _run_convert_captured(
    out_path: Path,
    sub_df,
    title: str,
    subtitle: str,
    encoding: str,
    template_path: str | None,
    column_layout: dict[str, str],
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
            column_layout=column_layout or None,
        )
    finally:
        sys.stdout = old_stdout
    return buf_tr.getvalue()


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
    can_login = bool((pw or "").strip())
    if st.button("ログイン", disabled=not can_login):
        if pw == expected:
            st.session_state.logged_in = True
            st.rerun()
        else:
            st.error("パスワードが正しくありません。")
    st.stop()


_check_login()

with st.sidebar:
    if st.button("🚪 ログアウト"):
        st.session_state.logged_in = False
        st.rerun()

st.title("📊 アンケート結果 → PPTX 変換")
st.caption(
    "アップロード後にテンプレートと出力パターンを決めてから読み込み、列とレイアウトを設定して PowerPoint を生成します。"
)

# ── ステップ0: アップロード + 初期設定（読み込み前は中身を表示しない） ───────
uploaded = st.file_uploader(
    "1. アンケートファイルを選択（CSV / XLSX）",
    type=["csv", "xlsx"],
    help="この時点ではファイル内容は読み込みません。下の設定のあと「CSV を読み込む」を押してください。",
)

if uploaded is not None:
    sig = f"{uploaded.name}:{uploaded.size}"
    if st.session_state.wizard_upload_sig != sig:
        st.session_state.wizard_upload_sig = sig
        st.session_state.wizard_loaded = False

if not st.session_state.wizard_loaded:
    builtins = list_builtin_templates()
    if builtins:
        tmpl_labels = [x[0] for x in builtins]
        tmpl_pick = st.selectbox(
            "2. テンプレート（PowerPoint の見た目）",
            tmpl_labels,
            help="Ligare / Amane など、2 種類から選択します。",
        )
        chosen_template_path = dict(builtins)[tmpl_pick]
    else:
        chosen_template_path = _DEFAULT_TEMPLATE if _DEFAULT_TEMPLATE.exists() else None

    output_scope = st.radio(
        "3. 出力する資料のパターン数",
        [
            "1 種類だけ（1 本の PowerPoint でよい）",
            "複数パターン（内容違いの資料を最大 8 本まで）",
        ],
        help="まず「1 本で足りるか／役割・登壇者などで分けたいか」を選びます。複数のときだけ、次の項目で本数を指定します。",
    )
    if output_scope.startswith("複数"):
        num_patterns = st.number_input(
            "　└ パターン数（半角・2〜8）",
            min_value=2,
            max_value=8,
            value=2,
            step=1,
            help="例: 一般向け・協賛向けの 2 パターン、登壇者 A/B 向けの別表現など。",
        )
    else:
        num_patterns = 1

    delivery_mode = st.radio(
        "4. 納品・配布の形",
        [
            "パターン別：パターン数ぶんのファイルを出す（複数あるときは ZIP 1 個）",
            "登壇者別：登壇者ごとに行を絞り、指定パターンの資料を ZIP でまとめる",
        ],
        help="パターン別＝「作ったパターンぶんのファイル」を受け取るイメージ。登壇者別＝「人ごとに別ファイルが ZIP にはいる」イメージです。",
    )

    encoding = st.selectbox(
        "CSV の文字コード（CSV を読み込むときに使用）",
        ["utf-8", "utf-8-sig", "shift_jis", "cp932"],
        index=0,
    )

    template_path_str = (
        str(chosen_template_path)
        if chosen_template_path and chosen_template_path.exists()
        else None
    )

    can_load = uploaded is not None and template_path_str is not None
    if st.button("CSV / XLSX を読み込む", type="primary", disabled=not can_load):
        st.session_state.wizard_loaded = True
        st.session_state.upload_bytes = uploaded.getvalue()
        st.session_state.upload_name = uploaded.name
        st.session_state.upload_suffix = ".xlsx" if uploaded.name.lower().endswith(".xlsx") else ".csv"
        st.session_state.wizard_encoding = encoding
        st.session_state.wizard_num_patterns = int(num_patterns)
        st.session_state.wizard_delivery_mode = delivery_mode
        st.session_state.wizard_template = template_path_str
        st.rerun()

    st.info(
        "ファイルを選び、テンプレート・出力パターン・納品の形を指定してから **CSV / XLSX を読み込む** を押してください。"
        "（読み込み前はプレビューしません。）"
    )
    st.stop()

template_path_str = st.session_state.wizard_template

# ── ステップ1: メモリ上にデータ読込・プレビュー ─────────────────────────────
raw = st.session_state.upload_bytes
suffix = st.session_state.upload_suffix
enc = st.session_state.wizard_encoding

with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
    tmp.write(raw)
    tmp_path = tmp.name
try:
    df = load_data(tmp_path, encoding=enc)
    col_types = classify_columns(df)
finally:
    Path(tmp_path).unlink(missing_ok=True)

type_labels = {
    "personal": "⛔ 個人情報",
    "metadata": "ℹ️ メタデータ",
    "categorical": "📊 選択式",
    "high_cardinality": "📊 選択式（多め）",
    "text": "📝 自由記述",
    "empty": "— 空",
    "appendix_company": "📋 Appendix（会社名）",
}

def _fmt_col_option(c: str) -> str:
    return f"{c}  —  {type_labels.get(col_types.get(c, ''), col_types.get(c, '?'))}"


def _column_checkbox_grid(df, pi: int, default_include: set[str]) -> list:
    """プレビューと同じ列順でチェックを並べ、選択済み列名のリストを返す。"""
    st.caption("下のチェックは表の列の並びに対応しています。スライドに載せる列だけオンにしてください。")
    cols = list(df.columns)
    chunk = 3
    for i in range(0, len(cols), chunk):
        part = cols[i : i + chunk]
        row = st.columns(len(part))
        for j, col in enumerate(part):
            h = _col_widget_hash(col)
            key = f"p{pi}_chk_{h}"
            with row[j]:
                st.checkbox(
                    col,
                    value=col in default_include,
                    key=key,
                    help=_fmt_col_option(col),
                )
    return [
        c
        for c in cols
        if st.session_state.get(f"p{pi}_chk_{_col_widget_hash(c)}", c in default_include)
    ]


st.success(
    f"読み込み済み: **{st.session_state.upload_name}**（{len(df)} 行 / {len(df.columns)} 列）"
)
st.caption(
    f"読み込み時の設定: パターン数 **{st.session_state.wizard_num_patterns}** ／ **{st.session_state.wizard_delivery_mode}**。"
    "テンプレート・パターン数・納品の形を変える場合は、下の「設定をやり直す」から最初の手順に戻ってください。"
)
if st.button("設定をやり直す（読み込み前に戻る）"):
    st.session_state.wizard_loaded = False
    st.session_state.upload_bytes = None
    st.rerun()

st.subheader("データプレビュー")
preview_rows = min(50, len(df))
st.dataframe(
    df.head(preview_rows),
    use_container_width=True,
    height=min(400, 35 * preview_rows + 38),
)
if len(df) > preview_rows:
    st.caption(f"先頭 {preview_rows} 件を表示（全 {len(df)} 件）")

exclude_personal = st.checkbox(
    "個人情報列を列リストの初期候補から外す",
    value=True,
    help="チェックを外すと個人情報列も選べます（通常はオフのまま推奨）。",
)
if exclude_personal:
    default_cols = [c for c in df.columns if col_types.get(c) != "personal"]
else:
    default_cols = list(df.columns)

layout_keys = [x[0] for x in COLUMN_LAYOUT_CHOICES_UI]
layout_labels = [x[1] for x in COLUMN_LAYOUT_CHOICES_UI]
layout_label_to_key = dict(zip(layout_labels, layout_keys))

npat = int(st.session_state.wizard_num_patterns)
delivery = st.session_state.wizard_delivery_mode

# ── 各パターン: 列 multiselect + 列ごとのレイアウト ─────────────────────────
pattern_specs = []

if npat == 1:
    st.subheader("パターン 1 の内容")
    pn = st.text_input("パターン名（メモ用）", value="レポート", key="p0_name")
    pfn = st.text_input(
        "出力ファイル名（.pptx の前の部分）",
        value=_safe_pptx_filename(pn, "report").replace(".pptx", ""),
        key="p0_fn",
    )
    st.markdown("**スライドに含める列**")
    sel = _column_checkbox_grid(
        df,
        0,
        set(default_cols if default_cols else list(df.columns)),
    )
    layouts_map = {}
    if sel:
        st.markdown("**列ごとのレイアウト**")
        for col in sel:
            h = _col_widget_hash(col)
            opt_label = st.selectbox(
                f"`{col}`",
                layout_labels,
                index=0,
                key=f"p0_lay_{h}",
                label_visibility="visible",
            )
            layouts_map[col] = layout_label_to_key[opt_label]
    pattern_specs.append(
        {
            "name": pn,
            "title": "",
            "subtitle": "",
            "filename": _safe_pptx_filename(pfn, "report"),
            "columns": sel,
            "layouts": layouts_map,
        }
    )
else:
    st.subheader("パターンごとの内容")
    tabs = st.tabs([f"パターン {i + 1}" for i in range(npat)])
    for pi, tab in enumerate(tabs):
        with tab:
            pn = st.text_input("パターン名", value=f"パターン{pi + 1}", key=f"p{pi}_name")
            pfn = st.text_input(
                "出力ファイル名（.pptx の前の部分）",
                value=_safe_pptx_filename(pn, f"report_{pi + 1}").replace(".pptx", ""),
                key=f"p{pi}_fn",
            )
            st.markdown("**スライドに含める列**")
            sel = _column_checkbox_grid(
                df,
                pi,
                set(default_cols)
                if (default_cols and pi == 0)
                else set(),
            )
            layouts_map = {}
            if sel:
                st.markdown("**列ごとのレイアウト**")
                for col in sel:
                    h = _col_widget_hash(f"{pi}_{col}")
                    opt_label = st.selectbox(
                        f"`{col}`",
                        layout_labels,
                        index=0,
                        key=f"p{pi}_lay_{h}",
                    )
                    layouts_map[col] = layout_label_to_key[opt_label]
            pattern_specs.append(
                {
                    "name": pn,
                    "title": "",
                    "subtitle": "",
                    "filename": _safe_pptx_filename(pfn, f"report_{pi + 1}"),
                    "columns": sel,
                    "layouts": layouts_map,
                }
            )

if delivery.startswith("登壇者別"):
    st.subheader("登壇者ごとの設定")
    n_speakers = st.number_input("登壇者数", min_value=1, max_value=8, value=3, step=1)
    speaker_defs = []
    for i in range(int(n_speakers)):
        with st.expander(f"登壇者 {i + 1}", expanded=(i == 0)):
            sp_name = st.text_input("表示名", value=f"登壇者{i + 1}", key=f"sp_{i}_label")
            pat_pick = st.number_input(
                "使うパターン（1 始まり）",
                min_value=1,
                max_value=npat,
                value=1,
                step=1,
                key=f"sp_{i}_pat",
            )
            t1, t2 = st.columns(2)
            with t1:
                sp_title = st.text_input("スライドタイトル上書き（任意・空ならパターン設定）", key=f"sp_{i}_title", value="")
            with t2:
                sp_sub = st.text_input("サブタイトル上書き（任意）", key=f"sp_{i}_sub", value="")

            use_f = st.checkbox("行フィルタでセッション絞り込み", value=False, key=f"sp_{i}_use_filter")
            filter_col, filter_vals = None, []
            if use_f:
                filter_col = st.selectbox("フィルタ列", [""] + list(df.columns), key=f"sp_{i}_fcol")
                if filter_col and filter_col in df.columns:
                    uniques = sorted(
                        df[filter_col].dropna().astype(str).str.strip().unique().tolist(),
                        key=str,
                    )
                    if len(uniques) > 200:
                        st.caption(f"{len(uniques)} 件中 200 件まで表示")
                        uniques = uniques[:200]
                    filter_vals = st.multiselect("含める値", options=uniques, key=f"sp_{i}_fval")

            out_fn = st.text_input(
                "ZIP 内ファイル名",
                value=_safe_pptx_filename(sp_name, f"speaker_{i + 1}"),
                key=f"sp_{i}_file",
            )
            speaker_defs.append(
                {
                    "index": i,
                    "label": sp_name,
                    "pattern_1based": int(pat_pick),
                    "title_ov": sp_title,
                    "subtitle_ov": sp_sub,
                    "use_filter": use_f,
                    "filter_col": filter_col if filter_col else None,
                    "filter_values": filter_vals,
                    "out_name": out_fn,
                }
            )
else:
    speaker_defs = []

# ── プリセット / 監査（簡易）──────────────────────────────────────────────
with st.expander("列プリセット（パターン1の列名リストの保存・JSON）"):
    preset_name = st.text_input("プリセット名", key="preset_name_input")
    if st.button("パターン1の列を保存"):
        if not preset_name.strip():
            st.error("名前を入力してください。")
        elif not pattern_specs:
            st.error("パターンがありません。")
        else:
            st.session_state.column_presets[preset_name.strip()] = {
                "columns": list(pattern_specs[0]["columns"]),
                "layouts": dict(pattern_specs[0]["layouts"]),
            }
            st.success("保存しました。")
    names = list(st.session_state.column_presets.keys())
    lp = st.selectbox("読み込み", [""] + names, key="preset_load_pick")
    if lp and st.button("パターン1に適用"):
        p = st.session_state.column_presets[lp]
        valid = set(df.columns)
        wanted = {c for c in p.get("columns", []) if c in valid}
        for c in df.columns:
            h = _col_widget_hash(c)
            st.session_state[f"p0_chk_{h}"] = c in wanted
        for c, ly in p.get("layouts", {}).items():
            if c in valid and c in wanted:
                h = _col_widget_hash(c)
                lab = next((lbl for k, lbl in COLUMN_LAYOUT_CHOICES_UI if k == ly), layout_labels[0])
                st.session_state[f"p0_lay_{h}"] = lab
        st.rerun()
    st.download_button(
        "プリセット JSON をダウンロード",
        data=json.dumps(st.session_state.column_presets, ensure_ascii=False, indent=2).encode("utf-8"),
        file_name="survey_pptx_column_presets.json",
        mime="application/json",
    )
    ju = st.file_uploader("JSON を取り込み", type=["json"], key="preset_upload")
    if ju is not None:
        try:
            obj = json.loads(ju.getvalue().decode("utf-8"))
            if isinstance(obj, dict):
                st.session_state.column_presets.update(obj)
                st.success("マージしました。")
            else:
                st.error("オブジェクト形式の JSON が必要です。")
        except json.JSONDecodeError as e:
            st.error(str(e))

with st.expander("監査ログ（セッション）"):
    for entry in reversed(st.session_state.audit_log[-20:]):
        st.json(entry)
    if st.button("監査ログをクリア"):
        st.session_state.audit_log = []
        st.rerun()

# ── 生成ボタン ─────────────────────────────────────────────────────────────
all_patterns_ok = all(spec["columns"] for spec in pattern_specs)
if not all_patterns_ok:
    st.warning("各パターンで少なくとも 1 列を選んでください。")

if delivery.startswith("パターン別"):
    gen = st.button("PowerPoint を生成", type="primary", disabled=not all_patterns_ok)
    if gen and all_patterns_ok:
        if npat == 1:
            spec = pattern_specs[0]
            sub = df[[c for c in spec["columns"] if c in df.columns]]
            lay = {c: spec["layouts"].get(c, COLUMN_LAYOUT_AUTO) for c in sub.columns}
            with tempfile.TemporaryDirectory() as tmpd:
                out_p = Path(tmpd) / spec["filename"]
                log = _run_convert_captured(
                    out_p, sub, spec["title"], spec["subtitle"], enc, template_path_str, lay
                )
                if out_p.exists():
                    st.success("生成しました。")
                    st.download_button(
                        "📥 ダウンロード",
                        data=out_p.read_bytes(),
                        file_name=spec["filename"],
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    )
                    st.session_state.audit_log.append(
                        {
                            "utc": datetime.now(timezone.utc).isoformat(),
                            "mode": "pattern",
                            "file": spec["filename"],
                            "columns": spec["columns"],
                        }
                    )
                else:
                    st.error("生成に失敗しました。")
                with st.expander("ログ"):
                    st.text(log)
        else:
            zip_buf = io.BytesIO()
            logs = []
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for spec in pattern_specs:
                    sub = df[[c for c in spec["columns"] if c in df.columns]]
                    lay = {c: spec["layouts"].get(c, COLUMN_LAYOUT_AUTO) for c in sub.columns}
                    with tempfile.TemporaryDirectory() as tmpd:
                        out_p = Path(tmpd) / "out.pptx"
                        log = _run_convert_captured(
                            out_p, sub, spec["title"], spec["subtitle"], enc, template_path_str, lay
                        )
                        logs.append(f"=== {spec['name']} ===\n{log}\n")
                        if out_p.exists():
                            zf.writestr(spec["filename"], out_p.read_bytes())
            st.success("ZIP を作成しました。")
            st.download_button(
                "📦 ZIP をダウンロード",
                data=zip_buf.getvalue(),
                file_name=f"reports_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M')}.zip",
                mime="application/zip",
            )
            st.session_state.audit_log.append(
                {
                    "utc": datetime.now(timezone.utc).isoformat(),
                    "mode": "pattern_zip",
                    "patterns": [s["filename"] for s in pattern_specs],
                }
            )
            with st.expander("ログ"):
                st.text("".join(logs))

else:
    gen = st.button("登壇者分を ZIP 生成", type="primary", disabled=not all_patterns_ok)
    if gen and all_patterns_ok and speaker_defs:
        zip_buf = io.BytesIO()
        results = []
        logs = []
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for sp in speaker_defs:
                tag = sp["label"] or f"speaker{sp['index'] + 1}"
                pidx = sp["pattern_1based"] - 1
                if pidx < 0 or pidx >= len(pattern_specs):
                    results.append((tag, False, "パターン番号が無効です"))
                    continue
                spec = pattern_specs[pidx]
                try:
                    sub_full = _apply_row_filter(
                        df, sp["use_filter"], sp["filter_col"], sp["filter_values"]
                    )
                    if len(sub_full) == 0:
                        results.append((tag, False, "フィルタ後 0 件"))
                        continue
                    use_cols = [c for c in spec["columns"] if c in sub_full.columns]
                    if not use_cols:
                        results.append((tag, False, "列がありません"))
                        continue
                    sub = sub_full[use_cols]
                    lay = {c: spec["layouts"].get(c, COLUMN_LAYOUT_AUTO) for c in sub.columns}
                    title_u = sp["title_ov"].strip() or spec["title"]
                    sub_u = sp["subtitle_ov"].strip() or spec["subtitle"]
                    with tempfile.TemporaryDirectory() as tmpd:
                        out_p = Path(tmpd) / "out.pptx"
                        log_piece = _run_convert_captured(
                            out_p, sub, title_u, sub_u, enc, template_path_str, lay
                        )
                        logs.append(f"=== {tag} ===\n{log_piece}\n")
                        if not out_p.exists():
                            results.append((tag, False, "PPTX が生成されませんでした"))
                            continue
                        name_in_zip = _safe_pptx_filename(
                            sp["out_name"].replace(".pptx", ""), f"sp_{sp['index'] + 1}"
                        )
                        zf.writestr(name_in_zip, out_p.read_bytes())
                        results.append((tag, True, f"{len(sub_full)} 行"))
                except Exception as ex:
                    results.append((tag, False, str(ex)))
        for tag, ok, msg in results:
            if ok:
                st.success(f"{tag}: {msg}")
            else:
                st.error(f"{tag}: {msg}")
        if any(ok for _, ok, _ in results):
            st.download_button(
                "📦 ZIP をダウンロード",
                data=zip_buf.getvalue(),
                file_name=f"seminar_reports_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M')}.zip",
                mime="application/zip",
            )
            st.session_state.audit_log.append(
                {
                    "utc": datetime.now(timezone.utc).isoformat(),
                    "mode": "speakers_zip",
                    "results": [{"tag": t, "ok": o, "msg": m} for t, o, m in results],
                }
            )
        with st.expander("ログ"):
            st.text("".join(logs))
