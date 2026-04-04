# API 仕様書

## 1. 概要

本プロダクトは **HTTP API を提供しない**。以下を「外部契約」として定義する。

1. **CLI**（`survey_to_pptx.py` のコマンドライン引数）
2. **Python モジュール API**（`survey_to_pptx` の公開関数・定数）

Web UI は Streamlit のサーバ内部から上記モジュールを直接呼び出す。

---

## 2. CLI API

### 2.1 実行形式

```bash
python survey_to_pptx.py <file> [options]
```

### 2.2 位置引数

| 引数 | 説明 |
|------|------|
| `file` | 入力パス。`.csv` または `.xlsx`（Excel は拡張子で判定） |

### 2.3 オプション

| オプション | 必須 | デフォルト | 説明 |
|------------|------|------------|------|
| `-o`, `--output` | 否 | 入力と同名の `.pptx` | 出力ファイルパス |
| `--title` | 否 | 空（`Survey Name` 列やファイル名から補完） | スライドタイトル |
| `--subtitle` | 否 | 空 | サブタイトル |
| `--encoding` | 否 | `utf-8` | CSV 読み込みエンコーディング |
| `--template` | 否 | 空 | テンプレ PPTX のパス（最優先） |
| `--builtin-template` | 否 | なし | 内蔵テンプレファイル名。`BUILTIN_PPTX_SPECS` の第1要素のみ選択可 |

`--template` が非空のときは `--builtin-template` より優先。両方なしの場合は既定テンプレートパスを使用。

### 2.4 終了コード

| コード | 条件 |
|--------|------|
| 0 | 正常終了 |
| 非 0 | ファイル不存在など（標準エラーにメッセージ） |

### 2.5 CLI でないもの

- `column_layout`（列ごとレイアウト指定）は **CLI からは指定不可**。常に自動判定に基づく。

---

## 3. Python モジュール API（survey_to_pptx）

### 3.1 `load_data`

```python
def load_data(path: str, encoding: str = "utf-8") -> pd.DataFrame
```

| 引数 | 説明 |
|------|------|
| `path` | ファイルパス |
| `encoding` | CSV 時の第1候補。失敗時は内部で別エンコーディングを試行 |

**戻り値**: 読み込み済み DataFrame  

**例外**: CSV でエンコーディング判定不能時 `ValueError`

---

### 3.2 `classify_columns`

```python
def classify_columns(df: pd.DataFrame) -> dict[str, str]
```

**戻り値**: 列名 → タイプ文字列（`personal` / `metadata` / `appendix_company` / `categorical` / `high_cardinality` / `text` / `empty`）

---

### 3.3 `list_builtin_templates`

```python
def list_builtin_templates() -> list[tuple[str, Path]]
```

**戻り値**: `(表示ラベル, 解決済みパス)` のリスト。ファイルが存在する内蔵テンプレのみ。

---

### 3.4 `convert`

```python
def convert(
    data_path: str | None,
    output_path: str,
    title: str = "",
    subtitle: str = "",
    encoding: str = "utf-8",
    template_path: str | Path | None = None,
    df: pd.DataFrame | None = None,
    column_layout: dict[str, str] | None = None,
) -> None
```

| 引数 | 説明 |
|------|------|
| `data_path` | `df` が `None` のとき必須。読み込み元パス |
| `output_path` | 出力 `.pptx` パス |
| `title` / `subtitle` | 表紙等に利用（テンプレ有無で挙動が異なる） |
| `encoding` | `data_path` から CSV 読む際に使用 |
| `template_path` | テンプレ PPTX。`None` または不存在時は内蔵既定 |
| `df` | 指定時は `data_path` を読まず、この DataFrame を変換 |
| `column_layout` | 列名 → レイアウトキー。`None` または各列未指定は `auto` 相当 |

**レイアウトキー**（定数と同一文字列）:

- `COLUMN_LAYOUT_AUTO` → `"auto"`
- `COLUMN_LAYOUT_SELECTION_CHART` → `"selection_chart"`
- `COLUMN_LAYOUT_PIE_ONLY` → `"pie_only"`
- `COLUMN_LAYOUT_FREE_TEXT` → `"free_text_table"`
- `COLUMN_LAYOUT_APPENDIX_LIST` → `"appendix_list"`

**副作用**: `output_path` に PPTX を書き込み。ログを `print` する。

**前提**: `df` を渡す場合でも列は変換対象のみに絞った subset 可。

---

### 3.5 UI 向け定数

```python
COLUMN_LAYOUT_CHOICES_UI: tuple[tuple[str, str], ...]
```

各要素は `(レイアウトキー, 日本語ラベル)`。

---

### 3.6 その他の関数

以下は主に内部利用だが import 可能。

| 名前 | 用途 |
|------|------|
| `is_personal_col` / `is_metadata_col` / `is_appendix_company_col` | 列名ルール |
| `detect_col_type` | 系列から型推定 |
| `add_*_slide` 系 | スライド生成の部品 |

安定した外部契約としては **3.1〜3.5** を推奨。スライド関数の署名変更は破壊的になり得る。

---

## 4. Web UI と「API」の関係

| UI 操作 | モジュール呼び出し |
|---------|-------------------|
| 読み込み | `load_data`, `classify_columns` |
| テンプレ一覧 | `list_builtin_templates` |
| 生成 | `convert(..., df=..., column_layout=..., template_path=...)` |

認証・セッション・ファイルダウンロードは Streamlit の仕様に依存。OpenAPI 等のスキーマは **該当なし**。

---

## 5. 環境変数（運用契約）

| 変数名 | 利用箇所 | 説明 |
|--------|----------|------|
| `STREAMLIT_PASSWORD` | `app.py` | 設定時のみ簡易ログインを有効化 |

---

## 6. 改訂履歴

| 日付 | 版 | 内容 |
|------|-----|------|
| 2026-04-04 | 1.0 | 初版 |
