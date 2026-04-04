# DB 設計書

## 1. 前提

本システムは **リレーショナルデータベース（RDB）を使用しない**。アンケートデータの想定保持場所は **ユーザーがアップロードする CSV/XLSX ファイル**および **メモリ上の pandas.DataFrame** に限定される。

この文書では、(1) 永続化がないことの明示、(2) アプリ内で扱う **論理データ構造**、(3) 将来 DB を導入する場合の参考を記載する。

## 2. 永続化の現状

| 種別 | 保存先 | ライフサイクル |
|------|--------|----------------|
| アンケート原データ | ブラウザ → サーバメモリ（一時処理） | リクエスト／セッション単位。サーバにファイルとして保存しない実装。 |
| 生成結果 PPTX | ブラウザダウンロード | ユーザー端末のみ。 |
| 列プリセット | `st.session_state.column_presets` | 同一ブラウザセッション。 |
| JSON プリセットファイル | ユーザーが手動で DL/UL | ファイルシステム（クライアント側）。 |
| 監査ログ | `st.session_state.audit_log` | 同一ブラウザセッション。 |
| ログインフラグ | `.streamlit` / セッション | セッション単位。 |

## 3. 論理データ構造（エンティティ相当）

実装に現れない「概念モデル」として整理する。

### 3.1 アンケートデータセット（SurveyDataset）

| 論理属性 | 説明 |
|----------|------|
| columns | 列名のリスト |
| rows | 行データ（表形式） |
| source_filename | 元ファイル名 |
| format | `csv` / `xlsx` |

### 3.2 列メタデータ（ColumnMeta）

`classify_columns` の結果。

| 論理属性 | 説明 |
|----------|------|
| column_name | 列名 |
| inferred_type | personal / metadata / appendix_company / categorical / high_cardinality / text / empty |

### 3.3 出力パターン（OutputPattern）

| 論理属性 | 説明 |
|----------|----------|
| pattern_index | 0 〜 N-1 |
| display_name | UI 上のパターン名 |
| slide_title / slide_subtitle | 表紙・サマリ用 |
| output_filename | `.pptx` 名 |
| included_columns | 含める列名のリスト |
| column_layouts | 列名 → レイアウトキー |

### 3.4 登壇者納品設定（SpeakerDelivery）

登壇者別 ZIP 時のみ。

| 論理属性 | 説明 |
|----------|------|
| speaker_label | 表示名 |
| pattern_index_1based | 参照するパターン（1 始まり） |
| title_override / subtitle_override | 空ならパターン既定を使用 |
| row_filter_enabled | 真偽 |
| filter_column / filter_values | 行絞り込み条件 |
| zip_entry_filename | ZIP 内の PPTX 名 |

### 3.5 列プリセット（ColumnPreset）

JSON / セッションで保持する形。

```json
{
  "プリセット名": {
    "columns": ["列A", "列B"],
    "layouts": {
      "列A": "auto",
      "列B": "selection_chart"
    }
  }
}
```

### 3.6 監査ログエントリ（AuditLogEntry）

| 論理属性 | 例 |
|----------|-----|
| utc | ISO8601 時刻（UTC） |
| mode | `pattern` / `pattern_zip` / `speakers_zip` |
| その他 | 生成ファイル名、パターン一覧、登壇者ごとの成否など（実装依存） |

## 4. RDB を導入する場合の参考（非必須）

将来的に複数ユーザー・履歴保管が必要になった場合の例。**現在未実装**。

| テーブル | 主なカラム案 |
|----------|----------------|
| users | id, email, … |
| jobs | id, user_id, created_at, mode, status |
| job_patterns | job_id, pattern_index, title, filename, columns_json, layouts_json |
| audit_events | id, job_id, payload_json, created_at |

正規化の詳細は要件確定後に別途設計する。

## 5. 改訂履歴

| 日付 | 版 | 内容 |
|------|-----|------|
| 2026-04-04 | 1.0 | 初版（ no-DB 方針の明示） |
