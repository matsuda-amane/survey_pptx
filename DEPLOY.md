# Streamlit Community Cloud へのデプロイ手順

## 1. GitHub にリポジトリを作る

### 1-1. リポジトリを初期化（まだの場合）

```bash
cd /Users/matsuda/develop/survey_pptx
git init
git add .
git commit -m "Initial commit: アンケート→PPTX変換 + Streamlitアプリ"
```

### 1-2. GitHub で新規リポジトリ作成

1. [GitHub](https://github.com) にログイン
2. **New repository** でリポジトリ作成（例: `survey_pptx`）
3. **Private** または **Public** を選択（Cloud ではどちらでもデプロイ可能）
4. **Create repository** まで進み、表示されるリモート URL を控える

### 1-3. プッシュ

```bash
git remote add origin https://github.com/<あなたのユーザー名>/survey_pptx.git
git branch -M main
git push -u origin main
```

> **テンプレートの置き場所**: このリポジトリ直下の `template/` に置くか、後述の **git submodule** でプライベートリポジトリから取り込みます。`data/` 内の実データ（xlsx/csv/pptx）は .gitignore で除外されます。

---

## 1 bis. テンプレートをプライベートリポジトリで管理する（git submodule）

パブリックなメインリポジトリに `.pptx` を載せたくないときの定番のやり方です。**`template` 用に別リポジトリを作り、このリポジトリでは `template` フォルダを submodule として参照**します。

### 手順の流れ（比喩）

- **メインの箱**（survey_pptx）＝ コードや Streamlit。みんなに見せてよいもの。
- **テンプレだけの鍵付きの箱**（private repo）＝ 中身の pptx。ここだけプライベート。
- **submodule**＝「メインの箱の `template` という棚は、あっちの鍵付き箱の中身をそのまま指す印」です。

### A. プライベート側リポジトリを用意する

1. GitHub などで **新しい Private リポジトリ** を作る（例: `survey_pptx-templates`）。
2. いまの `template/` と同じファイル名で、**ルートに** `.pptx` を push する。  
   （`template_ligare.pptx` / `template_amane.pptx` など。アプリは `template/` 直下＝ submodule のルートを読みます。サブフォルダにしないこと。）

### B. メインリポジトリで submodule を登録する

リポジトリのルートで:

```bash
chmod +x scripts/setup_template_submodule.sh
./scripts/setup_template_submodule.sh git@github.com:<あなた>/<survey_pptx-templates>.git
```

スクリプト後に表示どおり `git add` / `git commit` / `git push` する。

**初めて clone する人**は、次のどちらかが必要です。

```bash
git clone --recurse-submodules <メインリポジトリのURL>
# 既に clone 済みなら
git submodule update --init --recursive
```

### C. Streamlit Community Cloud での注意

- デプロイ時に **サブモジュールを取得できるか** は、ホストの clone 設定と **サブモジュール側リポジトリへのアクセス権** に依存します。
- **メインがパブリック・テンプレだけプライベート** のとき、無料の Community Cloud だけではサブモジュール用の認証が足りず、`template/` が空のままビルドに失敗することがあります。
- その場合の例: **メインも Private にして** Streamlit にアクセス許可する、別ホスト（Render など）でデプロイ鍵を設定する、テンプレをビルド前にコピーする CI を別途用意する、など。

どうなるかは必ず **再デプロイしてログを確認** してください。

### D. テンプレの場所を環境変数で指定する（任意）

サブモジュール以外に、コンテナ上で別パスにテンプレがある場合:

```text
SURVEY_PPTX_TEMPLATE_DIR=/path/to/dir
```

そのディレクトリに `template_ligare.pptx` などがある前提です（`survey_to_pptx` と `app` の両方で共通）。

---

## 2. Streamlit Community Cloud でデプロイ

### 2-1. デプロイ開始

1. [https://share.streamlit.io](https://share.streamlit.io) を開く
2. **Sign up / Log in** で **GitHub** アカウントでログイン
3. **New app** をクリック

### 2-2. アプリ設定

| 項目 | 入力例 |
|------|--------|
| **Repository** | `あなたのユーザー名/survey_pptx` |
| **Branch** | `main` |
| **Main file path** | `app.py` |
| **App URL** | 任意（例: `survey-pptx`）→ `https://survey-pptx.streamlit.app` など |

### 2-3. パスワード（シークレット）の設定

1. **Advanced settings** を開く
2. **Secrets** に次を入力（パスワードは任意の文字列に変更）：

```toml
STREAMLIT_PASSWORD = "ここにログインパスワードを設定"
```

3. **Deploy** をクリック

### 2-4. 初回ビルド

- 数分かかることがあります
- 完了するとアプリの URL が表示されます
- その URL を開き、設定したパスワードでログインして動作確認してください

---

## 3. デプロイ後の更新

コードを変更したら、GitHub に push するだけで自動で再デプロイされます。

```bash
git add .
git commit -m "説明メッセージ"
git push
```

---

## 4. 注意事項

- **無料枠**: 一定時間アクセスがないとアプリがスリープし、次回アクセス時に起動まで数十秒かかることがあります
- **パスワード**: シークレットは Streamlit の画面から再設定できます（Settings → Secrets）
- **テンプレート**: 実行環境の `template/`（または submodule・`SURVEY_PPTX_TEMPLATE_DIR`）に、`template_ligare.pptx` / `template_amane.pptx` などが存在する必要があります
