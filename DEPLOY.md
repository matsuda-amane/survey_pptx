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

> **テンプレート**: `template/template_ligare.pptx`・`template/template_amane.pptx` などをこのリポジトリに含めて push してください（Streamlit Cloud は通常の `git clone` だけで `template/` が揃います）。`data/` 内の実データは .gitignore で除外されます。  
> 別パスにテンプレを置く必要があるだけのときは、環境変数 **`SURVEY_PPTX_TEMPLATE_DIR`** を指定できます（そのディレクトリに `.pptx` がある前提）。

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
- **テンプレート**: リポジトリの `template/` に `template_ligare.pptx` / `template_amane.pptx` などが含まれている必要があります（上書きで `SURVEY_PPTX_TEMPLATE_DIR` を使うことも可）
