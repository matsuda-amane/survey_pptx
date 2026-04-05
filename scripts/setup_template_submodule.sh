#!/usr/bin/env bash
# テンプレート用のプライベートリポジトリを git submodule として template/ に結び付ける。
#
# 事前準備:
#   1. GitHub / GitLab などでプライベートリポジトリを新規作成する（例: survey_pptx-templates）
#   2. いまの template/ と同じ構成で、.pptx をリポジトリのルートに push する
#      （template_amane.pptx / template_ligare.pptx など。このアプリは TEMPLATE_DIR 直下を見る）
#   3. SSH か HTTPS で clone 可能な URL を用意する
#
# 使い方（メインの survey_pptx リポジトリのルートで）:
#   ./scripts/setup_template_submodule.sh git@github.com:ORG/survey_pptx-templates.git
#
set -euo pipefail

ROOT="$(git rev-parse --show-toplevel)"
cd "$ROOT"

URL="${1:-}"
if [[ -z "$URL" ]]; then
  echo "Usage: $0 <private-templates-repo-url>"
  echo "Example: $0 git@github.com:you/survey_pptx-templates.git"
  exit 1
fi

if [[ -f template/.git || -d template/.git ]]; then
  echo "template/ はすでに submodule のようです。.gitmodules を確認してください。"
  exit 0
fi

BACKUP_DIR="${ROOT}/../survey_pptx_template_backup_$(date +%Y%m%d_%H%M%S)"

if [[ -d template ]]; then
  echo "既存の template/ を退避します: $BACKUP_DIR"
  mv template "$BACKUP_DIR"
fi

# インデックスから template 配下を外す（未追跡なら何もしない）
git rm -r --cached template 2>/dev/null || true

git submodule add "$URL" template
git submodule update --init --recursive

echo ""
echo "完了。次を実行してコミットしてください:"
echo "  git add .gitmodules template"
echo "  git commit -m \"chore: template をプライベート repo を submodule に\""
echo ""
echo "他の開発者は次でテンプレ付き clone になります:"
echo "  git clone --recurse-submodules <メインrepoのURL>"
echo "  既に clone 済みなら: git submodule update --init --recursive"
echo ""
echo "退避した古い template/ は次にあります（リモートと差分確認後に削除してよいです）:"
echo "  $BACKUP_DIR"
