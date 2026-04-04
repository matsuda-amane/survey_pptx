#!/usr/bin/env bash
# Docker なしでローカル起動（初回は venv 作成と pip install）
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

if [[ -z "${STREAMLIT_PASSWORD:-}" ]]; then
  echo "STREAMLIT_PASSWORD が未設定です。例: STREAMLIT_PASSWORD=dev $0" >&2
  exit 1
fi

VENV="$ROOT/.venv"
if [[ ! -d "$VENV" ]]; then
  python3 -m venv "$VENV"
fi
# shellcheck source=/dev/null
source "$VENV/bin/activate"
pip install -q -r requirements.txt
exec streamlit run app.py --server.port 8501
