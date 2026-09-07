#!/bin/bash
# Claude Code on the web 的 SessionStart hook：安裝腳本需要的 Python 套件。
# 只在遠端環境執行；本機不動。
set -euo pipefail

if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

cd "${CLAUDE_PROJECT_DIR:-$(pwd)}"
python3 -m pip install --quiet --disable-pip-version-check -r requirements.txt

# 容器內建的系統版 cryptography（Debian 套件）會讓 pypdf import 時當掉；
# 無法解除安裝，改用 --ignore-installed 另裝新版到 /usr/local 蓋過。
if ! python3 -c "import pypdf" >/dev/null 2>&1; then
  python3 -m pip install --quiet --disable-pip-version-check --ignore-installed cryptography
fi

python3 - <<'PY'
import openpyxl, pypdf
print(f"deps ok: openpyxl {openpyxl.__version__}, pypdf {pypdf.__version__}")
PY
