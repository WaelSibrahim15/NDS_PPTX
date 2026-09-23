#!/bin/zsh
# NDS — Narrated Deck Studio launcher (double-click to start)
cd "$(dirname "$0")"
# A .venv only works in the folder it was created in; rebuild it if the
# folder was moved/renamed or Python was upgraded.
if [ -d .venv ] && ! ./.venv/bin/python -c "import uvicorn" >/dev/null 2>&1; then
  echo "Python environment is out of date (folder moved or Python updated) — rebuilding…"
  rm -rf .venv
fi
if [ ! -d .venv ]; then
  echo "First run — setting up NDS (this takes a minute)…"
  python3 -m venv .venv
  ./.venv/bin/pip install -q -r requirements.txt
fi
echo "Starting NDS at http://localhost:8765  (leave this window open)"
( sleep 2 && open "http://localhost:8765" ) &
./.venv/bin/python -m uvicorn app.main:app --port 8765
