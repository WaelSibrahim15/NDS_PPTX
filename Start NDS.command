#!/bin/zsh
# NDS — Narrated Deck Studio launcher (double-click to start)
cd "$(dirname "$0")"
if [ ! -d .venv ]; then
  echo "First run — setting up NDS (this takes a minute)…"
  python3 -m venv .venv
  ./.venv/bin/pip install -q -r requirements.txt
fi
echo "Starting NDS at http://localhost:8765  (leave this window open)"
( sleep 2 && open "http://localhost:8765" ) &
./.venv/bin/uvicorn app.main:app --port 8765
