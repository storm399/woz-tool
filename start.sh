#!/usr/bin/env bash
# Start de WOZ-tool en open in browser.
set -e
DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$DIR"

# Open de browser zodra de server live is (kleine sleep zodat Flask klaar staat)
( sleep 1.5 && open "http://127.0.0.1:5005" ) &

python3 app.py
