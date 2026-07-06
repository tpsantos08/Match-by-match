#!/bin/bash
# ── Update dashboard with new match data ──────────────────────────────────────
# 1. Drop updated Allsvenskan_2026.xlsx into source-data/ first
# 2. Run this script from the repo folder

set -e
cd "$(dirname "$0")"

echo "Running preprocess..."
python3 preprocess.py

echo "Pushing to GitHub..."
git add .
git commit -m "update data $(date +%Y-%m-%d)"
git push

echo "Done — Netlify will deploy in ~1 minute."
