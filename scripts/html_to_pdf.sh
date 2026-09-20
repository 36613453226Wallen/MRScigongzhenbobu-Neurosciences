#!/bin/bash
set -e
CHROME=/usr/local/bin/google-chrome
EXPORT=/workspace/docs/export
for f in "慢性应激大脑改变与睾酮保护_完整收录" "长期结构改变_文献笔记" "睾酮保护原理_文献笔记"; do
  udir="/tmp/chrome-pdf-$(date +%s)-$RANDOM"
  mkdir -p "$udir"
  echo "=== printing $f ==="
  timeout 45 "$CHROME" \
    --headless --disable-gpu --no-sandbox --allow-file-access-from-files \
    --user-data-dir="$udir" \
    --no-pdf-header-footer \
    --virtual-time-budget=12000 \
    --print-to-pdf="$EXPORT/${f}.pdf" \
    "file://$EXPORT/${f}.html" || true
  pkill -9 -f "user-data-dir=$udir" 2>/dev/null || true
  ls -la "$EXPORT/${f}.pdf"
  cp -f "$EXPORT/${f}.pdf" "/workspace/收录/${f}.pdf"
  rm -rf "$udir"
done
echo ALL_OK
ls -la "$EXPORT"/*.pdf /workspace/收录/*.pdf
