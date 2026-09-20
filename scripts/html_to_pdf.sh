#!/bin/bash
set -e
CHROME=/usr/local/bin/google-chrome
EXPORT=/workspace/docs/export
ARCHIVE=/workspace/收录
SUB="$ARCHIVE/长期结构改变2-近20年代补充"
mkdir -p "$ARCHIVE" "$SUB"
for f in "慢性应激大脑改变与睾酮保护_完整收录" "长期结构改变_文献笔记" "睾酮保护原理_文献笔记" "长期结构改变2_近20年代补充_文献笔记"; do
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
  if [ "$f" = "长期结构改变2_近20年代补充_文献笔记" ]; then
    cp -f "$EXPORT/${f}.pdf" "$SUB/${f}.pdf"
  else
    cp -f "$EXPORT/${f}.pdf" "$ARCHIVE/${f}.pdf"
  fi
  rm -rf "$udir"
done
echo ALL_OK
ls -la "$EXPORT"/*.pdf "$ARCHIVE"/*.pdf "$SUB"
