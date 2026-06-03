#!/bin/bash
# post-app（投稿本数アプリ）の本番デプロイ
# 使い方: ./deploy-postapp.sh
#
# 何をするか:
#   1. ローカル post-app.html を Xserver giver.work/post-app/index.html に配信
#   2. 配信前に Xserver 上の現行ファイルをバックアップ
#   3. JSの構文チェックして OK の場合のみアップロード

set -e

REMOTE="xserver"
BASE="/home/kodaidai/giver.work/public_html"
LOCAL_HTML="/Users/kodai/eisei/post-app.html"
REMOTE_HTML="$BASE/post-app/index.html"
TS=$(date +%Y%m%d_%H%M%S)

echo "================================================"
echo " post-app デプロイ"
echo "================================================"

# ① ローカルファイル存在確認
if [ ! -f "$LOCAL_HTML" ]; then
  echo "❌ ローカルファイルなし: $LOCAL_HTML"
  exit 1
fi

# ② JS構文チェック
echo "[1/4] JS構文チェック..."
node -e "
const fs = require('fs');
const html = fs.readFileSync('$LOCAL_HTML', 'utf8');
const m = html.match(/<script>([\s\S]*?)<\/script>/g) || [];
let ok = true;
m.forEach((s, i) => {
  try { new Function(s.replace(/<\/?script>/g, '')); }
  catch (e) { console.error('script #' + i + ' ERROR: ' + e.message); ok = false; }
});
process.exit(ok ? 0 : 1);
" || { echo "❌ JS構文エラー。デプロイ中止。"; exit 1; }
echo "  → OK"

# ③ 現行ファイルをバックアップ
echo "[2/4] 現行ファイルをバックアップ → index.html.bak.$TS"
ssh "$REMOTE" "cp '$REMOTE_HTML' '$REMOTE_HTML.bak.$TS'" || { echo "❌ バックアップ失敗"; exit 1; }
echo "  → OK"

# ④ ローカル → Xserver
echo "[3/4] アップロード: $LOCAL_HTML → $REMOTE_HTML"
scp "$LOCAL_HTML" "$REMOTE:$REMOTE_HTML"
echo "  → OK"

# ⑤ サイズ確認
echo "[4/4] デプロイ後の確認"
LOCAL_SIZE=$(wc -c < "$LOCAL_HTML" | tr -d ' ')
REMOTE_SIZE=$(ssh "$REMOTE" "wc -c < '$REMOTE_HTML'" | tr -d ' ')
echo "  ローカル: $LOCAL_SIZE bytes / 本番: $REMOTE_SIZE bytes"

if [ "$LOCAL_SIZE" = "$REMOTE_SIZE" ]; then
  echo "  → サイズ一致 OK"
else
  echo "❌ サイズ不一致！アップロード失敗の可能性"
  exit 1
fi

echo ""
echo "================================================"
echo "🟢 post-app デプロイ完了"
echo ""
echo "次のアクション:"
echo "  1. ブラウザで https://giver.work/post-app/ を開いて確認"
echo "  2. GAS側もデプロイした場合は ./scripts/deploy-check.sh で疎通確認"
echo "================================================"
