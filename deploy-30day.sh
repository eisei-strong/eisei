#!/bin/bash
# giver.work 一括デプロイ
# 使い方: ./deploy-30day.sh

REMOTE="xserver"
BASE="/home/kodaidai/giver.work/public_html"

echo "== 30日間講座デプロイ =="
scp /Users/kodai/営業講座/30day_program.html "$REMOTE:$BASE/30day/index.html"
scp /Users/kodai/営業講座/30day_admin.html "$REMOTE:$BASE/30day/admin.html"
scp /Users/kodai/営業講座/30day_scripts.html "$REMOTE:$BASE/30day/scripts.html"

echo "== ホープ数ダッシュボードデプロイ =="
HOPE_DIR="/Users/kodai/Hope's Dashboard"
rsync -avz "$HOPE_DIR/index.html" "$HOPE_DIR/data.json" "$REMOTE:$BASE/hope-dashboard/"
rsync -avz "$HOPE_DIR/icons/" "$REMOTE:$BASE/hope-dashboard/icons/"

echo "== 投稿ランキングデプロイ =="
ssh "$REMOTE" "mkdir -p $BASE/post-ranking"
scp "/Users/kodai/Hope's Dashboard/ranking.html" "$REMOTE:$BASE/post-ranking/index.html"

echo "== 営業ダッシュボードデプロイ =="
scp /Users/kodai/営業講座/Dashboard-wp.html "$REMOTE:$BASE/sales-dashboard/index.html"

echo ""
echo "== 完了 =="
echo "  30日間講座:     https://giver.work/30day/"
echo "  台本スクリプト:   https://giver.work/30day/scripts.html"
echo "  講座管理:       https://giver.work/30day/admin.html"
echo "  ホープ数:       https://giver.work/hope-dashboard/"
echo "  投稿ランキング:   https://giver.work/post-ranking/"
echo "  営業ダッシュボード: https://giver.work/sales-dashboard/"
