#!/bin/bash
# eisei GASデプロイ@665 ヘルスチェック
# 使い方: ./scripts/deploy-check.sh
# clasp deploy → GAS Editor で権限再設定 → このスクリプトで疎通確認

set -e

# 共有エンドポイント @665 (4システム共有)
URL="https://script.google.com/macros/s/AKfycbw2tvPqcuJttb09OuuCDKvi5mQMwcCDqJLFRPJk3pc4w0IIAOyDPEPTRnUKPrMDPgGE4A/exec"

PASS=0
FAIL=0
RESULTS=""

check() {
  local name=$1
  local action=$2
  local extra=$3
  local expect=${4:-json}  # "json" or "html"
  local response
  local http_code
  local ts=$(date +%s)

  response=$(curl -s -L "${URL}?action=${action}${extra}&_=${ts}" 2>&1)
  http_code=$(curl -s -o /dev/null -w "%{http_code}" -L "${URL}?action=${action}${extra}&_=${ts}_h" 2>&1)

  # 認証要求 = 致命的（権限リセット中）
  if [[ "$response" =~ "accounts.google.com" ]] || [[ "$response" =~ "ServiceLogin" ]]; then
    RESULTS+="🔴 ${name}: 認証要求ページ返却 (HTTP=${http_code}) ← アクセス権限がリセットされてる！\n"
    FAIL=$((FAIL+1))
    return
  fi

  if [[ "$response" =~ "ウェブ ワープロ" ]]; then
    RESULTS+="🔴 ${name}: アクセス権限なし (HTTP=${http_code}) ← デプロイIDが間違ってるかも\n"
    FAIL=$((FAIL+1))
    return
  fi

  if [ "$expect" = "json" ]; then
    if [[ "$response" =~ ^\{ ]] || [[ "$response" =~ ^\[ ]]; then
      RESULTS+="✅ ${name}: JSON応答OK (HTTP=${http_code})\n   → ${response:0:80}\n"
      PASS=$((PASS+1))
    else
      RESULTS+="⚠️ ${name}: 期待=JSON、実際=非JSON (HTTP=${http_code})\n   → ${response:0:100}\n"
      FAIL=$((FAIL+1))
    fi
  else
    # HTML応答が期待
    if [[ "$response" =~ "<html" ]] || [[ "$response" =~ "<!doctype" ]] || [[ "$response" =~ "<!DOCTYPE" ]]; then
      RESULTS+="✅ ${name}: HTML応答OK (HTTP=${http_code})\n"
      PASS=$((PASS+1))
    else
      RESULTS+="⚠️ ${name}: 期待=HTML、実際=非HTML (HTTP=${http_code})\n   → ${response:0:100}\n"
      FAIL=$((FAIL+1))
    fi
  fi
}

echo "================================================"
echo " eisei GAS @665 ヘルスチェック"
echo " URL: ${URL:0:80}..."
echo "================================================"
echo ""

# 4システム共有エンドポイントを叩いて、各システムのAPIが応答するか確認
echo "[1/4] 投稿本数アプリ (post-app) - postCheckId"
check "post-app/postCheckId" "postCheckId" "&id=5740"

echo "[2/4] 投稿本数アプリ (post-app) - postGetHope"
check "post-app/postGetHope" "postGetHope" "&token=invalid&year=2026&month=5"

echo "[3/4] 投稿本数アプリ (post-app) - postGetPush"
check "post-app/postGetPush" "postGetPush" "&token=invalid&year=2026&month=5"

echo "[4/4] 営業/学生ダッシュボード - dashboard (HTML応答が正常)"
check "営業ダッシュボード/dashboard" "dashboard" "" "html"

echo ""
echo "================================================"
echo -e "$RESULTS"
echo "================================================"
echo "結果: ${PASS} 成功 / ${FAIL} 失敗"
echo ""

if [ "$FAIL" -gt 0 ]; then
  echo "🚨 失敗あり！受講生は使えない状態の可能性。対応:"
  echo ""
  echo "1. GAS Editorで権限再設定:"
  echo "   cd ~/eisei && clasp open"
  echo "   右上「デプロイ」→「デプロイを管理」→対象デプロイ→✏️鉛筆"
  echo "   「次のユーザーとして実行」=「自分(kuta310k@gmail.com)」"
  echo "   「アクセスできるユーザー」=「全員」（NOT「Google アカウントを持つ全員」）"
  echo "   →「デプロイ」ボタン"
  echo ""
  echo "2. 再度このスクリプトで確認:"
  echo "   ./scripts/deploy-check.sh"
  exit 1
else
  echo "🟢 全システム正常応答。本番OK。"
  exit 0
fi
