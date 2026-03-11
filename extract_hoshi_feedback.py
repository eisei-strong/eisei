#!/usr/bin/env python3
"""
ほしのフィードバックメッセージをChatworkログから抽出するスクリプト
出力: hoshi_feedback_kb.csv（Google Sheetsにアップロード用）
"""

import csv
import os
import glob
import re
import sys

CHATWORK_LOGS = "/Users/kodai/chatwork-logs"
HOSHI_ACCOUNT_ID = "4867412"
OUTPUT_FILE = "/Users/kodai/sales-dashboard/hoshi_feedback_kb.csv"

# 対象ルーム
TARGET_ROOMS = {
    "229293499": {"name": "Namaka商談", "path": "rooms/86/229293499"},
    "227966133": {"name": "アポを取る", "path": "rooms/16/227966133"},
    "205589572": {"name": "チームNamaka", "path": "rooms/ed/205589572"},
}

# 無視パターン
IGNORE_PATTERNS = [
    r"^\[info\]\[title\]ファイルをアップロード",
    r"^\[info\]\[title\]タスクを(追加|完了)",
    r"^\[info\]\[title\]新しくグループチャット",
    r"^\[info\]\[title\]\[dtext:",
    r"^\[deleted\]$",
    r"^\[info\]\[title\]メンバー",
]

# カテゴリ判定
def categorize(body):
    if re.search(r"商談|成約|見送り|保留|クロージング|ドタキャン|リスケ", body):
        return "商談コーチング"
    if re.search(r"結論|理由|文章|日本語|見やすく|書き方|伝え方", body):
        return "コミュニケーション改善"
    if re.search(r"メッセージ|ステップ|エルメ|LINE|送信|配信", body):
        return "メッセージ最適化"
    if re.search(r"提案|施策|戦略|テスト|企画", body):
        return "提案・施策"
    if re.search(r"ナイス|素晴らしい|いいね|最高|やりましょう|good", body, re.IGNORECASE):
        return "称賛・激励"
    if re.search(r"データ|数字|％|%|率|件数|ランキング", body):
        return "データ駆動改善"
    return "日常業務"

# キーワード抽出
KEYWORD_PATTERNS = [
    (r"商談", "商談"), (r"成約", "成約"), (r"改善", "改善"),
    (r"提案", "提案"), (r"依頼", "依頼"), (r"クロージング", "クロージング"),
    (r"リスケ", "リスケ"), (r"返金", "返金"), (r"LINE", "LINE"),
    (r"エルメ", "エルメ"), (r"TikTok", "TikTok"), (r"データ", "データ"),
    (r"報告", "報告"), (r"質問", "質問"), (r"フォロー", "フォロー"),
    (r"アポ", "アポ"), (r"見送り", "見送り"), (r"保留", "保留"),
    (r"戦略", "戦略"), (r"コンテンツ", "コンテンツ"),
]

def extract_keywords(body):
    keywords = []
    for pattern, kw in KEYWORD_PATTERNS:
        if re.search(pattern, body):
            keywords.append(kw)
    return ",".join(keywords)

def should_ignore(body):
    for pattern in IGNORE_PATTERNS:
        if re.search(pattern, body):
            return True
    return False

def read_room_messages(room_id, room_info):
    """ルームの全メッセージファイルを読み込み、時系列順で返す"""
    room_path = os.path.join(CHATWORK_LOGS, room_info["path"])
    message_files = sorted(glob.glob(os.path.join(room_path, "message*.csv")))

    all_messages = []
    for mf in message_files:
        try:
            with open(mf, "r", encoding="utf-8-sig") as f:
                reader = csv.reader(f)
                header = next(reader, None)
                if not header:
                    continue
                for row in reader:
                    if len(row) >= 7:
                        all_messages.append({
                            "datetime": row[0],
                            "room_id": row[2],
                            "room_name": row[3],
                            "account_id": row[4],
                            "account_name": row[5],
                            "body": row[6],
                        })
        except Exception as e:
            print(f"  Warning: Error reading {mf}: {e}", file=sys.stderr)

    # 時系列ソート
    all_messages.sort(key=lambda m: m["datetime"])
    return all_messages

def extract_feedback():
    """メイン抽出処理"""
    results = []

    for room_id, room_info in TARGET_ROOMS.items():
        print(f"Processing room: {room_info['name']} ({room_id})...")
        messages = read_room_messages(room_id, room_info)
        print(f"  Total messages: {len(messages)}")

        hoshi_count = 0
        extracted = 0

        for i, msg in enumerate(messages):
            # ほしのメッセージのみ
            if msg["account_id"] != HOSHI_ACCOUNT_ID:
                continue
            hoshi_count += 1

            body = msg["body"]

            # 短すぎるメッセージスキップ
            if len(body) < 80:
                continue

            # 無視パターンスキップ
            if should_ignore(body):
                continue

            # トリガーメッセージ（直前の他者メッセージ）を取得
            trigger_msg = ""
            for j in range(i - 1, max(i - 5, -1), -1):
                if j < 0:
                    break
                if messages[j]["account_id"] != HOSHI_ACCOUNT_ID:
                    trigger_msg = messages[j]["body"][:500]
                    break

            category = categorize(body)
            keywords = extract_keywords(body)

            results.append({
                "datetime": msg["datetime"],
                "room_id": room_id,
                "room_name": room_info["name"],
                "category": category,
                "keywords": keywords,
                "trigger_message": trigger_msg,
                "response": body[:2000],
                "response_length": len(body),
            })
            extracted += 1

        print(f"  ほしメッセージ: {hoshi_count}, 抽出: {extracted}")

    return results

def write_csv(results):
    """CSV出力"""
    with open(OUTPUT_FILE, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([
            "日時", "ルームID", "ルーム名", "カテゴリ",
            "キーワード", "トリガーメッセージ", "ほしの返答", "文字数"
        ])
        for r in results:
            writer.writerow([
                r["datetime"], r["room_id"], r["room_name"],
                r["category"], r["keywords"], r["trigger_message"],
                r["response"], r["response_length"],
            ])

def main():
    print("=== ほしフィードバック抽出 ===")
    results = extract_feedback()
    print(f"\n合計抽出件数: {len(results)}")

    # カテゴリ別集計
    cats = {}
    for r in results:
        cats[r["category"]] = cats.get(r["category"], 0) + 1
    print("\nカテゴリ別:")
    for cat, count in sorted(cats.items(), key=lambda x: -x[1]):
        print(f"  {cat}: {count}")

    write_csv(results)
    print(f"\n出力: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
