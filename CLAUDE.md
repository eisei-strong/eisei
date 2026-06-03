# eisei プロジェクト用 Claude 運用ルール

このファイルは `eisei-strong/eisei` リポジトリ専用のルール。
全プロジェクト共通ルールは `~/.claude/CLAUDE.md` 参照。

---

## このリポジトリの全体像

複数のWebアプリ・GAS機能が1つのGASプロジェクト・1つのリポジトリに同居している。

| システム | フロントHTML | 配信先 | API URL |
|---|---|---|---|
| 投稿本数アプリ (post-app) | `post-app.html` | **Xserver** `giver.work/post-app/index.html` | 共有@665 |
| 営業ダッシュボード (Dashboard) | `Dashboard-wp.html` | **Xserver** `giver.work/sales-dashboard/` | 共有@665 |
| ガーディアンダッシュボード | `guardian-dashboard.html` | **Xserver** | 共有@665 |
| 学生ダッシュボード (経由) | `api-proxy.php` | **Xserver** (PHPプロキシ) | 共有@665 |
| 30日間講座 | `30day_*.html` | **Xserver** `giver.work/30day/` | 別 |
| 営業切り返しマスター | `sales_faq.html` | **Xserver** `giver.work/sales-faq/` | 別 |
| GAS バックエンド | `*.js` 全部 | **GAS Apps Script** (同一プロジェクト) | - |

### 共有エンドポイント@665について

**`AKfycbw2tvPqcuJttb09OuuCDKvi5mQMwcCDqJLFRPJk3pc4w0IIAOyDPEPTRnUKPrMDPgGE4A`** は**4システム共有**:
- post-app
- 営業ダッシュボード
- ガーディアン
- 学生ダッシュボード経由

→ **このURL @665 がぶっ壊れると4システム全部死亡する**。最重要エンドポイント。

---

## 🚨 デプロイ運用ルール（最重要・絶対遵守）

### 鉄則A: GAS デプロイ後は必ずアクセス権限を手動で再設定

`clasp deploy -i <既存デプロイID> -d "..."` を実行すると、**Google Apps Script の仕様により、以下の2項目が初期値にリセットされる**:

- 「次のユーザーとして実行」→ デフォルト「ウェブ アプリケーションにアクセスしているユーザー」
- 「アクセスできるユーザー」→ デフォルト「Google アカウントを持つ全員」

**これだと無認証ユーザー（受講生）からは認証要求HTML（302→Googleログイン画面）が返ってしまい、API全停止する。**

#### clasp deploy 後の必須手順

```
1. cd ~/eisei && clasp open  ← GASエディタを開く
2. 右上「デプロイ」→「デプロイを管理」
3. 対象デプロイ（@xxx）を選択
4. 右上の ✏️ 鉛筆マークをクリック（編集モードに入る）
5. 「次のユーザーとして実行」 → 「自分（kuta310k@gmail.com）」に変更
6. 「アクセスできるユーザー」 → 「全員」を選択 ← 「Google アカウントを持つ全員」ではない
7. 右下の青「デプロイ」ボタンを押す
8. 「デプロイが更新されました」を確認
9. ターミナルで scripts/deploy-check.sh を実行 → 全システム疎通確認
```

#### 「全員」と「Google アカウントを持つ全員」の違い

| 設定 | 動作 |
|---|---|
| **全員** | 認証なしで誰でもアクセス可能 ← これが正解 |
| Google アカウントを持つ全員 | Googleにログインが必要、しかも受講生が異なるアカウントだと毎回認証要求 ← NG |
| 自分のみ | スクリプトオーナーだけ ← NG（受講生使えない） |

### 鉄則B: Xserver 配信HTMLは別途デプロイが必要

post-app.html / Dashboard-wp.html / guardian-dashboard.html などは **Xserver にscp配置されてる**。
ローカル編集→clasp pushしてもブラウザに反映されない（GASに上がるだけ）。

#### Xserver 配信HTML一覧

| ローカルファイル | Xserver パス |
|---|---|
| `post-app.html` | `xserver:/home/kodaidai/giver.work/public_html/post-app/index.html` |
| `30day_program.html` | `xserver:/home/kodaidai/giver.work/public_html/30day/index.html` |
| 他は `deploy-30day.sh` 参照 | |

#### 配信スクリプト

- 30日間講座系: `./deploy-30day.sh`
- post-app: `./deploy-postapp.sh`

### 鉄則C: 共有エンドポイント @665 を触る時は4システム全部チェック

@665 の変更は4システムに影響する。デプロイ後は必ず `scripts/deploy-check.sh` を実行して全システム疎通を確認する。

---

## 過去の事故事例

### 2026-05-17 (今日): clasp deploy で4システム全停止
- 商談数機能追加のため `clasp deploy -i ...PgGE4A` 実行
- アクセス権限がデフォルトにリセット → 投稿本数アプリ・営業ダッシュボード・ガーディアン・学生ダッシュボード経由 が全部認証要求ページを返す状態に
- 受講生から「ログインできない」「投稿入力できない」「通信エラー」報告
- 復旧: GAS Editor で「自分」+「全員」を手動再設定（10分間障害）
- 教訓: **clasp deploy 後は権限が必ずリセットされる** → 直後に必ず手動再設定

### 2026-05-02: プッシュ数シートのメタ列構造を確認せず破壊
- 詳細は `~/.claude/CLAUDE.md` 「過去の事故例」参照

---

## このリポジトリで作業する時のチェックリスト

### コード変更前
- [ ] `clasp pull` で最新を取得（別端末・別セッションでpush済みかも）
- [ ] `git status` で未コミット変更確認
- [ ] 影響範囲を把握（共有エンドポイント @665 を触るか？）

### コード変更後
- [ ] ローカル構文チェック: `node --check 〇〇.js`
- [ ] HTMLのJS構文チェック (script タグ抽出してnew Function)
- [ ] `clasp push` で GAS に反映

### デプロイ時（@665 を更新する場合）
- [ ] 「変更は1回に1個」を守る（複数機能を1デプロイにまとめない）
- [ ] `clasp deploy -i AKfycbw2tvPqcuJttb09OuuCDKvi5mQMwcCDqJLFRPJk3pc4w0IIAOyDPEPTRnUKPrMDPgGE4A -d "..."`
- [ ] **GAS Editorで権限再設定（鉄則A）**
- [ ] `scripts/deploy-check.sh` で全システム疎通確認
- [ ] Xserver配信HTMLも変更してたら `./deploy-postapp.sh` 等で配置
- [ ] ブラウザで実際に各システムを開いて動作確認

### コミット時
- [ ] feature ブランチを切る（main直push禁止）
- [ ] PR作成 → mainへのマージは Namaka 確認後
