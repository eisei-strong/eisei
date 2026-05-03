## 🚨 絶対禁止事項

### スプレッドシートへの書き込み厳禁
- GASのsetValues(), setValue(), setFormula(), clearContents() など、スプシへの書き込み処理を一切実行・追加・修正してはいけない
- syncFromOldSheet() を呼び出す処理を追加・修正してはいけない
- updateSummary() をトリガーや自動実行から呼び出す処理を追加してはいけない
- quickFixFormulas() など数式を修正するトリガーを有効化してはいけない
- スプシの数式・値・フォーマットを変更するGAS関数は、明示的に許可を得るまで一切触れてはいけない

### データ取得は読み取り専用のみ
- GAS・PHPからスプシを参照する場合はgetValues(), getValue() など読み取りのみ許可
- debugSheetByGid はスプシを読むだけなのでOK

### マスターシートへの書き込み・削除を伴う関数は必ず別シートで検証
- clearContent(), deleteRow(), setValues() 等でマスターデータを変更する関数は、**本番シートに直接実行してはいけない**
- 必ず別シート（テスト用コピー）に対して先に実行し、結果を確認してから本番に適用すること
- clearContent() → setValues() のように「消してから書く」パターンは、途中エラーでデータ消失するリスクがあるため特に注意
- 入力規則（Data Validation）がある列は、書き込み前に必ず clearDataValidations() を実行すること

## マスター管理（マスターシート = gid:326094286）

### マスタースプレッドシート
- ID: `1KxHeLmrpdaw1IUhBaQ46UWSHu-8SCRZqcrHOE2hMwDo`
- マスターシート GID: `326094286`
- フォーム回答シート GID: `260080737`

### 重複除去・ソートのルール
- 重複キー: C列(アポ日時) + D列(商談者名) + E列(相手名)
- **JSの`.sort()`は使用禁止** → `resultSheet.sort(3, true)` のみ使う
- **Date型変換をJSでやらない** → シートのソート機能に任せる
- clearContent()の前に必ず clearDataValidations() を実行
- 本番シートへの書き込みは必ず別シート（プレビュー）で確認してから

### 関数一覧
| 関数 | 役割 |
|---|---|
| fixMasterPreview() | 重複除去+タイムスタンプ修正+フォーム未転記追加+ソート+装飾 → 「マスター修正プレビュー」に出力 |
| applyFixToMaster() | プレビュー確認後に本番適用 |
| removeDuplicateRows() | 重複除去のみ → 「重複チェック結果」に出力 |
| applyDedup() | 重複チェック結果を本番適用 |
| syncFormToMaster() | フォーム未転記行をマスターに追記（1時間おき自動 or メニューから手動） |
| installSyncFormTrigger() | syncFormToMasterの定期トリガー設定（1時間おき） |

### 装飾ルール
- 成約行（L列に「成約」を含む）: 背景色 `#FF9999`（赤）
- それ以外: 背景色なし（白）
- 月またぎ: 太い黒罫線で区切り
- A列タイムスタンプ: `yyyy/MM/dd HH:mm:ss` フォーマット

### フォーム→マスター列マッピング
- フォーム[0-23] → マスター[0-23] 直接コピー
- フォーム[24] → マスター[26]（支払方法③）
- フォーム[25] → マスター[27]（代理店TikTok）
- フォーム[26] → マスター[29]（支払③金額）
- フォーム[27] → マスター[30]（契約書郵送）
- フォーム[30] → マスター[32]（着金率）
- フォーム[31] → マスター[33]（ステータス）
- フォーム[32] → マスター[34]（契約アドレス）
- フォーム[33] → マスター[35]（契約日）
- フォーム[36] → マスター[37]（日付）

## ステータス集計ルール（api-proxy.php）

| ステータス | 扱い | 成約数 | 着金 | 備考 |
|---|---|---|---|---|
| `成約` | 成約 | ✅ | ✅ | |
| `成約➔CO` | CO | ❌ | ❌ | coCount/coAmountのみ記録 |
| `成約➔キャンセル` | 失注 | ❌ | ❌ | lost++としてカウント |
| `成約➔失注` | 失注 | ❌ | ❌ | lost++としてカウント |
| `顧客情報に記入` | 継続中 | ❌ | ❌ | 部分一致（メモ付きも対応） |
| `失注` | 失注 | ❌ | ❌ | LF/CBS否決は別カウント |

- `$isExcludedStatus` は `$isCurrentMonth` の外で定義（過去月の着金除外にも使用）
- 内訳チェック: 成約 + 失注 + LF否決 + 継続 + CO = 商談数 で一致すること

## 着金列ルール（マスターシート）

着金データは以下の列に入る。**全て実際の着金（入金確認済み）として扱う**。

### フォーム由来（成約時の支払い金額）
| 列 | Index | ヘッダー | 用途 |
|---|---|---|---|
| P | 15 | 今回の支払い金額 | 支払①の着金額 |
| V | 21 | 今回の支払い②金額 | 支払②の着金額 |

### 支払スロット（分割払い・後日入金の追跡）
| 列 | Index | ヘッダー | 日付列 | 手段列 |
|---|---|---|---|---|
| AP | 41 | 着金額（万円） | AM[38] | AO[40] |
| AW | 48 | 着金額 | AT[45] | AV[47] |
| BD | 55 | 着金額 | BA[52] | BC[54] |
| BO | 66 | 着金額 | BL[63] | BN[65] |
| BV | 73 | 着金額 | BS[70] | BU[72] |
| CC | 80 | 着金額 | BZ[77] | CB[79] |
| CJ | 87 | 着金額 | CG[84] | CI[86] |

### api-proxy.php の着金計算ロジック
1. **支払スロット（AP〜CJ）を先に集計** → 金額>0なら着金計上
2. **スロットが空の成約 → P+Vでフォールバック** → P+V>0なら着金計上
3. **P+Vも0 → 契約金額フォールバック**（未着金W列=0 かつ 支払方法②なしの場合のみ）
4. 着金速報: 支払日付が空の場合はアポ日（C列）にフォールバック

## デプロイ手順

### PHPプロキシ（api-proxy.php）
1. ローカルで編集
2. scp /Users/kodai/営業ダッシュボード/api-proxy.php xserver:/home/kodaidai/giver.work/public_html/sales-dashboard/api-proxy.php
3. キャッシュクリア: ssh xserver "rm -f /home/kodaidai/giver.work/public_html/sales-dashboard/cache/*.json"
4. 動作確認: ssh xserver "curl -s -H 'Host: giver.work' 'http://giver.work/sales-dashboard/api-proxy.php?action=api'"

### GASコード（コード.js等）
1. ローカルで編集
2. cd /Users/kodai/eisei && clasp push
3. ※ `clasp push` のみでトリガーは即反映（デプロイ不要）
4. Webアプリ（exec URL）の更新が必要な場合は **kuta310k@gmail.com でGASエディタからデプロイ**

### post-app デプロイ運用ルール

#### 絶対ルール
1. **本番 deploy は `kuta310k@gmail.com`（本番権限アカウント）でしかやらない** — `clasp deploy` を namaka.hoshi@gmail.com で実行すると保存エラーになる
2. **`clasp push` と `clasp deploy` は別物** — push=HEAD更新、deploy=本番公開。pushだけでは本番反映されない
3. **サーバーHTML上書き前に必ず本番との差分確認** — ローカル版にない本番専用機能が消える事故が実際に発生した

#### GASデプロイ手順
1. ローカル修正 → `clasp push`
2. `clasp deployments` でデプロイ前のバージョン番号を記録
3. GASエディタ（kuta310k@gmail.com）で「デプロイを管理」→ **`AKfycby6qaai...qg` のデプロイ** を選ぶ（他のデプロイを更新しても意味がない）→ 鉛筆 → **必ず「新しいバージョン」を選択** → デプロイ
4. `clasp deployments` でバージョン番号が **上がったことを確認**（上がってなければデプロイ失敗）
5. API直接テスト: `curl` で postLogin → postSave（10本含む）を確認
6. ブラウザでスーパーリロード→ログアウト→再ログイン→保存確認

#### デプロイ時の注意（過去の事故パターン）
- **別のデプロイIDを更新してしまう** → 必ず末尾 `...qg` を確認
- **「新しいバージョン」を選ばずにデプロイボタンを押す** → 旧バージョンのまま反映されない
- **clasp push だけで完了と思い込む** → pushはHEAD更新のみ。exec URLには反映されない

#### サーバーHTML反映手順
1. 本番HTMLバックアップ: `ssh xserver "cp /home/kodaidai/giver.work/public_html/post-app/index.html /home/kodaidai/giver.work/public_html/post-app/index.html.bak"`
2. ローカル版と本番版を `diff` 比較
3. 既存機能確認（grep: `tabHope`, `hopeScreen`, `loadHope`, `renderHope`, `postGetHope`）
4. `scp` で反映
5. ブラウザ確認（スーパーリロード→ログアウト→再ログイン→保存→リスト数タブ）

#### 完了条件（全てOKで完了）
- ログインできる / カレンダー表示正常 / 保存できる（10本含む）
- リスト数タブが出る（該当者のみ）/ 本番URLで動く
- ログアウト→再ログイン後も保存できる
- `clasp deployments` でバージョン番号が更新されている

#### 禁止事項
- 権限確認なしで `clasp deploy`
- 差分確認なしで本番HTML丸ごと上書き
- APIテストなしで「直った」と判断
- ローカルで動いただけで本番OK判定

### GAS関数の手動実行をユーザーに案内する場合
1. Apps Scriptエディタ（https://script.google.com/）を開く
2. 左のファイル一覧から対象ファイル（例: `コード.js`）をクリック
3. 上部のドロップダウン（▶ボタンの左）で実行したい関数名を選択
4. ▶ 実行ボタンをクリック
5. ※ 初回は権限の許可が必要な場合がある

### ダッシュボードHTML（Dashboard-wp.html）
1. ローカルで編集
2. scp /Users/kodai/営業ダッシュボード/Dashboard-wp.html xserver:/home/kodaidai/giver.work/public_html/sales-dashboard/index.html

## サーバー情報

### Xserver
- SSH接続: ssh xserver
- ドメイン: giver.work
- ダッシュボードパス: /home/kodaidai/giver.work/public_html/sales-dashboard/
- キャッシュディレクトリ: /home/kodaidai/giver.work/public_html/sales-dashboard/cache/
- curl確認用: curl -s -H 'Host: giver.work' 'http://giver.work/sales-dashboard/api-proxy.php?action=api'

### GAS
- exec URL（v487、フォールバック用のみ）: https://script.google.com/macros/s/AKfycby6qaaiUoadCBnxHlUNKd-RkHxarE0WBGiitkdV0IbzL6ninM-df0FFx4SYRYVfdwcxqg/exec
- ※ 旧exec URL（AKfycbwoj...）は死亡(404)。使用禁止
- ※ `clasp deploy` 禁止。`clasp push` のみ（トリガーはHEADコードで動く）
- スプレッドシートID: 1k_x3aNRTbojmhJZGMS6JGNiTNJLQR4sD5zyJCBh1YqY
- ウォーリアーズシートGID: 1235299010

### 休日データパイプライン
- GASトリガー（1時間ごと）: `writeHolidayDataToSheet()` → CalendarAppで休日取得 → マスタースプシ「休日データ」シートに書き出し + PHPにJSON POST
- PHP受信: `api-proxy.php` POST `action=updateHoliday` → `holiday-data.json` に保存
- API応答: `holidayByDate` フィールドに休日データを含む

### マスターCSV
- URL: https://docs.google.com/spreadsheets/d/1KxHeLmrpdaw1IUhBaQ46UWSHu-8SCRZqcrHOE2hMwDo/export?format=csv&gid=326094286

## 作業完了報告のルール

### 必ずダブルチェックしてから完了報告すること
- デプロイ後は必ずcurlで実際のAPIレスポンスを確認してから「完了」と報告する
- 数値の変更を伴う作業は、変更前と変更後の値を両方確認してから報告する
- 「デプロイ済み」「修正完了」などの報告は、実際に動作確認が取れた後にのみ行う
- 確認コマンドを省略して完了報告してはいけない

### 数値ロジック変更時の最終チェック（必須）
api-proxy.phpの集計ロジックを変更した場合、デプロイ後に以下を実行:
1. マスターCSVをダウンロードしてローカルでPythonで独立計算
2. APIレスポンスと全メンバーの全フィールドを突合（deals, closed, lost, continuing, revenue, sales, shinpan, creditCard, coAmount, coRevenue, lifety, cbs 等）
3. 差異がある場合は原因を特定してから完了報告
4. 合計着金額(totalRevenue)も一致確認
- ※ 数字がずれるとボトルネック分析を間違える原因になるため、省略厳禁

## メンバー名マッピング

| 本名 | 表示名 | 別名・旧名 |
|---|---|---|
| 阿部 | AをAでやる | |
| 伊東 | ポジティブ | ドライ、勝友美、勝 |
| 久保田 | ヒトコト | |
| 辻阪 | ビッグマウス | |
| 五十嵐 | ぜんぶり | |
| 新居 | スクリプトくん | |
| 佐々木心雪 | ワントーン | 佐々木 |
| 福島 | けつだん | |
| 大久保友佑悟 | ゴン | 大久保 |
| 矢吹友一 | トニー | 矢吹 |
| 川合 | リヴァイ | 退職済み・ダッシュボード除外 |
| 大内 | ガロウ | 退職済み・ダッシュボード除外 |

## UIルール

- スマホ表示で一文が改行されないようにフォントサイズ・余白を調整すること
- **UI変更（CSS/HTML/SVG/レイアウト/演出）を入れたら毎回スマホ最適化までセットで実施すること**
  - PCで動くだけで完了報告しない。`@media (max-width: 768px)` でスマホ調整を必ず追加
  - 対象: アイコンサイズ/フォント/余白/marker・SVG要素サイズ/アニメ強度/横はみ出し
  - スマホで見たら崩れてた → 直してから出す。「言われたら直す」じゃなくデフォルト

## アイコン画像

### 取得元
- Chatworkルーム `rid349937583`（営業全員ルーム）から取得
- API: `GET https://api.chatwork.com/v2/rooms/349937583/members` （`X-ChatWorkToken` ヘッダ）
- Avatar URLは `avatar_image_url` フィールド

### ローカル保管: `~/eisei/icons/`
### 本番配信: `https://giver.work/sales-dashboard/icons/<name>.png`

### 5月以降のメンバー15人マッピング（2026-05-02更新）

| 表示名 | 本名 | Chatwork名 | ファイル名 | Chatwork avatar |
|---|---|---|---|---|
| 意思決定 | 阿部 | 人を巻き込んだ意思決定を1日3回... | `ishikettei.png` | `w7zBRgQg7l.png` |
| ポジティブ | 伊東 | 2〜5AM休み【ポジティブ】... | `positive.png` | `372J2Gw875.png` |
| ヒトコト | 久保田 | ヒトコト「どんな時も一言で言う」 | `hitokoto.png` | `2Akb3xE9q0.rsz.png` |
| ありのまま | 辻阪 | ありのままを捨てる | `arinomama.png` | `Oqaob3GO76.png` |
| ぜんぶり | 五十嵐 | ぜんぶり【Namaka全振り】 | `zenburi.png` | `VqPDr9Vwqg.png` |
| 1日1more | 新居 | more(伸び代)を1回は言うくん セナ | `ichinichi1more.png` | `372J8ve875.rsz.png` |
| けつだん | 福島 | 【けつだん】/結果出すために手段を選ばない | `ketsudan.png` | `4MlnyN9eA5.png` |
| 言い切り | 大久保 | 言い切り「説得する時は断定」ゴン | `iikiri.png` | `374Bk2XBqn.png` |
| 週1休みくん | 矢吹 | 週1休みくん（4/未定 | `shu1yasumi.png` | `JqnWXrb4AD.png` |
| ゴジータ | 吉崎 | ゴジータ5/30までに着金1500万 | `gojita.png` | `374BdNVXqn.png` |
| 夜神月 | 中市 | 夜神月 | `yagami.png` | `d7gaP1pKqp.png` |
| サンウォン | 笹山楓太 | サンウォン【楓太】 | `sungwon.png` | `EMyJ9da8qB.png` |
| 司波 | 坂野宙輝 | 司波達也【坂野】 | `shiba.png` | `Vq3W22Jbql.png` |
| 信 | 吉田羚虹 | 信　(吉田 羚虹) | `shin.png` | `372J83k975.png` |
| 悟空 | 荒木泰人 | 悟空 | `goku.png` | `dqvokz6JAz.rsz.png` |

### 5月から非表示
- L（鍋嶋・蒙恬）: 営業外れた
- スマイル（佐々木心雪）: 営業外れた
- ※ 4月以前のダッシュボードでは引き続き表示

### 新規メンバーの稼働開始日
| 表示名 | 稼働開始 |
|---|---|
| 悟空（荒木泰人） | 2026/03/28 |
| 信（吉田羚虹） | 2026/04/12 |
| 司波（坂野宙輝） | 2026/04/17 |
| サンウォン（笹山楓太） | 2026/05/01 |

※ 4月以前にデータがあっても表示しない（5月以降のみ表示）。CLAUDE.mdの「5月以前は同じにしないで」方針に従う。

## 新しいPCでのセットアップ手順

### 前提
- Node.js がインストール済み
- Git がインストール済み
- Claude Code がインストール済み

### 1. リポジトリクローン
```bash
git clone https://github.com/eisei-strong/eisei.git
cd eisei
```

### 2. clasp（GASデプロイツール）セットアップ
```bash
npm install -g @google/clasp
clasp login
```
- ブラウザが開くので、GASプロジェクトのオーナーアカウントでログイン
- `.clasp.json` はリポジトリに含まれている（scriptId設定済み）
- GAS Script ID: `1bu0jv_4kJ-ht9xByW43EWH2Acw01GIu8JP5B1t-2XTl0qQAsfX6cCbtg`

### 3. Xserver SSH接続セットアップ
```bash
# SSH鍵を配置（既存PCから ~/.ssh/xserver_key をコピー）
chmod 600 ~/.ssh/xserver_key

# SSH config に追加
cat >> ~/.ssh/config << 'EOF'
Host xserver
  HostName sv3032.xserver.jp
  Port 10022
  User kodaidai
  IdentityFile ~/.ssh/xserver_key
  StrictHostKeyChecking no
EOF
chmod 600 ~/.ssh/config

# 接続テスト
ssh xserver "echo 'SSH OK'"
```

### 4. Git ユーザー設定（リポジトリ単位）
```bash
git config user.name "こーだい"
git config user.email "kodai@Mac-Studio-2.local"
```

### 5. 動作確認
```bash
clasp push --force   # GASにコードをpush
clasp deployments    # デプロイ一覧を確認
ssh xserver "ls /home/kodaidai/giver.work/public_html/sales-dashboard/"  # Xserver接続確認
```

### 主要デプロイID
| 用途 | デプロイID（末尾） | 備考 |
|---|---|---|
| 営業ダッシュボード + PostApp API + 経費API + ガーディアン | `...qg` (AKfycby6qaai...) | メインのexec URL（本番フロントは全てこれを叩く） |
| 万バズ台本API | `...J-` (AKfycbzZedrv...) | @200固定 |
| 経費API（旧・冗長系） | `...Hw` (AKfycbw75mQ4...) | @513 ※過去に経費専用として立てたデプロイ。現在は本番フロント (giver.work/expense-dashboard/) は `...qg` を叩いている。撤去保留・障害時の切り戻し先として残置 |

### 注意事項
- `clasp deploy` で新バージョンを作ると**Webアプリの再認可が必要**になる場合がある
- 再認可が必要な場合: GASエディタ（https://script.google.com/home/projects/1bu0jv_4kJ-ht9xByW43EWH2Acw01GIu8JP5B1t-2XTl0qQAsfX6cCbtg/edit）からデプロイを管理→該当デプロイを編集→新しいバージョン→デプロイ→認可承認
- **重要: `clasp deploy` 禁止** — 必ず kuta310k@gmail.com でGASエディタからデプロイ（上記「post-app デプロイ運用ルール」参照）
- GASのバージョン上限は200。超えた場合はGASエディタから古いバージョンを削除してからデプロイ
- `appsscript.json` のOAuthスコープを変更すると再認可が**必ず**必要。v582互換のスコープ（documents/calendar.readonly なし）を維持すること
- SSH鍵（xserver_key）は**絶対にGitHubにプッシュしない**（.gitignoreに含まれていないので注意）
