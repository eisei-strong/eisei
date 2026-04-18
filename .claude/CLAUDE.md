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
4. Webアプリ（exec URL）の更新が必要な場合のみ `clasp deploy -i <デプロイID>` を実行

### ⚠️ clasp deploy 禁止 → GASエディタからデプロイすること
- GASのWebアプリは `executeAs: USER_DEPLOYING`（デプロイした人の権限で実行）
- `clasp deploy` すると**デプロイユーザーが namaka.hoshi@gmail.com に変わり、スプシの書き込み権限がないため保存エラーになる**
- **必ず `kuta310k@gmail.com` でGASエディタからデプロイすること**
- 手順: ① `clasp push` でHEADを更新 → ② GASエディタ（kuta310k@gmail.com）で「デプロイを管理」→ 鉛筆アイコン → 「新しいバージョン」→ デプロイ
- ※ `clasp push` はOK（HEADコード更新のみ、デプロイユーザーに影響なし）

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

## アイコン画像
- スクリプトくん: https://giver.work/sales-dashboard/icons/script-kun.png（元画像: /Users/kodai/Downloads/4AVg98WR7V.png）
- 他メンバー: Google Drive（lh3.googleusercontent.com）経由

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
| 営業ダッシュボード + PostApp API | `...qg` (AKfycby6qaai...) | メインのexec URL |
| 万バズ台本API | `...J-` (AKfycbzZedrv...) | @200固定 |
| 経費API | `...Hw` (AKfycbw75mQ4...) | @513 |

### 注意事項
- `clasp deploy` で新バージョンを作ると**Webアプリの再認可が必要**になる場合がある
- 再認可が必要な場合: GASエディタ（https://script.google.com/home/projects/1bu0jv_4kJ-ht9xByW43EWH2Acw01GIu8JP5B1t-2XTl0qQAsfX6cCbtg/edit）からデプロイを管理→該当デプロイを編集→新しいバージョン→デプロイ→認可承認
- **重要: `clasp deploy` 禁止** — デプロイユーザーが namaka.hoshi@gmail.com に変わり書き込みエラーになる（過去に複数回発生）。必ず kuta310k@gmail.com でGASエディタからデプロイすること
- GASのバージョン上限は200。超えた場合はGASエディタの「プロジェクトの設定」から古いバージョンを削除してからデプロイ
- `appsscript.json` のOAuthスコープを変更すると再認可が**必ず**必要。v582互換のスコープ（documents/calendar.readonly なし）を維持すること
- SSH鍵（xserver_key）は**絶対にGitHubにプッシュしない**（.gitignoreに含まれていないので注意）
