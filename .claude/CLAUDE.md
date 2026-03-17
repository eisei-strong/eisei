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

## デプロイ手順

### PHPプロキシ（api-proxy.php）
1. ローカルで編集
2. scp /Users/kodai/営業ダッシュボード/api-proxy.php xserver:/home/kodaidai/giver.work/public_html/sales-dashboard/api-proxy.php
3. キャッシュクリア: ssh xserver "rm -f /home/kodaidai/giver.work/public_html/sales-dashboard/cache/*.json"
4. 動作確認: ssh xserver "curl -s -H 'Host: giver.work' 'http://giver.work/sales-dashboard/api-proxy.php?action=api'"

### GASコード（コード.js）
1. ローカルで編集
2. cd /Users/kodai/営業ダッシュボード && clasp push
3. ※ clasp pushはexec URL（v200）には反映されない。反映するにはGASコンソールで手動デプロイが必要

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
- exec URL: https://script.google.com/macros/s/AKfycbwojGHuvzycc07FJKwBdbBJJQZpssF6lYz0DbNJlu6zsVuXkAj8V8w3XNBPieo2wsYbFg/exec
- スプレッドシートID: 1k_x3aNRTbojmhJZGMS6JGNiTNJLQR4sD5zyJCBh1YqY
- ウォーリアーズシートGID: 1235299010

### マスターCSV
- URL: https://docs.google.com/spreadsheets/d/1KxHeLmrpdaw1IUhBaQ46UWSHu-8SCRZqcrHOE2hMwDo/export?format=csv&gid=326094286
