// ============================================
// PaymentAggregator.js
// 経理が貼り付けた一次データ（ユニヴァ・ライフティ・銀振）から
// 商談者別 × 商談日 の着金額を集計し、「集計済み」タブに書き出す
//
// 着金の起点は「商談日（マスターのプッシュ日時）」：
// - 4月商談 → 5月着金でも「4月の着金」として計上
// - 取引日（決済日/完了日/入金日）ではなく、商談日で集計
// ============================================

var PA_MASTER_ID = '1KxHeLmrpdaw1IUhBaQ46UWSHu-8SCRZqcrHOE2hMwDo';
var PA_TAB_MASTER = '👑商談マスターデータ';
var PA_TAB_UNIVA  = 'ユニヴァ';
var PA_TAB_LIFTY  = 'ライフティ';
var PA_TAB_BANK   = '銀振';
var PA_TAB_AGG    = '集計済み';      // 当月商談分
var PA_TAB_PAST   = '過去分割';      // 過去成約者の今月入金分

// マスター本体カラム（0-based）
var PA_COL_PUSH    = 2;   // プッシュ日時
var PA_COL_SALES   = 3;   // 商談者名
var PA_COL_LINE    = 4;   // LINE名
var PA_COL_STATUS  = 11;  // 成約状況
var PA_COL_REVENUE = 16;  // 売上（万円）
var PA_COL_NAME    = 17;  // 顧客名（漢字フルネーム）
var PA_COL_EMAIL   = 34;  // 契約アドレス

// 手数料率
var PA_FEE_RATE = {
  'ユニバペイ':  0.042,
  'ユニヴァペイ': 0.042,
  'ユニヴァ':    0.042,
  'ライフテイ':  0.05,
  'ライフティ':  0.05,
  'MOSH':       0.06,
  'CBS':        0.05,
  '銀振':       0,
  '銀行振込':   0
};

// 銀振の決済代行・除外キーワード（顧客直接振込以外）
var PA_BANK_PROCESSOR_KW = [
  'ユニヴアペイ', 'ユニバペイ', 'ライフテイ', 'ジ－エムオ－', 'シ－ビ－エス',
  'ＧＭＯアオゾラ', 'GMO', 'MOSH', '振込手数料', 'Visaデビット', 'ｾｿﾞﾝ',
  '振込資金返却'
];

/**
 * メインエントリ：トリガーから1日1回呼ばれる
 *
 * 動作仕様：
 * - 一次データタブ（ユニヴァ・ライフティ・銀振）には経理が直近データを毎朝貼り付け
 * - 各取引を商談者にマッチ → その商談者の「商談日」に着金として計上
 * - 対象月の商談日 → 「集計済み」タブ
 * - 対象月以外の商談日 → 「過去分割」タブ（過去成約者の今月入金など）
 * - 両タブとも (商談者, 商談日) 単位で upsert（履歴保持）
 *
 * @param {string} [targetMonth] - 'YYYY-MM' 形式。省略時は実行時の当月（JST）
 */
function aggregatePrimaryData(targetMonth) {
  var ss = SpreadsheetApp.openById(PA_MASTER_ID);
  if (!targetMonth) {
    targetMonth = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
  }
  Logger.log('=== aggregatePrimaryData 開始: ' + new Date().toISOString() + ' / 対象月: ' + targetMonth);

  // マスター本体から商談済み顧客リスト（テスト除く全件）
  var seiyaku = readSeiyakuFromMaster_(ss);
  Logger.log('商談者数（マスター登録、テスト除外）: ' + seiyaku.length);

  // 一次データ読み込み
  var univaTxs = readUnivaTab_(ss);
  var liftyTxs = readLiftyTab_(ss);
  var bankTxs  = readBankTab_(ss);
  Logger.log('一次データ: ユニヴァ' + univaTxs.length + '件 / ライフティ' + liftyTxs.length + '件 / 銀振' + bankTxs.length + '件');

  // インデックス作成
  var indexes = buildSeiyakuIndexes_(seiyaku);

  // 商談者×商談日で全件集計（フィルタなし）
  var aggResult = aggregateBySalesAndPushDate_(univaTxs, liftyTxs, bankTxs, indexes);
  var aggregated = aggResult.result;
  var stats = aggResult.stats;

  // マッチ統計をログ出力
  Logger.log('--- マッチ統計 ---');
  ['univa', 'lifty', 'bank'].forEach(function(k) {
    var s = stats[k];
    var label = (k === 'univa' ? 'ユニヴァ' : k === 'lifty' ? 'ライフティ' : '銀振');
    Logger.log('  ' + label + ': ' + s.total + '件 → マッチ ' + s.match + '件 ¥' + Math.round(s.matchAmt).toLocaleString() + ' / 不一致 ' + s.miss + '件 ¥' + Math.round(s.missAmt).toLocaleString());
  });

  // 当月商談と過去商談に振り分け
  var split = splitByTargetMonth_(aggregated, targetMonth);
  var currentDates = collectPushDates_(split.current);
  var pastDates = collectPushDates_(split.past);
  Logger.log('当月(' + targetMonth + ')商談日: ' + currentDates.length + '件');
  Logger.log('過去商談日: ' + pastDates.length + '件');

  // 各タブに upsert
  writeSheet_(ss, PA_TAB_AGG, split.current, currentDates);
  writeSheet_(ss, PA_TAB_PAST, split.past, pastDates);

  Logger.log('=== aggregatePrimaryData 完了');
  return {
    ok: true,
    targetMonth: targetMonth,
    currentDateCount: currentDates.length,
    pastDateCount: pastDates.length,
    salesCount: countSales_(aggregated)
  };
}

/**
 * 集計結果を当月（targetMonth で始まる商談日）と過去に振り分け
 */
function splitByTargetMonth_(aggregated, targetMonth) {
  var current = {};
  var past = {};
  for (var sales in aggregated) {
    for (var d in aggregated[sales]) {
      var bucket = (d.indexOf(targetMonth) === 0) ? current : past;
      if (!bucket[sales]) bucket[sales] = {};
      bucket[sales][d] = aggregated[sales][d];
    }
  }
  return { current: current, past: past };
}

function collectPushDates_(aggregated) {
  var set = {};
  for (var sales in aggregated) {
    for (var d in aggregated[sales]) set[d] = true;
  }
  var arr = [];
  for (var d in set) arr.push(d);
  arr.sort();
  arr.reverse();
  return arr;
}

/**
 * マスター本体から商談済み顧客リストを読み取る
 * - 商談者・顧客名が両方ある全行を対象（成約/CO/失注/継続すべて含む）
 * - 「テスト」のみ除外
 * - status は保持（後段で必要に応じて使用）
 *
 * 失注やCOも対象にする理由：
 * 失注後の返金や、デポジット課金の途中で離脱したケースでも
 * 一次データには取引が現れる。商談記録としてマスター登録があるなら
 * 紐付けて事業内取引として扱う
 */
function readSeiyakuFromMaster_(ss) {
  var sheet = ss.getSheetByName(PA_TAB_MASTER);
  if (!sheet) throw new Error('マスター本体タブが見つかりません: ' + PA_TAB_MASTER);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var lastCol = Math.max(sheet.getLastColumn(), PA_COL_EMAIL + 1);
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var out = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var status = String(row[PA_COL_STATUS] || '').trim();
    if (status === 'テスト') continue;

    var sales = String(row[PA_COL_SALES] || '').trim();
    var name = String(row[PA_COL_NAME] || '').trim();
    if (!sales || !name) continue;

    out.push({
      push: String(row[PA_COL_PUSH] || '').trim(),
      pushDate: pushDateKey_(row[PA_COL_PUSH]),  // YYYY-MM-DD
      sales: sales,
      name: name,
      line: String(row[PA_COL_LINE] || '').trim(),
      email: String(row[PA_COL_EMAIL] || '').toLowerCase().trim(),
      status: status
    });
  }
  return out;
}

/**
 * セル値（Date or 文字列）から "YYYY-MM-DD" を抽出
 */
function pushDateKey_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  var s = String(v || '').trim();
  var m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  return '';
}

/**
 * 「ユニヴァ」タブ読み取り
 * 期待カラム: ID, 日付, 店舗, 支払い詳細, 金額, タイプ, モード, ステータス
 */
function readUnivaTab_(ss) {
  var sheet = ss.getSheetByName(PA_TAB_UNIVA);
  if (!sheet) { Logger.log('ユニヴァタブなし'); return []; }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();

  var out = [];
  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    var status = String(r[7] || '').trim();
    if (status && status !== '成功') continue;
    var typ = String(r[5] || '').trim();
    var amt = parsePAAmount_(r[4]);
    if (amt === null) continue;
    if (typ === '返金') amt = -Math.abs(amt);
    var detail = String(r[3] || '').trim();
    var parsed = parsePAUnivaDetail_(detail);
    var dateStr = parsePAUnivaDate_(r[1]);
    if (!dateStr) continue;
    out.push({
      date: dateStr,
      name: parsed.name,
      email: parsed.email,
      amt: amt,
      type: typ
    });
  }
  return out;
}

/**
 * 「ライフティ」タブ読み取り
 * 期待カラム: 申込ID, 申込日時, 加盟店名, 加盟店支店名, 担当者, 加盟店顧客ID,
 *           申込者氏名, 金額, 回数, 承認番号, お客様ﾀﾞｳﾝﾛｰﾄﾞ日時, 受付ｽﾃｰﾀｽ, ...
 */
function readLiftyTab_(ss) {
  var sheet = ss.getSheetByName(PA_TAB_LIFTY);
  if (!sheet) { Logger.log('ライフティタブなし'); return []; }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 17).getValues();

  var out = [];
  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    var receipt = String(r[11] || '').trim();
    if (receipt && receipt !== '完了') continue;
    var amt = parsePAAmount_(r[7]);
    if (amt === null || amt <= 0) continue;
    var name = String(r[6] || '').trim();
    if (!name) continue;
    var applyDate = parsePALiftyDate_(r[1]);
    var completeDate = parsePALiftyDate_(r[10]);
    if (!completeDate) continue;
    out.push({
      apply_date: applyDate,
      complete_date: completeDate,
      name: name,
      amt: amt,
      sales: String(r[4] || '').trim()
    });
  }
  return out;
}

/**
 * 「銀振」タブ読み取り
 * 期待カラム: 日付, 摘要, 入金金額, 出金金額, 残高, メモ
 * - 入金金額があるレコードのみ
 * - 摘要に「振込」を含む
 * - 決済代行・手数料は除外
 */
function readBankTab_(ss) {
  var sheet = ss.getSheetByName(PA_TAB_BANK);
  if (!sheet) { Logger.log('銀振タブなし'); return []; }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

  var out = [];
  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    var inAmt = parsePAAmount_(r[2]);
    if (inAmt === null || inAmt <= 0) continue;
    var desc = String(r[1] || '').trim();
    if (desc.indexOf('振込') === -1) continue;
    if (desc.indexOf('手数料') !== -1) continue;
    var isProcessor = false;
    for (var k = 0; k < PA_BANK_PROCESSOR_KW.length; k++) {
      if (desc.indexOf(PA_BANK_PROCESSOR_KW[k]) !== -1) { isProcessor = true; break; }
    }
    if (isProcessor) continue;
    var dateStr = parsePABankDate_(r[0]);
    if (!dateStr) continue;
    var nameKana = desc.replace(/^振込\s+/, '').replace(/^入金\s+/, '').trim();
    out.push({
      date: dateStr,
      name_kana: nameKana,
      amt: inAmt
    });
  }
  return out;
}

/**
 * 成約者リストからマッチング用インデックスを構築
 */
function buildSeiyakuIndexes_(seiyaku) {
  var emailIdx = {};
  var nameIdx = {};
  var lineIdx = {};
  for (var i = 0; i < seiyaku.length; i++) {
    var s = seiyaku[i];
    if (s.email) emailIdx[s.email] = s;
    var n = normalizeName_(s.name);
    if (n) {
      if (!nameIdx[n]) nameIdx[n] = [];
      nameIdx[n].push(s);
    }
    if (s.name.indexOf('/') !== -1) {
      var parts = s.name.split('/');
      for (var p = 0; p < parts.length; p++) {
        var np = normalizeName_(parts[p]);
        if (np) {
          if (!nameIdx[np]) nameIdx[np] = [];
          nameIdx[np].push(s);
        }
      }
    }
    var l = normalizeName_(s.line);
    if (l) {
      if (!lineIdx[l]) lineIdx[l] = [];
      lineIdx[l].push(s);
    }
  }
  return { email: emailIdx, name: nameIdx, line: lineIdx };
}

/**
 * 商談者×商談日別の着金額集計
 * - 各取引を商談者にマッチ → その商談者の商談日に紐付けて計上
 * - 取引日（決済日/完了日/入金日）は使用しない（商談日が起点）
 */
function aggregateBySalesAndPushDate_(univaTxs, liftyTxs, bankTxs, idx) {
  var result = {};
  var stats = {
    univa: { total: univaTxs.length, match: 0, matchAmt: 0, miss: 0, missAmt: 0 },
    lifty: { total: liftyTxs.length, match: 0, matchAmt: 0, miss: 0, missAmt: 0 },
    bank:  { total: bankTxs.length,  match: 0, matchAmt: 0, miss: 0, missAmt: 0 }
  };
  function ensure(sales, pushDate) {
    if (!result[sales]) result[sales] = {};
    if (!result[sales][pushDate]) result[sales][pushDate] = { univa: 0, lifty: 0, bank: 0, total: 0, count: 0 };
    return result[sales][pushDate];
  }

  // ユニヴァ
  for (var i = 0; i < univaTxs.length; i++) {
    var tx = univaTxs[i];
    var s = matchUnivaSeiyaku_(tx, idx);
    if (!s || !s.pushDate) {
      stats.univa.miss++;
      stats.univa.missAmt += tx.amt;
      continue;
    }
    stats.univa.match++;
    stats.univa.matchAmt += tx.amt;
    var bucket = ensure(shortSales_(s.sales), s.pushDate);
    bucket.univa += tx.amt;
    bucket.total += tx.amt;
    bucket.count += 1;
  }

  // ライフティ
  for (var i = 0; i < liftyTxs.length; i++) {
    var tx = liftyTxs[i];
    var s = matchLiftySeiyaku_(tx, idx);
    if (!s || !s.pushDate) {
      stats.lifty.miss++;
      stats.lifty.missAmt += tx.amt;
      continue;
    }
    stats.lifty.match++;
    stats.lifty.matchAmt += tx.amt;
    var bucket = ensure(shortSales_(s.sales), s.pushDate);
    bucket.lifty += tx.amt;
    bucket.total += tx.amt;
    bucket.count += 1;
  }

  // 銀振
  for (var i = 0; i < bankTxs.length; i++) {
    var tx = bankTxs[i];
    var s = matchBankSeiyaku_(tx, idx);
    if (!s || !s.pushDate) {
      stats.bank.miss++;
      stats.bank.missAmt += tx.amt;
      continue;
    }
    stats.bank.match++;
    stats.bank.matchAmt += tx.amt;
    var bucket = ensure(shortSales_(s.sales), s.pushDate);
    bucket.bank += tx.amt;
    bucket.total += tx.amt;
    bucket.count += 1;
  }

  return { result: result, stats: stats };
}

/**
 * 指定タブに書き込み（毎回全クリア→指定範囲のみ書き戻し）
 *
 * - 集計済みタブには「当月商談分のみ」、過去分割タブには「当月以外のみ」を
 *   呼び出し側で振り分けて渡す
 * - 各タブは毎回全クリアして書き込むため、過去の旧フォーマット行や
 *   想定外の行は残らない
 * - 経理が直近6ヶ月分のCSVを毎日貼る運用なら、その期間のデータは毎日再計算される
 *
 * 商談日列(A列)は強制テキスト型で書き込み（"2026-04-13" がDate型に自動変換されると
 * 比較ロジックが破綻するため）
 *
 * 金額は万円単位・小数1位に丸めて書き込み（ダッシュボードの単位に合わせる）
 */
function writeSheet_(ss, tabName, aggregated, processedDates) {
  var sheet = ss.getSheetByName(tabName);
  var headerRow = ['商談日', '商談者', '着金額', 'ユニヴァ', 'ライフティ', '銀振', '件数', '更新日時'];
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  if (!sheet) {
    sheet = ss.insertSheet(tabName);
  }

  // タブを全クリア
  sheet.clearContents();

  // ヘッダー書き込み
  sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
  sheet.setFrozenRows(1);
  sheet.getRange('A:A').setNumberFormat('@');

  // 新規データ生成（金額は万円・小数1位に丸め）
  var rows = [];
  for (var i = 0; i < processedDates.length; i++) {
    var d = processedDates[i];
    var salesList = [];
    for (var sales in aggregated) {
      if (aggregated[sales][d] && aggregated[sales][d].total !== 0) {
        salesList.push(sales);
      }
    }
    salesList.sort(function(a, b) {
      return aggregated[b][d].total - aggregated[a][d].total;
    });
    for (var j = 0; j < salesList.length; j++) {
      var sales = salesList[j];
      var b = aggregated[sales][d];
      rows.push([
        d,
        sales,
        toMan_(b.total),
        toMan_(b.univa),
        toMan_(b.lifty),
        toMan_(b.bank),
        b.count,
        now
      ]);
    }
  }

  // 商談日（新→旧）→ 着金額（多→少）でソート
  rows.sort(function(a, b) {
    if (a[0] !== b[0]) return String(b[0]).localeCompare(String(a[0]));
    return Number(b[2]) - Number(a[2]);
  });

  // 書き込み
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headerRow.length).setValues(rows);
    sheet.getRange(2, 1, rows.length, 1).setNumberFormat('@');
  }
}

/**
 * 円 → 万円（小数1位、四捨五入）
 */
function toMan_(yen) {
  return Math.round(yen / 1000) / 10;
}

// =========== ヘルパー関数 ===========

function parsePAAmount_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return v;
  var s = String(v).replace(/[¥￥,JPY円\s]/g, '').replace(/,/g, '');
  if (s === '') return null;
  var n = Number(s);
  return isNaN(n) ? null : n;
}

/**
 * ユニヴァの「支払い詳細」セルから name と email を抽出
 * 例: "misa kikuchi  manmaru-misakichi@t.vodafone.ne.jp"
 */
function parsePAUnivaDetail_(s) {
  var emailMatch = s.match(/[\w.+-]+@[\w.-]+\.[\w]+/);
  var email = emailMatch ? emailMatch[0].toLowerCase() : '';
  var name = email ? s.replace(email, '').trim() : s.trim();
  // 改行で分かれている場合の処理
  name = name.replace(/\s+/g, ' ').trim();
  return { name: name, email: email };
}

/**
 * ユニヴァの日付列を YYYY-MM-DD 文字列に変換
 * 例: "2026-04-30, 21:10:15" or Date オブジェクト
 */
function parsePAUnivaDate_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  var s = String(v || '').trim();
  var m = s.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  return null;
}

/**
 * ライフティの日付（"2026年4月30日 21:11:44" 形式）を YYYY-MM-DD 文字列に変換
 */
function parsePALiftyDate_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  var s = String(v || '').trim();
  var m = s.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  m = s.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  return null;
}

/**
 * 銀振の日付（"20260403" 形式）を YYYY-MM-DD 文字列に変換
 */
function parsePABankDate_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  var s = String(v || '').trim();
  if (/^\d{8}$/.test(s)) {
    return s.substring(0, 4) + '-' + s.substring(4, 6) + '-' + s.substring(6, 8);
  }
  var m = s.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  return null;
}

/**
 * 名前正規化：カッコ・/・スペース・全角半角を除去、小文字化
 */
function normalizeName_(s) {
  if (!s) return '';
  var x = String(s);
  x = x.replace(/\([^)]*\)/g, '');
  x = x.replace(/（[^）]*）/g, '');
  x = x.replace(/\//g, '');
  x = x.replace(/\s+/g, '');
  x = x.replace(/　/g, '');
  return x.toLowerCase().trim();
}

function matchUnivaSeiyaku_(tx, idx) {
  if (tx.email && idx.email[tx.email]) return idx.email[tx.email];
  var nl = tx.name.toLowerCase().replace(/\s+/g, '');
  for (var k in idx.line) {
    if (k.length >= 4 && (nl.indexOf(k) !== -1 || k.indexOf(nl) !== -1)) {
      return latestEntry_(idx.line[k]);
    }
  }
  return null;
}

function matchLiftySeiyaku_(tx, idx) {
  var n = normalizeName_(tx.name);
  if (idx.name[n]) return latestEntry_(idx.name[n]);
  for (var k in idx.name) {
    if (n.length >= 3 && (n.indexOf(k) !== -1 || k.indexOf(n) !== -1)) {
      return latestEntry_(idx.name[k]);
    }
  }
  return null;
}

/**
 * 同名顧客の複数商談行から、pushDate が最新のものを選ぶ
 * 過去同名顧客がいる場合に「最古商談」が選ばれて4月商談が漏れる問題を防止
 */
function latestEntry_(entries) {
  if (!entries || entries.length === 0) return null;
  if (entries.length === 1) return entries[0];
  var latest = entries[0];
  for (var i = 1; i < entries.length; i++) {
    if (entries[i].pushDate && entries[i].pushDate > (latest.pushDate || '')) {
      latest = entries[i];
    }
  }
  return latest;
}

/**
 * 銀振：カナ→漢字のヒントマップ（手動メンテ）
 */
var PA_KANA_HINTS = {
  'ヒラカタ ヒデキ': ['平方'],
  'アダチ サオリ': ['安達'],
  'ノジリ クミコ': ['野尻'],
  'コクブ アケミ': ['国分'],
  'キムラ カオリ': ['木村'],
  'カ） カノメイド': ['Kanomade'],
  'ナカノ ヨウイチ': ['中野'],
  'ミナガワ タケヒロ': ['皆川'],
  'ヤマシタ カオリ': ['山下'],
  'コイケ クミコ': ['小池'],
  'オクムラ リエ': ['奥村'],
  'ヤギ ユウコ': ['八木', '裕子'],
  'オクムラコウイチ': ['奥村', '公一'],
  'ナカガワ サユミ': ['中川'],
  'タナカヨシカズ': ['田中', '吉一'],
  'フルタ サチコ': ['古田'],
  'イトウ ユリカ': ['伊藤'],
  'フルカワ サトシ': ['古川'],
  'アリマ シホ': ['有馬'],
  'イワシマ ユミ': ['岩島'],
  'ナメリカワ タケシ': ['滑川'],
  'クスヤ ユミ': ['楠谷'],
  'タチオカ シンゴ': ['立岡'],
  'ヨシカワ サトシ': ['吉川'],
  'ハセガワ リヨウ': ['長谷川'],
  'アオキ ユウマ': ['青木'],
  'タナカ エツヨ': ['田中', '悦代'],
  'ハシモト フミコ': ['橋本', '布美子'],
  'ヤマイシ ミホ': ['山石'],
  'ウエノ ユウイチ': ['上野'],
  'オオミナト メグミ': ['大湊'],
  'マツカワ ヒロヤ': ['松川'],
  'マツモト シズカ': ['松本', '静香'],
  'オオオカ ヒロカズ': ['大岡'],
  'スギサワ ダイスケ': ['杉澤'],
  'イイダ マサノリ': ['飯田'],
  'イシイ エリコ': ['石井'],
  'オノデラカナ': ['小野寺'],
  'ヒラヤマ ユウスケ': ['平山'],
  'モリ タカアキ': ['森', '貴明']
};

function matchBankSeiyaku_(tx, idx) {
  var hints = PA_KANA_HINTS[tx.name_kana];
  if (!hints || hints.length === 0) return null;
  for (var k in idx.name) {
    var allMatch = true;
    for (var h = 0; h < hints.length; h++) {
      if (k.indexOf(hints[h].toLowerCase()) === -1) { allMatch = false; break; }
    }
    if (allMatch) return latestEntry_(idx.name[k]);
  }
  return null;
}

/**
 * 商談者名の正規化（フルネーム→姓のみ）
 */
function shortSales_(s) {
  var prefixes = ['阿部', '大久保', '新居', '伊東', '中市', '辻阪', '辻坂',
                  '吉崎', '五十嵐', '鍋嶋', '森本', '久保田', '福島', '佐々木',
                  '矢吹', '川合', '大内', '前村', 'スズカ', 'セナ', '関', '奥'];
  for (var i = 0; i < prefixes.length; i++) {
    if (s.indexOf(prefixes[i]) === 0) {
      return prefixes[i] === '辻坂' ? '辻阪' : prefixes[i];
    }
  }
  return s;
}

function countSales_(aggregated) {
  var c = 0;
  for (var k in aggregated) c++;
  return c;
}

// =========== トリガー設定（手動で1回だけ実行）===========

/**
 * 毎日朝8時にaggregatePrimaryDataを実行するトリガーを設定
 * GASエディタから手動で1回だけ実行する
 * 経理が朝7時に貼付け→8時にGAS実行
 */
function setupPaymentAggregatorTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'aggregatePrimaryData') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('aggregatePrimaryData')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();
  Logger.log('aggregatePrimaryData トリガー設定完了: 毎日朝8時');
}

// =========== 過去月集計用ヘルパー ===========
// GASエディタの関数ドロップダウンから引数なしで実行できるよう、
// よく使う過去月をハードコード関数として用意する

function aggregateApril2026()    { return aggregatePrimaryData('2026-04'); }
function aggregateMarch2026()    { return aggregatePrimaryData('2026-03'); }
function aggregateFebruary2026() { return aggregatePrimaryData('2026-02'); }
function aggregateJanuary2026()  { return aggregatePrimaryData('2026-01'); }

// =========== デバッグ用関数 ===========

/**
 * 指定日に商談したマスター行を全件ログに出し、一次データから
 * その商談者にマッチした取引を可視化する
 *
 * 例: debugByPushDate('2026-04-30')
 *
 * GASエディタでは引数なしで実行する関数として
 * debugApril30() を用意。
 */
function debugByPushDate(targetDate) {
  var ss = SpreadsheetApp.openById(PA_MASTER_ID);
  var sheet = ss.getSheetByName(PA_TAB_MASTER);
  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), PA_COL_EMAIL + 1);
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  Logger.log('=== マスター内 ' + targetDate + ' 商談行 ===');
  var hits = 0;
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var pd = pushDateKey_(row[PA_COL_PUSH]);
    if (pd === targetDate) {
      hits++;
      Logger.log('  row=' + (i + 2)
        + ' / 商談者=' + row[PA_COL_SALES]
        + ' / 顧客=' + row[PA_COL_NAME]
        + ' / line=' + row[PA_COL_LINE]
        + ' / status=' + row[PA_COL_STATUS]
        + ' / email=' + row[PA_COL_EMAIL]);
    }
  }
  Logger.log('合計: ' + hits + '件');
  Logger.log('');

  // 一次データを読み、targetDate商談者にマッチするか確認
  var seiyaku = readSeiyakuFromMaster_(ss);
  var indexes = buildSeiyakuIndexes_(seiyaku);
  var liftyTxs = readLiftyTab_(ss);
  var univaTxs = readUnivaTab_(ss);
  var bankTxs = readBankTab_(ss);

  Logger.log('=== 一次データ → ' + targetDate + ' 商談者へのマッチ ===');
  var matchCount = 0;
  for (var i = 0; i < liftyTxs.length; i++) {
    var tx = liftyTxs[i];
    var s = matchLiftySeiyaku_(tx, indexes);
    if (s && s.pushDate === targetDate) {
      matchCount++;
      Logger.log('  ライフティ: 顧客=' + tx.name + ' ¥' + tx.amt + ' → matched: ' + s.name + ' (商談者:' + s.sales + ')');
    }
  }
  for (var i = 0; i < univaTxs.length; i++) {
    var tx = univaTxs[i];
    var s = matchUnivaSeiyaku_(tx, indexes);
    if (s && s.pushDate === targetDate) {
      matchCount++;
      Logger.log('  ユニヴァ: 顧客=' + tx.name + ' / ' + tx.email + ' ¥' + tx.amt + ' → matched: ' + s.name + ' (商談者:' + s.sales + ')');
    }
  }
  for (var i = 0; i < bankTxs.length; i++) {
    var tx = bankTxs[i];
    var s = matchBankSeiyaku_(tx, indexes);
    if (s && s.pushDate === targetDate) {
      matchCount++;
      Logger.log('  銀振: 顧客=' + tx.name_kana + ' ¥' + tx.amt + ' → matched: ' + s.name + ' (商談者:' + s.sales + ')');
    }
  }
  Logger.log('マッチ合計: ' + matchCount + '件');
}

function debugApril30() { return debugByPushDate('2026-04-30'); }
