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
var PA_TAB_MOSH   = 'MOSH';
var PA_TAB_AGG    = '集計済み';        // マスター紐付け済み全月
var PA_TAB_PAST   = '過去分割';        // 廃止（互換のためタブ自体は残す）
var PA_TAB_UNMATCHED = 'マスター登録漏れ'; // 事業内（ユニヴァ・ライフティ・MOSH）の不一致一覧
var PA_FORM_GID   = 260080737;       // フォーム回答シートのGID（マスター本体未転記の商談記録）

// マスター本体カラム（0-based）
var PA_COL_PUSH    = 2;   // プッシュ日時
var PA_COL_SALES   = 3;   // 商談者名
var PA_COL_LINE    = 4;   // LINE名
var PA_COL_STATUS  = 11;  // 成約状況
var PA_COL_REVENUE = 16;  // 売上（万円）
var PA_COL_NAME       = 17;  // 顧客名（漢字フルネーム）
var PA_COL_SUPPLEMENT = 23;  // 支払い予定の補足（X列、契約者別人時の備考）
var PA_COL_EMAIL      = 34;  // 契約アドレス（マスター本体）

// フォーム回答シートのカラム（CLAUDE.md「フォーム→マスター列マッピング」より）
// フォーム[0-23] → マスター[0-23] 直接コピー、フォーム[32] → マスター[34]（契約アドレス）
var PA_FORM_COL_EMAIL = 32;  // フォーム上の契約アドレス

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
  // targetMonth は互換のため受け取るが、もう絞り込みには使わない
  // 集計済みタブには全期間（マスターに紐付くもの全部）を書き込む
  Logger.log('=== aggregatePrimaryData 開始: ' + new Date().toISOString() + (targetMonth ? ' / 旧引数: ' + targetMonth + ' (無視)' : ''));

  // マスター本体 + フォーム回答シート から商談済み顧客リスト（テスト除く、重複排除）
  var seiyaku = readAllSeiyaku_(ss);
  Logger.log('商談者数（マスター本体+フォーム回答、テスト除外）: ' + seiyaku.length);

  // 一次データ読み込み
  var univaTxs = readUnivaTab_(ss);
  var liftyTxs = readLiftyTab_(ss);
  var bankTxs  = readBankTab_(ss);
  var moshTxs  = readMoshTab_(ss);
  Logger.log('一次データ: ユニヴァ' + univaTxs.length + '件 / ライフティ' + liftyTxs.length + '件 / 銀振' + bankTxs.length + '件 / MOSH' + moshTxs.length + '件');

  // インデックス作成
  var indexes = buildSeiyakuIndexes_(seiyaku);

  // 商談者×商談日で全件集計
  var aggResult = aggregateBySalesAndPushDate_(univaTxs, liftyTxs, bankTxs, moshTxs, indexes);
  var aggregated = aggResult.result;
  var stats = aggResult.stats;

  // マッチ統計をログ出力
  Logger.log('--- マッチ統計 ---');
  ['univa', 'lifty', 'mosh', 'bank'].forEach(function(k) {
    var s = stats[k];
    var label = (k === 'univa' ? 'ユニヴァ' : k === 'lifty' ? 'ライフティ' : k === 'mosh' ? 'MOSH' : '銀振');
    Logger.log('  ' + label + ': ' + s.total + '件 → マッチ ' + s.match + '件 ¥' + Math.round(s.matchAmt).toLocaleString()
      + ' / 不一致 ' + s.miss + '件 ¥' + Math.round(s.missAmt).toLocaleString()
      + ' / CO除外 ' + s.coSkip + '件 ¥' + Math.round(s.coSkipAmt).toLocaleString());
  });

  // 全マッチ商談日を集計済みタブに書き出し
  var allDates = collectPushDates_(aggregated);
  Logger.log('マッチ商談日: ' + allDates.length + '件');

  writeSheet_(ss, PA_TAB_AGG, aggregated, allDates);

  // 事業内不一致（ユニヴァ・ライフティ・MOSH）を「マスター登録漏れ」タブに書き出す
  writeUnmatchedSheet_(ss, univaTxs, liftyTxs, moshTxs, indexes);

  // 過去分割タブは廃止（クリアして注記行のみ残す）
  var pastSheet = ss.getSheetByName(PA_TAB_PAST);
  if (pastSheet) {
    pastSheet.clearContents();
    pastSheet.getRange(1, 1).setValue('（このタブは廃止されました。集計済みタブに全月分が書き込まれています）');
  }

  Logger.log('=== aggregatePrimaryData 完了');
  return {
    ok: true,
    matchedDateCount: allDates.length,
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

    // X列補足に「契約者は○○」がある場合、契約者を別エントリ（エイリアス）として追加
    var contractor = extractContractor_(row[PA_COL_SUPPLEMENT]);
    if (contractor) {
      out.push({
        push: String(row[PA_COL_PUSH] || '').trim(),
        pushDate: pushDateKey_(row[PA_COL_PUSH]),
        sales: sales,
        name: contractor,
        line: contractor,
        email: '',
        status: status,
        _source: 'contractor_alias',
        _origName: name
      });
    }
    // X列補足のローマ字決済名義も追加（ユニヴァ name とのmatching用）
    var romaji = extractRomajiName_(row[PA_COL_SUPPLEMENT]);
    if (romaji) {
      out.push({
        push: String(row[PA_COL_PUSH] || '').trim(),
        pushDate: pushDateKey_(row[PA_COL_PUSH]),
        sales: sales,
        name: romaji,
        line: romaji,
        email: '',
        status: status,
        _source: 'romaji_alias',
        _origName: name
      });
    }
    // X列補足の決済者email（守口信一タイプ：契約者と別人の決済者emailが書いてある）
    var primaryEmail = String(row[PA_COL_EMAIL] || '').toLowerCase().trim();
    var emailAliases = extractEmailAliases_(row[PA_COL_SUPPLEMENT]);
    for (var ea = 0; ea < emailAliases.length; ea++) {
      var aliasEmail = emailAliases[ea];
      if (aliasEmail === primaryEmail) continue;
      out.push({
        push: String(row[PA_COL_PUSH] || '').trim(),
        pushDate: pushDateKey_(row[PA_COL_PUSH]),
        sales: sales,
        name: name,
        line: String(row[PA_COL_LINE] || '').trim(),
        email: aliasEmail,
        status: status,
        _source: 'email_alias',
        _origName: name
      });
    }
  }
  return out;
}

/**
 * 「支払い予定の補足」セルから「契約者：○○」のような契約者名を抽出
 * 例: '⚠️補足：契約者は「石田　葵花」様 ...' → '石田　葵花'
 */
function extractContractor_(supplement) {
  if (!supplement) return null;
  var s = String(supplement);
  var m = s.match(/契約者[はは:：]\s*[「『]([^」』]+)[」』]/);
  if (m) return m[1].trim();
  m = s.match(/契約者[はは:：]\s*([^。、，,\n]{2,20}?)[様氏]/);
  if (m) return m[1].trim();
  return null;
}

/**
 * 「支払い予定の補足」セルからローマ字名を抽出（決済名義）
 * 例: '15万ユニバペイ着金（AOI ISHIDA名義）' → 'aoi ishida'
 *
 * ユニヴァ管理画面の name はローマ字なので、X列補足にローマ字決済名義が
 * あれば matching用エイリアスとして登録できる
 */
function extractRomajiName_(supplement) {
  if (!supplement) return null;
  var s = String(supplement);
  // 「○○ ○○名義」パターン（ローマ字大文字、全角/半角スペース許容）
  var m = s.match(/([A-Z]{2,}[\s　]+[A-Z]{2,})[\s　]*名義/);
  if (m) return m[1].toLowerCase().replace(/[\s　]+/g, ' ').trim();
  return null;
}

/**
 * 「支払い予定の補足」セルから決済者emailを抽出（守口信一タイプ対応）
 * 例: 'クレカで全額着金 守口信一s.moriguchi@accord-m.co.jp 3/10✅'
 *      → ['s.moriguchi@accord-m.co.jp']
 *
 * 契約者本人とは別人の決済者emailが補足に書かれているケースを救済
 */
function extractEmailAliases_(supplement) {
  if (!supplement) return [];
  var s = String(supplement);
  var out = [];
  var matches = s.match(/[\w.+-]+@[\w.-]+\.[\w]+/g);
  if (matches) {
    for (var i = 0; i < matches.length; i++) {
      out.push(matches[i].toLowerCase());
    }
  }
  return out;
}

/**
 * フォーム回答シート（gid=260080737）から商談記録を読み取る
 * 商談者がフォーム送信したがマスター本体に未転記の商談を拾う
 *
 * フォーム→マスター 列マッピング (CLAUDE.md)：
 * - col[0-23] 直接コピー
 * - col[32] → マスター[34]（契約アドレス）
 */
function readSeiyakuFromFormAnswers_(ss) {
  var sheet = getSheetByGid_(ss, PA_FORM_GID);
  if (!sheet) { Logger.log('フォーム回答シート(gid=' + PA_FORM_GID + ')が見つかりません'); return []; }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var lastCol = Math.max(sheet.getLastColumn(), PA_FORM_COL_EMAIL + 1);
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
      pushDate: pushDateKey_(row[PA_COL_PUSH]),
      sales: sales,
      name: name,
      line: String(row[PA_COL_LINE] || '').trim(),
      email: String(row[PA_FORM_COL_EMAIL] || '').toLowerCase().trim(),
      status: status,
      _source: 'form'
    });

    var contractor = extractContractor_(row[PA_COL_SUPPLEMENT]);
    if (contractor) {
      out.push({
        push: String(row[PA_COL_PUSH] || '').trim(),
        pushDate: pushDateKey_(row[PA_COL_PUSH]),
        sales: sales,
        name: contractor,
        line: contractor,
        email: '',
        status: status,
        _source: 'form_contractor_alias',
        _origName: name
      });
    }
    var romaji = extractRomajiName_(row[PA_COL_SUPPLEMENT]);
    if (romaji) {
      out.push({
        push: String(row[PA_COL_PUSH] || '').trim(),
        pushDate: pushDateKey_(row[PA_COL_PUSH]),
        sales: sales,
        name: romaji,
        line: romaji,
        email: '',
        status: status,
        _source: 'form_romaji_alias',
        _origName: name
      });
    }
    var primaryEmail = String(row[PA_FORM_COL_EMAIL] || '').toLowerCase().trim();
    var emailAliases = extractEmailAliases_(row[PA_COL_SUPPLEMENT]);
    for (var ea = 0; ea < emailAliases.length; ea++) {
      var aliasEmail = emailAliases[ea];
      if (aliasEmail === primaryEmail) continue;
      out.push({
        push: String(row[PA_COL_PUSH] || '').trim(),
        pushDate: pushDateKey_(row[PA_COL_PUSH]),
        sales: sales,
        name: name,
        line: String(row[PA_COL_LINE] || '').trim(),
        email: aliasEmail,
        status: status,
        _source: 'form_email_alias',
        _origName: name
      });
    }
  }
  return out;
}

/**
 * マスター本体 + フォーム回答 を統合して商談者リストを返す
 * 重複は (pushDate + 顧客名) で排除（マスター優先）
 */
function readAllSeiyaku_(ss) {
  var master = readSeiyakuFromMaster_(ss);
  var form = readSeiyakuFromFormAnswers_(ss);
  var seen = {};
  master.forEach(function(s) {
    var key = (s.pushDate || '') + '|' + normalizeName_(s.name);
    seen[key] = true;
  });
  var added = 0;
  form.forEach(function(s) {
    var key = (s.pushDate || '') + '|' + normalizeName_(s.name);
    if (!seen[key]) {
      master.push(s);
      seen[key] = true;
      added++;
    }
  });
  Logger.log('  内訳: マスター本体=' + (master.length - added) + '件 / フォーム回答(未転記)=' + added + '件');
  return master;
}

/**
 * GID指定でシートを取得
 */
function getSheetByGid_(ss, gid) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === gid) return sheets[i];
  }
  return null;
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
  // 55列フォーマット: メールアドレス col27, カード名義 col38
  var lastCol = Math.max(sheet.getLastColumn(), 39);
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var out = [];
  for (var i = 0; i < data.length; i++) {
    var r = data[i];

    // 課金ステータス col9（成功のみ採用、空も除外）
    var status = String(r[9] || '').trim();
    if (status !== '成功') continue;

    // イベント col16（売上/返金/リカーリングトークン発行 等）
    var typ = String(r[16] || '').trim();
    if (typ === 'リカーリングトークン発行') continue; // トークン発行は取引でない

    // 課金金額 col7
    var amt = parsePAAmount_(r[7]);
    if (amt === null || amt === 0) continue;
    if (typ === '返金') amt = -Math.abs(amt);

    // メールアドレス col27
    var email = String(r[27] || '').trim().toLowerCase();

    // トークンメタデータ col11 から "univapay-name"（日本語名）を抽出
    var meta = String(r[11] || '');
    var nameJp = '';
    var m = meta.match(/"univapay-name":\s*"([^"]+)"/);
    if (m) nameJp = m[1].trim();

    // カード名義 col38（ローマ字フォールバック）
    var nameRoma = String(r[38] || '').trim();

    var name = nameJp || nameRoma;
    if (!name && !email) continue;

    // 日付 col4（イベント作成日時、ISO形式）
    var dateStr = parsePAUnivaDate_(r[4]);
    if (!dateStr) continue;

    out.push({
      date: dateStr,
      name: name,
      email: email,
      amt: amt,
      type: typ
    });
  }
  return out;
}

/**
 * 「ライフティ」タブ読み取り
 * 期待カラム: 0:申込ID, 1:申込日時, 2:加盟店名, 3:加盟店支店名, 4:担当者, 5:加盟店顧客ID,
 *           6:申込者氏名, 7:金額, 8:回数, 9:承認番号, 10:お客様ﾀﾞｳﾝﾛｰﾄﾞ日時,
 *           11:受付ｽﾃｰﾀｽ, 12:審査ｽﾃｰﾀｽ, 13:審査ｻﾌﾞｽﾃｰﾀｽ, 14:加盟店メモ,
 *           15:連絡事項ｽﾃｰﾀｽ, 16:集計
 *
 * - r[16] (集計) === '集計済み' のみ採用（=ライフティ側が承認＋確定した取引のみ）
 *   受付=完了 でも 審査=否決/本人確認中/審査中 のものは「集計済み」にならない
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
    var aggStatus = String(r[16] || '').trim();
    if (aggStatus !== '集計済み') continue;
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
 * 「MOSH」タブ読み取り
 * 期待カラム:
 *   0:サービスID, 1:サービス名, 2:ゲストID, 3:email, 4:ゲスト名,
 *   5:申し込み日, 6:決済日, 7:支払い方法, 8:支払い種別, 9:クーポン名,
 *  10:申し込み総額(税込), 11:クーポン割引額, 12:その他費用, 13:特別割引額,
 *  14:決済額(税込), 15:決済額(税抜), 16:決済消費税, 17:決済ステータス,
 *  18:総支払い回数, 19:X回目の支払い, 20:分割ステータス, 21:キャンセル日時,
 *  22:対象期間(サブスク)
 *
 * - 決済ステータス === '支払い済み' のみ
 * - 金額: 決済額(税込) col 14
 * - 日付: 決済日 col 6
 * - email + ゲスト名 でマッチ
 */
function readMoshTab_(ss) {
  var sheet = ss.getSheetByName(PA_TAB_MOSH);
  if (!sheet) { Logger.log('MOSHタブなし'); return []; }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 23).getValues();

  var out = [];
  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    var status = String(r[17] || '').trim();
    if (status !== '支払い済み') continue;
    var amt = parsePAAmount_(r[14]);
    if (amt === null || amt <= 0) continue;
    var dateStr = parsePAUnivaDate_(r[6]); // ISO 形式なのでユニヴァのパーサで OK
    if (!dateStr) continue;
    var email = String(r[3] || '').trim().toLowerCase();
    var name = String(r[4] || '').trim();
    out.push({
      date: dateStr,
      name: name,
      email: email,
      amt: amt
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
/**
 * マスターステータスが集計除外対象かどうか
 * 該当パターン:
 *   - 成約➔CO / 成約→CO （クーリングオフ）
 *   - 成約➔キャンセル / 成約→キャンセル
 *   - 成約➔失注 / 成約→失注 / 失注 / 継続失注 / 継続→失注
 *   - 継続（商談継続中で成約に至ってない＝着金扱いしない）
 *
 * 計上対象（誤除外しない）:
 *   - 成約
 *   - 継続n→成約（継続商談を経て最終的に成約）
 */
function isCOStatus_(status) {
  if (!status) return false;
  if (status.indexOf('CO') !== -1) return true;
  if (status.indexOf('キャンセル') !== -1) return true;
  if (status.indexOf('失注') !== -1) return true;
  // 「継続」を含むが「成約」を含まない場合は除外（継続中で未成約）
  if (status.indexOf('継続') !== -1 && status.indexOf('成約') === -1) return true;
  return false;
}

function aggregateBySalesAndPushDate_(univaTxs, liftyTxs, bankTxs, moshTxs, idx) {
  var result = {};
  var stats = {
    univa: { total: univaTxs.length, match: 0, matchAmt: 0, miss: 0, missAmt: 0, coSkip: 0, coSkipAmt: 0 },
    lifty: { total: liftyTxs.length, match: 0, matchAmt: 0, miss: 0, missAmt: 0, coSkip: 0, coSkipAmt: 0 },
    bank:  { total: bankTxs.length,  match: 0, matchAmt: 0, miss: 0, missAmt: 0, coSkip: 0, coSkipAmt: 0 },
    mosh:  { total: moshTxs.length,  match: 0, matchAmt: 0, miss: 0, missAmt: 0, coSkip: 0, coSkipAmt: 0 }
  };
  // 3層ネスト: result[sales][pushDate][customerName] = bucket
  function ensure(sales, pushDate, name) {
    if (!result[sales]) result[sales] = {};
    if (!result[sales][pushDate]) result[sales][pushDate] = {};
    if (!result[sales][pushDate][name]) result[sales][pushDate][name] = { univa: 0, lifty: 0, mosh: 0, bank: 0, total: 0, count: 0 };
    return result[sales][pushDate][name];
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
    if (isCOStatus_(s.status)) {
      stats.univa.coSkip++;
      stats.univa.coSkipAmt += tx.amt;
      continue;
    }
    stats.univa.match++;
    stats.univa.matchAmt += tx.amt;
    var bucket = ensure(shortSales_(s.sales), s.pushDate, s.name);
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
    if (isCOStatus_(s.status)) {
      stats.lifty.coSkip++;
      stats.lifty.coSkipAmt += tx.amt;
      continue;
    }
    stats.lifty.match++;
    stats.lifty.matchAmt += tx.amt;
    var bucket = ensure(shortSales_(s.sales), s.pushDate, s.name);
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
    if (isCOStatus_(s.status)) {
      stats.bank.coSkip++;
      stats.bank.coSkipAmt += tx.amt;
      continue;
    }
    stats.bank.match++;
    stats.bank.matchAmt += tx.amt;
    var bucket = ensure(shortSales_(s.sales), s.pushDate, s.name);
    bucket.bank += tx.amt;
    bucket.total += tx.amt;
    bucket.count += 1;
  }

  // MOSH
  for (var i = 0; i < moshTxs.length; i++) {
    var tx = moshTxs[i];
    var s = matchMoshSeiyaku_(tx, idx);
    if (!s || !s.pushDate) {
      stats.mosh.miss++;
      stats.mosh.missAmt += tx.amt;
      continue;
    }
    if (isCOStatus_(s.status)) {
      stats.mosh.coSkip++;
      stats.mosh.coSkipAmt += tx.amt;
      continue;
    }
    stats.mosh.match++;
    stats.mosh.matchAmt += tx.amt;
    var bucket = ensure(shortSales_(s.sales), s.pushDate, s.name);
    bucket.mosh += tx.amt;
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
  var headerRow = ['商談日', '商談者', '顧客', '着金額', 'ユニヴァ', 'ライフティ', 'MOSH', '銀振', '件数', '更新日時'];
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

  // 新規データ生成（金額は万円・小数1位に丸め、顧客ごとに行を作る）
  var rows = [];
  for (var i = 0; i < processedDates.length; i++) {
    var d = processedDates[i];
    for (var sales in aggregated) {
      if (!aggregated[sales][d]) continue;
      var byName = aggregated[sales][d];
      for (var name in byName) {
        var b = byName[name];
        if (b.total === 0) continue;
        rows.push([
          d,
          sales,
          name,
          toMan_(b.total),
          toMan_(b.univa),
          toMan_(b.lifty),
          toMan_(b.mosh || 0),
          toMan_(b.bank),
          b.count,
          now
        ]);
      }
    }
  }

  // 商談日（新→旧）→ 着金額（多→少）でソート
  rows.sort(function(a, b) {
    if (a[0] !== b[0]) return String(b[0]).localeCompare(String(a[0]));
    return Number(b[3]) - Number(a[3]);  // C列が顧客、D列(index 3)が着金額
  });

  // 書き込み（書式も明示的にリセット）
  if (rows.length > 0) {
    var range = sheet.getRange(2, 1, rows.length, headerRow.length);
    range.setNumberFormat('General');
    range.setValues(rows);
    // A列（商談日）は強制テキスト
    sheet.getRange(2, 1, rows.length, 1).setNumberFormat('@');
    // D-H: 着金額/ユニヴァ/ライフティ/MOSH/銀振 = 5列
    sheet.getRange(2, 4, rows.length, 5).setNumberFormat('0.0');
    // I: 件数（整数）
    sheet.getRange(2, 9, rows.length, 1).setNumberFormat('0');
    // J: 更新日時（テキスト）
    sheet.getRange(2, 10, rows.length, 1).setNumberFormat('@');
  }
}

/**
 * 円 → 万円（小数1位、四捨五入）
 */
function toMan_(yen) {
  return Math.round(yen / 1000) / 10;
}

/**
 * 「マスター登録漏れ」タブに事業内の不一致取引（ユニヴァ・ライフティ・MOSH）を書き出す
 * 銀振は別事業前提のため除外
 *
 * このタブを見て、リーさんor営業がマスターに登録漏れを追加すれば、次回再集計で正しくマッチする
 */
function writeUnmatchedSheet_(ss, univaTxs, liftyTxs, moshTxs, idx) {
  var sheet = ss.getSheetByName(PA_TAB_UNMATCHED);
  if (!sheet) sheet = ss.insertSheet(PA_TAB_UNMATCHED);
  sheet.clearContents();

  var header = ['日付', 'ソース', '顧客名', 'email', '担当者(ライフティのみ)', '金額(円)', '更新日時'];
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.setFrozenRows(1);
  sheet.getRange('A:A').setNumberFormat('@');

  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  var rows = [];

  for (var i = 0; i < univaTxs.length; i++) {
    var tx = univaTxs[i];
    var s = matchUnivaSeiyaku_(tx, idx);
    if (!s || !s.pushDate) {
      rows.push([tx.date, 'ユニヴァ', tx.name, tx.email || '', '', Math.round(tx.amt), now]);
    }
  }
  for (var j = 0; j < liftyTxs.length; j++) {
    var tx2 = liftyTxs[j];
    var s2 = matchLiftySeiyaku_(tx2, idx);
    if (!s2 || !s2.pushDate) {
      rows.push([tx2.complete_date || tx2.apply_date, 'ライフティ', tx2.name, '', tx2.sales || '', Math.round(tx2.amt), now]);
    }
  }
  for (var k = 0; k < moshTxs.length; k++) {
    var tx3 = moshTxs[k];
    var s3 = matchMoshSeiyaku_(tx3, idx);
    if (!s3 || !s3.pushDate) {
      rows.push([tx3.date, 'MOSH', tx3.name, tx3.email || '', '', Math.round(tx3.amt), now]);
    }
  }

  // 日付降順 → 金額降順
  rows.sort(function(a, b) {
    if (a[0] !== b[0]) return String(b[0]).localeCompare(String(a[0]));
    return Number(b[5]) - Number(a[5]);
  });

  if (rows.length > 0) {
    var range = sheet.getRange(2, 1, rows.length, header.length);
    range.setNumberFormat('General');
    range.setValues(rows);
    sheet.getRange(2, 1, rows.length, 1).setNumberFormat('@');     // 日付テキスト
    sheet.getRange(2, 6, rows.length, 1).setNumberFormat('#,##0'); // 金額(円)
    sheet.getRange(2, 7, rows.length, 1).setNumberFormat('@');     // 更新日時
  }

  Logger.log('マスター登録漏れタブに ' + rows.length + ' 件書き出し');
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
 * 名前正規化：カッコ・/・スペース・全角半角を除去、装飾記号削除、小文字化
 *
 * LINE名に紛れ込んでいる装飾記号 (®️ ⭐︎ ☆ ★ 等) があると
 * normalize 結果がローマ字決済名 (rikuo shimizu 等) と一致せず
 * matching 漏れの原因になるため、特殊文字を吸収する
 */
function normalizeName_(s) {
  if (!s) return '';
  var x = String(s);
  x = x.replace(/\([^)]*\)/g, '');
  x = x.replace(/（[^）]*）/g, '');
  x = x.replace(/\//g, '');
  // 装飾記号削除（清水陸雄の "®️ikuo SHIMIZU" 問題対策）
  // 「®」を「r」に変換、その他の装飾記号は削除
  x = x.replace(/®️|®/g, 'r');
  x = x.replace(/[⭐⭐︎☆★♪♡♥◎◯○●◆◇■□▲△▼▽✨💫🌟]/g, '');
  // 異体字セレクタ・結合用記号
  x = x.replace(/[​-‏︀-️‪-‮]/g, '');
  x = x.replace(/\s+/g, '');
  x = x.replace(/　/g, '');
  return x.toLowerCase().trim();
}

function matchUnivaSeiyaku_(tx, idx) {
  // 1. email完全一致
  if (tx.email && idx.email[tx.email]) return idx.email[tx.email];

  // 2. email の local part 共通プレフィックス（5文字以上一致）
  //    例: hitomi10e ⇔ hitomin12yuki31 → 'hitomi' 6文字共通 → match
  if (tx.email) {
    var txLocal = tx.email.split('@')[0].toLowerCase();
    if (txLocal.length >= 5) {
      for (var key in idx.email) {
        var keyLocal = key.split('@')[0].toLowerCase();
        if (commonPrefixLen_(txLocal, keyLocal) >= 5) {
          return idx.email[key];
        }
      }
    }
  }

  // 3. 正規化名 完全一致（深澤文代タイプ: フォーム回答行で本名登録あり、emailなし）
  //    新フォーマットのユニヴァは col11 トークンメタデータから日本語名を取得済み
  var nNorm = normalizeName_(tx.name);
  if (nNorm && idx.name[nNorm]) return latestEntry_(idx.name[nNorm]);

  // 4. 正規化名 部分一致（3文字以上）
  if (nNorm && nNorm.length >= 3) {
    for (var k0 in idx.name) {
      if (k0.length >= 2 && (nNorm.indexOf(k0) !== -1 || k0.indexOf(nNorm) !== -1)) {
        return latestEntry_(idx.name[k0]);
      }
    }
    // 正規化名 → line_idx 部分一致（line=LINE名/ニックネーム）
    for (var kl in idx.line) {
      if (kl.length >= 2 && (nNorm.indexOf(kl) !== -1 || kl.indexOf(nNorm) !== -1)) {
        return latestEntry_(idx.line[kl]);
      }
    }
  }

  // 5. ローマ字 name → line_idx 部分一致（既存: aoi ishida 等）
  var nl = tx.name.toLowerCase().replace(/\s+/g, '');
  for (var k in idx.line) {
    if (k.length >= 4 && (nl.indexOf(k) !== -1 || k.indexOf(nl) !== -1)) {
      return latestEntry_(idx.line[k]);
    }
  }

  // 6. ローマ字 → カナ変換 → line_idx / name_idx 部分一致
  //    例: fumiyo fukazawa → フミヨフカザワ → マスターLINE「フミヨ」と一致
  var kana = romajiToKatakana_(tx.name);
  if (kana && kana.length >= 4) {
    for (var k2 in idx.line) {
      if (k2.length >= 3 && (kana.indexOf(k2) !== -1 || k2.indexOf(kana) !== -1)) {
        return latestEntry_(idx.line[k2]);
      }
    }
    for (var k3 in idx.name) {
      if (k3.length >= 3 && (kana.indexOf(k3) !== -1 || k3.indexOf(kana) !== -1)) {
        return latestEntry_(idx.name[k3]);
      }
    }
  }

  return null;
}

/**
 * 2 つの文字列の共通プレフィックス長
 */
function commonPrefixLen_(a, b) {
  var n = Math.min(a.length, b.length);
  for (var i = 0; i < n; i++) {
    if (a.charAt(i) !== b.charAt(i)) return i;
  }
  return n;
}

/**
 * デバッグ用: マスター内で query を含む顧客行を全て出す
 */
function debugCustomer(query) {
  var ss = SpreadsheetApp.openById(PA_MASTER_ID);
  var seiyaku = readAllSeiyaku_(ss);
  Logger.log('=== マスター本体+フォーム回答 内 「' + query + '」を含む顧客 ===');
  var hits = 0;
  var q = String(query).toLowerCase();
  for (var i = 0; i < seiyaku.length; i++) {
    var s = seiyaku[i];
    var hay = (s.name + ' ' + s.line + ' ' + s.email).toLowerCase();
    if (hay.indexOf(q) !== -1) {
      hits++;
      Logger.log('  ' + s.pushDate + ' / 商談者=' + s.sales + ' / 名:' + s.name + ' / line:' + s.line + ' / email:' + s.email + ' / status:' + s.status);
    }
  }
  Logger.log('合計: ' + hits + '件');
}

/**
 * 簡易ヘボン式ローマ字→カタカナ変換
 * 商談者のマスター LINE 名が「フミヨ」のようにカナで登録されている時、
 * ユニヴァのローマ字 name 「fumiyo fukazawa」と紐付けるために使う
 */
function romajiToKatakana_(s) {
  if (!s) return '';
  s = String(s).toLowerCase().replace(/[\s_\-\.]+/g, '');
  var conv = [
    // 3文字
    ['kya','キャ'],['kyu','キュ'],['kyo','キョ'],
    ['sha','シャ'],['shu','シュ'],['sho','ショ'],['shi','シ'],
    ['cha','チャ'],['chu','チュ'],['cho','チョ'],['chi','チ'],['tsu','ツ'],
    ['nya','ニャ'],['nyu','ニュ'],['nyo','ニョ'],
    ['hya','ヒャ'],['hyu','ヒュ'],['hyo','ヒョ'],
    ['mya','ミャ'],['myu','ミュ'],['myo','ミョ'],
    ['rya','リャ'],['ryu','リュ'],['ryo','リョ'],
    ['gya','ギャ'],['gyu','ギュ'],['gyo','ギョ'],
    ['jya','ジャ'],['jyu','ジュ'],['jyo','ジョ'],
    ['bya','ビャ'],['byu','ビュ'],['byo','ビョ'],
    ['pya','ピャ'],['pyu','ピュ'],['pyo','ピョ'],
    // 促音 (kk, ss, tt, pp 等)
    ['kk','ッk'],['ss','ッs'],['tt','ッt'],['pp','ッp'],
    // 2文字
    ['ka','カ'],['ki','キ'],['ku','ク'],['ke','ケ'],['ko','コ'],
    ['sa','サ'],['su','ス'],['se','セ'],['so','ソ'],
    ['ta','タ'],['te','テ'],['to','ト'],['ti','チ'],
    ['na','ナ'],['ni','ニ'],['nu','ヌ'],['ne','ネ'],['no','ノ'],
    ['ha','ハ'],['hi','ヒ'],['hu','フ'],['he','ヘ'],['ho','ホ'],
    ['fa','ファ'],['fi','フィ'],['fu','フ'],['fe','フェ'],['fo','フォ'],
    ['ma','マ'],['mi','ミ'],['mu','ム'],['me','メ'],['mo','モ'],
    ['ya','ヤ'],['yu','ユ'],['yo','ヨ'],
    ['ra','ラ'],['ri','リ'],['ru','ル'],['re','レ'],['ro','ロ'],
    ['wa','ワ'],['wi','ウィ'],['we','ウェ'],['wo','ヲ'],
    ['ga','ガ'],['gi','ギ'],['gu','グ'],['ge','ゲ'],['go','ゴ'],
    ['za','ザ'],['zi','ジ'],['zu','ズ'],['ze','ゼ'],['zo','ゾ'],
    ['ja','ジャ'],['ji','ジ'],['ju','ジュ'],['je','ジェ'],['jo','ジョ'],
    ['da','ダ'],['di','ジ'],['du','ヅ'],['de','デ'],['do','ド'],
    ['ba','バ'],['bi','ビ'],['bu','ブ'],['be','ベ'],['bo','ボ'],
    ['pa','パ'],['pi','ピ'],['pu','プ'],['pe','ペ'],['po','ポ'],
    ['vu','ヴ'],
    // 1文字
    ['a','ア'],['i','イ'],['u','ウ'],['e','エ'],['o','オ'],
    ['n','ン']
  ];
  var result = '';
  while (s.length > 0) {
    var matched = false;
    for (var i = 0; i < conv.length; i++) {
      if (s.indexOf(conv[i][0]) === 0) {
        result += conv[i][1];
        s = s.substring(conv[i][0].length);
        matched = true;
        break;
      }
    }
    if (!matched) {
      s = s.substring(1);  // 不明文字スキップ
    }
  }
  // 促音マーカーを正規ッに置換
  result = result.replace(/ッ[a-z]/g, function(m) {
    var c = m.charAt(1);
    var map = {k:'カ',s:'サ',t:'タ',p:'パ'};
    return 'ッ' + (map[c] || '');
  });
  return result;
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
 * MOSHマッチング: email完全 → 共通プレフィックス → 正規化名 → line fuzzy
 * MOSHのゲスト名は日本語フルネーム（"水谷　康晃"）の前提
 */
function matchMoshSeiyaku_(tx, idx) {
  // 1. email完全一致
  if (tx.email && idx.email[tx.email]) return idx.email[tx.email];

  // 2. email 共通プレフィックス（5文字以上一致）
  if (tx.email) {
    var txLocal = tx.email.split('@')[0].toLowerCase();
    if (txLocal.length >= 5) {
      for (var key in idx.email) {
        var keyLocal = key.split('@')[0].toLowerCase();
        if (commonPrefixLen_(txLocal, keyLocal) >= 5) {
          return idx.email[key];
        }
      }
    }
  }

  // 3. 正規化名 完全一致
  var n = normalizeName_(tx.name);
  if (n && idx.name[n]) return latestEntry_(idx.name[n]);

  // 4. 正規化名 部分一致（3文字以上）
  if (n && n.length >= 3) {
    for (var k in idx.name) {
      if (n.indexOf(k) !== -1 || k.indexOf(n) !== -1) {
        return latestEntry_(idx.name[k]);
      }
    }
    // line_idx fuzzy
    for (var k2 in idx.line) {
      if (k2.length >= 3 && (n.indexOf(k2) !== -1 || k2.indexOf(n) !== -1)) {
        return latestEntry_(idx.line[k2]);
      }
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
  // pushDate が空でないものを優先（空だと aggregate で除外されて取りこぼす）
  var withDate = entries.filter(function(e) { return !!e.pushDate; });
  var pool = withDate.length > 0 ? withDate : entries;
  if (pool.length === 1) return pool[0];
  var latest = pool[0];
  for (var i = 1; i < pool.length; i++) {
    if (pool[i].pushDate > (latest.pushDate || '')) {
      latest = pool[i];
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
  var seiyaku = readAllSeiyaku_(ss);
  var indexes = buildSeiyakuIndexes_(seiyaku);
  var liftyTxs = readLiftyTab_(ss);
  var univaTxs = readUnivaTab_(ss);
  var bankTxs = readBankTab_(ss);
  var moshTxs = readMoshTab_(ss);

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
  for (var i = 0; i < moshTxs.length; i++) {
    var tx = moshTxs[i];
    var s = matchMoshSeiyaku_(tx, indexes);
    if (s && s.pushDate === targetDate) {
      matchCount++;
      Logger.log('  MOSH: 顧客=' + tx.name + ' / ' + tx.email + ' ¥' + tx.amt + ' → matched: ' + s.name + ' (商談者:' + s.sales + ')');
    }
  }
  Logger.log('マッチ合計: ' + matchCount + '件');
}

function debugApril30() { return debugByPushDate('2026-04-30'); }

/**
 * 不一致取引（一次データに出てるがマスター商談者にマッチしない取引）を全件ログ出し
 * これらが「マスター未掲載だけど事業の取引」候補
 */
function debugUnmatched() {
  var ss = SpreadsheetApp.openById(PA_MASTER_ID);
  var seiyaku = readAllSeiyaku_(ss);
  var indexes = buildSeiyakuIndexes_(seiyaku);

  var univaTxs = readUnivaTab_(ss);
  var liftyTxs = readLiftyTab_(ss);
  var bankTxs = readBankTab_(ss);
  var moshTxs = readMoshTab_(ss);

  Logger.log('=== ユニヴァ不一致 ===');
  for (var i = 0; i < univaTxs.length; i++) {
    var tx = univaTxs[i];
    var s = matchUnivaSeiyaku_(tx, indexes);
    if (!s || !s.pushDate) {
      Logger.log('  ' + tx.date + ' / ' + tx.name + ' / ' + tx.email + ' ¥' + tx.amt);
    }
  }

  Logger.log('');
  Logger.log('=== ライフティ不一致 ===');
  for (var i = 0; i < liftyTxs.length; i++) {
    var tx = liftyTxs[i];
    var s = matchLiftySeiyaku_(tx, indexes);
    if (!s || !s.pushDate) {
      Logger.log('  申込=' + tx.apply_date + ' 完了=' + tx.complete_date + ' / ' + tx.name + ' ¥' + tx.amt);
    }
  }

  Logger.log('');
  Logger.log('=== 銀振不一致（事業外候補多数なので決済代行除外後の純粋な不一致）===');
  for (var i = 0; i < bankTxs.length; i++) {
    var tx = bankTxs[i];
    var s = matchBankSeiyaku_(tx, indexes);
    if (!s || !s.pushDate) {
      Logger.log('  ' + tx.date + ' / ' + tx.name_kana + ' ¥' + tx.amt);
    }
  }

  Logger.log('');
  Logger.log('=== MOSH不一致 ===');
  for (var i = 0; i < moshTxs.length; i++) {
    var tx = moshTxs[i];
    var s = matchMoshSeiyaku_(tx, indexes);
    if (!s || !s.pushDate) {
      Logger.log('  ' + tx.date + ' / ' + tx.name + ' / ' + tx.email + ' ¥' + tx.amt);
    }
  }
}

// =========== debugCustomer 用ラッパー（GASエディタの関数ドロップダウンから実行できるよう） ===========

function debugFukazawa()   { return debugCustomer('深澤'); }
function debugFumiyo()     { return debugCustomer('fumiyo'); }
function debugIshidaAoi()  { return debugCustomer('石田葵'); }
function debugAoiIshida()  { return debugCustomer('aoi ishida'); }
function debugOhkuma()     { return debugCustomer('大熊'); }
function debugIshidaAoik() { return debugCustomer('aoi'); }
function debugKotaki()     { return debugCustomer('小滝'); }
function debugIimura()     { return debugCustomer('iimura'); }

function debugApril28() { return debugByPushDate('2026-04-28'); }

/**
 * 指定月の不一致取引を一次データ別に金額降順でログ出力
 * 銀振は別事業前提なので除外（ユニヴァ・ライフティ・MOSHのみ事業内不一致）
 */
function debugUnmatchedByMonth_(targetMonth, limit) {
  if (!limit) limit = 30;
  var ss = SpreadsheetApp.openById(PA_MASTER_ID);
  var seiyaku = readAllSeiyaku_(ss);
  var indexes = buildSeiyakuIndexes_(seiyaku);

  var univaTxs = readUnivaTab_(ss);
  var liftyTxs = readLiftyTab_(ss);
  var moshTxs  = readMoshTab_(ss);

  function inMonth(d) { return d && d.indexOf(targetMonth) === 0; }

  var groups = [
    { name: 'ユニヴァ', txs: univaTxs, match: matchUnivaSeiyaku_, dateField: 'date' },
    { name: 'ライフティ', txs: liftyTxs, match: matchLiftySeiyaku_, dateField: 'complete_date' },
    { name: 'MOSH', txs: moshTxs, match: matchMoshSeiyaku_, dateField: 'date' }
  ];

  groups.forEach(function(g) {
    var miss = [];
    var total = 0;
    for (var i = 0; i < g.txs.length; i++) {
      var tx = g.txs[i];
      if (!inMonth(tx[g.dateField])) continue;
      var s = g.match(tx, indexes);
      if (!s || !s.pushDate) {
        miss.push(tx);
        total += tx.amt;
      }
    }
    miss.sort(function(a, b) { return b.amt - a.amt; });
    Logger.log('=== ' + g.name + ' 不一致 (' + targetMonth + '): ' + miss.length + '件 ¥' + Math.round(total).toLocaleString() + ' ===');
    var show = Math.min(miss.length, limit);
    for (var j = 0; j < show; j++) {
      var tx = miss[j];
      var line = '  ¥' + Math.round(tx.amt).toLocaleString();
      if (g.name === 'ライフティ') {
        line += ' / 完了=' + tx.complete_date + ' / ' + tx.name + ' / 担当=' + tx.sales;
      } else {
        line += ' / ' + tx[g.dateField] + ' / ' + tx.name + ' / ' + (tx.email || '(emailなし)');
      }
      Logger.log(line);
    }
    if (miss.length > show) Logger.log('  ... 他 ' + (miss.length - show) + ' 件');
    Logger.log('');
  });
}

function debugUnmatchedApril()    { return debugUnmatchedByMonth_('2026-04', 50); }
function debugUnmatchedMarch()    { return debugUnmatchedByMonth_('2026-03', 50); }
function debugUnmatchedFebruary() { return debugUnmatchedByMonth_('2026-02', 50); }

// =========== タブ構造スキャン（読み取り専用） ===========
/**
 * 集約スプシの全タブ + ヘッダー行 + 行数 + (検出できれば)日付範囲をログ出力
 * MOSH対応や過去月対応の前段で「データがどこに何件あるか」を可視化するため
 */
function paDebugScanTabs() {
  var ss = SpreadsheetApp.openById(PA_MASTER_ID);
  var sheets = ss.getSheets();
  Logger.log('=== 全タブスキャン (' + sheets.length + 'タブ) ===');
  for (var i = 0; i < sheets.length; i++) {
    var s = sheets[i];
    var name = s.getName();
    var lastRow = s.getLastRow();
    var lastCol = s.getLastColumn();
    Logger.log('---');
    Logger.log('[' + (i + 1) + '] ' + name + ' (rows=' + lastRow + ', cols=' + lastCol + ', gid=' + s.getSheetId() + ')');
    if (lastRow >= 1 && lastCol >= 1) {
      var header = s.getRange(1, 1, 1, lastCol).getValues()[0];
      Logger.log('  header: ' + JSON.stringify(header));
    }
    if (lastRow >= 2 && lastCol >= 1) {
      var first = s.getRange(2, 1, 1, lastCol).getValues()[0];
      Logger.log('  row2  : ' + JSON.stringify(first));
    }
    if (lastRow >= 3 && lastCol >= 1) {
      var last = s.getRange(lastRow, 1, 1, lastCol).getValues()[0];
      Logger.log('  rowN  : ' + JSON.stringify(last));
    }
  }
}
