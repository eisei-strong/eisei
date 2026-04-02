// ========================================
// ExpenseAPI.js — 経費ダッシュボード バックエンド
// ========================================

var EXP_SS_ID = '1jXzkVWuSbpR7VelHppHil3IbuQOgJ8m_jEj0mB3THhw';
var EXP_CALC_SUFFIX = '-計算用';
var EXP_WASTEFUL_SHEET = '🚀削減したい経費';

// 列定数 (0-based)
var EXP_COL_CAT  = 0;
var EXP_COL_DATE = 1;
var EXP_COL_DESC = 2;
var EXP_COL_INC  = 3;
var EXP_COL_OUT  = 4;
var EXP_COL_IEXC = 5;
var EXP_COL_OEXC = 6;
var EXP_COL_SRC  = 7;

// カテゴリ正規化マップ（完全一致）
var EXP_CAT_MAP = {
  // 人件費系
  '社員人件費(営業)':             '営業',
  '社員人件費(営業以外)':         'バックオフィス',
  // Namaka系
  'Namaka（事務）':               'バックオフィス',
  'Namaka（動画）':               '動画',
  // 動画系
  '動画編集':                     '動画',
  '動画（台本作成／台本修正）':   '動画',
  '動画（素材集め／外注依頼）':   '動画',
  '動画（AIアフレコ編集業務）':   '動画',
  // 代理店報酬系
  '受講生報酬（代理店報酬）':     '代理店報酬',
  '受講生報酬（登録月キャッシュバック）': 'キャッシュバック',
  '受講生報酬（インセンティブ）': '代理店報酬',
  // マーケ系
  '広告業務サポート':             'マーケ',
  '広告業務サポート（UTAGE）':    'マーケ',
  // 営業系
  '営業代行/アフィリエイト':      '営業代行/アフィリ',
  '営業サポート（架電業務）':     '営業',
  '営業サポート（事務）':         '営業',
  '営業サポート（ライフティー申請サポート対応）': '営業',
  '営業サポート（アポ取り）':     '営業',
  '営業サポート（練習相手）':     '営業',
  // 事務/サポート系
  '事務サポート（採用業務）':     'バックオフィス',
  '事務サポート':                 'バックオフィス',
  '事務サポート（物販業務）':     'バックオフィス',
  '秘書業務':                     'バックオフィス',
  '秘書業務（KPI入力‧各SNSのlineリスト入力）': 'バックオフィス',
  'LINEサポート':                 'バックオフィス',
  // 通信費系
  '通信費（LINE・メッセージ系費用）':  '通信費',
  '通信費 （業務SaaS系）':        '通信費',
  '通信費（AI・自動化系）':       '通信費',
  '通信費 （クラウド系）':        '通信費',
  '通信費（通信ツール他）':       '通信費',
  // その他経費
  '消耗品費':                     '消耗品',
  '広告宣伝費':                   '広告宣伝費',
  '旅費交通費':                   '旅費交通費',
  '交際費':                       '交際費',
  '会議費':                       '会議費',
  '支払手数料':                   '支払手数料',
  '支払報酬':                     '士業',
  '研修費':                       '研修費',
  '新聞図書費':                   '新聞図書費',
  '売上返金':                     '売上返金',
  '地代家賃':                     '家賃水道光熱費',
  '水道光熱費':                   '家賃水道光熱費',
  'カード利用料':                 'カード利用料',
  '家事代行':                     '家事代行',
  '弁当代':                       '弁当代',
  '画像':                         '画像',
  '画像制作':                     '画像',
  '販促費':                       '販促費',
  '講師':                         '講師'
};

// 部分一致でカテゴリを解決するマップ（indexOf >= 0）
var EXP_CAT_PREFIX = [
  ['営業サポート', '営業'],
  ['営業代行',     '営業代行/アフィリ'],
  ['アフィリ',     '営業代行/アフィリ'],
  ['動画',         '動画'],
  ['事務サポート', 'バックオフィス'],
  ['秘書',         'バックオフィス'],
  ['LINEサポート', 'バックオフィス'],
  ['LINE対応',     'バックオフィス'],
  ['採用',         'バックオフィス'],
  ['物販',         'バックオフィス'],
  ['通信費',       '通信費'],
  ['Namaka',       'バックオフィス'],
  ['広告業務',     'マーケ'],
  ['広告事務',     'マーケ'],
  ['受講生報酬',   '代理店報酬'],
  ['社員人件費',   'バックオフィス'],
  ['幹部人件費',   '幹部報酬'],
  ['画像',         '画像'],
  ['税理士',       '士業'],
  ['弁護士',       '士業'],
  ['販促',         '広告宣伝費'],
  ['水道光熱',     '家賃水道光熱費'],
  ['地代家賃',     '家賃水道光熱費']
];

// ☆マーク項目の摘要パターンで分類
var EXP_DESC_PATTERNS = [
  [/ﾋﾞﾂﾄﾎﾟｲﾝﾄ|ビットポイント|bitpoint/i, '内部留保(BTC)'],
  [/ｼﾔｶｲﾎｹﾝ|社会保険/i,                   '税金/社保'],
  [/源泉|ゲンセン|所得税/i,                 '税金/社保'],
  [/住民税|ジュウミンゼイ/i,                '税金/社保'],
  [/厚生年金|コウセイネンキン/i,            '税金/社保'],
  [/ふみや|エフシ－エヌ/i,                  '幹部報酬']
];

// ========================================
// ユーティリティ
// ========================================
function expSS_() { return SpreadsheetApp.openById(EXP_SS_ID); }

function expParseNum_(val) {
  if (typeof val === 'number') return val;
  if (!val || val === '-' || val === '') return 0;
  return Number(String(val).replace(/[¥￥,、円万\s%％]/g, '')) || 0;
}

function expDateStr_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return val.getFullYear() + '/' + ('0' + (val.getMonth()+1)).slice(-2) + '/' + ('0' + val.getDate()).slice(-2);
  }
  var s = String(val).replace(/\D/g, '');
  if (s.length === 8) return s.substr(0,4) + '/' + s.substr(4,2) + '/' + s.substr(6,2);
  return String(val);
}

function expNormCat_(raw) {
  var t = String(raw || '').trim();
  if (!t || t === '☆') return 'その他';
  return t;
}

/**
 * 括弧を統一し、前方一致でカテゴリを解決
 */
function expParentCat_(cat, desc) {
  // 括弧を全角に統一してからマッチング
  var n = cat.replace(/\(/g, '（').replace(/\)/g, '）').replace(/\s+/g, '');

  // 1. 完全一致（括弧統一後）
  if (EXP_CAT_MAP[cat]) return EXP_CAT_MAP[cat];
  if (EXP_CAT_MAP[n]) return EXP_CAT_MAP[n];

  // 2. 前方一致（部分文字列マッチ）
  for (var i = 0; i < EXP_CAT_PREFIX.length; i++) {
    if (n.indexOf(EXP_CAT_PREFIX[i][0]) >= 0) return EXP_CAT_PREFIX[i][1];
  }

  // 3. ☆/その他 → 摘要パターンで判定
  if (cat === 'その他' && desc) {
    for (var j = 0; j < EXP_DESC_PATTERNS.length; j++) {
      if (EXP_DESC_PATTERNS[j][0].test(desc)) return EXP_DESC_PATTERNS[j][1];
    }
  }

  return cat;
}

// ========================================
// シート検出・解析
// ========================================
function expFindMonths_(ss) {
  var sheets = ss.getSheets();
  var months = [], seen = {};
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    var idx = name.indexOf(EXP_CALC_SUFFIX);
    if (idx < 0) continue;
    var match = name.substring(0, idx).match(/^(\d{4})年(\d{1,2})月$/);
    if (!match) continue;
    var y = parseInt(match[1]), m = parseInt(match[2]);
    var key = y + '-' + m;
    if (!seen[key]) { seen[key] = true; months.push({year:y, month:m}); }
  }
  months.sort(function(a,b) { return (a.year*12+a.month) - (b.year*12+b.month); });
  return months;
}

function expFindCalcSheet_(ss, year, month) {
  return ss.getSheetByName(year + '年' + month + '月' + EXP_CALC_SUFFIX);
}

function expParseSheet_(sheet) {
  if (!sheet) return null;
  var data = sheet.getRange(1, 1, sheet.getLastRow(), Math.max(sheet.getLastColumn(), 8)).getValues();
  var expenses = [], income = [];
  var totalExp = 0, totalInc = 0;
  var inIncome = false;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var colC = String(row[EXP_COL_DESC] || '').trim();

    if (colC === '支出計') { totalExp = expParseNum_(row[EXP_COL_OUT]); continue; }
    if (colC === '入金一覧') { inIncome = true; continue; }
    if (colC === '入金計') { totalInc = expParseNum_(row[EXP_COL_INC]); inIncome = false; continue; }
    if (inIncome && colC.indexOf('項目') === 0) continue;

    if (inIncome) {
      var ia = expParseNum_(row[EXP_COL_INC]);
      if (ia > 0) income.push({ category: expNormCat_(row[EXP_COL_CAT]), date: expDateStr_(row[EXP_COL_DATE]), description: colC, amount: ia, source: String(row[EXP_COL_SRC]||'') });
    } else {
      var ea = expParseNum_(row[EXP_COL_OUT]);
      if (ea > 0) {
        var rawCat = expNormCat_(row[EXP_COL_CAT]);
        var pCat = expParentCat_(rawCat, colC);
        var dispCat = (rawCat === 'その他') ? pCat : rawCat;
        expenses.push({ category: dispCat, parentCategory: pCat, date: expDateStr_(row[EXP_COL_DATE]), description: colC, amount: ea, source: String(row[EXP_COL_SRC]||'') });
      }
    }
  }
  return { expenses: expenses, income: income, totalExpense: totalExp, totalIncome: totalInc };
}

/**
 * 小分類（元カテゴリ）で集計
 */
function expAggregate_(expenses) {
  var map = {};
  for (var i = 0; i < expenses.length; i++) {
    var e = expenses[i], cat = e.category;
    if (!map[cat]) map[cat] = { name: cat, parent: e.parentCategory, amount: 0, count: 0 };
    map[cat].amount += e.amount;
    map[cat].count++;
  }
  var result = [];
  for (var k in map) result.push(map[k]);
  result.sort(function(a,b) { return b.amount - a.amount; });
  var total = 0;
  for (var j = 0; j < result.length; j++) total += result[j].amount;
  for (var l = 0; l < result.length; l++) result[l].pct = total > 0 ? Math.round((result[l].amount/total)*1000)/10 : 0;
  return result;
}

/**
 * 大分類（parentCategory）で集計
 */
function expAggregateParent_(expenses) {
  var map = {};
  for (var i = 0; i < expenses.length; i++) {
    var e = expenses[i], cat = e.parentCategory;
    if (!map[cat]) map[cat] = { name: cat, amount: 0, count: 0 };
    map[cat].amount += e.amount;
    map[cat].count++;
  }
  var result = [];
  for (var k in map) result.push(map[k]);
  result.sort(function(a,b) { return b.amount - a.amount; });
  var total = 0;
  for (var j = 0; j < result.length; j++) total += result[j].amount;
  for (var l = 0; l < result.length; l++) result[l].pct = total > 0 ? Math.round((result[l].amount/total)*1000)/10 : 0;
  return result;
}

// ========================================
// APIハンドラ
// ========================================
function expGetMonths_() {
  var ss = expSS_();
  var months = expFindMonths_(ss);
  return { months: months, current: months.length > 0 ? months[months.length-1] : {year:2026,month:2} };
}

function expGetData_(year, month) {
  var ss = expSS_();
  var months = expFindMonths_(ss);
  if (!year || !month) {
    if (months.length > 0) { var l = months[months.length-1]; year = l.year; month = l.month; }
    else return { error: 'No data' };
  }

  var cacheKey = 'exp_' + year + '_' + month;
  var cache = CacheService.getScriptCache();
  var cached = cache.get(cacheKey);
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }

  var sheet = expFindCalcSheet_(ss, year, month);
  if (!sheet) return { error: 'Sheet not found: ' + year + '/' + month };
  var parsed = expParseSheet_(sheet);
  if (!parsed) return { error: 'Parse failed' };

  var categories = expAggregate_(parsed.expenses);
  var parentCategories = expAggregateParent_(parsed.expenses);
  var topSorted = parsed.expenses.slice().sort(function(a,b){return b.amount-a.amount;});
  var top30 = topSorted.map(function(e){
    return {category:e.parentCategory, subCategory:e.category, date:e.date, description:e.description, amount:e.amount, source:e.source};
  });

  // 前月
  var pm = month-1, py = year;
  if (pm < 1) { pm = 12; py--; }
  var prev = {totalExpense:0, totalIncome:0};
  var ps = expFindCalcSheet_(ss, py, pm);
  if (ps) { var pp = expParseSheet_(ps); if (pp) { prev.totalExpense = pp.totalExpense; prev.totalIncome = pp.totalIncome; } }

  var profit = parsed.totalIncome - parsed.totalExpense;
  var prevProfit = prev.totalIncome - prev.totalExpense;

  // 削減したい経費
  var wasteful = [];
  try {
    var ws = ss.getSheetByName(EXP_WASTEFUL_SHEET);
    if (ws) {
      var wd = ws.getDataRange().getValues();
      for (var w = 1; w < wd.length; w++) {
        var desc = String(wd[w][2]||'').trim();
        var amt = expParseNum_(wd[w][4] || wd[w][3]);
        if (desc || amt) wasteful.push({category:String(wd[w][0]||''), description:desc, amount:amt, source:String(wd[w][7]||wd[w][5]||'')});
      }
    }
  } catch(e) {}

  var result = {
    revenue: parsed.totalIncome,
    totalExpenses: parsed.totalExpense,
    profit: profit,
    profitRate: parsed.totalIncome > 0 ? Math.round((profit/parsed.totalIncome)*1000)/10 : 0,
    categories: categories,
    parentCategories: parentCategories,
    topExpenses: top30,
    wasteful: wasteful,
    currentMonth: month,
    currentYear: year,
    prevRevenue: prev.totalIncome,
    prevExpenses: prev.totalExpense,
    prevProfit: prevProfit,
    diffRevenue: parsed.totalIncome - prev.totalIncome,
    diffExpenses: parsed.totalExpense - prev.totalExpense,
    diffProfit: profit - prevProfit,
    updatedAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')
  };

  try { var j = JSON.stringify(result); if (j.length < 100000) cache.put(cacheKey, j, 300); } catch(e) {}
  return result;
}

function expGetTrends_() {
  var ss = expSS_();
  var months = expFindMonths_(ss);
  var labels = [], revenues = [], expenses = [], profits = [];
  for (var i = 0; i < months.length; i++) {
    var m = months[i];
    var sheet = expFindCalcSheet_(ss, m.year, m.month);
    if (!sheet) continue;
    var parsed = expParseSheet_(sheet);
    if (!parsed) continue;
    labels.push(m.month + '月');
    revenues.push(parsed.totalIncome);
    expenses.push(parsed.totalExpense);
    profits.push(parsed.totalIncome - parsed.totalExpense);
  }
  return { labels: labels, revenue: revenues, expenses: expenses, profit: profits };
}

function expGetDetail_(year, month, catName) {
  var ss = expSS_();
  var sheet = expFindCalcSheet_(ss, year, month);
  if (!sheet) return { error: 'Sheet not found' };
  var parsed = expParseSheet_(sheet);
  if (!parsed) return { error: 'Parse failed' };

  var items = [], total = 0;
  for (var i = 0; i < parsed.expenses.length; i++) {
    var e = parsed.expenses[i];
    if (e.parentCategory === catName || e.category === catName) {
      items.push({category:e.category, date:e.date, description:e.description, amount:e.amount, source:e.source});
      total += e.amount;
    }
  }
  items.sort(function(a,b){return b.amount-a.amount;});
  return { category: catName, total: total, count: items.length, items: items };
}
