/**
 * 投稿本数 1〜100 の検証。
 * - サーバーが受け付ける値
 * - 旧COUNTIF（1〜10本）と厳密な合計が、❌ と 1〜10本では同じになること
 * - 11〜100本が足されること
 * - シートへ書く合計式は SUMPRODUCT。アプリが保存する値では厳密な合計と一致する
 * - COUNTIF と重い MAP 式は、同じ範囲の軽い式に差し替わる
 * - 日付列へプルダウンの入力規則を付けない
 */
var fs = require('fs');
var path = require('path');
var vm = require('vm');

var root = path.join(__dirname, '..');
var code = fs.readFileSync(path.join(root, 'PostApp.js'), 'utf8');
var sandbox = {
  SpreadsheetApp: {},
  Utilities: { formatDate: function() { return ''; } },
  PropertiesService: {},
  Logger: { log: function() {} },
  ContentService: {},
  UrlFetchApp: {},
  ScriptApp: {},
  HtmlService: {},
  Session: {},
  CacheService: {},
  DriveApp: {},
  console: console
};
vm.createContext(sandbox);
vm.runInContext(code, sandbox);

var failed = 0;
function assert(cond, msg) {
  if (!cond) {
    failed++;
    console.error('FAIL: ' + msg);
  }
}

assert(sandbox.POST_APP_MAX_DAILY_POSTS === 100, '上限は100');

// 受け付ける / 弾く
['❌'].concat(range(1, 100).map(function(n) { return n + '本'; })).forEach(function(v) {
  assert(sandbox.isValidPostAppValue_(v) === true, 'valid ' + v);
});
['', '0本', '01本', '101本', '1000本', '10', '10本です', '初投稿', '１０本', ' 5本'].forEach(function(v) {
  assert(sandbox.isValidPostAppValue_(v) === false, 'invalid raw ' + JSON.stringify(v));
});
assert(sandbox.normalizePostAppValue_(' ５本 ') === '5本', '全角数字と空白');
assert(sandbox.normalizePostAppValue_('１００本') === '100本', '全角100');
assert(sandbox.isValidPostAppValue_(sandbox.normalizePostAppValue_('１００本')) === true, '正規化後の100本');
assert(sandbox.isValidPostAppValue_(sandbox.normalizePostAppValue_('１０１本')) === false, '正規化後の101本');

// 旧式（1〜10）と新式が、11本以上を含まないデータで一致する
var legacyPool = ['❌', '', '初投稿', '10', '10本 ', '0本', '01本', '101本'];
for (var n = 1; n <= 10; n++) legacyPool.push(n + '本');
for (var trial = 0; trial < 200; trial++) {
  var row = [];
  var len = 1 + Math.floor(Math.random() * 40);
  for (var i = 0; i < len; i++) row.push(legacyPool[Math.floor(Math.random() * legacyPool.length)]);
  var legacy = sandbox.postAppLegacyCountifSum_(row, 10);
  var next = sandbox.postAppSumPostCells_(row);
  assert(legacy === next, 'trial ' + trial + ' legacy=' + legacy + ' next=' + next + ' row=' + JSON.stringify(row));
}

// 1本〜10本を1つずつ、❌、空欄
var sample = ['❌', '', '1本', '2本', '3本', '4本', '5本', '6本', '7本', '8本', '9本', '10本'];
assert(sandbox.postAppSumPostCells_(sample) === 55, '1〜10の合計は55');
assert(sandbox.postAppLegacyCountifSum_(sample, 10) === 55, '旧式も55');

// 11〜100 が足され、既存分は崩れない
var withHigh = sample.concat(['11本', '20本', '50本', '100本', '101本', '0本']);
assert(sandbox.postAppSumPostCells_(withHigh) === 55 + 11 + 20 + 50 + 100, '11以上を加算');
assert(sandbox.postAppLegacyCountifSum_(withHigh, 10) === 55, '旧式は11以上を無視');
assert(sandbox.postAppCountCell_('100本') === 100, '100本');
assert(sandbox.postAppCountCell_('11本') === 11, '11本');
assert(sandbox.postAppCountCell_('❌') === 0, '❌は0');
assert(sandbox.postAppCountCell_('初投稿') === 0, '初投稿は0');

// 式の差し替えは範囲を保つ。SUM(月別) と、すでに軽い式は触らない
var light = function(range) {
  return '=SUMPRODUCT(IFERROR(VALUE(SUBSTITUTE(' + range + ',"本","")),0))';
};
var oldFormula = '=COUNTIF($O2:AR2,"1本")+COUNTIF($O2:AR2,"10本")*10';
var upgraded = sandbox.postAppUpgradeCountFormula_(oldFormula);
assert(upgraded === light('$O2:AR2'), 'COUNTIFを軽い式へ。実際=' + upgraded);
assert(upgraded.indexOf('COUNTIF') < 0, 'COUNTIFを残さない');
assert(upgraded.indexOf('MAP(') < 0, 'MAPを書かない');
assert(upgraded.indexOf('REGEXMATCH') < 0, '正規表現を書かない');
assert(sandbox.postAppUpgradeCountFormula_('=SUM(F2:N2)') === '', '全期間SUMは対象外');
assert(sandbox.postAppUpgradeCountFormula_(upgraded) === '', '新しい式は再変換しない');

var heavy = '=SUM(MAP($O2:AR2,LAMBDA(c,IFERROR(IF(REGEXMATCH(TO_TEXT(c),"^([1-9]|[1-9][0-9]|100)本$"),VALUE(REGEXEXTRACT(TO_TEXT(c),"^[0-9]+")),0),0))))';
assert(sandbox.postAppUpgradeCountFormula_(heavy) === light('$O2:AR2'), '重いMAP式も同じ範囲の軽い式へ');
var heavySpaced = '=SUM(MAP( O2:AR2 , LAMBDA(c, IFERROR(IF(REGEXMATCH(TO_TEXT(c), "^([1-9]|[1-9][0-9]|100)本$"), VALUE(REGEXEXTRACT(TO_TEXT(c), "^[0-9]+")), 0), 0))))';
assert(sandbox.postAppUpgradeCountFormula_(heavySpaced) === light('O2:AR2'), '空白入りの重い式も範囲を保つ');

var direct = sandbox.postAppPostsSumFormula_('F12:AI12');
assert(direct === light('F12:AI12'), '直接生成も軽い式。実際=' + direct);

// 全期間合計（F〜N）は当月の式で上書きしない。日付列を直接足す式は上書き対象
assert(sandbox.postAppFormulaIsAllPeriodTotal_('=SUM(F2:N2)') === true, 'SUM(F:N)は全期間');
assert(sandbox.postAppFormulaIsAllPeriodTotal_('=SUM($F2:$N2)') === true, '絶対参照のSUM(F:N)');
assert(sandbox.postAppFormulaIsAllPeriodTotal_(light('F2:N2')) === true, 'SUMPRODUCT(F:N)は全期間');
assert(sandbox.postAppFormulaIsAllPeriodTotal_(light('$F$2:$N$2')) === true, '絶対参照のSUMPRODUCT(F:N)');
assert(sandbox.postAppFormulaIsAllPeriodTotal_(light('O2:IX2')) === false, '日付範囲のSUMPRODUCTは当月更新の対象');
assert(sandbox.postAppFormulaIsAllPeriodTotal_(heavy) === false, '重いMAP式は全期間扱いしない');
assert(sandbox.postAppFormulaIsAllPeriodTotal_('=COUNTIF(O2:AR2,"1本")') === false, 'COUNTIFは全期間扱いしない');
assert(sandbox.postAppLetterToCol_('F') === 6 && sandbox.postAppLetterToCol_('N') === 14, '月別合計列');
assert(sandbox.postAppLetterToCol_('O') === 15 && sandbox.postAppLetterToCol_('KC') === 289, '日付列 O〜KC');

// アプリが保存する値では、厳密な合計とシート式の合計が一致する
var storedPool = ['❌', '', '初投稿'];
for (var s = 1; s <= 100; s++) storedPool.push(s + '本');
for (var storedTrial = 0; storedTrial < 200; storedTrial++) {
  var storedRow = [];
  var storedLen = 1 + Math.floor(Math.random() * 40);
  for (var si = 0; si < storedLen; si++) storedRow.push(storedPool[Math.floor(Math.random() * storedPool.length)]);
  var strictSum = sandbox.postAppSumPostCells_(storedRow);
  var sheetSum = sandbox.postAppSheetFormulaSum_(storedRow);
  assert(strictSum === sheetSum, 'stored trial ' + storedTrial + ' strict=' + strictSum + ' sheet=' + sheetSum);
}
assert(sandbox.postAppSheetFormulaSum_(sample) === 55, 'シート式でも1〜10は55');
assert(sandbox.postAppSheetFormulaSum_(['❌', '', '11本', '20本', '50本', '100本', '0本', '初投稿']) === 11 + 20 + 50 + 100, 'シート式は0本と初投稿を足さない');
// サーバーが保存しない文字は、シート式だけ数えることがある。保存チェックは変えない。
assert(sandbox.postAppSheetFormulaSumCell_('10') === 10 && sandbox.postAppCountCell_('10') === 0, '「10」はシート式だけ');
assert(sandbox.postAppSheetFormulaSumCell_('101本') === 101 && sandbox.postAppCountCell_('101本') === 0, '101本はシート式だけ');
assert(sandbox.postAppSheetFormulaSumCell_('01本') === 1 && sandbox.postAppCountCell_('01本') === 0, '01本はシート式だけ');
assert(sandbox.postAppSheetFormulaSumCell_('0本') === 0 && sandbox.postAppCountCell_('0本') === 0, '0本はどちらも0');
assert(sandbox.postAppSheetFormulaSumCell_('10本 ') === 10 && sandbox.postAppCountCell_('10本 ') === 0, '末尾空白はシート式だけ');

assert(code.indexOf('setDataValidation') < 0, 'setDataValidation を呼ばない');
assert(code.indexOf('requireValueInList') < 0, 'プルダウンのリストを作らない');
assert(code.indexOf('applyPostCountDropdown_') < 0, '日付列への一括プルダウンを残さない');
assert(code.indexOf('postAppDropdownRule_') < 0, 'プルダウン規則を残さない');
assert(code.indexOf('clearDataValidations') >= 0, '残った入力規則は外す');
assert(code.indexOf('clearPostCountDropdown_') >= 0, '列追加でコピーされた入力規則を外す');

// 画面側: 1〜10の色クラスは今まで通り。11以上は別クラス。合計は parseInt で100まで足せる
var html = fs.readFileSync(path.join(root, 'post-app.html'), 'utf8');
var script = html.match(/<script>([\s\S]*)<\/script>/);
if (!script) throw new Error('script tag not found');
new Function(script[1]);
var ui = {};
vm.createContext(ui);
vm.runInContext(script[1].replace(
  "document.addEventListener('DOMContentLoaded',function(){playOpening(function(){initApp()})});",
  ''
), ui);
assert(typeof ui.postCountFromValue === 'function', 'postCountFromValue');
assert(typeof ui.postDayClass === 'function', 'postDayClass');
assert(typeof ui.cheerMessages === 'function', 'cheerMessages');
for (var d = 1; d <= 10; d++) {
  assert(ui.postDayClass(d + '本') === ' done' + d, '色クラス ' + d);
  var cheer = ui.cheerMessages(d);
  assert(cheer && cheer.length === 3, '1〜10のお祝い文 ' + d);
}
assert(ui.postDayClass('❌') === ' miss', '❌は未記録');
assert(ui.postDayClass('初投稿') === ' done1', '初投稿は従来の薄い青');
assert(ui.postDayClass('11本') === ' done15', '11本');
assert(ui.postDayClass('19本') === ' done15', '19本');
assert(ui.postDayClass('20本') === ' done20', '20本');
assert(ui.postDayClass('49本') === ' done20', '49本');
assert(ui.postDayClass('50本') === ' done50', '50本');
assert(ui.postDayClass('99本') === ' done50', '99本');
assert(ui.postDayClass('100本') === ' done100', '100本');
assert(ui.cheerMessages(10).indexOf('🌟 10本！完全体！！') >= 0, '10本のお祝いを残す');
assert(ui.cheerMessages(100)[0].indexOf('100本') >= 0, '100本のお祝い');
assert(ui.POST_MAX === 100, '画面の上限');

// 画面の合計は「N本」の数字を足す。11本以上も含まれる
function clientTotal(values) {
  var total = 0;
  values.forEach(function(v) {
    if (v && v !== '❌') total += (parseInt(v, 10) || 0);
  });
  return total;
}
assert(clientTotal(['❌', '1本', '10本', '11本', '100本']) === 122, '画面の合計');
assert(clientTotal(['❌', '1本', '10本']) === 11, '既存の画面合計は変わらない');

// ボタン 1〜10 は残っている
for (var b = 1; b <= 10; b++) {
  assert(html.indexOf('data-v="' + b + '本"') >= 0, 'ボタン ' + b);
}
assert(html.indexOf('data-v="❌"') >= 0, '❌ボタン');
assert(html.indexOf('id="postCountInput"') >= 0, '数字入力');

if (failed) {
  console.error(failed + ' failed');
  process.exit(1);
}
console.log('ok');

function range(a, b) {
  var out = [];
  for (var i = a; i <= b; i++) out.push(i);
  return out;
}
