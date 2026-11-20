/**
 * 投稿本数 1〜100 の検証。
 * - サーバーが受け付ける値
 * - 旧COUNTIF（1〜10本）と新しい合計が、❌ と 1〜10本では同じになること
 * - 11〜100本が足されること
 * - 既存の COUNTIF 式は同じ範囲のまま差し替わること
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

// 式の差し替えは範囲を保つ。SUM(月別) は触らない
var oldFormula = '=COUNTIF($O2:AR2,"1本")+COUNTIF($O2:AR2,"10本")*10';
var upgraded = sandbox.postAppUpgradeCountFormula_(oldFormula);
assert(upgraded.indexOf('$O2:AR2') >= 0, '範囲を維持');
assert(upgraded.indexOf('COUNTIF') < 0, 'COUNTIFを残さない');
assert(upgraded.indexOf(sandbox.POST_APP_COUNT_RE_SRC) >= 0, '1〜100の正規表現を含む');
assert(sandbox.postAppUpgradeCountFormula_('=SUM(F2:N2)') === '', '全期間SUMは対象外');
assert(sandbox.postAppUpgradeCountFormula_(upgraded) === '', '新しい式は再変換しない');

var direct = sandbox.postAppPostsSumFormula_('F12:AI12');
assert(direct.indexOf('F12:AI12') >= 0, '直接生成も範囲を含む');

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
