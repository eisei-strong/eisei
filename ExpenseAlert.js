// ========================================
// ExpenseAlert.js — 経費アラート Chatwork定期通知
// ========================================

var EXP_ALERT_CW_ROOM = '381039351';
var EXP_ALERT_EXCLUDE_CATS = ['幹部報酬'];

/**
 * 毎週月曜に実行するトリガー関数
 * 当月の経費を分析してChatworkに通知
 */
function expenseAlertWeekly() {
  var ss = expSS_();
  var months = expFindMonths_(ss);
  if (months.length === 0) return;

  var cur = months[months.length - 1];
  var sheet = expFindCalcSheet_(ss, cur.year, cur.month);
  if (!sheet) return;

  var parsed = expParseSheet_(sheet);
  if (!parsed) return;

  // 幹部報酬を除外した経費合計
  var filteredExpTotal = 0;
  for (var i = 0; i < parsed.expenses.length; i++) {
    var pCat = expParentCat_(parsed.expenses[i].category, parsed.expenses[i].description);
    if (EXP_ALERT_EXCLUDE_CATS.indexOf(pCat) < 0) {
      filteredExpTotal += parsed.expenses[i].amount;
    }
  }

  var profit = parsed.totalIncome - filteredExpTotal;
  var profitRate = parsed.totalIncome > 0 ? Math.round((profit / parsed.totalIncome) * 1000) / 10 : 0;

  var lines = [];
  lines.push('[To:7769958]');
  lines.push('[info][title]' + cur.year + '年' + cur.month + '月 経費レポート[/title]');
  lines.push('着金: ' + formatYenCW_(parsed.totalIncome));
  lines.push('経費（幹部報酬除く）: ' + formatYenCW_(filteredExpTotal));
  lines.push('利益: ' + formatYenCW_(profit) + ' (利益率 ' + profitRate + '%)');
  lines.push('');

  // 削減検討リスト
  var wasteful = [];
  try {
    var ws = ss.getSheetByName(EXP_WASTEFUL_SHEET);
    if (ws) {
      var wd = ws.getDataRange().getValues();
      for (var w = 1; w < wd.length; w++) {
        var desc = String(wd[w][2] || '').trim();
        var amt = expParseNum_(wd[w][4] || wd[w][3]);
        if (desc || amt) wasteful.push({ description: desc, amount: amt });
      }
    }
  } catch (e) {}

  if (wasteful.length > 0) {
    lines.push('【削減検討リスト】');
    lines.push('');
    for (var k = 0; k < wasteful.length; k++) {
      lines.push('・' + wasteful[k].description);
      lines.push('　　→ ' + formatYenCW_(wasteful[k].amount));
      lines.push('');
    }
  } else {
    lines.push('削減検討リストは現在空です');
  }

  // ダブル課金チェック（同じ摘要＋同じ金額が複数回）
  var dupeMap = {};
  for (var d = 0; d < parsed.expenses.length; d++) {
    var ex = parsed.expenses[d];
    var key = ex.description + '|' + ex.amount;
    if (!dupeMap[key]) dupeMap[key] = [];
    dupeMap[key].push(ex);
  }
  var dupes = [];
  for (var dk in dupeMap) {
    if (dupeMap[dk].length >= 2 && dupeMap[dk][0].amount >= 10000) dupes.push(dupeMap[dk]);
  }
  if (dupes.length > 0) {
    lines.push('【⚠️ ダブル課金の可能性】');
    lines.push('');
    for (var di = 0; di < dupes.length; di++) {
      var dup = dupes[di];
      lines.push('⚠️ ' + dup[0].description + ' × ' + dup.length + '回');
      lines.push('　　→ ' + formatYenCW_(dup[0].amount) + ' / 回（計 ' + formatYenCW_(dup[0].amount * dup.length) + '）');
      lines.push('');
    }
  }

  lines.push('詳細: https://giver.work/expense-dashboard/');
  lines.push('[/info]');
  lines.push('');
  lines.push('※このメッセージは自動送信です。');
  lines.push('');
  lines.push('🔸アクションプラン');
  lines.push('上記の経費削減を検討し、実行したものを✅で報告してください！');

  var body = lines.join('\n');
  sendExpenseAlertCW_(body);
}

/**
 * Chatwork送信
 */
function sendExpenseAlertCW_(body) {
  var url = 'https://api.chatwork.com/v2/rooms/' + EXP_ALERT_CW_ROOM + '/messages';
  var options = {
    method: 'post',
    headers: { 'X-ChatWorkToken': CW_API_TOKEN },
    payload: { body: body },
    muteHttpExceptions: true
  };
  var res = UrlFetchApp.fetch(url, options);
  Logger.log('ExpenseAlert CW response: ' + res.getResponseCode() + ' ' + res.getContentText());
}

function formatYenCW_(val) {
  if (!val && val !== 0) return '¥0';
  var abs = Math.abs(val);
  var sign = val < 0 ? '-' : '';
  if (abs >= 10000) return sign + '¥' + Math.round(abs / 10000).toLocaleString() + '万';
  return sign + '¥' + Math.round(abs).toLocaleString();
}

/**
 * トリガー設定（1回だけ実行）
 */
/**
 * トリガー設定（1回だけ実行）
 * 毎日9時にチェックし、1日と15日のみ通知を送信
 */
function setupExpenseAlertTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'expenseAlertWeekly' || fn === 'expenseAlertBimonthly_') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('expenseAlertBimonthly_')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
  Logger.log('ExpenseAlert trigger created: 1st & 15th at 9:00');
}

/**
 * 毎日実行され、1日と15日のみ通知を送る
 */
function expenseAlertBimonthly_() {
  var today = new Date();
  var day = today.getDate();
  if (day === 1 || day === 15) {
    expenseAlertWeekly();
  }
}
