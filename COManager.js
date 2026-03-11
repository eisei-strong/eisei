// ============================================
// COManager.js — CO残管理シートロジック (v2)
// 詳細CO追跡: メンバー別のCO案件管理
// ============================================

/**
 * CO残管理シートの集計を更新
 * データ入力テーブル(行8〜)から退職者のCO未回収額を集計し、サマリーセクションに反映
 */
function updateCORemaining() {
  var ss = getSpreadsheet_();
  var coSheet = getCOManageSheet_(ss);
  if (!coSheet) {
    Logger.log('CO残管理シートが見つかりません');
    return;
  }

  var lastRow = coSheet.getLastRow();
  if (lastRow < CO_ROW_DATA_START) {
    // データなし — サマリーをゼロに
    clearCOSummary_(coSheet);
    return;
  }

  // データ読み取り（行8〜）
  var rowCount = lastRow - CO_ROW_DATA_START + 1;
  var data = coSheet.getRange(CO_ROW_DATA_START, 1, rowCount, CO_COL_COUNT).getValues();

  var memberTotals = {}; // { name: { coTotal, collected, remaining } }

  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][CO_COL_MEMBER - 1] || '').trim();
    if (!name) continue;

    var coAmount = round1_(parseNum_(data[i][CO_COL_CO_AMOUNT - 1]));
    var collected = round1_(parseNum_(data[i][CO_COL_COLLECTED - 1]));
    var remaining = round1_(coAmount - collected);

    if (!memberTotals[name]) {
      memberTotals[name] = { coTotal: 0, collected: 0, remaining: 0 };
    }

    memberTotals[name].coTotal += coAmount;
    memberTotals[name].collected += collected;
    memberTotals[name].remaining += remaining;
  }

  // サマリーセクション更新（行3〜: メンバー, CO総額, 回収済, 未回収残高）
  var memberNames = Object.keys(memberTotals);
  var summaryRows = [];
  var grandTotal = { coTotal: 0, collected: 0, remaining: 0 };

  for (var m = 0; m < memberNames.length; m++) {
    var mt = memberTotals[memberNames[m]];
    summaryRows.push([
      memberNames[m],
      round1_(mt.coTotal),
      round1_(mt.collected),
      round1_(mt.remaining)
    ]);
    grandTotal.coTotal += mt.coTotal;
    grandTotal.collected += mt.collected;
    grandTotal.remaining += mt.remaining;
  }

  // クリアしてから書き込み（行3-4）
  coSheet.getRange(CO_ROW_SUMMARY_START, 1, 2, 4).clearContent();
  if (summaryRows.length > 0) {
    var writeRows = Math.min(summaryRows.length, 2); // 最大2行（退職者は通常2人）
    coSheet.getRange(CO_ROW_SUMMARY_START, 1, writeRows, 4).setValues(summaryRows.slice(0, writeRows));
  }

  // 合計行（行5）
  coSheet.getRange(CO_ROW_SUMMARY_TOTAL, 1, 1, 4).setValues([
    ['合計', round1_(grandTotal.coTotal), round1_(grandTotal.collected), round1_(grandTotal.remaining)]
  ]);

  Logger.log('CO残更新完了: 未回収合計=' + round1_(grandTotal.remaining) + '万円');
  return {
    totalRemaining: round1_(grandTotal.remaining),
    memberTotals: memberTotals
  };
}

/**
 * CO残管理シートのサマリーをゼロクリア
 */
function clearCOSummary_(coSheet) {
  coSheet.getRange(CO_ROW_SUMMARY_START, 1, 2, 4).clearContent();
  coSheet.getRange(CO_ROW_SUMMARY_TOTAL, 1, 1, 4).setValues([['合計', 0, 0, 0]]);
}

/**
 * 現役メンバーのCO残を取得（サマリー表示用）
 */
function updateActiveMemberCO_() {
  var ss = getSpreadsheet_();
  var members = getActiveMembers_(ss);
  var coData = [];

  for (var i = 0; i < members.length; i++) {
    var dailySheet = getDailySheet_(ss, members[i].name);
    if (!dailySheet) continue;

    var data = readDailySheetData_(dailySheet);
    if (!data) continue;

    if (data.totals.coAmount > 0) {
      coData.push({
        name: members[i].name,
        displayName: members[i].displayName,
        coAmount: data.totals.coAmount,
        coCount: data.totals.coCount
      });
    }
  }

  return coData;
}
