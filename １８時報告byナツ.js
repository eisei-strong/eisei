// ============================================
// 18時報告byナツ.js — Chatwork自動報告
// サマリーシートから読み取り
// ============================================

/**
 * 18時のウォーリアーズ報告をChatworkに送信
 */
function sendWarriorsReportToChatwork() {
  var chatworkApiToken = getChatworkToken_();
  var roomId = "406745569";
  if (!chatworkApiToken) { Logger.log('CHATWORK_API_TOKEN未設定'); return; }

  var ss = getSpreadsheet_();

  // 新構造が有効ならサマリーから読み取り、なければ旧構造から
  var ranking;
  if (isNewStructureReady_(ss)) {
    ranking = getRankingFromSummary_(ss);
  } else {
    ranking = getRankingFromOldSheet_(ss);
  }

  if (!ranking || ranking.length === 0) {
    Logger.log('ランキングデータが取得できません');
    return;
  }

  // 着金ランキング
  var amountRanking = ranking.slice().sort(function(a, b) { return b.amount - a.amount; });

  // 成約率ランキング（商談10以上）
  var rateRanking = ranking
    .filter(function(r) { return r.deals >= 10; })
    .sort(function(a, b) { return b.rate - a.rate; });

  // メッセージ作成
  var msg = "👑 着金ランキング（自動送信）\n";
  for (var i = 0; i < amountRanking.length; i++) {
    msg += (i + 1) + ". " + amountRanking[i].name + " " + amountRanking[i].amount.toFixed(1) + "万円\n";
  }

  msg += "\n👑 成約率ランキング（商談10以上）\n";
  for (var j = 0; j < rateRanking.length; j++) {
    msg += (j + 1) + ". " + rateRanking[j].name + " " + (rateRanking[j].rate * 100).toFixed(1) + "%\n";
  }

  // 送信
  UrlFetchApp.fetch(
    "https://api.chatwork.com/v2/rooms/" + roomId + "/messages",
    {
      method: "post",
      headers: { "X-ChatWorkToken": chatworkApiToken },
      payload: { body: msg }
    }
  );

  Logger.log("送信完了\n" + msg);
}

/**
 * サマリーシートからランキングデータを取得（新構造）
 */
function getRankingFromSummary_(ss) {
  var summarySheet = getSummarySheet_(ss);
  if (!summarySheet) return null;

  var data = summarySheet.getRange(SM_ROW_MEMBER_START, 1, 8, SM_COL_GAP_TO_TOP).getValues();
  var ranking = [];

  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][SM_COL_DISPLAY - 1] || '').trim();
    if (!name) continue;

    var revenue = parseNum_(data[i][SM_COL_REVENUE - 1]);
    var deals = parseNum_(data[i][SM_COL_DEALS - 1]);
    var closed = parseNum_(data[i][SM_COL_CLOSED - 1]);
    var closeRate = parseNum_(data[i][SM_COL_CLOSE_RATE - 1]);

    ranking.push({
      name: name,
      amount: revenue,
      deals: deals,
      rate: closeRate / 100 // 報告用に0-1スケール
    });
  }

  return ranking;
}

/**
 * 旧シートからランキングデータを取得（フォールバック）
 */
function getRankingFromOldSheet_(ss) {
  var now = new Date();
  var currentMonth = now.getMonth() + 1;
  var sheet = getSheetByMonth_(ss, currentMonth);
  if (!sheet) return null;

  var allData = sheet.getDataRange().getValues();
  var ranking = [];

  for (var i = 0; i < OLD_MEMBER_COLS.length; i++) {
    var col = OLD_MEMBER_COLS[i];
    var name = OLD_MEMBER_NAME_MAP[col];
    if (!name) continue;

    var revenue = round1_(parseNum_(allData[OLD_ROW_REVENUE][col]));
    var deals = parseNum_(allData[OLD_ROW_DEALS][col]);
    var rateRaw = allData[OLD_ROW_RATE][col];
    var closeRate = typeof rateRaw === 'number' ? rateRaw : 0;

    ranking.push({
      name: displayName_(name),
      amount: revenue,
      deals: deals,
      rate: closeRate
    });
  }

  return ranking;
}
