// ============================================
// コード.js — ダッシュボード サーバーサイド (v2)
// getDashboardData + doGet + メニュー
// ============================================

/**
 * スプシを開いた時にメニューを追加
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚔️ ダッシュボード')
    .addItem('📊 ダッシュボードを開く', 'openDashboard')
    .addSeparator()
    .addItem('🔄 サマリー更新', 'updateSummary')
    .addItem('📅 月締め処理', 'monthlyArchive')
    .addItem('🆕 新月セットアップ', 'newMonth')
    .addSeparator()
    .addItem('⚙️ シート再生成', 'setupAllSheets')
    .addItem('📦 データ移行（v1→v2）', 'migrateV1toV2')
    .addItem('💳 クレカ/信販更新', 'updateCreditShinpan')
    .addItem('🔄 旧シート同期', 'syncAndUpdate')
    .addItem('🙈 旧シート非表示', 'hideOldSheets')
    .addSeparator()
    .addItem('🤖 FBぼっと開始', 'setupBotTrigger')
    .addItem('⏹ FBぼっと停止', 'stopBot')
    .addItem('📋 Bot用シート作成', 'setupBotSheets')
    .addToUi();
}

function openDashboard() {
  var html = HtmlService.createHtmlOutput(
    '<script>window.open("https://script.google.com/macros/s/AKfycbwojGHuvzycc07FJKwBdbBJJQZpssF6lYz0DbNJlu6zsVuXkAj8V8w3XNBPieo2wsYbFg/exec");google.script.host.close();</script>'
  ).setWidth(90).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, '開いています...');
}

/**
 * 「営業ダッシュボード」タブを作成（初回のみ実行）
 */
function createDashboardTab() {
  var ss = getSpreadsheet_();
  var tabName = '営業ダッシュボード';
  var url = 'https://script.google.com/macros/s/AKfycbwojGHuvzycc07FJKwBdbBJJQZpssF6lYz0DbNJlu6zsVuXkAj8V8w3XNBPieo2wsYbFg/exec';

  var existing = ss.getSheetByName(tabName);
  if (existing) ss.deleteSheet(existing);

  var sheet = ss.insertSheet(tabName, 0);

  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 500);
  sheet.setRowHeight(1, 20);
  sheet.setRowHeight(2, 60);
  sheet.setRowHeight(3, 50);
  sheet.setRowHeight(4, 40);

  var titleCell = sheet.getRange('B2');
  titleCell.setValue('⚔️ Namaka ウォーリアーズ 営業ダッシュボード');
  titleCell.setFontSize(22).setFontWeight('bold').setFontColor('#7C2D12');

  var linkCell = sheet.getRange('B3');
  linkCell.setFormula('=HYPERLINK("' + url + '", "📊 ダッシュボードを開く →")');
  linkCell.setFontSize(16).setFontWeight('bold').setFontColor('#FFFFFF');
  linkCell.setBackground('#EA580C');
  linkCell.setHorizontalAlignment('center');
  linkCell.setVerticalAlignment('middle');

  var descCell = sheet.getRange('B4');
  descCell.setValue('↑ クリックすると営業ダッシュボードが開きます（スマホでもOK）');
  descCell.setFontSize(11).setFontColor('#94A3B8');

  sheet.setTabColor('#EA580C');

  sheet.hideColumns(3, sheet.getMaxColumns() - 2);
  sheet.hideRows(6, sheet.getMaxRows() - 5);
}

// ============================================
// ウェブアプリ エントリーポイント
// ============================================

/**
 * POST: KBデータ受信用（チャンク送信対応）
 */
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action || '';

    if (action === 'kb_import') {
      var ss = getSpreadsheet_();
      var sheetName = payload.sheet || 'KB_Staging';
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
      }
      var rows = payload.rows || [];
      if (rows.length > 0) {
        var startRow = sheet.getLastRow() + 1;
        sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
      }
      return ContentService.createTextOutput(JSON.stringify({
        ok: true, appended: rows.length, totalRows: sheet.getLastRow()
      })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ error: 'unknown action' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  var params = e ? e.parameter : {};

  if (params.action === 'api') {
    try {
      var result;
      if (params.type === 'kpi') {
        result = getKpiDashboardData();
      } else if (params.type === 'months') {
        result = getAvailableMonths();
      } else if (params.month && params.year) {
        result = getDashboardData(parseInt(params.month), parseInt(params.year));
      } else {
        result = getDashboardData();
      }
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (params.action === 'run') {
    try {
      var fn = params.fn || '';
      var fnResult;
      if (fn === 'updateSummary') fnResult = updateSummary();
      else if (fn === 'archive') fnResult = monthlyArchive();
      else if (fn === 'newMonth') fnResult = newMonth();
      else if (fn === 'setupTriggers') fnResult = setupSummaryTrigger();
      else if (fn === 'showOldSheets') fnResult = showOldSheets();
      else if (fn === 'migrateV1toV2') fnResult = migrateV1toV2_api();
      else if (fn === 'updateCreditShinpan') fnResult = updateCreditShinpan();
      else if (fn === 'clearCarryover') fnResult = clearCarryoverRows();
      else if (fn === 'fixDailyFormulas') fnResult = fixDailyFormulas();
      else if (fn === 'debugSheetInfo') fnResult = debugFreezeBlockers();
      else if (fn === 'syncOldSheet') fnResult = syncAndUpdate();
      else if (fn === 'debugOldSheet') fnResult = debugOldSheet();
      else if (fn === 'fixRefErrors') fnResult = fixRefErrors();
      else if (fn === 'unmergeAndFreeze') fnResult = unmergeAndFreeze();
      else if (fn === 'formatOldSheet') fnResult = formatOldSheet();
      else if (fn === 'addGonMember') fnResult = addGonMember();
      else if (fn === 'addTonyMember') fnResult = addTonyMember();
      else if (fn === 'debugPrevMonth') fnResult = debugPrevMonth();
      else if (fn === 'listSheets') fnResult = listSheets();
      else if (fn === 'debugCBSLifety') fnResult = debugCBSLifety();
      else if (fn === 'debugShinpanEJ') fnResult = debugShinpanEJ();
      else if (fn === 'refreshSummaryFormat') fnResult = refreshSummaryFormat();
      else if (fn === 'archiveFromOldSheet') fnResult = archiveFromOldSheet(params.month, params.year);
      else if (fn === 'deleteArchive') fnResult = deleteArchive(params.month, params.year);
      else if (fn === 'rebuildAllArchives') fnResult = rebuildAllArchives();
      else if (fn === 'debugDashboardDeals') fnResult = debugDashboardDeals();
      else if (fn === 'debugArchiveAll') fnResult = debugArchiveAll();
      else if (fn === 'getAvailableMonths') fnResult = getAvailableMonths();
      else if (fn === 'debugOldSheetRows') fnResult = debugOldSheetRows();
      else if (fn === 'fixOldSheetFormulas') fnResult = fixOldSheetFormulas();
      else if (fn === 'debugSectionMapping') fnResult = debugSectionMapping();
      else if (fn === 'debugSectionCarryover') fnResult = debugSectionCarryover();
      else if (fn === 'insertMemberImages') fnResult = insertMemberImages();
      else if (fn === 'setNameLinks') fnResult = setNameLinks();
      else if (fn === 'debugSheetByGid') fnResult = debugSheetByGid(params.gid, params.rows, params.formulas === '1');
      else if (fn === 'repairKpiCalcSheet') fnResult = repairKpiCalcSheet();
      else if (fn === 'getKpiDashboardData') fnResult = getKpiDashboardData();
      else if (fn === 'rebuildSummarySheet') {
        var ss2 = getSpreadsheet_();
        createSummarySheet_(ss2);
        updateSummary();
        fnResult = { status: 'rebuilt' };
      }
      else if (fn === 'startBot') fnResult = setupBotTrigger();
      else if (fn === 'setBotToken') fnResult = setBotToken();
      else if (fn === 'stopBot') fnResult = stopBot();
      else if (fn === 'botStatus') fnResult = getBotStatus();
      else if (fn === 'setupBotSheets') fnResult = setupBotSheets();
      else if (fn === 'importKB') fnResult = processStaging(params.sheet);
      else if (fn === 'testFeedback') fnResult = testFeedbackGeneration(params.roomId, params.msgBody);
      else if (fn === 'getPaymentNewsOnly') fnResult = getPaymentNewsOnly_();
      else if (fn === 'renameMember') fnResult = renameMember(params.from, params.to);
      else if (fn === 'fixCell') fnResult = fixCell(params.gid, params.row, params.col, params.val, params.formula);
      else if (fn === 'fixRankOnly') fnResult = fixRankOnly();
      else fnResult = { error: 'unknown fn: ' + fn };
      return ContentService.createTextOutput(JSON.stringify({ ok: true, result: fnResult }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (params.action === 'createTab') {
    try {
      createDashboardTab();
      return ContentService.createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('営業ダッシュボード')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================
// getDashboardData — デュアルモード
// ============================================

function getDashboardData(targetMonth, targetYear) {
  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);

  // 引数あり & 当月以外 → アーカイブから取得
  if (targetMonth && targetYear) {
    targetMonth = parseNum_(targetMonth);
    targetYear = parseNum_(targetYear);
    if (targetMonth !== settings.month || targetYear !== settings.year) {
      return getDashboardDataFromArchive_(ss, targetYear, targetMonth);
    }
  }

  // 当月 → 既存のリアルタイム処理
  var settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  if (settingsSheet) {
    // NOTE: サマリーが空でもupdateSummary()を呼ばない（syncFromOldSheet回避）
    // updateSummaryはメニューまたはトリガーから手動実行すること

    if (isNewStructureReady_(ss)) {
      return getDashboardData_new_(ss);
    }
  }

  return getDashboardData_legacy_(ss);
}

/**
 * アーカイブに存在する月一覧+当月を返す
 */
function getAvailableMonths() {
  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);
  var archiveSheet = getArchiveSheet_(ss);

  var months = [];
  var seen = {};

  if (archiveSheet && archiveSheet.getLastRow() > 1) {
    var data = archiveSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var y = parseNum_(data[i][0]);
      var m = parseNum_(data[i][1]);
      var key = y + '-' + m;
      if (y > 0 && m > 0 && !seen[key]) {
        seen[key] = true;
        months.push({ year: y, month: m });
      }
    }
  }

  // 当月を追加
  var currentKey = settings.year + '-' + settings.month;
  if (!seen[currentKey]) {
    months.push({ year: settings.year, month: settings.month });
  }

  // ソート（古い順）
  months.sort(function(a, b) { return (a.year * 12 + a.month) - (b.year * 12 + b.month); });

  return { months: months, current: { year: settings.year, month: settings.month } };
}

/**
 * アーカイブから過去月のダッシュボードデータを取得
 */
function getDashboardDataFromArchive_(ss, year, month) {
  var archiveSheet = getArchiveSheet_(ss);
  if (!archiveSheet || archiveSheet.getLastRow() <= 1) {
    return { error: 'アーカイブデータなし', members: [], totalRevenue: 0 };
  }

  var allData = archiveSheet.getDataRange().getValues();
  var settings = getGlobalSettings_(ss);

  // メンバー名→表示名のマップを事前に構築
  var allSettingsMembers = getMembersFromSettings_(ss);
  var nameToDisplay = {};
  var nameToIcon = {};
  for (var j = 0; j < allSettingsMembers.length; j++) {
    nameToDisplay[allSettingsMembers[j].name] = allSettingsMembers[j].displayName;
    nameToIcon[allSettingsMembers[j].name] = allSettingsMembers[j].iconUrl;
  }

  // 対象月のデータを収集
  var memberMap = {};
  for (var i = 1; i < allData.length; i++) {
    if (parseNum_(allData[i][ARC_COL_YEAR - 1]) === year &&
        parseNum_(allData[i][ARC_COL_MONTH - 1]) === month) {
      var name = String(allData[i][ARC_COL_NAME - 1] || '').trim();
      if (!name) continue;

      var displayName = nameToDisplay[name] || DISPLAY_NAME_MAP[name] || name;

      memberMap[name] = {
        name: displayName,
        icon: nameToIcon[name] || ICON_MAP[name] || ICON_MAP[displayName] || '',
        fundedDeals: parseNum_(allData[i][ARC_COL_FUNDED_DEALS - 1]),
        closed: parseNum_(allData[i][ARC_COL_FUNDED_CLOSED - 1]),
        deals: parseNum_(allData[i][ARC_COL_TOTAL_DEALS - 1]),
        closeRate: round1_(parseNum_(allData[i][ARC_COL_CLOSE_RATE - 1])),
        sales: round1_(parseNum_(allData[i][ARC_COL_SALES - 1])),
        revenue: round1_(parseNum_(allData[i][ARC_COL_REVENUE - 1])),
        coRevenue: parseNum_(allData[i][ARC_COL_CO_COUNT - 1]),
        coAmount: round1_(parseNum_(allData[i][ARC_COL_CO_AMOUNT - 1])),
        creditCard: round1_(parseNum_(allData[i][ARC_COL_CREDIT_CARD - 1])),
        shinpan: round1_(parseNum_(allData[i][ARC_COL_SHINPAN - 1])),
        cbs: String(allData[i][ARC_COL_CBS - 1] || '-'),
        lifety: String(allData[i][ARC_COL_LIFETY - 1] || '-'),
        avgPrice: 0
      };
    }
  }

  // 前月データ（さらに1つ前の月）を取得
  var prevMonth = month === 1 ? 12 : month - 1;
  var prevYear = month === 1 ? year - 1 : year;
  var prevData = {};
  for (var i = 1; i < allData.length; i++) {
    if (parseNum_(allData[i][ARC_COL_YEAR - 1]) === prevYear &&
        parseNum_(allData[i][ARC_COL_MONTH - 1]) === prevMonth) {
      var pName = String(allData[i][ARC_COL_NAME - 1] || '').trim();
      if (pName) {
        prevData[pName] = {
          revenue: round1_(parseNum_(allData[i][ARC_COL_REVENUE - 1])),
          deals: parseNum_(allData[i][ARC_COL_TOTAL_DEALS - 1]),
          closed: parseNum_(allData[i][ARC_COL_FUNDED_CLOSED - 1]),
          closeRate: round1_(parseNum_(allData[i][ARC_COL_CLOSE_RATE - 1]))
        };
      }
    }
  }

  // members配列に変換 + 前月比計算
  var members = [];
  for (var key in memberMap) {
    var m = memberMap[key];
    var prev = prevData[key] || prevData[m.name];
    if (!prev) {
      for (var oldN in LEGACY_TO_V2_NAME) {
        if (LEGACY_TO_V2_NAME[oldN] === key && prevData[oldN]) { prev = prevData[oldN]; break; }
      }
    }
    if (!prev) prev = { revenue: 0, deals: 0, closed: 0, closeRate: 0 };
    m.prevRevenue = prev.revenue;
    m.diffRevenue = round1_(m.revenue - prev.revenue);
    m.prevDeals = prev.deals;
    m.diffDeals = m.deals - prev.deals;
    m.prevClosed = prev.closed;
    m.diffClosed = m.closed - prev.closed;
    m.prevCloseRate = prev.closeRate;
    m.diffCloseRate = round1_(m.closeRate - prev.closeRate);
    members.push(m);
  }

  // 着金額で降順ソート
  members.sort(function(a, b) { return b.revenue - a.revenue; });

  // ランク計算
  for (var r = 0; r < members.length; r++) {
    if (r > 0 && members[r].revenue === members[r - 1].revenue) {
      members[r].rank = members[r - 1].rank;
    } else {
      members[r].rank = r + 1;
    }
    members[r].gapToTop = r > 0 ? round1_(members[0].revenue - members[r].revenue) : 0;
  }

  var totalRevenue = round1_(members.reduce(function(s, m) { return s + m.revenue; }, 0));
  var teamGoal = settings.teamGoal;

  // 前月チーム合計（退職メンバー含む全員分）
  var prevTotalDeals = 0, prevTotalRevenue = 0, prevTotalClosed = 0;
  for (var pKey in prevData) {
    prevTotalDeals += prevData[pKey].deals || 0;
    prevTotalRevenue += round1_(prevData[pKey].revenue || 0);
    prevTotalClosed += prevData[pKey].closed || 0;
  }

  return {
    members: members,
    totalRevenue: totalRevenue,
    teamGoal: teamGoal,
    remaining: round1_(Math.max(0, teamGoal - totalRevenue)),
    progressRate: Math.min(100, round1_((totalRevenue / teamGoal) * 100)),
    dailyTarget: 0,
    daysLeft: 0,
    currentMonth: month,
    paymentNews: [],
    prevTotalDeals: prevTotalDeals,
    prevTotalRevenue: round1_(prevTotalRevenue),
    prevTotalClosed: prevTotalClosed,
    updatedAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
    isArchive: true,
    archiveYear: year
  };
}

// ============================================
// 新構造からデータ取得 (v2)
// ============================================

function getDashboardData_new_(ss) {
  var settings = getGlobalSettings_(ss);
  var activeMembers = getActiveMembers_(ss);
  var currentMonth = settings.month;

  var prevMonthData = getPrevMonthFromArchive_(ss, settings);

  var members = [];

  for (var i = 0; i < activeMembers.length; i++) {
    var m = activeMembers[i];
    var dailySheet = getDailySheet_(ss, m.name);
    var sheetData = readDailySheetData_(dailySheet);

    if (!sheetData) {
      members.push(createEmptyMember_(m, prevMonthData));
      continue;
    }

    var d = sheetData.totals;
    var prev = prevMonthData[m.name] || prevMonthData[m.displayName];
    if (!prev) {
      for (var oldN in LEGACY_TO_V2_NAME) {
        if (LEGACY_TO_V2_NAME[oldN] === m.name && prevMonthData[oldN]) { prev = prevMonthData[oldN]; break; }
      }
    }
    if (!prev) prev = { revenue: 0, deals: 0, closed: 0, closeRate: 0 };

    members.push({
      name: m.displayName,
      icon: m.iconUrl,
      deals: d.totalDeals,
      closed: d.fundedClosed,
      revenue: d.revenue,
      closeRate: d.closeRate,
      fundedDeals: d.fundedDeals,
      sales: d.sales,
      avgPrice: d.avgPrice,
      coRevenue: d.coCount,  // Dashboard互換: coRevenue→CO数
      coAmount: d.coAmount,
      creditCard: d.creditCard,
      shinpan: d.shinpan,
      cbs: sheetData.cbs,
      lifety: sheetData.lifety,
      prevRevenue: prev.revenue,
      diffRevenue: round1_(d.revenue - prev.revenue),
      prevDeals: prev.deals,
      diffDeals: d.totalDeals - prev.deals,
      prevClosed: prev.closed,
      diffClosed: d.fundedClosed - prev.closed,
      prevCloseRate: prev.closeRate,
      diffCloseRate: round1_(d.closeRate - prev.closeRate)
    });
  }

  // 着金額で降順ソート
  members.sort(function(a, b) { return b.revenue - a.revenue; });

  // ランク・逆転差額を計算
  for (var r = 0; r < members.length; r++) {
    if (r > 0 && members[r].revenue === members[r - 1].revenue) {
      members[r].rank = members[r - 1].rank;
    } else {
      members[r].rank = r + 1;
    }
    members[r].gapToTop = r > 0
      ? round1_(members[0].revenue - members[r].revenue)
      : 0;
  }

  // チーム合計
  var totalRevenue = round1_(members.reduce(function(sum, m) { return sum + m.revenue; }, 0));
  var teamGoal = settings.teamGoal;
  var remaining = round1_(Math.max(0, teamGoal - totalRevenue));

  // 前月チーム合計（退職メンバー含む全員分）
  var prevTotalDeals = 0, prevTotalRevenue = 0, prevTotalClosed = 0;
  for (var pKey in prevMonthData) {
    prevTotalDeals += prevMonthData[pKey].deals || 0;
    prevTotalRevenue += round1_(prevMonthData[pKey].revenue || 0);
    prevTotalClosed += prevMonthData[pKey].closed || 0;
  }

  // 日割り
  var now = new Date();
  var today = now.getDate();
  var lastDay = new Date(settings.year, currentMonth, 0).getDate();
  var daysLeft = Math.max(1, lastDay - today);
  var dailyTarget = remaining > 0 ? round1_(remaining / daysLeft) : 0;

  // 着金速報
  var paymentNews = getPaymentNews_new_(ss, activeMembers);

  return {
    members: members,
    totalRevenue: totalRevenue,
    teamGoal: teamGoal,
    remaining: remaining,
    progressRate: Math.min(100, round1_((totalRevenue / teamGoal) * 100)),
    dailyTarget: dailyTarget,
    daysLeft: daysLeft,
    currentMonth: currentMonth,
    paymentNews: paymentNews,
    prevTotalDeals: prevTotalDeals,
    prevTotalRevenue: round1_(prevTotalRevenue),
    prevTotalClosed: prevTotalClosed,
    updatedAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')
  };
}

/**
 * 空のメンバーデータを生成
 */
function createEmptyMember_(m, prevMonthData) {
  var prev = prevMonthData[m.name] || prevMonthData[m.displayName];
  if (!prev) {
    for (var oldN in LEGACY_TO_V2_NAME) {
      if (LEGACY_TO_V2_NAME[oldN] === m.name && prevMonthData[oldN]) { prev = prevMonthData[oldN]; break; }
    }
  }
  if (!prev) prev = { revenue: 0, deals: 0, closed: 0, closeRate: 0 };
  return {
    name: m.displayName, icon: m.iconUrl,
    deals: 0, closed: 0, revenue: 0, closeRate: 0,
    fundedDeals: 0, sales: 0, avgPrice: 0, coRevenue: 0, coAmount: 0,
    creditCard: 0, shinpan: 0, cbs: '-', lifety: '-',
    prevRevenue: prev.revenue, diffRevenue: round1_(0 - prev.revenue),
    prevDeals: prev.deals, diffDeals: 0 - prev.deals,
    prevClosed: prev.closed, diffClosed: 0 - prev.closed,
    prevCloseRate: prev.closeRate, diffCloseRate: round1_(0 - prev.closeRate)
  };
}

/**
 * 新構造: 着金速報データを日別入力シートから取得
 */
function getPaymentNews_new_(ss, activeMembers) {
  var news = [];

  for (var i = 0; i < activeMembers.length; i++) {
    var m = activeMembers[i];
    var dailySheet = getDailySheet_(ss, m.name);
    if (!dailySheet) continue;

    var data = dailySheet.getRange(DAILY_ROW_DATA_START, 1, DAILY_ROW_DATA_END - DAILY_ROW_DATA_START + 1, COL_REVENUE + 1).getValues();

    for (var r = 0; r < data.length; r++) {
      var dateVal = data[r][COL_DATE];
      var revVal = data[r][COL_REVENUE];
      if (!(dateVal instanceof Date)) continue;

      var amount = round1_(parseNum_(revVal));
      if (amount > 0) {
        news.push({
          date: Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy-MM-dd'),
          dateShort: (dateVal.getMonth() + 1) + '/' + dateVal.getDate(),
          name: m.displayName,
          icon: m.iconUrl,
          amount: amount
        });
      }
    }
  }

  news.sort(function(a, b) {
    if (a.date > b.date) return -1;
    if (a.date < b.date) return 1;
    return b.amount - a.amount;
  });

  return news;
}

/**
 * 着金速報だけ返す軽量API（syncを呼ばない）
 */
function getPaymentNewsOnly_() {
  var ss = getSpreadsheet_();
  var now = new Date();
  var currentMonth = now.getMonth() + 1;
  var sheet = getSheetByMonth_(ss, currentMonth);
  if (!sheet) return [];
  var allData = sheet.getDataRange().getValues();
  return getPaymentNews_legacy_(sheet, allData);
}

// ============================================
// 旧構造からデータ取得（フォールバック）
// ============================================

function getDashboardData_legacy_(ss) {
  var now = new Date();
  var currentMonth = now.getMonth() + 1;
  var prevMonth = currentMonth === 1 ? 12 : currentMonth - 1;

  var sheet = getSheetByMonth_(ss, currentMonth);
  if (!sheet) {
    throw new Error(currentMonth + '月のシートが見つかりません');
  }

  var allData = sheet.getDataRange().getValues();
  var prevData = getPrevMonthData_legacy_(ss, prevMonth);

  var members = [];

  for (var i = 0; i < OLD_MEMBER_COLS.length; i++) {
    var col = OLD_MEMBER_COLS[i];
    var name = OLD_MEMBER_NAME_MAP[col];
    if (!name) continue;

    var fundedDeals = parseNum_(allData[OLD_ROW_FUNDED_DEALS][col]);
    var deals   = parseNum_(allData[OLD_ROW_DEALS][col]);
    var closed  = parseNum_(allData[OLD_ROW_CLOSED][col]);
    var revenue = round1_(parseNum_(allData[OLD_ROW_REVENUE][col]));
    var sales   = round1_(parseNum_(allData[OLD_ROW_SALES][col]));
    var avgPrice = round1_(parseNum_(allData[OLD_ROW_AVG_PRICE][col]));
    var coRevenue = round1_(parseNum_(allData[OLD_ROW_CO_REVENUE][col]));
    var coAmount = round1_(parseNum_(allData[OLD_ROW_CO_AMOUNT][col]));
    var creditCard = round1_(parseNum_(allData[OLD_ROW_CREDIT_CARD][col]));
    var shinpan = round1_(parseNum_(allData[OLD_ROW_SHINPAN][col]));

    var cbsRaw = allData[OLD_ROW_CBS][col];
    var cbs = (!cbsRaw && cbsRaw !== 0) ? '-' : String(cbsRaw);
    var lifetyRaw = allData[OLD_ROW_LIFETY][col];
    var lifety = (!lifetyRaw && lifetyRaw !== 0) ? '-' : String(lifetyRaw);

    var rateRaw = allData[OLD_ROW_RATE][col];
    var closeRate = typeof rateRaw === 'number'
      ? round1_(rateRaw * 100)
      : (deals > 0 ? round1_((closed / deals) * 100) : 0);

    // v2メンバー名で前月データ検索
    var v2Name = LEGACY_TO_V2_NAME[name] || displayName_(name);
    var prev = prevData[v2Name] || prevData[name] || { revenue: 0, deals: 0, closed: 0, closeRate: 0 };

    members.push({
      name: v2Name, icon: iconUrl_(v2Name),
      deals: deals, closed: closed, revenue: revenue, closeRate: closeRate,
      fundedDeals: fundedDeals, sales: sales, avgPrice: avgPrice,
      coRevenue: coRevenue, coAmount: coAmount,
      creditCard: creditCard, shinpan: shinpan, cbs: cbs, lifety: lifety,
      prevRevenue: prev.revenue, diffRevenue: round1_(revenue - prev.revenue),
      prevDeals: prev.deals, diffDeals: deals - prev.deals,
      prevClosed: prev.closed, diffClosed: closed - prev.closed,
      prevCloseRate: prev.closeRate, diffCloseRate: round1_(closeRate - prev.closeRate)
    });
  }

  members.sort(function(a, b) { return b.revenue - a.revenue; });

  for (var r = 0; r < members.length; r++) {
    if (r > 0 && members[r].revenue === members[r - 1].revenue) {
      members[r].rank = members[r - 1].rank;
    } else {
      members[r].rank = r + 1;
    }
    members[r].gapToTop = r > 0
      ? round1_(members[0].revenue - members[r].revenue)
      : 0;
  }

  var totalRevenue = round1_(members.reduce(function(sum, m) { return sum + m.revenue; }, 0));
  var remaining = round1_(Math.max(0, TEAM_GOAL - totalRevenue));

  var today = now.getDate();
  var lastDay = new Date(now.getFullYear(), currentMonth, 0).getDate();
  var daysLeft = Math.max(1, lastDay - today);
  var dailyTarget = remaining > 0 ? round1_(remaining / daysLeft) : 0;

  var paymentNews = getPaymentNews_legacy_(sheet, allData);

  return {
    members: members,
    totalRevenue: totalRevenue,
    teamGoal: TEAM_GOAL,
    remaining: remaining,
    progressRate: Math.min(100, round1_((totalRevenue / TEAM_GOAL) * 100)),
    dailyTarget: dailyTarget,
    daysLeft: daysLeft,
    currentMonth: currentMonth,
    paymentNews: paymentNews,
    updatedAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')
  };
}

/**
 * 旧構造: 着金速報データを取得（レガシー）
 */
function getPaymentNews_legacy_(sheet, allData) {
  var maxCol = allData[0].length;
  var news = [];

  var headerRows = [];
  for (var r = 40; r < allData.length; r++) {
    if (String(allData[r][1] || '').replace(/\s+/g, '') === '日別') {
      headerRows.push(r);
    }
  }

  for (var h = 0; h < headerRows.length; h++) {
    var hr = headerRows[h];

    var sections = [];
    for (var c = 0; c < maxCol; c++) {
      var label = String(allData[hr][c] || '').replace(/\s+/g, '');
      if (label === '着金額') {
        var dateCol = -1;
        for (var dc = c - 1; dc >= 0; dc--) {
          if (String(allData[hr][dc] || '').replace(/\s+/g, '') === '日別') {
            dateCol = dc;
            break;
          }
        }
        if (dateCol >= 0) {
          sections.push({ dateCol: dateCol, revCol: c });
        }
      }
    }

    var rankRow = hr - 2;
    var rankFormulas = [];
    if (rankRow >= 0) {
      rankFormulas = sheet.getRange(rankRow + 1, 1, 1, maxCol).getFormulas()[0];
    }

    for (var s = 0; s < sections.length; s++) {
      var sec = sections[s];

      var memberCol = -1;
      for (var fc = sec.revCol; fc <= Math.min(sec.revCol + 8, maxCol - 1); fc++) {
        var formula = rankFormulas[fc] || '';
        if (formula.indexOf('RANK') !== -1) {
          var match = formula.match(/RANK\(([A-Z]+)\d+/i);
          if (match) {
            var letter = match[1].toUpperCase();
            var colIdx = 0;
            for (var li = 0; li < letter.length; li++) {
              colIdx = colIdx * 26 + (letter.charCodeAt(li) - 64);
            }
            memberCol = colIdx - 1;
            break;
          }
        }
      }
      if (memberCol < 0) continue;

      var rawName = OLD_MEMBER_NAME_MAP[memberCol];
      if (!rawName) {
        rawName = String(allData[5][memberCol] || '').trim();
        if (!rawName || rawName === '\\' || rawName === '合計') continue;
        rawName = normalizeName_(rawName);
      }

      var v2Name = LEGACY_TO_V2_NAME[rawName] || displayName_(rawName);
      var icon = iconUrl_(v2Name);

      for (var dr = hr + 1; dr < Math.min(hr + 32, allData.length); dr++) {
        var dateVal = allData[dr][sec.dateCol];
        var revVal = allData[dr][sec.revCol];
        if (dateVal instanceof Date && revVal && revVal !== 0) {
          var amount = round1_(parseNum_(revVal));
          if (amount > 0) {
            news.push({
              date: Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy-MM-dd'),
              dateShort: (dateVal.getMonth() + 1) + '/' + dateVal.getDate(),
              name: v2Name,
              icon: icon,
              amount: amount
            });
          }
        }
      }
    }
  }

  news.sort(function(a, b) {
    if (a.date > b.date) return -1;
    if (a.date < b.date) return 1;
    return b.amount - a.amount;
  });

  return news;
}
