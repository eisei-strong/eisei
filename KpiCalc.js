// ============================================
// KpiCalc.js — 🧮数値計算シート修復 + KPI API
// ============================================

var KPI_SHEET_NAME = '🧮数値計算(マイキー）';

// 🧮シートの列→月マッピング (C=3月, D=2月, ... L=6月)
var KPI_MONTH_COLS = [
  { col: 3,  month: 3,  year: 2026 },  // C
  { col: 4,  month: 2,  year: 2026 },  // D
  { col: 5,  month: 1,  year: 2026 },  // E
  { col: 6,  month: 12, year: 2025 },  // F
  { col: 7,  month: 11, year: 2025 },  // G
  { col: 8,  month: 10, year: 2025 },  // H
  { col: 9,  month: 9,  year: 2025 },  // I
  { col: 10, month: 8,  year: 2025 },  // J
  { col: 11, month: 7,  year: 2025 },  // K
  { col: 12, month: 6,  year: 2025 }   // L
];

// 修復対象の行 (1-based)
var KPI_ROW_LP_VIEWS = 5;   // LP閲覧数
var KPI_ROW_SEATED  = 10;  // 着座数（商談数）
var KPI_ROW_DEALS   = 19;  // 成約数
var KPI_ROW_SALES   = 25;  // 売上額（万円）
var KPI_ROW_REVENUE = 30;  // 着金額（万円）

// 旧シート サマリー行番号 (1-based) — 3月以降フォーマット
var OLD_SUM_ROW_REVENUE       = 6;   // 合計着金額
var OLD_SUM_ROW_DEALS         = 7;   // 成約数
var OLD_SUM_ROW_SALES         = 13;  // 合計売上
var OLD_SUM_ROW_TOTAL_DEALS   = 21;  // 合計商談数（当月のみ）
var OLD_SUM_ROW_FUNDED_DEALS  = 25;  // 資金有の合計商談数
var OLD_SUM_ROW_UNFUNDED_DEALS= 26;  // 資金無し商談数
var OLD_SUM_ROW_FUNDED_CLOSED = 27;  // 資金有成約数
var OLD_SUM_ROW_CLOSE_RATE    = 14;  // 成約率
var OLD_SUM_ROW_AVG_PRICE     = 12;  // 平均単価
var OLD_SUM_COL_TOTAL         = 33;  // AG列 = 合計列 (1-based)

/**
 * 🧮数値計算シートの#REF!エラーを修復
 * - 月次アーカイブから過去月データを取得
 * - 旧シートから当月データを直接取得
 * - 着座数、成約数、売上額、着金額の4行を値で上書き
 */
function repairKpiCalcSheet() {
  var ss = getSpreadsheet_();
  var kpiSheet = ss.getSheetByName(KPI_SHEET_NAME);
  if (!kpiSheet) return { error: 'KPIシートが見つかりません: ' + KPI_SHEET_NAME };

  var settings = getGlobalSettings_(ss);
  var results = { fixed: 0, details: [] };

  // 月次アーカイブから全データ取得
  var archiveData = getArchiveAggregated_(ss);

  // 当月データ（旧シートから直接取得）
  var currentMonthData = getCurrentMonthFromOldSheet_(ss);

  for (var i = 0; i < KPI_MONTH_COLS.length; i++) {
    var mc = KPI_MONTH_COLS[i];
    var key = mc.year + '-' + mc.month;
    var data;

    // 当月なら旧シートデータ、過去月ならアーカイブ
    var isCurrent = (mc.year === settings.year && mc.month === settings.month);
    if (isCurrent && currentMonthData) {
      data = currentMonthData;
    } else {
      data = archiveData[key] || null;
    }

    if (!data) {
      // データがない月も0で埋めてREFエラーを解消
      data = { seated: 0, deals: 0, sales: 0, revenue: 0 };
      results.details.push(key + ': データなし（0で補完）');
    }

    // 値をセット
    kpiSheet.getRange(KPI_ROW_SEATED, mc.col).setValue(data.seated);
    kpiSheet.getRange(KPI_ROW_DEALS, mc.col).setValue(data.deals);
    kpiSheet.getRange(KPI_ROW_SALES, mc.col).setValue(round1_(data.sales / 10000));
    kpiSheet.getRange(KPI_ROW_REVENUE, mc.col).setValue(round1_(data.revenue / 10000));

    results.fixed += 4;
    results.details.push(key + ': 着座=' + data.seated + ' 成約=' + data.deals +
      ' 売上=' + round1_(data.sales / 10000) + '万 着金=' + round1_(data.revenue / 10000) + '万');
  }

  return results;
}

/**
 * 月次アーカイブから月ごとの集計データを取得
 */
function getArchiveAggregated_(ss) {
  var archiveSheet = getArchiveSheet_(ss);
  if (!archiveSheet || archiveSheet.getLastRow() <= 1) return {};

  var allData = archiveSheet.getDataRange().getValues();
  var aggregated = {};

  for (var i = 1; i < allData.length; i++) {
    var y = parseNum_(allData[i][ARC_COL_YEAR - 1]);
    var m = parseNum_(allData[i][ARC_COL_MONTH - 1]);
    if (y <= 0 || m <= 0) continue;

    var key = y + '-' + m;
    if (!aggregated[key]) {
      aggregated[key] = { seated: 0, deals: 0, sales: 0, revenue: 0 };
    }

    // total_deals = 着座数（商談数）
    aggregated[key].seated += parseNum_(allData[i][ARC_COL_TOTAL_DEALS - 1]);
    // funded_closed = 成約数（失注/CO除く）
    aggregated[key].deals += parseNum_(allData[i][ARC_COL_FUNDED_CLOSED - 1]);
    // sales = 売上（円）- アーカイブは万円で格納
    aggregated[key].sales += parseNum_(allData[i][ARC_COL_SALES - 1]) * 10000;
    // revenue = 着金（円）- アーカイブは万円で格納
    aggregated[key].revenue += parseNum_(allData[i][ARC_COL_REVENUE - 1]) * 10000;
  }

  return aggregated;
}

/**
 * 当月データを旧シートのサマリー（AG列）から直接取得
 * v2日別シートは同期ズレがあるため、旧シートを正とする
 */
function getCurrentMonthFromOldSheet_(ss) {
  var settings = getGlobalSettings_(ss);
  var sheet = getSheetByMonth_(ss, settings.month);
  if (!sheet) return null;

  var lastRow = Math.max(sheet.getLastRow(), 30);
  var lastCol = Math.max(sheet.getLastColumn(), OLD_SUM_COL_TOTAL);
  var data = sheet.getRange(1, 1, Math.min(lastRow, 30), lastCol).getValues();

  var agIdx = OLD_SUM_COL_TOTAL - 1; // 0-based

  // 合計商談数（Row 21）— #VALUE!の場合は資金有+資金なしで代替
  var seated = parseNum_(data[OLD_SUM_ROW_TOTAL_DEALS - 1][agIdx]);
  if (seated === 0 || isNaN(seated)) {
    seated = parseNum_(data[OLD_SUM_ROW_FUNDED_DEALS - 1][agIdx]) +
             parseNum_(data[OLD_SUM_ROW_UNFUNDED_DEALS - 1][agIdx]);
  }

  // 成約数（Row 7）
  var deals = parseNum_(data[OLD_SUM_ROW_DEALS - 1][agIdx]);

  // 合計売上（Row 13, 万円）
  var salesMan = parseNum_(data[OLD_SUM_ROW_SALES - 1][agIdx]);

  // 合計着金額（Row 6, 万円）
  var revenueMan = parseNum_(data[OLD_SUM_ROW_REVENUE - 1][agIdx]);

  return {
    seated: seated,
    deals: deals,
    sales: salesMan * 10000,    // 万→円（repairKpiCalcSheetが /10000 するため）
    revenue: revenueMan * 10000 // 万→円
  };
}

// ============================================
// KPI API エンドポイント
// ============================================

/**
 * KPIダッシュボード用のファネルデータを返す
 * 全月の LINE追加→予約→着座→成約→売上→着金 データ
 */
function getKpiDashboardData() {
  var ss = getSpreadsheet_();
  var kpiSheet = ss.getSheetByName(KPI_SHEET_NAME);
  var settings = getGlobalSettings_(ss);

  // Chatworkから上流KPIデータ同期（LP閲覧・予約・リスト）
  var cwSync = syncKpiFromChatwork();

  // 下流KPIデータ修復（着座・成約・売上・着金）
  repairKpiCalcSheet();

  // KPIシートから全データ読み取り
  var lastRow = kpiSheet.getLastRow();
  var lastCol = kpiSheet.getLastColumn();
  var data = kpiSheet.getRange(1, 1, Math.max(lastRow, 33), Math.max(lastCol, 13)).getValues();

  var months = [];

  for (var i = 0; i < KPI_MONTH_COLS.length; i++) {
    var mc = KPI_MONTH_COLS[i];
    var c = mc.col - 1; // 0-based index

    // LINE追加数 (row 3) - 「件」を除去してパース
    var lineAdds = parseNum_(data[2][c]);
    // LP閲覧数 (row 5)
    var lpViews = parseNum_(data[KPI_ROW_LP_VIEWS - 1][c]);
    // カレンダー予約数 (row 6) = LP予約数
    var bookings = parseNum_(data[5][c]);
    // 着座数 (row 10)
    var seated = parseNum_(data[9][c]);
    // 資金あり商談数 (row 16)
    var fundedDeals = parseNum_(data[15][c]);
    // 成約数 (row 19)
    var deals = parseNum_(data[18][c]);
    // 売上額万 (row 25)
    var salesMan = parseNum_(data[24][c]);
    // 着金額万 (row 30)
    var revenueMan = parseNum_(data[29][c]);

    // 転換率
    var bookingRate = lineAdds > 0 ? round1_(bookings / lineAdds * 100) : 0;
    var seatRate = bookings > 0 ? round1_(seated / bookings * 100) : 0;
    var closeRate = seated > 0 ? round1_(deals / seated * 100) : 0;
    var fundedCloseRate = fundedDeals > 0 ? round1_(deals / fundedDeals * 100) : 0;
    var avgDeal = deals > 0 ? round1_(salesMan / deals) : 0;
    var collectionRate = salesMan > 0 ? round1_(revenueMan / salesMan * 100) : 0;

    months.push({
      year: mc.year,
      month: mc.month,
      lineAdds: lineAdds,
      lpViews: lpViews,
      bookings: bookings,
      seated: seated,
      fundedDeals: fundedDeals,
      deals: deals,
      sales: salesMan,
      revenue: revenueMan,
      bookingRate: bookingRate,
      seatRate: seatRate,
      closeRate: closeRate,
      fundedCloseRate: fundedCloseRate,
      avgDeal: avgDeal,
      collectionRate: collectionRate,
      isCurrent: (mc.year === settings.year && mc.month === settings.month)
    });
  }

  // 当月データ詳細（旧シートから）
  var currentDetail = getCurrentMonthDetailFromOldSheet_(ss);

  return {
    months: months,
    currentMonth: settings.month,
    currentYear: settings.year,
    teamSize: getActiveMembers_(ss).length,
    currentDetail: currentDetail,
    chatworkKpi: cwSync && cwSync.parsed ? cwSync.parsed : null
  };
}

/**
 * 当月の詳細データ（メンバー別）— 旧シートから直接読み取り
 */
function getCurrentMonthDetailFromOldSheet_(ss) {
  var settings = getGlobalSettings_(ss);
  var sheet = getSheetByMonth_(ss, settings.month);
  if (!sheet) return [];

  var lastRow = Math.max(sheet.getLastRow(), 30);
  var lastCol = Math.max(sheet.getLastColumn(), OLD_SUM_COL_TOTAL);
  var data = sheet.getRange(1, 1, Math.min(lastRow, 30), lastCol).getValues();

  var members = [];

  for (var i = 0; i < MEMBER_SECTIONS.length; i++) {
    var sec = MEMBER_SECTIONS[i];
    var c = sec.summaryCol - 1; // 0-based
    if (c >= data[0].length) continue;

    var displayName = DISPLAY_NAME_MAP[sec.name] || sec.name;

    var totalDeals = parseNum_(data[OLD_SUM_ROW_TOTAL_DEALS - 1][c]);
    if (totalDeals === 0 || isNaN(totalDeals)) {
      totalDeals = parseNum_(data[OLD_SUM_ROW_FUNDED_DEALS - 1][c]) +
                   parseNum_(data[OLD_SUM_ROW_UNFUNDED_DEALS - 1][c]);
    }

    var closed = parseNum_(data[OLD_SUM_ROW_DEALS - 1][c]);
    var salesVal = parseNum_(data[OLD_SUM_ROW_SALES - 1][c]);
    var revenueVal = parseNum_(data[OLD_SUM_ROW_REVENUE - 1][c]);
    var closeRateVal = parseNum_(data[OLD_SUM_ROW_CLOSE_RATE - 1][c]);
    var avgPriceVal = parseNum_(data[OLD_SUM_ROW_AVG_PRICE - 1][c]);

    // 成約率は小数 (0.8) → パーセント (80)
    if (closeRateVal > 0 && closeRateVal <= 1) {
      closeRateVal = round1_(closeRateVal * 100);
    }

    members.push({
      name: displayName,
      deals: totalDeals,
      closed: closed,
      sales: round1_(salesVal),
      revenue: round1_(revenueVal),
      closeRate: round1_(closeRateVal),
      avgPrice: round1_(avgPriceVal)
    });
  }

  return members;
}
