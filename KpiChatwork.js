// ============================================
// KpiChatwork.js — ChatworkからKPIデータ自動取得
// ============================================

var KPI_REPORT_ROOM_ID = '418585408';

/**
 * ChatworkのKPIレポートから上流指標を取得して🧮シートに書込み
 * - LP流入数（UU）→ LP閲覧数 (Row 5)
 * - カレンダー申込 → カレンダー予約数 (Row 6)
 * - リスト数 → LINE追加数 (Row 3)
 * - 申込率 → 転換率データとして返却
 */
function syncKpiFromChatwork() {
  // 5分キャッシュ（API負荷軽減）
  var cache = CacheService.getScriptCache();
  var cached = cache.get('kpi_cw_sync');
  if (cached) return JSON.parse(cached);

  var token = getChatworkToken_();
  if (!token) return { error: 'token未設定' };

  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);

  // メッセージ取得（force=1で直近100件）
  var messages = fetchChatworkKpiMessages_(token);
  if (!messages || messages.length === 0) return { error: 'メッセージなし' };

  // 最新のKPIレポートを検索（新→古）
  var report = null;
  for (var i = messages.length - 1; i >= 0; i--) {
    var body = messages[i].body || '';
    if (body.indexOf('LP流入数') !== -1 && body.indexOf('月合計') !== -1) {
      report = messages[i];
      break;
    }
  }
  if (!report) return { error: 'レポート未検出' };

  // パース
  var parsed = parseKpiChatworkReport_(report.body);

  // 当月列を特定
  var kpiSheet = ss.getSheetByName(KPI_SHEET_NAME);
  if (!kpiSheet) return { error: 'KPIシートなし' };

  var currentCol = null;
  for (var i = 0; i < KPI_MONTH_COLS.length; i++) {
    if (KPI_MONTH_COLS[i].year === settings.year && KPI_MONTH_COLS[i].month === settings.month) {
      currentCol = KPI_MONTH_COLS[i].col;
      break;
    }
  }
  if (!currentCol) return { error: '当月列なし' };

  // 🧮シート書込み
  var updated = [];
  if (parsed.lpViews > 0) {
    kpiSheet.getRange(KPI_ROW_LP_VIEWS, currentCol).setValue(parsed.lpViews);
    updated.push('LP閲覧=' + parsed.lpViews);
  }
  if (parsed.bookings > 0) {
    kpiSheet.getRange(6, currentCol).setValue(parsed.bookings);
    updated.push('予約=' + parsed.bookings);
  }
  if (parsed.listCount > 0) {
    kpiSheet.getRange(3, currentCol).setValue(parsed.listCount);
    updated.push('リスト=' + parsed.listCount);
  }

  var result = {
    success: true,
    updated: updated,
    parsed: parsed
  };

  // 5分キャッシュ
  cache.put('kpi_cw_sync', JSON.stringify(result), 300);

  return result;
}

/**
 * Chatwork KPIルームからメッセージ取得
 */
function fetchChatworkKpiMessages_(token) {
  var url = 'https://api.chatwork.com/v2/rooms/' + KPI_REPORT_ROOM_ID + '/messages?force=1';
  try {
    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'X-ChatWorkToken': token },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) return [];
    return JSON.parse(response.getContentText());
  } catch (e) {
    Logger.log('KPI Chatwork取得エラー: ' + e.message);
    return [];
  }
}

/**
 * KPIレポートメッセージをパース
 *
 * 期待フォーマット:
 * 🔴KGI（カレンダー申込）：月2000　65申込/1日
 * 月合計：461（前月：1012）
 * 3/1：67（不足：0）（前月：38）
 * ...
 * ⭕️KPI（申込率）
 * 月平均：10.7％（前月：7.2％）
 * ...
 * 🚨KKPI（LP流入数（UU））
 * 月合計：4367（前月：13950）
 * ...
 * ◎KKKPI：リスト数
 * 月合計：3411（前月：11875）
 */
function parseKpiChatworkReport_(body) {
  var result = {
    bookings: 0,
    bookingsPrev: 0,
    bookingTarget: 0,
    bookingRate: 0,
    bookingRatePrev: 0,
    lpViews: 0,
    lpViewsPrev: 0,
    listCount: 0,
    listCountPrev: 0,
    dailyBookings: {},
    dailyLpViews: {},
    dailyList: {},
    dailyBookingRate: {}
  };

  // --- カレンダー申込 ---
  var bookSec = body.match(/カレンダー申込[）\)]?[：:]月(\d[\d,]*)/);
  if (bookSec) result.bookingTarget = parseInt(bookSec[1].replace(/,/g, ''), 10);

  var bookTotal = body.match(/カレンダー申込[\s\S]*?月合計[：:](\d[\d,]*)(?:（前月[：:](\d[\d,]*)）)?/);
  if (bookTotal) {
    result.bookings = parseInt(bookTotal[1].replace(/,/g, ''), 10);
    if (bookTotal[2]) result.bookingsPrev = parseInt(bookTotal[2].replace(/,/g, ''), 10);
  }

  // カレンダー申込 日次
  var bookDailySection = body.match(/カレンダー申込[\s\S]*?(?=⇧|⭕)/);
  if (bookDailySection) {
    result.dailyBookings = parseDailyValues_(bookDailySection[0]);
  }

  // --- 申込率 ---
  var rateMatch = body.match(/申込率[\s\S]*?月平均[：:](\d+\.?\d*)％(?:（前月[：:](\d+\.?\d*)％）)?/);
  if (rateMatch) {
    result.bookingRate = parseFloat(rateMatch[1]);
    if (rateMatch[2]) result.bookingRatePrev = parseFloat(rateMatch[2]);
  }

  // 申込率 日次
  var rateDailySection = body.match(/申込率[\s\S]*?(?=⇧|🚨)/);
  if (rateDailySection) {
    result.dailyBookingRate = parseDailyRates_(rateDailySection[0]);
  }

  // --- LP流入数（UU） ---
  var lpTotal = body.match(/LP流入数[\s\S]*?月合計[：:](\d[\d,]*)(?:（前月[：:](\d[\d,]*)）)?/);
  if (lpTotal) {
    result.lpViews = parseInt(lpTotal[1].replace(/,/g, ''), 10);
    if (lpTotal[2]) result.lpViewsPrev = parseInt(lpTotal[2].replace(/,/g, ''), 10);
  }

  // LP流入 日次
  var lpDailySection = body.match(/LP流入数[\s\S]*?(?=⇧|◎)/);
  if (lpDailySection) {
    result.dailyLpViews = parseDailyValues_(lpDailySection[0]);
  }

  // --- リスト数 ---
  var listTotal = body.match(/リスト数[\s\S]*?月合計[：:](\d[\d,]*)(?:（前月[：:](\d[\d,]*)）)?/);
  if (listTotal) {
    result.listCount = parseInt(listTotal[1].replace(/,/g, ''), 10);
    if (listTotal[2]) result.listCountPrev = parseInt(listTotal[2].replace(/,/g, ''), 10);
  }

  // リスト 日次
  var listDailySection = body.match(/リスト数[\s\S]*$/);
  if (listDailySection) {
    result.dailyList = parseDailyValues_(listDailySection[0]);
  }

  return result;
}

/**
 * 日次データ（整数）をパース
 * "3/1：618（前月：427）" → { 1: 618, 2: 672, ... }
 */
function parseDailyValues_(section) {
  var daily = {};
  var pattern = /(\d+)\/(\d+)[：:](\d[\d,]*)/g;
  var m;
  while ((m = pattern.exec(section)) !== null) {
    var day = parseInt(m[2], 10);
    var val = parseInt(m[3].replace(/,/g, ''), 10);
    // 「月合計」行のday値は除外（月=大きい数値は通常日データ）
    if (!isNaN(day) && day <= 31) {
      daily[day] = val;
    }
  }
  return daily;
}

/**
 * 日次データ（パーセント）をパース
 * "3/1：10.8％（前月：8.9％）" → { 1: 10.8, 2: 8.8, ... }
 */
function parseDailyRates_(section) {
  var daily = {};
  var pattern = /(\d+)\/(\d+)[：:](\d+\.?\d*)％/g;
  var m;
  while ((m = pattern.exec(section)) !== null) {
    var day = parseInt(m[2], 10);
    var val = parseFloat(m[3]);
    if (!isNaN(day) && day <= 31) {
      daily[day] = val;
    }
  }
  return daily;
}
