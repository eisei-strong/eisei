// ============================================
// Summary.js — 自動集計ロジック (v2)
// トリガー: 5分おき or 編集時
// ============================================

/**
 * サマリーシートを全日別入力シートから自動更新
 */
function updateSummary() {
  var ss = getSpreadsheet_();
  var summarySheet = getSummarySheet_(ss);
  if (!summarySheet) {
    Logger.log('サマリーシートが見つかりません');
    return;
  }

  var members = getActiveMembers_(ss);
  if (members.length === 0) {
    Logger.log('アクティブメンバーが見つかりません');
    return;
  }

  var settings = getGlobalSettings_(ss);

  // 旧ウォーリアーズ数値シートからデータを同期（列構造が異なるため動的検索）
  try {
    syncFromOldSheet();
  } catch (e) {
    Logger.log('旧シート同期スキップ: ' + e.message);
  }

  // クレカ/信販を⚔️信販会社割合シートから自動更新
  updateCreditShinpan_(ss, settings);

  // CBS/ライフティを⚔️信販会社割合シートから直接取得
  updateCBSLifetyFromShinpan_(ss, settings);

  // SUM数式を再計算してから読み取る
  SpreadsheetApp.flush();

  // 各メンバーのデータを収集
  var memberDataList = [];
  for (var i = 0; i < members.length; i++) {
    var m = members[i];
    var dailySheet = getDailySheet_(ss, m.name);
    var data = readDailySheetData_(dailySheet);

    memberDataList.push({
      name: m.name,
      displayName: m.displayName,
      iconUrl: m.iconUrl,
      data: data
    });
  }

  // 前月データ取得
  var prevMonthData = getPrevMonthFromArchive_(ss, settings);

  // ランキング計算（着金額降順）
  memberDataList.sort(function(a, b) {
    var revA = a.data ? a.data.totals.revenue : 0;
    var revB = b.data ? b.data.totals.revenue : 0;
    return revB - revA;
  });

  // ランクを計算
  for (var r = 0; r < memberDataList.length; r++) {
    var rev = memberDataList[r].data ? memberDataList[r].data.totals.revenue : 0;
    if (r === 0) {
      memberDataList[r].rank = 1;
    } else if (rev === (memberDataList[r - 1].data ? memberDataList[r - 1].data.totals.revenue : 0)) {
      memberDataList[r].rank = memberDataList[r - 1].rank;
    } else {
      memberDataList[r].rank = r + 1;
    }
  }

  // チームKPI計算
  var totalRevenue = 0, totalDeals = 0, totalClosed = 0, totalSales = 0, totalCO = 0;
  for (var t = 0; t < memberDataList.length; t++) {
    if (!memberDataList[t].data) continue;
    var d = memberDataList[t].data.totals;
    totalRevenue += d.revenue;
    totalDeals += d.totalDeals;
    totalClosed += d.fundedClosed;
    totalSales += d.sales;
    totalCO += d.coCount;
  }
  totalRevenue = round1_(totalRevenue);
  totalSales = round1_(totalSales);

  var teamGoal = settings.teamGoal;
  var progressRate = teamGoal > 0 ? round1_((totalRevenue / teamGoal) * 100) : 0;

  var now = new Date();

  // === サマリーシート書き込み ===

  // 最終更新
  summarySheet.getRange(SM_ROW_UPDATED, 2).setValue(
    Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')
  );

  // セクション1: チームKPI
  summarySheet.getRange(SM_ROW_KPI_REVENUE, 2).setValue(totalRevenue);
  summarySheet.getRange(SM_ROW_KPI_DEALS, 2).setValue(totalDeals);
  summarySheet.getRange(SM_ROW_KPI_CLOSED, 2).setValue(totalClosed);
  summarySheet.getRange(SM_ROW_KPI_SALES, 2).setValue(totalSales);
  summarySheet.getRange(SM_ROW_KPI_CO, 2).setValue(totalCO);
  summarySheet.getRange(SM_ROW_KPI_PROGRESS, 2).setValue(progressRate + '%');

  // セクション2: メンバー別成績（ランキング＋前月比）
  var topRevenue = memberDataList[0] && memberDataList[0].data
    ? memberDataList[0].data.totals.revenue : 0;

  var memberRows = [];
  for (var m = 0; m < memberDataList.length; m++) {
    var md = memberDataList[m];
    var dt = md.data ? md.data.totals : emptyTotals_();
    var lr = md.data ? md.data.lifetyRate : '-';
    var prev = prevMonthData[md.name] || prevMonthData[md.displayName] || { revenue: 0, deals: 0, closed: 0, closeRate: 0 };
    var currentRev = md.data ? md.data.totals.revenue : 0;
    var gapToTop = (md.rank === 1) ? 0 : round1_(topRevenue - currentRev);

    memberRows.push([
      md.displayName,
      md.rank,
      dt.fundedDeals,
      dt.unfundedDeals,
      dt.totalDeals,
      dt.fundedClosed,
      dt.closeRate > 0 ? dt.closeRate : '-',
      dt.sales,
      dt.revenue,
      dt.coCount,
      lr,
      prev.revenue,
      round1_(currentRev - prev.revenue),
      gapToTop,
      dt.totalDeals - prev.deals,
      dt.fundedClosed - prev.closed,
      round1_(dt.closeRate - prev.closeRate)
    ]);
  }

  // 空行でパディング（メンバー行数分）
  var memberSlots = SM_ROW_MEMBER_END - SM_ROW_MEMBER_START + 1;
  while (memberRows.length < memberSlots) {
    memberRows.push(new Array(SM_MEMBER_COL_COUNT).fill(''));
  }

  summarySheet.getRange(SM_ROW_MEMBER_START, 1, memberSlots, SM_MEMBER_COL_COUNT).setValues(memberRows);

  // 前月チーム合計（退職メンバー含む全員分）
  var prevTotalDeals = 0, prevTotalRevenue = 0, prevTotalClosed = 0;
  for (var pk in prevMonthData) {
    prevTotalDeals += prevMonthData[pk].deals || 0;
    prevTotalRevenue += round1_(prevMonthData[pk].revenue || 0);
    prevTotalClosed += prevMonthData[pk].closed || 0;
  }
  var prevTotalCloseRate = prevTotalDeals > 0 ? round1_((prevTotalClosed / prevTotalDeals) * 100) : 0;

  // 合計行
  var totalRow = [
    '合計', '', totalDeals > 0 ? '' : '', '', totalDeals, totalClosed,
    totalDeals > 0 ? round1_((totalClosed / totalDeals) * 100) : '-',
    totalSales, totalRevenue, totalCO, '',
    round1_(prevTotalRevenue),
    round1_(totalRevenue - prevTotalRevenue),
    '',
    totalDeals - prevTotalDeals,
    totalClosed - prevTotalClosed,
    round1_((totalDeals > 0 ? round1_((totalClosed / totalDeals) * 100) : 0) - prevTotalCloseRate)
  ];
  // 合計の資金有商談・資金なし商談
  var totalFunded = 0, totalUnfunded = 0;
  for (var s = 0; s < memberDataList.length; s++) {
    if (!memberDataList[s].data) continue;
    totalFunded += memberDataList[s].data.totals.fundedDeals;
    totalUnfunded += memberDataList[s].data.totals.unfundedDeals;
  }
  totalRow[2] = totalFunded;
  totalRow[3] = totalUnfunded;
  summarySheet.getRange(SM_ROW_MEMBER_TOTAL, 1, 1, SM_MEMBER_COL_COUNT).setValues([totalRow]);

  // セクション3: 期間比較
  var periods = calculatePeriods_(memberDataList);
  summarySheet.getRange(SM_ROW_PERIOD_1_10, 2).setValue(periods.p1.deals);
  summarySheet.getRange(SM_ROW_PERIOD_11_20, 2).setValue(periods.p2.deals);
  summarySheet.getRange(SM_ROW_PERIOD_21_END, 2).setValue(periods.p3.deals);

  // セクション4: 着金内訳
  var totalCredit = 0, totalShinpan = 0, totalCarryover = 0;
  for (var b = 0; b < memberDataList.length; b++) {
    if (!memberDataList[b].data) continue;
    totalCredit += memberDataList[b].data.totals.creditCard;
    totalShinpan += memberDataList[b].data.totals.shinpan;
    totalCarryover += memberDataList[b].data.carryover.revenue;
  }
  summarySheet.getRange(SM_ROW_CREDIT_TOTAL, 2).setValue(round1_(totalCredit));
  summarySheet.getRange(SM_ROW_SHINPAN_TOTAL, 2).setValue(round1_(totalShinpan));
  summarySheet.getRange(SM_ROW_CURRENT_REVENUE, 2).setValue(round1_(totalRevenue - totalCarryover));
  summarySheet.getRange(SM_ROW_CARRYOVER_REVENUE, 2).setValue(round1_(totalCarryover));

  // セクション5: 退職者CO残
  var coSheet = getCOManageSheet_(ss);
  if (coSheet) {
    var coTotal = coSheet.getRange(CO_ROW_SUMMARY_TOTAL, 4).getValue();
    summarySheet.getRange(SM_ROW_CO_START, 1, 1, 2).setValues([
      ['退職者 未回収CO合計', parseNum_(coTotal)]
    ]);
  }

  // ランキングを設定シートに書き戻し
  updateRankInSettings_(ss, memberDataList);

  // ダッシュボードキャッシュをクリア（次回API呼び出しで最新データを返す）
  clearDashboardCache_();

  Logger.log('サマリー更新完了: 着金合計=' + totalRevenue + '万円');
}

/**
 * ランキングを設定シートに書き戻し
 */
function updateRankInSettings_(ss, memberDataList) {
  var settingsSheet = getSettingsSheet_(ss);
  if (!settingsSheet) return;

  var allMembers = getMembersFromSettings_(ss);
  for (var i = 0; i < allMembers.length; i++) {
    for (var j = 0; j < memberDataList.length; j++) {
      if (allMembers[i].name === memberDataList[j].name ||
          allMembers[i].displayName === memberDataList[j].displayName) {
        settingsSheet.getRange(SETTINGS_ROW_DATA_START + i, SETTINGS_COL_RANK).setValue(memberDataList[j].rank);
        break;
      }
    }
  }
}

/**
 * 期間別集計を計算
 */
function calculatePeriods_(memberDataList) {
  var p1 = { revenue: 0, deals: 0, closed: 0 };
  var p2 = { revenue: 0, deals: 0, closed: 0 };
  var p3 = { revenue: 0, deals: 0, closed: 0 };

  for (var m = 0; m < memberDataList.length; m++) {
    if (!memberDataList[m].data) continue;
    var daily = memberDataList[m].data.daily;

    for (var d = 0; d < daily.length; d++) {
      var day = daily[d].day;
      var target;
      if (day <= 10) target = p1;
      else if (day <= 20) target = p2;
      else target = p3;

      target.revenue += daily[d].revenue;
    }
  }

  p1.revenue = round1_(p1.revenue);
  p2.revenue = round1_(p2.revenue);
  p3.revenue = round1_(p3.revenue);

  return { p1: p1, p2: p2, p3: p3 };
}

/**
 * 月次アーカイブから前月データを取得
 */
function getPrevMonthFromArchive_(ss, settings) {
  var prevMonth = settings.month === 1 ? 12 : settings.month - 1;
  var prevYear = settings.month === 1 ? settings.year - 1 : settings.year;

  var archiveSheet = getArchiveSheet_(ss);
  if (archiveSheet && archiveSheet.getLastRow() > 1) {
    var data = archiveSheet.getDataRange().getValues();
    var result = {};

    for (var i = 1; i < data.length; i++) {
      if (parseNum_(data[i][ARC_COL_YEAR - 1]) === prevYear &&
          parseNum_(data[i][ARC_COL_MONTH - 1]) === prevMonth) {
        var name = String(data[i][ARC_COL_NAME - 1] || '').trim();
        if (name) {
          result[name] = {
            revenue: round1_(parseNum_(data[i][ARC_COL_REVENUE - 1])),
            deals: parseNum_(data[i][ARC_COL_TOTAL_DEALS - 1]),
            closed: parseNum_(data[i][ARC_COL_FUNDED_CLOSED - 1]),
            closeRate: round1_(parseNum_(data[i][ARC_COL_CLOSE_RATE - 1]))
          };
        }
      }
    }

    if (Object.keys(result).length > 0) {
      return result;
    }
  }

  // フォールバック: 旧シートから読み取り
  return getPrevMonthData_legacy_(ss, prevMonth);
}

/**
 * 旧シートから前月データ取得（動的ラベル検索方式）
 * 2月/3月等の異なるレイアウトに対応
 */
function getPrevMonthData_legacy_(ss, prevMonth) {
  var sheet = getSheetByMonth_(ss, prevMonth);
  if (!sheet) return {};

  // #REF!エラーを事前修正（先頭50行×50列に限定して高速化）
  var scanRows = Math.min(50, sheet.getLastRow());
  var scanCols = Math.min(50, sheet.getLastColumn());
  if (scanRows < 5) return {};

  var displayVals = sheet.getRange(1, 1, scanRows, scanCols).getDisplayValues();
  var refFixed = 0;
  for (var ri = 0; ri < displayVals.length; ri++) {
    for (var ci = 0; ci < displayVals[ri].length; ci++) {
      var dv = displayVals[ri][ci];
      if (dv === '#REF!' || dv === '#NAME?' || dv === '#ERROR!' || dv === '#VALUE!') {
        sheet.getRange(ri + 1, ci + 1).setValue(0);
        refFixed++;
      }
    }
  }
  if (refFixed > 0) SpreadsheetApp.flush();

  var allData = sheet.getRange(1, 1, scanRows, scanCols).getValues();
  var result = {};

  // === Step 1: ラベルで行を動的検索 ===
  var rowMap = {};

  for (var r = 0; r < allData.length; r++) {
    var label = String(allData[r][0] || '').replace(/\s+/g, '').replace(/[\(\)（）]/g, '');
    if (!label) {
      label = String(allData[r][1] || '').replace(/\s+/g, '').replace(/[\(\)（）]/g, '');
    }
    if (!label) continue;

    // 着金額（最初にマッチしたものを使用 = セクション1のデータ）
    if (!rowMap.revenue && label.indexOf('合計着金額') !== -1) rowMap.revenue = r;
    // 商談数（「資金有り合計商談数」を除外して「合計商談数（当月のみ）」等にマッチ）
    if (!rowMap.deals && label.indexOf('合計商談数') !== -1 && label.indexOf('資金') === -1) rowMap.deals = r;
    // 成約数（「合計成約数」優先、なければ「成約数」をフォールバック）
    if (label.indexOf('合計成約数') !== -1) rowMap.closed = r;
    if (!rowMap.closed && label === '成約数') rowMap.closed = r;
    // 成約率
    if (!rowMap.closeRate && label.indexOf('成約率') !== -1) rowMap.closeRate = r;
    // 売上
    if (!rowMap.sales && (label.indexOf('合計売上') !== -1 || label === '売上')) rowMap.sales = r;
    // 資金有商談
    if (!rowMap.fundedDeals && label.indexOf('資金有') !== -1 && label.indexOf('商談') !== -1) rowMap.fundedDeals = r;
  }

  // === Step 2: 名前行を検索（【】3つ以上 or 既知メンバー名3つ以上） ===
  var nameRowIdx = -1;
  for (var r = 0; r < Math.min(15, allData.length); r++) {
    var bracketCount = 0;
    var knownCount = 0;
    for (var c = 0; c < allData[r].length; c++) {
      var v = String(allData[r][c] || '').trim();
      if (v.indexOf('【') !== -1 && v.indexOf('】') !== -1) bracketCount++;
      var clean = v.replace(/[【】\[\]]/g, '');
      if (LEGACY_TO_V2_NAME[clean] || ICON_MAP[clean] || OLD_DISPLAY_TO_V2[clean]) knownCount++;
    }
    if (bracketCount >= 3 || knownCount >= 3) {
      nameRowIdx = r;
      break;
    }
  }

  if (nameRowIdx < 0) {
    Logger.log('前月シート: 名前行が見つかりません');
    return {};
  }

  // === Step 3: メンバー列を名前から検出＆データ取得 ===
  for (var c = 0; c < allData[nameRowIdx].length; c++) {
    var rawName = String(allData[nameRowIdx][c] || '').trim();
    if (!rawName) continue;

    var cleanName = rawName.replace(/[【】\[\]]/g, '').trim();
    if (!cleanName || cleanName === '合計') continue;

    // 数値やCellImage等のゴミをスキップ
    if (!isNaN(Number(cleanName)) || cleanName === 'CellImage') continue;

    // NAME_MAP で正規化 → v2名に変換
    var normalized = NAME_MAP[cleanName] || cleanName;
    var v2Name = normalized;
    if (LEGACY_TO_V2_NAME[normalized]) v2Name = LEGACY_TO_V2_NAME[normalized];
    else if (DISPLAY_NAME_MAP[normalized]) v2Name = DISPLAY_NAME_MAP[normalized];
    else if (ICON_MAP[normalized]) v2Name = normalized;
    else if (OLD_DISPLAY_TO_V2[normalized]) v2Name = OLD_DISPLAY_TO_V2[normalized];
    else {
      for (var iconKey in ICON_MAP) {
        if (iconKey.indexOf(normalized) === 0 || normalized.indexOf(iconKey) === 0) {
          v2Name = iconKey;
          break;
        }
      }
    }

    var revenue = rowMap.revenue !== undefined ? round1_(parseNum_(allData[rowMap.revenue][c])) : 0;
    var deals = rowMap.deals !== undefined ? parseNum_(allData[rowMap.deals][c])
      : (rowMap.fundedDeals !== undefined ? parseNum_(allData[rowMap.fundedDeals][c]) : 0);
    var closed = rowMap.closed !== undefined ? parseNum_(allData[rowMap.closed][c]) : 0;
    var sales = rowMap.sales !== undefined ? round1_(parseNum_(allData[rowMap.sales][c])) : 0;

    var closeRate = 0;
    if (rowMap.closeRate !== undefined) {
      var rateRaw = allData[rowMap.closeRate][c];
      closeRate = typeof rateRaw === 'number'
        ? round1_(rateRaw * 100)
        : (deals > 0 ? round1_((closed / deals) * 100) : 0);
    } else if (deals > 0) {
      closeRate = round1_((closed / deals) * 100);
    }

    // データがあるメンバーのみ追加
    if (revenue > 0 || deals > 0 || closed > 0 || sales > 0) {
      result[v2Name] = { revenue: revenue, deals: deals, closed: closed, closeRate: closeRate };
    }
  }

  Logger.log('前月データ(' + prevMonth + '月): ' + Object.keys(result).length + '名, rowMap=' + JSON.stringify(rowMap));
  return result;
}

/**
 * 旧シートの特定行をデバッグ
 * API: ?action=run&fn=debugOldSheetRows
 */
function debugOldSheetRows() {
  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);
  var oldSheet = getSheetByMonth_(ss, settings.month);
  if (!oldSheet) return { error: 'シートなし' };

  var allData = oldSheet.getDataRange().getValues();
  var allFormulas = oldSheet.getDataRange().getFormulas();

  // rowMapを構築
  var rowMap = {};
  var labelPatterns = {
    'ranking': 'ランキング', 'revenue': '合計着金額', 'closed': '合計成約数',
    'currentRev': '当月着金額', 'carryoverRev': '前月繰り越し',
    'credit': 'クレジットカード', 'shinpan': '信販会社', 'avgPrice': '平均単価',
    'sales': '合計売上', 'closeRate': '全体の成約率', 'cbs': 'CBS', 'lifety': 'ライフティ',
    'deals': '合計商談数', 'fundedDeals': '資金有りの合計商談数',
    'unfundedDeals': '資金なしの合計商談数', 'fundedClosed': '資金有成約数',
    'unfundedClosed': '資金なし成約数', 'coCount': '合計CO数', 'coAmount': '合計CO金額'
  };
  for (var r = 0; r < Math.min(40, allData.length); r++) {
    var label = String(allData[r][0] || '').replace(/\s+/g, '').replace(/[\(\)（）]/g, '');
    if (!label) label = String(allData[r][1] || '').replace(/\s+/g, '').replace(/[\(\)（）]/g, '');
    for (var key in labelPatterns) {
      if (!rowMap[key] && label && label.indexOf(labelPatterns[key]) !== -1) rowMap[key] = r;
    }
  }

  // メンバー列を取得
  var nameRowIdx = -1;
  for (var r = 0; r < Math.min(10, allData.length); r++) {
    var bc = 0;
    for (var c = 0; c < allData[r].length; c++) {
      if (String(allData[r][c] || '').indexOf('【') !== -1) bc++;
    }
    if (bc >= 3) { nameRowIdx = r; break; }
  }

  // 行12-15のデータをダンプ（売上・成約率付近）
  var rowDump = {};
  var targetRows = [12, 13, 14, 15]; // 1-indexed
  for (var ri = 0; ri < targetRows.length; ri++) {
    var row = targetRows[ri] - 1; // 0-indexed
    if (row >= allData.length) continue;
    var label = String(allData[row][0] || allData[row][1] || '');
    var vals = {};
    for (var c = 0; c < Math.min(35, allData[row].length); c++) {
      var v = allData[row][c];
      var f = allFormulas[row][c];
      if (v !== '' && v !== null && v !== undefined) {
        vals['col' + (c+1)] = { value: v, formula: f || '' };
      }
    }
    rowDump['row' + targetRows[ri] + '_' + label.substring(0, 20)] = vals;
  }

  // 各メンバーの売上・成約率を表示
  var memberData = {};
  if (nameRowIdx >= 0) {
    for (var c = 0; c < allData[nameRowIdx].length; c++) {
      var raw = String(allData[nameRowIdx][c] || '').replace(/[【】]/g, '').trim();
      if (!raw || raw === '合計') continue;
      var normalized = NAME_MAP[raw] || raw;
      var v2 = LEGACY_TO_V2_NAME[normalized] || normalized;
      if (!ICON_MAP[v2]) continue;

      var salesVal = rowMap.sales !== undefined ? allData[rowMap.sales][c] : '?';
      var salesFormula = rowMap.sales !== undefined ? allFormulas[rowMap.sales][c] : '';
      var rateVal = rowMap.closeRate !== undefined ? allData[rowMap.closeRate][c] : '?';
      var rateFormula = rowMap.closeRate !== undefined ? allFormulas[rowMap.closeRate][c] : '';
      var closedVal = rowMap.closed !== undefined ? allData[rowMap.closed][c] : '?';
      var dealsVal = rowMap.deals !== undefined ? allData[rowMap.deals][c] : '?';
      var fundedDealsVal = rowMap.fundedDeals !== undefined ? allData[rowMap.fundedDeals][c] : '?';

      memberData[v2] = {
        col: c + 1,
        sales: salesVal, salesFormula: salesFormula,
        closeRate: rateVal, closeRateFormula: rateFormula,
        closed: closedVal, deals: dealsVal, fundedDeals: fundedDealsVal
      };
    }
  }

  // 下部セクション構造を調査
  var totalRows = oldSheet.getLastRow();
  var totalCols = oldSheet.getLastColumn();

  // 各セクションのヘッダー行を特定（名前が書いてある行）
  var sectionHeaders = [39, 75, 111, 147, 184, 221, 258, 295, 332]; // 0-indexed header rows
  var sectionInfo = [];
  if (totalRows > 40) {
    for (var si = 0; si < sectionHeaders.length; si++) {
      var hRow = sectionHeaders[si]; // 0-indexed
      if (hRow + 1 > totalRows) continue;
      var rowData = oldSheet.getRange(hRow + 1, 1, 1, Math.min(totalCols, 50)).getValues()[0];
      var headerBelow = oldSheet.getRange(hRow + 2, 1, 1, Math.min(totalCols, 50)).getValues()[0];
      var names = [];
      for (var c = 0; c < rowData.length; c++) {
        var v = String(rowData[c] || '').trim();
        if (v) names.push({ col: c + 1, val: v.substring(0, 20) });
      }
      var headers = [];
      for (var c = 0; c < headerBelow.length; c++) {
        var v = String(headerBelow[c] || '').trim();
        if (v) headers.push({ col: c + 1, val: v.substring(0, 20) });
      }
      sectionInfo.push({ sheetRow: hRow + 1, dataStart: hRow + 3, dataEnd: hRow + 32,
        nameRow: names.slice(0, 20), headerRow: headers.slice(0, 20) });
    }
  }

  // row 36 付近（繰り越しデータ）
  var row36Data = {};
  if (totalRows > 36) {
    var r36 = oldSheet.getRange(36, 1, 1, Math.min(totalCols, 35)).getValues()[0];
    for (var c = 0; c < r36.length; c++) {
      if (r36[c] !== '' && r36[c] !== null) row36Data['col' + (c+1)] = r36[c];
    }
  }

  // 各セクションの上の行でメンバー名を探す + セクション内の合計行を確認
  var sectionMapping = [];
  var sectionStartRows = [41, 77, 113, 149, 186, 223, 260, 297, 334]; // 1-indexed data start
  for (var si = 0; si < sectionStartRows.length; si++) {
    var startRow = sectionStartRows[si];
    var endRow = startRow + 30;
    var nameAbove = '';
    // 3行上まで名前を探す
    for (var nr = startRow - 4; nr < startRow - 1; nr++) {
      if (nr < 1) continue;
      var nameRowData = oldSheet.getRange(nr, 1, 1, Math.min(totalCols, 35)).getValues()[0];
      for (var c = 0; c < nameRowData.length; c++) {
        var v = String(nameRowData[c] || '').trim().replace(/[【】\[\]]/g, '');
        if (v && (LEGACY_TO_V2_NAME[v] || ICON_MAP[v] || LEGACY_TO_V2_NAME[NAME_MAP[v] || ''] || DISPLAY_NAME_MAP[v])) {
          nameAbove = v + ' (row' + nr + ' col' + (c+1) + ')';
        }
      }
    }
    // 合計行（データ最終行+2付近）のcol L (12) の値を確認
    var sumRow = endRow + 2; // 合計行推定
    var salesSum = 0;
    if (sumRow <= totalRows) {
      salesSum = oldSheet.getRange(sumRow, 12).getValue(); // col L = 売上合計
    }
    sectionMapping.push({
      section: si + 1, dataRows: startRow + '-' + endRow,
      nameFound: nameAbove || 'なし',
      salesSumColL: salesSum
    });
  }

  // 各セクションの実データを確認（名前特定用）
  var sectionData = [];
  var secStarts = [41, 77, 113, 149, 186, 223, 260, 297, 334];
  for (var si = 0; si < secStarts.length; si++) {
    var ds = secStarts[si];
    var de = ds + 30;
    if (de > totalRows) continue;
    var secRange = oldSheet.getRange(ds, 1, de - ds + 1, Math.min(totalCols, 30)).getValues();
    // 各データ列の合計を算出
    var colSums = {};
    var colLabels = {3:'資金有商談', 6:'資金有成約', 9:'着金額', 12:'売上', 15:'CO数', 25:'資金なし成約', 26:'資金なし商談'};
    for (var colKey in colLabels) {
      var c = parseInt(colKey) - 1; // 0-indexed
      var sum = 0;
      for (var r = 0; r < secRange.length; r++) {
        sum += parseNum_(secRange[r][c]);
      }
      if (sum !== 0) colSums[colLabels[colKey]] = round1_(sum);
    }
    // セクション上3行のテキストを読む
    var aboveText = '';
    if (ds > 3) {
      var above = oldSheet.getRange(ds - 3, 1, 3, Math.min(totalCols, 10)).getValues();
      for (var ar = 0; ar < above.length; ar++) {
        for (var ac = 0; ac < above[ar].length; ac++) {
          var v = String(above[ar][ac] || '').trim();
          if (v && v.length > 1 && isNaN(Number(v))) aboveText += v + ' ';
        }
      }
    }
    sectionData.push({ section: si + 1, rows: ds + '-' + de, data: colSums, above: aboveText.trim().substring(0, 60) });
  }

  return { rowMap: rowMap, nameRow: nameRowIdx + 1, rowDump: rowDump, memberData: memberData,
    totalRows: totalRows, totalCols: totalCols, sectionInfo: sectionInfo, row36: row36Data,
    sectionMapping: sectionMapping, sectionData: sectionData };
}

/**
 * ダッシュボード表示用データのデバッグ（前月比含む）
 * API: ?action=run&fn=debugDashboardDeals
 */
/**
 * アーカイブ全データダンプ
 * API: ?action=run&fn=debugArchiveAll
 */
function debugArchiveAll() {
  var ss = getSpreadsheet_();
  var archiveSheet = getArchiveSheet_(ss);
  if (!archiveSheet) return { error: 'no archive sheet' };

  var data = archiveSheet.getDataRange().getValues();
  var result = {};
  for (var i = 1; i < data.length; i++) {
    var year = parseNum_(data[i][0]);
    var month = parseNum_(data[i][1]);
    var name = String(data[i][2] || '').trim();
    var key = year + '/' + month;
    if (!result[key]) result[key] = {};
    result[key][name] = {
      fundedDeals: parseNum_(data[i][3]),
      fundedClosed: parseNum_(data[i][4]),
      unfundedDeals: parseNum_(data[i][5]),
      unfundedClosed: parseNum_(data[i][6]),
      totalDeals: parseNum_(data[i][7]),
      closeRate: parseNum_(data[i][8]),
      sales: parseNum_(data[i][9]),
      revenue: parseNum_(data[i][10])
    };
  }
  return result;
}

function debugDashboardDeals() {
  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);
  var activeMembers = getActiveMembers_(ss);
  var prevMonthData = getPrevMonthFromArchive_(ss, settings);

  var result = [];
  for (var i = 0; i < activeMembers.length; i++) {
    var m = activeMembers[i];
    var dailySheet = getDailySheet_(ss, m.name);
    var sheetData = readDailySheetData_(dailySheet);
    var d = sheetData ? sheetData.totals : emptyTotals_();
    var prev = prevMonthData[m.name] || prevMonthData[m.displayName] || null;

    result.push({
      name: m.name,
      displayName: m.displayName,
      currentDeals: d.totalDeals,
      currentFundedDeals: d.fundedDeals,
      currentUnfundedDeals: d.unfundedDeals,
      currentClosed: d.fundedClosed,
      currentRevenue: d.revenue,
      prevMatchKey: prev ? (prevMonthData[m.name] ? m.name : m.displayName) : 'NO_MATCH',
      prevDeals: prev ? prev.deals : 0,
      prevClosed: prev ? prev.closed : 0,
      prevRevenue: prev ? prev.revenue : 0,
      diffDeals: d.totalDeals - (prev ? prev.deals : 0),
      diffClosed: d.fundedClosed - (prev ? prev.closed : 0)
    });
  }

  return {
    currentMonth: settings.month,
    currentYear: settings.year,
    prevMonthKeys: Object.keys(prevMonthData),
    members: result
  };
}

/**
 * 前月データのデバッグ（API: ?action=run&fn=debugPrevMonth）
 */
function debugPrevMonth() {
  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);
  var prevMonth = settings.month === 1 ? 12 : settings.month - 1;

  // Step 1: アーカイブチェック
  var archiveSheet = getArchiveSheet_(ss);
  var archiveHasData = false;
  if (archiveSheet && archiveSheet.getLastRow() > 1) {
    var arcData = archiveSheet.getDataRange().getValues();
    for (var i = 1; i < arcData.length; i++) {
      if (parseNum_(arcData[i][ARC_COL_MONTH - 1]) === prevMonth) {
        archiveHasData = true;
        break;
      }
    }
  }

  // Step 2: 旧シート検索
  var oldSheet = getSheetByMonth_(ss, prevMonth);
  var sheetInfo = oldSheet ? { name: oldSheet.getName(), rows: oldSheet.getLastRow(), cols: oldSheet.getLastColumn() } : null;

  // Step 3: 旧シートの先頭40行のダンプ
  var rowDump = [];
  var nameRowCheck = null;
  if (oldSheet) {
    var scanRows = Math.min(45, oldSheet.getLastRow());
    var scanCols = Math.min(45, oldSheet.getLastColumn());
    var scanData = oldSheet.getRange(1, 1, scanRows, scanCols).getValues();

    for (var r = 0; r < scanData.length; r++) {
      var cells = [];
      for (var c = 0; c < Math.min(35, scanData[r].length); c++) {
        var v = scanData[r][c];
        if (v !== '' && v !== null && v !== undefined) {
          cells.push({ c: c, v: String(v).substring(0, 25) });
        }
      }
      if (cells.length > 0) rowDump.push({ r: r, d: cells });

      // 名前行検出（【】3つ以上、または既知メンバー名3つ以上）
      var bc = 0;
      var knownNames = 0;
      for (var c = 0; c < scanData[r].length; c++) {
        var sv = String(scanData[r][c] || '').trim();
        if (sv.indexOf('【') !== -1) bc++;
        // 既知メンバー名チェック
        var clean = sv.replace(/[【】\[\]]/g, '');
        if (LEGACY_TO_V2_NAME[clean] || DISPLAY_NAME_MAP[clean] || ICON_MAP[clean] || OLD_DISPLAY_TO_V2[clean]) knownNames++;
      }
      if ((bc >= 3 || knownNames >= 3) && !nameRowCheck) {
        nameRowCheck = { row: r, bracketCount: bc, knownNames: knownNames, samples: [] };
        for (var c = 0; c < scanData[r].length; c++) {
          var nm = String(scanData[r][c] || '').trim();
          if (nm) nameRowCheck.samples.push({ col: c, name: nm.substring(0, 20) });
        }
      }
    }
  }

  // Step 4: 実際のデータ取得
  var prevData = {};
  try {
    prevData = getPrevMonthFromArchive_(ss, settings);
  } catch (e) {
    prevData = { error: e.message };
  }

  return {
    month: prevMonth,
    archiveHasData: archiveHasData,
    oldSheet: sheetInfo,
    rowDump: rowDump,
    nameRowCheck: nameRowCheck,
    data: prevData
  };
}

/**
 * 空の合計データを返す
 */
function emptyTotals_() {
  return {
    fundedDeals: 0, fundedClosed: 0, revenue: 0, sales: 0,
    coCount: 0, coAmount: 0, fundedPassed: 0, unfundedPassed: 0,
    unfundedClosed: 0, unfundedDeals: 0, totalDeals: 0,
    closeRate: 0, avgPrice: 0, creditCard: 0, shinpan: 0
  };
}

/**
 * ⚔️信販会社割合シートからクレカ/信販着金を自動集計し日別シートに反映
 *
 * シート構造:
 * - ヘッダー行(row 0)に「クレカ、銀行振込X月」「信販会社X月」列がある
 * - 最初の月別セクション（例:9月決済金額ごと）の名前列でメンバーを特定
 * - 同じ行のクレカ/信販列から値を読み取る
 */
function updateCreditShinpan_(ss, settings) {
  ss = ss || getSpreadsheet_();
  settings = settings || getGlobalSettings_(ss);

  var shinpanSheet = ss.getSheetByName(SHEET_SHINPAN);
  if (!shinpanSheet) {
    Logger.log('⚔️信販会社割合シートが見つかりません');
    return;
  }

  var currentMonth = settings.month;
  var allData = shinpanSheet.getDataRange().getValues();
  if (allData.length < 3) return;

  // ヘッダー行(row 0)から「クレカ」列と「信販会社」列を検索
  var colCredit = -1;
  var colShinpan = -1;
  var colName = -1;

  for (var c = 0; c < allData[0].length; c++) {
    var h = String(allData[0][c] || '');
    if (h.indexOf('クレカ') !== -1 && h.indexOf(currentMonth + '月') !== -1) {
      colCredit = c;
    }
    if (h.indexOf('信販会社') !== -1 && h.indexOf(currentMonth + '月') !== -1) {
      colShinpan = c;
    }
    // 最初の月別セクションのラベルから名前列位置を特定
    if (h.indexOf('月決済金額ごと') !== -1 && colName < 0) {
      colName = c - 1; // ラベルの1つ左がセクション開始（月番号）、データ行では名前列
    }
  }

  // ヘッダー行で見つからない場合、月番号なしで「クレカ」を検索
  if (colCredit < 0) {
    for (var c = 0; c < allData[0].length; c++) {
      var h = String(allData[0][c] || '');
      if (h.indexOf('クレカ') !== -1) colCredit = c;
      if (h.indexOf('信販会社') !== -1 && colShinpan < 0) colShinpan = c;
    }
  }

  if (colCredit < 0 && colShinpan < 0) {
    Logger.log('クレカ/信販のヘッダー列が見つかりません');
    return;
  }

  // 名前列が見つからない場合、最初の月別セクションのサブヘッダーから推定
  if (colName < 0) {
    for (var c = 0; c < allData[1].length; c++) {
      if (String(allData[1][c] || '').trim() === '銀振') {
        colName = c - 1;
        break;
      }
    }
  }

  if (colName < 0) {
    Logger.log('名前列が見つかりません');
    return;
  }

  // データ行を読み取り（row 2〜、「合計」まで）
  var memberData = {};
  var hasAnyData = false;

  for (var r = 2; r < allData.length; r++) {
    var name = String(allData[r][colName] || '').trim();
    if (!name) continue;
    if (name === '合計') break;

    var credit = colCredit >= 0 ? parseNum_(allData[r][colCredit]) : 0;
    var shinpan = colShinpan >= 0 ? parseNum_(allData[r][colShinpan]) : 0;

    if (credit === 0 && shinpan === 0) continue;

    var v2Name = resolveRealName_(name);
    if (!v2Name) continue;

    if (!memberData[v2Name]) {
      memberData[v2Name] = { credit: 0, shinpan: 0 };
    }
    memberData[v2Name].credit += credit;
    memberData[v2Name].shinpan += shinpan;
    hasAnyData = true;
  }

  // データが全てゼロの場合は上書きしない（既存値を保護）
  if (!hasAnyData) {
    Logger.log('クレカ/信販: ソースデータが空のため更新スキップ');
    return;
  }

  // 各日別シートに書き込み
  var updated = [];
  for (var memberName in memberData) {
    var dailySheet = getDailySheet_(ss, memberName);
    if (!dailySheet) continue;

    var md = memberData[memberName];
    dailySheet.getRange(DAILY_ROW_CREDIT_TOTAL, 2).setValue(round1_(md.credit));
    dailySheet.getRange(DAILY_ROW_SHINPAN_TOTAL, 2).setValue(round1_(md.shinpan));
    updated.push(memberName + ':クレカ=' + round1_(md.credit) + '/信販=' + round1_(md.shinpan));
  }

  Logger.log('クレカ/信販更新: ' + updated.join(', '));
}

/**
 * 本名からv2メンバー名を解決（完全一致→部分一致）
 */
function resolveRealName_(name) {
  // 完全一致
  if (REAL_NAME_TO_V2[name]) return REAL_NAME_TO_V2[name];

  // 「：」や「:」の前の名前で検索（例: 佐々木心雪：やまと → 佐々木心雪）
  var baseName = name.split(/[：:]/)[0].trim();
  if (REAL_NAME_TO_V2[baseName]) return REAL_NAME_TO_V2[baseName];

  // 苗字部分一致（キーがnameの先頭に含まれるか）
  for (var key in REAL_NAME_TO_V2) {
    if (name.indexOf(key) === 0 || baseName.indexOf(key) === 0) {
      return REAL_NAME_TO_V2[key];
    }
  }

  return null;
}

/**
 * ⚔️信販会社割合シートからCBS/ライフティ承認・申請数を直接読み取り日別シートに反映
 * 旧ウォーリアーズ数値シート経由ではなく、ソースシートから直接取得する
 */
function updateCBSLifetyFromShinpan_(ss, settings) {
  ss = ss || getSpreadsheet_();
  settings = settings || getGlobalSettings_(ss);

  var shinpanSheet = ss.getSheetByName(SHEET_SHINPAN);
  if (!shinpanSheet) {
    Logger.log('⚔️信販会社割合シートが見つかりません');
    return;
  }

  var targetMonth = settings.month + '月';
  var lastRow = shinpanSheet.getLastRow();
  var lastCol = shinpanSheet.getLastColumn();
  var scanEnd = Math.min(lastRow, 300);
  if (scanEnd < 5) return;

  var allData = shinpanSheet.getRange(1, 1, scanEnd, Math.min(lastCol, 10)).getValues();

  // 対象月のライフティ/CBSセクション開始行を探す
  var lifetyStart = -1, cbsStart = -1;
  for (var r = 0; r < allData.length; r++) {
    var eVal = String(allData[r][4] || '').trim();
    if (eVal === targetMonth + 'ライフティ') lifetyStart = r;
    if (eVal === targetMonth + 'CBS') cbsStart = r;
  }

  if (lifetyStart < 0 && cbsStart < 0) {
    Logger.log('CBS/ライフティ: ' + targetMonth + 'セクションなし');
    return;
  }

  // セクションからメンバーデータを読み取る共通関数
  // 構造: E=名前, F=承認, G=否決, H=成約(=申請数), I=承認割合, J=否決割合
  function readSection_(startRow) {
    var result = {};
    if (startRow < 0) return result;
    for (var r = startRow + 2; r < allData.length; r++) {
      var name = String(allData[r][4] || '').trim();
      if (!name) break;
      if (name.indexOf('月') !== -1) break;  // 次セクション

      var approved = parseNum_(allData[r][5]);
      var applied = parseNum_(allData[r][7]);  // H列=成約=申請総数

      if (approved > 0 || applied > 0) {
        var v2Name = resolveRealName_(name);
        if (v2Name) {
          result[v2Name] = { approved: approved, applied: applied };
        }
      }
    }
    return result;
  }

  var lifetyData = readSection_(lifetyStart);
  var cbsData = readSection_(cbsStart);

  // 各メンバーの日別シートに書き込み
  var updated = [];
  var members = getActiveMembers_(ss);
  for (var i = 0; i < members.length; i++) {
    var m = members[i];
    var dailySheet = getDailySheet_(ss, m.name);
    if (!dailySheet) continue;

    var lf = lifetyData[m.name];
    if (lf) {
      dailySheet.getRange(DAILY_ROW_LF_APPROVED, 2).setValue(lf.approved);
      dailySheet.getRange(DAILY_ROW_LF_APPLIED, 2).setValue(lf.applied);
      updated.push(m.displayName + ':LF=' + lf.approved + '/' + lf.applied);
    }

    var cbs = cbsData[m.name];
    if (cbs) {
      dailySheet.getRange(DAILY_ROW_CBS_APPROVED, 2).setValue(cbs.approved);
      dailySheet.getRange(DAILY_ROW_CBS_APPLIED, 2).setValue(cbs.applied);
      updated.push(m.displayName + ':CBS=' + cbs.approved + '/' + cbs.applied);
    }
  }

  Logger.log('CBS/ライフティ更新(' + targetMonth + '): ' + (updated.length > 0 ? updated.join(', ') : 'データなし'));
}

/**
 * クレカ/信販を手動更新（API・メニューから実行可能）
 */
function updateCreditShinpan() {
  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);
  updateCreditShinpan_(ss, settings);
  SpreadsheetApp.flush();
  return { ok: true };
}

/**
 * サマリー更新トリガーを設定（5分おき）
 */
function setupSummaryTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'updateSummary') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('updateSummary')
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log('サマリー更新トリガー（5分おき）を設定しました');
}
