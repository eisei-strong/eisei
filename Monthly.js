// ============================================
// Monthly.js — 月次アーカイブ・新月準備 (v2)
// ============================================

/**
 * 月締め処理: サマリーデータを月次アーカイブに保存
 */
function monthlyArchive() {
  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);
  var members = getActiveMembers_(ss);

  if (members.length === 0) {
    Logger.log('アクティブメンバーが見つかりません');
    return;
  }

  var archiveSheet = getArchiveSheet_(ss);
  if (!archiveSheet) {
    Logger.log('月次アーカイブシートが見つかりません');
    return;
  }

  // 既にアーカイブ済みかチェック
  var existingData = archiveSheet.getDataRange().getValues();
  for (var i = 1; i < existingData.length; i++) {
    if (parseNum_(existingData[i][0]) === settings.year &&
        parseNum_(existingData[i][1]) === settings.month) {
      Logger.log(settings.year + '年' + settings.month + '月は既にアーカイブ済み');
      return { status: 'already_archived' };
    }
  }

  // 各メンバーのデータを収集してアーカイブに追記
  var archiveRows = [];

  for (var m = 0; m < members.length; m++) {
    var member = members[m];
    var dailySheet = getDailySheet_(ss, member.name);
    var data = readDailySheetData_(dailySheet);

    if (!data) {
      Logger.log(member.name + ': データなし → スキップ');
      continue;
    }

    var d = data.totals;

    archiveRows.push([
      settings.year,
      settings.month,
      member.name,
      d.fundedDeals,
      d.fundedClosed,
      d.unfundedDeals,
      d.unfundedClosed,
      d.totalDeals,
      d.closeRate,
      d.sales,
      d.revenue,
      d.coCount,
      d.coAmount,
      d.creditCard,
      d.shinpan,
      data.cbs,
      data.lifety
    ]);
  }

  if (archiveRows.length > 0) {
    var lastRow = archiveSheet.getLastRow();
    archiveSheet.getRange(lastRow + 1, 1, archiveRows.length, ARCHIVE_HEADERS.length)
      .setValues(archiveRows);
    Logger.log(settings.year + '年' + settings.month + '月のデータをアーカイブに保存 (' + archiveRows.length + '件)');
  }

  return { status: 'archived', count: archiveRows.length };
}

/**
 * 新月セットアップ: 対象月を+1し、日別入力シートをリセット
 */
function newMonth() {
  var ss = getSpreadsheet_();
  var ui = SpreadsheetApp.getUi();

  var settings = getGlobalSettings_(ss);
  var nextMonth = settings.month === 12 ? 1 : settings.month + 1;
  var nextYear = settings.month === 12 ? settings.year + 1 : settings.year;

  var response = ui.alert(
    '新月セットアップ',
    settings.year + '年' + settings.month + '月 → ' + nextYear + '年' + nextMonth + '月に切り替えます。\n' +
    '日別入力シートのデータがクリアされます。\n\n' +
    '先に「月締め処理」でアーカイブしましたか？',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  // 1. 設定シートの対象月を更新
  var settingsSheet = getSettingsSheet_(ss);
  if (settingsSheet) {
    var lastRow = settingsSheet.getLastRow();
    if (lastRow >= SETTINGS_ROW_GLOBAL_START) {
      var rowCount = lastRow - SETTINGS_ROW_GLOBAL_START + 1;
      var globalData = settingsSheet.getRange(SETTINGS_ROW_GLOBAL_START, 1, rowCount, 2).getValues();

      for (var i = 0; i < globalData.length; i++) {
        var label = String(globalData[i][0] || '').trim();
        var row = SETTINGS_ROW_GLOBAL_START + i;
        if (label === '対象年度' || label === '対象年') {
          settingsSheet.getRange(row, 2).setValue(nextYear);
        }
        if (label === '対象月') {
          settingsSheet.getRange(row, 2).setValue(nextMonth);
        }
      }
    }
  }

  // 2. 各日別入力シートをリセット
  var members = getActiveMembers_(ss);
  var daysInMonth = new Date(nextYear, nextMonth, 0).getDate();

  for (var m = 0; m < members.length; m++) {
    var dailySheet = getDailySheet_(ss, members[m].name);
    if (!dailySheet) continue;

    resetDailySheet_(dailySheet, members[m].name, nextYear, nextMonth, daysInMonth, ss);
  }

  // 3. サマリー更新
  updateSummary();

  ui.alert(
    '完了',
    nextYear + '年' + nextMonth + '月のセットアップが完了しました。',
    ui.ButtonSet.OK
  );
}

/**
 * 日別入力シートを新月用にリセット
 * - 前月繰越にCO関連の繰越データを転記
 * - 日付を新月に更新
 * - データ部分クリア（前月繰越は残す）
 * - 合計行のSUM数式を再設定
 */
function resetDailySheet_(dailySheet, memberName, year, month, daysInMonth, ss) {
  // 現在の合計を取得（前月繰越に転記用）
  var currentData = readDailySheetData_(dailySheet);

  // 前月繰越行に転記するデータを準備
  var carryoverValues = new Array(DAILY_COL_COUNT).fill('');
  carryoverValues[COL_DATE] = '前月繰越';

  if (currentData) {
    // CO関連は繰越対象
    carryoverValues[COL_CO_COUNT] = currentData.totals.coCount;
    carryoverValues[COL_CO_AMOUNT] = currentData.totals.coAmount;
  }

  // 日別データクリア (Row 3-33)
  dailySheet.getRange(DAILY_ROW_DATA_START, 1, 31, DAILY_COL_COUNT).clearContent();

  // 前月繰越行を更新 (Row 2)
  dailySheet.getRange(DAILY_ROW_CARRYOVER, 1, 1, DAILY_COL_COUNT).setValues([carryoverValues]);

  // 日付を新月に更新 (Row 3-33)
  for (var d = 1; d <= 31; d++) {
    var rowNum = DAILY_ROW_DATA_START + d - 1;
    if (d <= daysInMonth) {
      dailySheet.getRange(rowNum, 1).setValue(new Date(year, month - 1, d));
    } else {
      dailySheet.getRange(rowNum, 1).setValue('');
    }
  }

  // 合計行のSUM数式を再設定 (Row 35)
  var totalRow = ['合計'];
  for (var c = 1; c < DAILY_COL_COUNT; c++) {
    var colLetter = columnToLetter_(c + 1);
    totalRow.push('=SUM(' + colLetter + DAILY_ROW_CARRYOVER + ':' + colLetter + DAILY_ROW_DATA_END + ')');
  }
  dailySheet.getRange(DAILY_ROW_TOTAL, 1, 1, DAILY_COL_COUNT).setValues([totalRow]);

  // CBS/ライフティ/クレカ/信販はクリア
  dailySheet.getRange(DAILY_ROW_CBS_APPROVED, 2).clearContent();
  dailySheet.getRange(DAILY_ROW_CBS_APPLIED, 2).clearContent();
  dailySheet.getRange(DAILY_ROW_LF_APPROVED, 2).clearContent();
  dailySheet.getRange(DAILY_ROW_LF_APPLIED, 2).clearContent();
  dailySheet.getRange(DAILY_ROW_CREDIT_TOTAL, 2).clearContent();
  dailySheet.getRange(DAILY_ROW_SHINPAN_TOTAL, 2).clearContent();

  Logger.log(memberName + ': ' + year + '年' + month + '月にリセット完了');
}

/**
 * 旧ウォーリアーズ数値シートから指定月のデータをアーカイブに保存
 * API用: ?action=run&fn=archiveFromOldSheet&month=2&year=2026
 */
function archiveFromOldSheet(targetMonth, targetYear) {
  var ss = getSpreadsheet_();

  if (!targetMonth || !targetYear) {
    var settings = getGlobalSettings_(ss);
    targetMonth = targetMonth || (settings.month === 1 ? 12 : settings.month - 1);
    targetYear = targetYear || (settings.month === 1 ? settings.year - 1 : settings.year);
  }
  targetMonth = parseNum_(targetMonth);
  targetYear = parseNum_(targetYear);

  var archiveSheet = getArchiveSheet_(ss);
  if (!archiveSheet) return { error: 'アーカイブシートなし' };

  // 既にアーカイブ済みかチェック
  var existingData = archiveSheet.getDataRange().getValues();
  for (var i = 1; i < existingData.length; i++) {
    if (parseNum_(existingData[i][0]) === targetYear &&
        parseNum_(existingData[i][1]) === targetMonth) {
      return { status: 'already_archived', month: targetMonth, year: targetYear };
    }
  }

  // 旧シートからデータ取得（getPrevMonthData_legacy_と同じロジック）
  var oldSheet = getSheetByMonth_(ss, targetMonth);
  if (!oldSheet) return { error: targetMonth + '月のシートが見つかりません' };

  var scanRows = Math.min(50, oldSheet.getLastRow());
  var scanCols = Math.min(50, oldSheet.getLastColumn());
  if (scanRows < 5) return { error: 'データ不足' };

  var allData = oldSheet.getRange(1, 1, scanRows, scanCols).getValues();

  // ラベルで行を動的検索
  var rowMap = {};
  for (var r = 0; r < allData.length; r++) {
    var label = String(allData[r][0] || '').replace(/\s+/g, '').replace(/[\(\)（）]/g, '');
    if (!label) label = String(allData[r][1] || '').replace(/\s+/g, '').replace(/[\(\)（）]/g, '');
    if (!label) continue;

    if (!rowMap.revenue && label.indexOf('合計着金額') !== -1) rowMap.revenue = r;
    if (!rowMap.deals && label.indexOf('合計商談数') !== -1 && label.indexOf('資金') === -1) rowMap.deals = r;
    if (label.indexOf('合計成約数') !== -1) rowMap.closed = r;
    if (!rowMap.closed && label === '成約数') rowMap.closed = r;
    if (!rowMap.closeRate && label.indexOf('成約率') !== -1) rowMap.closeRate = r;
    if (!rowMap.sales && (label.indexOf('合計売上') !== -1 || label === '売上')) rowMap.sales = r;
    if (!rowMap.fundedDeals && label.indexOf('資金有') !== -1 && label.indexOf('商談') !== -1) rowMap.fundedDeals = r;
    if (!rowMap.coCount && label.indexOf('CO数') !== -1) rowMap.coCount = r;
    if (!rowMap.coAmount && label.indexOf('CO額') !== -1 || label.indexOf('CO金額') !== -1) rowMap.coAmount = r;
    if (!rowMap.creditCard && label.indexOf('クレカ') !== -1) rowMap.creditCard = r;
    if (!rowMap.shinpan && label.indexOf('信販') !== -1 && label.indexOf('内') !== -1) rowMap.shinpan = r;
  }

  // 名前行を検索
  var nameRowIdx = -1;
  for (var r = 0; r < Math.min(15, allData.length); r++) {
    var knownCount = 0;
    for (var c = 0; c < allData[r].length; c++) {
      var v = String(allData[r][c] || '').trim().replace(/[【】\[\]]/g, '');
      if (LEGACY_TO_V2_NAME[v] || ICON_MAP[v] || OLD_DISPLAY_TO_V2[v]) knownCount++;
    }
    if (knownCount >= 3) { nameRowIdx = r; break; }
  }
  if (nameRowIdx < 0) return { error: '名前行が見つかりません' };

  // メンバー別にデータを取得してアーカイブ行を作成
  var archiveRows = [];
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

    var revenue = rowMap.revenue !== undefined ? round1_(parseNum_(allData[rowMap.revenue][c])) : 0;
    var deals = rowMap.deals !== undefined ? parseNum_(allData[rowMap.deals][c]) : 0;
    var closed = rowMap.closed !== undefined ? parseNum_(allData[rowMap.closed][c]) : 0;
    var sales = rowMap.sales !== undefined ? round1_(parseNum_(allData[rowMap.sales][c])) : 0;
    var fundedDeals = rowMap.fundedDeals !== undefined ? parseNum_(allData[rowMap.fundedDeals][c]) : 0;
    var coCount = rowMap.coCount !== undefined ? parseNum_(allData[rowMap.coCount][c]) : 0;
    var coAmount = rowMap.coAmount !== undefined ? round1_(parseNum_(allData[rowMap.coAmount][c])) : 0;
    var creditCard = rowMap.creditCard !== undefined ? round1_(parseNum_(allData[rowMap.creditCard][c])) : 0;
    var shinpan = rowMap.shinpan !== undefined ? round1_(parseNum_(allData[rowMap.shinpan][c])) : 0;

    var closeRate = 0;
    if (rowMap.closeRate !== undefined) {
      var rateRaw = allData[rowMap.closeRate][c];
      closeRate = typeof rateRaw === 'number' ? round1_(rateRaw * 100)
        : (deals > 0 ? round1_((closed / deals) * 100) : 0);
    } else if (deals > 0) {
      closeRate = round1_((closed / deals) * 100);
    }

    if (revenue > 0 || deals > 0 || closed > 0 || sales > 0) {
      var unfundedDeals = Math.max(0, deals - fundedDeals);
      archiveRows.push([
        targetYear, targetMonth, v2Name,
        fundedDeals, closed, unfundedDeals, 0,
        deals, closeRate, sales, revenue,
        coCount, coAmount, creditCard, shinpan,
        '-', '-'
      ]);
    }
  }

  if (archiveRows.length === 0) return { error: 'データなし' };

  var lastRow = archiveSheet.getLastRow();
  archiveSheet.getRange(lastRow + 1, 1, archiveRows.length, ARCHIVE_HEADERS.length)
    .setValues(archiveRows);

  return { status: 'archived', month: targetMonth, year: targetYear, count: archiveRows.length, rowMap: rowMap };
}

/**
 * 指定月のアーカイブを削除（再アーカイブ用）
 * API: ?action=run&fn=deleteArchive&month=1&year=2026
 */
function deleteArchive(targetMonth, targetYear) {
  targetMonth = parseNum_(targetMonth);
  targetYear = parseNum_(targetYear);

  var ss = getSpreadsheet_();
  var archiveSheet = getArchiveSheet_(ss);
  if (!archiveSheet) return { error: 'アーカイブシートなし' };

  var data = archiveSheet.getDataRange().getValues();
  var rowsToDelete = [];
  for (var i = data.length - 1; i >= 1; i--) {
    if (parseNum_(data[i][0]) === targetYear && parseNum_(data[i][1]) === targetMonth) {
      rowsToDelete.push(i + 1);
    }
  }

  for (var j = 0; j < rowsToDelete.length; j++) {
    archiveSheet.deleteRow(rowsToDelete[j]);
  }

  return { deleted: rowsToDelete.length, month: targetMonth, year: targetYear };
}

/**
 * 全月のアーカイブを再構築（削除→再アーカイブ）
 * API: ?action=run&fn=rebuildAllArchives
 */
function rebuildAllArchives() {
  var ss = getSpreadsheet_();
  var results = [];

  // 旧シートからアーカイブ可能な月を検出
  var months = [
    { year: 2025, month: 9 },
    { year: 2025, month: 10 },
    { year: 2025, month: 11 },
    { year: 2025, month: 12 },
    { year: 2026, month: 1 },
    { year: 2026, month: 2 }
  ];

  for (var i = 0; i < months.length; i++) {
    var m = months[i];
    // 既存データを削除
    var del = deleteArchive(m.month, m.year);
    // 再アーカイブ
    var arc = archiveFromOldSheet(m.month, m.year);
    results.push({ month: m.month, year: m.year, deleted: del.deleted, archive: arc.status || arc.error, count: arc.count || 0 });
  }

  return results;
}

/**
 * 月次処理トリガーを設定（毎月1日 9:00）
 */
function setupMonthlyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'monthlyArchive') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('monthlyArchive')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();
  Logger.log('月次アーカイブトリガー（毎月1日9時）を設定しました');
}
