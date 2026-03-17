// ============================================
// Migration.js — 旧→新データ移行 (v2)
// v1→v2 構造変更に対応
// ============================================

// ============================================
// v1→v2 構造移行
// ============================================

/**
 * v1→v2 移行（メニューから実行、UI確認あり）
 */
function migrateV1toV2() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'v1→v2 構造移行',
    'v1構造からv2構造へ移行します。\n\n' +
    '以下の変更が行われます:\n' +
    '- 日別入力シートの列ヘッダー更新\n' +
    '- 意味が変わった列(F,H,I,J)のデータクリア\n' +
    '- 設定シートを新構造に再構築\n' +
    '- サマリーシートを新レイアウトに再構築\n' +
    '- CO残管理シートを詳細構造に再構築\n\n' +
    '続行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  var result = doMigrateV1toV2_();

  ui.alert('移行完了', 'v2構造への移行が完了しました。\n' + JSON.stringify(result), ui.ButtonSet.OK);
}

/**
 * v1→v2 移行（API経由、UI確認なし）
 */
function migrateV1toV2_api() {
  return doMigrateV1toV2_();
}

/**
 * 全日別シートの前月繰越行をクリア（3月データのみにリセット）
 * API経由で呼び出し可能
 */
function clearCarryoverRows() {
  var ss = getSpreadsheet_();
  var sheets = ss.getSheets();
  var cleared = [];

  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name.indexOf(SHEET_DAILY_PREFIX) !== 0) continue;

    // 前月繰越行（Row 2）のB〜K列をクリア（A列の「前月繰越」ラベルは残す）
    sheets[i].getRange(DAILY_ROW_CARRYOVER, 2, 1, DAILY_COL_COUNT - 1).clearContent();
    cleared.push(name);
  }

  // SUM数式を再計算
  SpreadsheetApp.flush();

  // サマリー更新
  updateSummary();

  Logger.log('前月繰越クリア完了: ' + cleared.join(', '));
  return { cleared: cleared };
}

/**
 * v1→v2 移行の実体
 */
function doMigrateV1toV2_() {
  var ss = getSpreadsheet_();
  var result = { renamedSheets: [], updatedHeaders: [], clearedCols: [], rebuiltSheets: [] };

  // =======================================
  // Step 1: 日別入力シートのリネーム（v1内部名→v2表示名）
  // =======================================
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sheetName = sheets[i].getName();
    if (sheetName.indexOf(SHEET_DAILY_PREFIX) !== 0) continue;

    var memberPart = sheetName.replace(SHEET_DAILY_PREFIX, '');
    var v2Name = LEGACY_TO_V2_NAME[memberPart];
    if (v2Name && memberPart !== v2Name) {
      // v2名の日別シートが既にないか確認
      var existing = ss.getSheetByName(SHEET_DAILY_PREFIX + v2Name);
      if (!existing) {
        sheets[i].setName(SHEET_DAILY_PREFIX + v2Name);
        result.renamedSheets.push(memberPart + ' → ' + v2Name);
      }
    }
  }

  // =======================================
  // Step 2: 日別入力シートのヘッダー更新 & 意味変更列クリア
  // =======================================
  var allSheets = ss.getSheets();
  for (var j = 0; j < allSheets.length; j++) {
    var name = allSheets[j].getName();
    if (name.indexOf(SHEET_DAILY_PREFIX) !== 0) continue;

    var dailySheet = allSheets[j];

    // ヘッダー行を新v2ヘッダーに更新
    dailySheet.getRange(DAILY_ROW_HEADER, 1, 1, DAILY_COL_COUNT).setValues([DAILY_HEADERS]);
    result.updatedHeaders.push(name);

    // 意味が変わった列のデータをクリア（F=CO数, H=資金有見送り, I=資金なし見送り, J=資金なし成約）
    // v1: F=CO含着金額, H=商談数, I=クレカ着金, J=信販着金
    var colsToClear = [COL_CO_COUNT, COL_FUNDED_PASSED, COL_UNFUNDED_PASSED, COL_UNFUNDED_CLOSED];
    for (var c = 0; c < colsToClear.length; c++) {
      var col = colsToClear[c] + 1; // 1-based
      // 前月繰越行 (row 2) と日別データ行 (rows 3-33) をクリア
      dailySheet.getRange(DAILY_ROW_CARRYOVER, col).clearContent();
      dailySheet.getRange(DAILY_ROW_DATA_START, col, 31, 1).clearContent();
    }
    result.clearedCols.push(name);

    // 合計行のSUM数式を再設定（念のため）
    var totalRow = ['合計'];
    for (var tc = 1; tc < DAILY_COL_COUNT; tc++) {
      var colLetter = columnToLetter_(tc + 1);
      totalRow.push('=SUM(' + colLetter + DAILY_ROW_CARRYOVER + ':' + colLetter + DAILY_ROW_DATA_END + ')');
    }
    dailySheet.getRange(DAILY_ROW_TOTAL, 1, 1, DAILY_COL_COUNT).setValues([totalRow]);
  }

  // =======================================
  // Step 3: 設定シートを新構造に再構築
  // =======================================
  var oldSettings = readOldSettingsForMigration_(ss);
  var existingSettings = ss.getSheetByName(SHEET_SETTINGS);
  if (existingSettings) {
    ss.deleteSheet(existingSettings);
  }
  createSettingsSheet_(ss);
  // 旧設定から引き継げるデータを書き戻す
  writeBackMigratedSettings_(ss, oldSettings);
  result.rebuiltSheets.push(SHEET_SETTINGS);

  // =======================================
  // Step 4: サマリーシートを新レイアウトに再構築
  // =======================================
  var existingSummary = ss.getSheetByName(SHEET_SUMMARY);
  if (existingSummary) {
    ss.deleteSheet(existingSummary);
  }
  createSummarySheet_(ss);
  result.rebuiltSheets.push(SHEET_SUMMARY);

  // =======================================
  // Step 5: CO残管理シートを新構造に再構築
  // =======================================
  var existingCO = ss.getSheetByName(SHEET_CO_MANAGE);
  if (existingCO) {
    ss.deleteSheet(existingCO);
  }
  createCOManageSheet_(ss);
  result.rebuiltSheets.push(SHEET_CO_MANAGE);

  // =======================================
  // Step 6: サマリー更新
  // =======================================
  SpreadsheetApp.flush();
  updateSummary();

  Logger.log('v1→v2 移行完了: ' + JSON.stringify(result));
  return result;
}

/**
 * 旧設定シートのデータを読み取り（移行前に保存）
 */
function readOldSettingsForMigration_(ss) {
  var sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) return { members: [], globals: {} };

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 2) return { members: [], globals: {} };

  var allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  var members = [];
  var globals = {};

  // メンバーデータを読む（行2〜）
  for (var i = 1; i < allData.length; i++) {
    var name = String(allData[i][0] || '').trim();
    if (!name) continue;

    // グローバル設定行かチェック
    if (name === '年間着金目標' || name === 'チーム目標' || name === '対象年度' ||
        name === '対象年' || name === '対象月') {
      globals[name] = allData[i][1];
      continue;
    }

    // メンバー行
    var displayName = String(allData[i][1] || '').trim();
    var status, color;

    // v1構造: A=内部名, B=表示名, C=ステータス, D=カラー
    // v2構造: A=メンバー名, B=表示名, C=ランキング, D=ステータス, E=カラー
    if (lastCol >= 5 && allData[i][3] && (String(allData[i][3]).indexOf('アクティブ') !== -1 || String(allData[i][3]).indexOf('退職') !== -1)) {
      // v2構造の可能性
      status = String(allData[i][3] || '').trim();
      color = String(allData[i][4] || '').trim();
    } else {
      // v1構造
      status = String(allData[i][2] || '').trim();
      color = String(allData[i][3] || '').trim();
    }

    members.push({
      name: name,
      displayName: displayName,
      status: status,
      color: color
    });
  }

  return { members: members, globals: globals };
}

/**
 * 移行した設定データを新構造の設定シートに書き戻す
 */
function writeBackMigratedSettings_(ss, oldSettings) {
  var sheet = getSettingsSheet_(ss);
  if (!sheet) return;

  // グローバル設定を書き戻す
  if (oldSettings.globals['対象年度'] || oldSettings.globals['対象年']) {
    var yearVal = oldSettings.globals['対象年度'] || oldSettings.globals['対象年'];
    sheet.getRange(SETTINGS_ROW_GLOBAL_START + 1, 2).setValue(yearVal);
  }
  if (oldSettings.globals['対象月']) {
    sheet.getRange(SETTINGS_ROW_GLOBAL_START + 2, 2).setValue(oldSettings.globals['対象月']);
  }
}

// ============================================
// 旧v1移行関連（フォールバック用に残す）
// ============================================

/**
 * 旧シートから新構造へ3月データを移行（メニューから実行）
 */
function migrateCurrentMonthData() {
  var ss = getSpreadsheet_();
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert(
    'データ移行',
    '旧シートから新しい日別入力シートへデータを移行します。\n既存のデータは上書きされます。続行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  var settings = getGlobalSettings_(ss);
  var currentMonth = settings.month;

  var oldSheet = getSheetByMonth_(ss, currentMonth);
  if (!oldSheet) {
    ui.alert('エラー', currentMonth + '月の旧シートが見つかりません。', ui.ButtonSet.OK);
    return;
  }

  Logger.log('移行元シート: ' + oldSheet.getName());

  var allData = oldSheet.getDataRange().getValues();
  var allFormulas = oldSheet.getDataRange().getFormulas();
  var maxCol = allData[0].length;

  var migratedMembers = [];
  for (var i = 0; i < OLD_MEMBER_COLS.length; i++) {
    var col = OLD_MEMBER_COLS[i];
    var memberName = OLD_MEMBER_NAME_MAP[col];
    if (!memberName) continue;

    // v2名でシートを探す
    var v2Name = LEGACY_TO_V2_NAME[memberName] || memberName;
    var dailySheet = getDailySheet_(ss, v2Name) || getDailySheet_(ss, memberName);
    if (!dailySheet) {
      Logger.log('日別入力シートなし: ' + memberName + ' / ' + v2Name + ' → スキップ');
      continue;
    }

    migrateMonthlyTotals_(dailySheet, allData, col);
    migrateCBSLifety_(dailySheet, allData, col);
    migrateCreditShinpan_(dailySheet, allData, allFormulas, col);
    migrateDailyData_(dailySheet, oldSheet, allData, col, v2Name, maxCol);

    migratedMembers.push(v2Name);
    Logger.log('移行完了: ' + v2Name);
  }

  migratePrevMonthToArchive_(ss, settings);
  updateSummary();

  ui.alert(
    '移行完了',
    '以下のメンバーのデータを移行しました:\n' + migratedMembers.join(', ') +
    '\n\nサマリーも更新しました。',
    ui.ButtonSet.OK
  );
}

/**
 * 月次合計データを日別入力シートの前月繰越行に書き込み
 */
function migrateMonthlyTotals_(dailySheet, allData, col) {
  var revenue     = round1_(parseNum_(allData[OLD_ROW_REVENUE][col]));
  var fundedDeals = parseNum_(allData[OLD_ROW_FUNDED_DEALS][col]);
  var deals       = parseNum_(allData[OLD_ROW_DEALS][col]);
  var closed      = parseNum_(allData[OLD_ROW_CLOSED][col]);
  var sales       = round1_(parseNum_(allData[OLD_ROW_SALES][col]));
  var coAmount    = round1_(parseNum_(allData[OLD_ROW_CO_AMOUNT][col]));

  // v2列構造に合わせた繰越データ
  // F(CO数), H(資金有見送り), I(資金なし見送り), J(資金なし成約) は旧データから対応がないので0
  var carryoverData = [
    fundedDeals, // B: 資金有商談数
    closed,      // C: 資金有成約数（v1の成約数をそのまま）
    revenue,     // D: 着金額（後で日別移行時に調整）
    sales,       // E: 売上
    0,           // F: CO数（v2新規、旧データなし）
    coAmount,    // G: CO金額
    0,           // H: 資金有で見送り（v2新規）
    0,           // I: 資金なし見送り数（v2新規）
    0,           // J: 資金なし成約数（v2新規）
    0            // K: 資金なし商談数
  ];
  dailySheet.getRange(DAILY_ROW_CARRYOVER, 2, 1, DAILY_COL_COUNT - 1)
    .setValues([carryoverData]);
}

/**
 * CBS/ライフティデータを移行
 */
function migrateCBSLifety_(dailySheet, allData, col) {
  var cbsRaw = allData[OLD_ROW_CBS][col];
  var cbsStr = String(cbsRaw || '');

  if (cbsStr && cbsStr !== '-') {
    var cbsParts = cbsStr.split('/');
    if (cbsParts.length === 2) {
      dailySheet.getRange(DAILY_ROW_CBS_APPROVED, 2).setValue(parseNum_(cbsParts[0]));
      dailySheet.getRange(DAILY_ROW_CBS_APPLIED, 2).setValue(parseNum_(cbsParts[1]));
    } else {
      dailySheet.getRange(DAILY_ROW_CBS_APPROVED, 2).setValue(cbsStr);
    }
  }

  var lfRaw = allData[OLD_ROW_LIFETY][col];
  var lfStr = String(lfRaw || '');

  if (lfStr && lfStr !== '-') {
    var lfParts = lfStr.split('/');
    if (lfParts.length === 2) {
      dailySheet.getRange(DAILY_ROW_LF_APPROVED, 2).setValue(parseNum_(lfParts[0]));
      dailySheet.getRange(DAILY_ROW_LF_APPLIED, 2).setValue(parseNum_(lfParts[1]));
    } else {
      dailySheet.getRange(DAILY_ROW_LF_APPROVED, 2).setValue(lfStr);
    }
  }
}

/**
 * クレカ・信販の数式/値を移行
 */
function migrateCreditShinpan_(dailySheet, allData, allFormulas, col) {
  var creditFormula = allFormulas[OLD_ROW_CREDIT_CARD][col];
  var creditValue = round1_(parseNum_(allData[OLD_ROW_CREDIT_CARD][col]));

  if (creditFormula) {
    dailySheet.getRange(DAILY_ROW_CREDIT_TOTAL, 2).setFormula(creditFormula);
  } else if (creditValue > 0) {
    dailySheet.getRange(DAILY_ROW_CREDIT_TOTAL, 2).setValue(creditValue);
  }

  var shinpanFormula = allFormulas[OLD_ROW_SHINPAN][col];
  var shinpanValue = round1_(parseNum_(allData[OLD_ROW_SHINPAN][col]));

  if (shinpanFormula) {
    dailySheet.getRange(DAILY_ROW_SHINPAN_TOTAL, 2).setFormula(shinpanFormula);
  } else if (shinpanValue > 0) {
    dailySheet.getRange(DAILY_ROW_SHINPAN_TOTAL, 2).setValue(shinpanValue);
  }
}

/**
 * 日別ブロックから着金額等の日次データを移行
 */
function migrateDailyData_(dailySheet, oldSheet, allData, memberCol, memberName, maxCol) {
  var headerRows = [];
  for (var r = 40; r < allData.length; r++) {
    if (String(allData[r][1] || '').replace(/\s+/g, '') === '日別') {
      headerRows.push(r);
    }
  }

  var dailyRevenues = {};

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
      rankFormulas = oldSheet.getRange(rankRow + 1, 1, 1, maxCol).getFormulas()[0];
    }

    for (var s = 0; s < sections.length; s++) {
      var sec = sections[s];
      var detectedCol = -1;
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
            detectedCol = colIdx - 1;
            break;
          }
        }
      }

      if (detectedCol !== memberCol) continue;

      for (var dr = hr + 1; dr < Math.min(hr + 32, allData.length); dr++) {
        var dateVal = allData[dr][sec.dateCol];
        var revVal = allData[dr][sec.revCol];
        if (dateVal instanceof Date && revVal && revVal !== 0) {
          var amount = round1_(parseNum_(revVal));
          if (amount > 0) {
            var day = dateVal.getDate();
            dailyRevenues[day] = (dailyRevenues[day] || 0) + amount;
          }
        }
      }
    }
  }

  if (Object.keys(dailyRevenues).length > 0) {
    for (var day in dailyRevenues) {
      var rowNum = DAILY_ROW_DATA_START + (parseInt(day) - 1);
      dailySheet.getRange(rowNum, COL_REVENUE + 1).setValue(dailyRevenues[day]);
    }

    // 前月繰越はsyncFromOldSheet()で旧シートから直接取得済み → 調整不要
    Logger.log(memberName + ': 日別着金 ' + Object.keys(dailyRevenues).length + '日分移行');
  }
}

/**
 * 旧シートからメンバーセクションの行範囲を動的に検出
 * 前月繰越着金額でセクション→メンバーを照合する堅牢な方式
 */
function detectMemberSections_(oldSheet) {
  var allData = oldSheet.getDataRange().getValues();

  // Step 1: 「日別」ヘッダー行を全て見つける（row 25以降）
  var sectionHeaders = []; // [row 0-indexed]
  for (var r = 25; r < allData.length; r++) {
    for (var c = 0; c < Math.min(30, allData[r].length); c++) {
      if (String(allData[r][c] || '').replace(/\s+/g, '') === '日別') {
        sectionHeaders.push(r);
        break;
      }
    }
  }

  Logger.log('セクションヘッダー検出: ' + sectionHeaders.length + '件');

  // Step 2: 各セクションの前月繰越着金額を読み取る
  // 前月繰越はセクションデータの後（次のセクションヘッダーの前）にある
  var sectionCarryovers = []; // [carryoverRev]
  for (var si = 0; si < sectionHeaders.length; si++) {
    var hdr = sectionHeaders[si]; // 0-indexed
    var dataEnd = hdr + 31; // 31日分のデータ末尾 (0-indexed)
    var carryRev = 0;

    // dataEnd+1 〜 次のセクションヘッダー（or +6）の範囲で前月繰越を探す
    var searchEnd = (si + 1 < sectionHeaders.length) ? sectionHeaders[si + 1] : Math.min(dataEnd + 8, allData.length - 1);
    for (var cr = dataEnd + 1; cr < searchEnd; cr++) {
      for (var cc = 0; cc < Math.min(10, allData[cr].length); cc++) {
        if (String(allData[cr][cc] || '').replace(/\s+/g, '') === '前月繰越') {
          if (cr + 1 < allData.length) {
            carryRev = parseNum_(allData[cr + 1][8]); // I列 = 着金額
          }
          break;
        }
      }
      if (carryRev !== 0) break;
    }
    sectionCarryovers.push(carryRev);
  }

  // Step 3: サマリー行の前月繰越値を読み取る
  var carryoverRowIdx = -1;
  for (var r = 0; r < Math.min(40, allData.length); r++) {
    var label = String(allData[r][0] || '').replace(/\s+/g, '') + String(allData[r][1] || '').replace(/\s+/g, '');
    if (label.indexOf('前月繰り越し') !== -1 || label.indexOf('前月繰越') !== -1) {
      carryoverRowIdx = r;
      break;
    }
  }

  // Step 4: 前月繰越値でセクション→メンバーを照合
  var sections = [];
  for (var i = 0; i < MEMBER_SECTIONS.length; i++) {
    sections.push({
      name: MEMBER_SECTIONS[i].name,
      summaryCol: MEMBER_SECTIONS[i].summaryCol,
      dataStart: MEMBER_SECTIONS[i].dataStart,
      dataEnd: MEMBER_SECTIONS[i].dataEnd
    });
  }

  if (sectionHeaders.length !== MEMBER_SECTIONS.length) {
    Logger.log('WARNING: セクション数(' + sectionHeaders.length + ')≠メンバー数(' + MEMBER_SECTIONS.length + ') → 静的値使用');
    return sections;
  }

  // メンバーのcarryover値を取得
  var memberCarries = []; // [{idx, carry}]
  for (var i = 0; i < MEMBER_SECTIONS.length; i++) {
    var sc = MEMBER_SECTIONS[i].summaryCol - 1; // 0-indexed
    var carry = (carryoverRowIdx >= 0) ? parseNum_(allData[carryoverRowIdx][sc]) : 0;
    memberCarries.push({ idx: i, carry: Math.round(carry * 10) / 10 });
  }

  // セクションとメンバーを照合（carryover値でマッチ）
  var matched = {}; // sectionIdx → memberIdx
  var usedMembers = {};

  // 非ゼロcarryover同士のマッチング（一意に特定できる）
  for (var si = 0; si < sectionHeaders.length; si++) {
    var secCarry = Math.round(sectionCarryovers[si] * 10) / 10;
    if (secCarry === 0) continue;
    for (var mi = 0; mi < memberCarries.length; mi++) {
      if (usedMembers[mi]) continue;
      if (Math.abs(secCarry - memberCarries[mi].carry) < 0.5) {
        matched[si] = mi;
        usedMembers[mi] = true;
        Logger.log('セクション' + (si+1) + ' (carry=' + secCarry + ') → ' + MEMBER_SECTIONS[mi].name + ' (carry=' + memberCarries[mi].carry + ')');
        break;
      }
    }
  }

  // 残りのセクション（carry=0）は元のMEMBER_SECTIONSのdataStart近接で対応付け
  var unmatchedSections = [];
  var unmatchedMembers = [];
  for (var si = 0; si < sectionHeaders.length; si++) {
    if (matched[si] === undefined) unmatchedSections.push(si);
  }
  for (var mi = 0; mi < MEMBER_SECTIONS.length; mi++) {
    if (!usedMembers[mi]) unmatchedMembers.push(mi);
  }

  // 未マッチセクション/メンバーをdataStart近接で対応付け
  for (var ui = 0; ui < unmatchedSections.length; ui++) {
    var si = unmatchedSections[ui];
    var secDataStart = sectionHeaders[si] + 2; // 1-indexed data start
    var bestMi = -1;
    var bestDist = Infinity;
    for (var uj = 0; uj < unmatchedMembers.length; uj++) {
      var mi = unmatchedMembers[uj];
      if (usedMembers[mi]) continue;
      var dist = Math.abs(MEMBER_SECTIONS[mi].dataStart - secDataStart);
      if (dist < bestDist) {
        bestDist = dist;
        bestMi = mi;
      }
    }
    if (bestMi >= 0) {
      matched[si] = bestMi;
      usedMembers[bestMi] = true;
      Logger.log('セクション' + (si+1) + ' (carry=0, row=' + secDataStart + ') → ' + MEMBER_SECTIONS[bestMi].name + ' (近接マッチ dist=' + bestDist + ')');
    }
  }

  // Step 5: マッチング結果でdataStart/dataEndを更新
  for (var si = 0; si < sectionHeaders.length; si++) {
    if (matched[si] === undefined) continue;
    var mi = matched[si];
    var headerRow = sectionHeaders[si] + 1; // 1-indexed
    sections[mi].dataStart = headerRow + 1;
    sections[mi].dataEnd = headerRow + 31;
  }

  return sections;
}

/**
 * 各セクションの前月繰越着金額を読み取り、メンバー列の前月繰越値と照合して
 * セクション→メンバーのマッピングを特定する
 */
function debugSectionCarryover() {
  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);
  var oldSheet = getSheetByMonth_(ss, settings.month);
  if (!oldSheet) return { error: 'シートなし' };

  var allData = oldSheet.getDataRange().getValues();

  // 1. サマリー行のcarryoverRev行を特定
  var carryoverRow = -1;
  for (var r = 0; r < Math.min(40, allData.length); r++) {
    var label = String(allData[r][0] || '').replace(/\s+/g, '') + String(allData[r][1] || '').replace(/\s+/g, '');
    if (label.indexOf('前月繰り越し') !== -1 || label.indexOf('前月繰越') !== -1) {
      carryoverRow = r;
      break;
    }
  }

  // 2. 名前行を検出してメンバー列を特定
  var nameRowIdx = -1;
  for (var r = 0; r < Math.min(10, allData.length); r++) {
    var bracketCount = 0;
    for (var c = 0; c < allData[r].length; c++) {
      if (String(allData[r][c] || '').indexOf('【') !== -1) bracketCount++;
    }
    if (bracketCount >= 3) { nameRowIdx = r; break; }
  }

  // 3. 各メンバーのサマリーcarryover値を取得
  var memberCarryovers = [];
  for (var i = 0; i < MEMBER_SECTIONS.length; i++) {
    var sc = MEMBER_SECTIONS[i].summaryCol - 1; // 0-indexed
    var carryVal = carryoverRow >= 0 ? parseNum_(allData[carryoverRow][sc]) : 0;
    memberCarryovers.push({
      name: MEMBER_SECTIONS[i].name,
      summaryCol: MEMBER_SECTIONS[i].summaryCol,
      carryoverRev: carryVal
    });
  }

  // 4. セクションヘッダーを検出し、各セクションの前月繰越着金額を読み取る
  var sectionHeaders = [];
  for (var r = 25; r < allData.length; r++) {
    for (var c = 0; c < Math.min(30, allData[r].length); c++) {
      if (String(allData[r][c] || '').replace(/\s+/g, '') === '日別') {
        sectionHeaders.push(r);
        break;
      }
    }
  }

  var sectionInfo = [];
  for (var si = 0; si < sectionHeaders.length; si++) {
    var hdr = sectionHeaders[si]; // 0-indexed header row
    var carryRev = 0;
    // 前月繰越は header の 3-6行前にある
    for (var cr = hdr - 6; cr < hdr; cr++) {
      if (cr < 0) continue;
      for (var cc = 0; cc < Math.min(30, allData[cr].length); cc++) {
        var clabel = String(allData[cr][cc] || '').replace(/\s+/g, '');
        if (clabel === '前月繰越') {
          // 次の行の着金額列を読む
          // 着金額列はヘッダー行のcol9 (0-indexed 8) と同じ
          if (cr + 1 < allData.length) {
            carryRev = parseNum_(allData[cr + 1][8]); // I列 = col9 (0-indexed 8)
          }
          break;
        }
      }
    }

    // セクションデータの着金額合計
    var revTotal = 0;
    var dealsTotal = 0;
    var closedTotal = 0;
    var salesTotal = 0;
    for (var dr = hdr + 1; dr <= Math.min(hdr + 31, allData.length - 1); dr++) {
      if (allData[dr][1] instanceof Date) { // col B = date column
        revTotal += parseNum_(allData[dr][8]);    // I列
        dealsTotal += parseNum_(allData[dr][2]);   // C列 = 商談数
        closedTotal += parseNum_(allData[dr][5]);  // F列 = 成約数
        salesTotal += parseNum_(allData[dr][11]);  // L列 = 売上
      }
    }

    sectionInfo.push({
      index: si + 1,
      headerRow: hdr + 1,
      dataStart: hdr + 2,
      dataEnd: hdr + 32,
      carryoverRev: carryRev,
      revTotal: Math.round(revTotal * 10) / 10,
      dealsTotal: dealsTotal,
      closedTotal: closedTotal,
      salesTotal: Math.round(salesTotal * 10) / 10
    });
  }

  return {
    carryoverRow: carryoverRow >= 0 ? carryoverRow + 1 : 'not found',
    memberCarryovers: memberCarryovers,
    sections: sectionInfo
  };
}

/**
 * 旧シートのセクション構造をダンプして正しいメンバーマッピングを特定する
 */
function debugSectionMapping() {
  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);
  var oldSheet = getSheetByMonth_(ss, settings.month);
  if (!oldSheet) return { error: 'シートなし' };

  var allData = oldSheet.getDataRange().getValues();
  var result = { totalRows: allData.length, sections: [], memberSections: [] };

  // 1. 全ての「日別」ヘッダー行を検出
  for (var r = 25; r < allData.length; r++) {
    for (var c = 0; c < Math.min(30, allData[r].length); c++) {
      var val = String(allData[r][c] || '').replace(/\s+/g, '');
      if (val === '日別') {
        // この行のヘッダー全体を取得
        var headerCells = [];
        for (var hc = 0; hc < Math.min(30, allData[r].length); hc++) {
          var hv = allData[r][hc];
          if (hv !== '' && hv !== null && hv !== undefined) {
            headerCells.push('col' + (hc+1) + '=' + String(hv));
          }
        }

        // 上5行の内容をダンプ
        var aboveRows = [];
        for (var ar = Math.max(0, r - 5); ar < r; ar++) {
          var rowCells = [];
          for (var ac = 0; ac < Math.min(30, allData[ar].length); ac++) {
            var av = allData[ar][ac];
            if (av !== '' && av !== null && av !== undefined) {
              var avStr = (av instanceof Date) ? 'DATE:' + av.toLocaleDateString() : String(av);
              if (avStr.length > 30) avStr = avStr.substring(0, 30) + '...';
              rowCells.push('col' + (ac+1) + '=' + avStr);
            }
          }
          if (rowCells.length > 0) {
            aboveRows.push({ row: ar + 1, cells: rowCells });
          }
        }

        // 下2行のデータをダンプ
        var belowRows = [];
        for (var br = r + 1; br <= Math.min(r + 2, allData.length - 1); br++) {
          var rowCells = [];
          for (var bc = 0; bc < Math.min(30, allData[br].length); bc++) {
            var bv = allData[br][bc];
            if (bv !== '' && bv !== null && bv !== undefined) {
              var bvStr = (bv instanceof Date) ? 'DATE:' + bv.toLocaleDateString() : String(bv);
              if (bvStr.length > 30) bvStr = bvStr.substring(0, 30) + '...';
              rowCells.push('col' + (bc+1) + '=' + bvStr);
            }
          }
          if (rowCells.length > 0) {
            belowRows.push({ row: br + 1, cells: rowCells });
          }
        }

        // セクションの着金額(I列=col9)合計を計算
        var revTotal = 0;
        var dataCount = 0;
        for (var dr = r + 1; dr <= Math.min(r + 35, allData.length - 1); dr++) {
          if (allData[dr][c] instanceof Date) { // dateCol=c にDate値あり
            var revVal = parseNum_(allData[dr][8]); // I列=index8
            revTotal += revVal;
            dataCount++;
          }
        }

        result.sections.push({
          headerRow: r + 1,
          dateCol: c + 1,
          header: headerCells,
          above: aboveRows,
          below: belowRows,
          revTotal: Math.round(revTotal * 10) / 10,
          dataRows: dataCount
        });
        break; // この行で1つ見つけたら次の行へ
      }
    }
  }

  // 2. 現在のMEMBER_SECTIONSの各エントリについて、実際のデータを確認
  for (var i = 0; i < MEMBER_SECTIONS.length; i++) {
    var ms = MEMBER_SECTIONS[i];
    var ds = ms.dataStart - 1; // 0-indexed
    var de = ms.dataEnd - 1;

    // dataStart付近の内容
    var nearCells = [];
    for (var nr = Math.max(0, ds - 3); nr <= Math.min(ds + 1, allData.length - 1); nr++) {
      var rowCells = [];
      for (var nc = 0; nc < Math.min(20, allData[nr].length); nc++) {
        var nv = allData[nr][nc];
        if (nv !== '' && nv !== null && nv !== undefined) {
          var nvStr = (nv instanceof Date) ? 'DATE:' + nv.toLocaleDateString() : String(nv);
          if (nvStr.length > 30) nvStr = nvStr.substring(0, 30) + '...';
          rowCells.push('col' + (nc+1) + '=' + nvStr);
        }
      }
      if (rowCells.length > 0) {
        nearCells.push({ row: nr + 1, cells: rowCells });
      }
    }

    // セクション内着金額合計
    var secRevTotal = 0;
    for (var sr = ds; sr <= Math.min(de, allData.length - 1); sr++) {
      secRevTotal += parseNum_(allData[sr][8]); // I列
    }

    result.memberSections.push({
      name: ms.name,
      summaryCol: ms.summaryCol,
      dataStart: ms.dataStart,
      dataEnd: ms.dataEnd,
      nearData: nearCells,
      revTotal: Math.round(secRevTotal * 10) / 10
    });
  }

  return result;
}

/**
 * 動的検出されたセクション情報を使って日別着金データを読み取り
 * sectionsMapは { memberName: {dataStart, dataEnd} } 形式（1-indexed）
 */
function migrateDailyDataDirect_(dailySheet, allData, memberName, sectionsMap) {
  // sectionsMapから該当メンバーのセクションを取得
  var section = sectionsMap ? sectionsMap[memberName] : null;
  if (!section) {
    // フォールバック: MEMBER_SECTIONS
    for (var i = 0; i < MEMBER_SECTIONS.length; i++) {
      var secName = MEMBER_SECTIONS[i].name;
      if (secName === memberName || memberName.indexOf(secName) === 0 || secName.indexOf(memberName) === 0) {
        section = MEMBER_SECTIONS[i];
        break;
      }
    }
  }
  if (!section || section.dataStart < 0) return;

  var sectionStart = section.dataStart - 1; // 0-indexed
  var sectionEnd = section.dataEnd - 1;

  // セクションヘッダー行から日別/着金額列を検出
  var dateCol = -1;
  var revCol = -1;
  var hdrRow = Math.max(0, sectionStart - 1); // ヘッダーはデータ開始の1行前
  for (var c = 0; c < allData[hdrRow].length; c++) {
    var label = String(allData[hdrRow][c] || '').replace(/\s+/g, '');
    if (label === '日別') dateCol = c;
    if (label === '着金額') revCol = c;
  }
  if (dateCol < 0 || revCol < 0) {
    Logger.log(memberName + ': 日別/着金額ヘッダーなし (hdrRow=' + (hdrRow+1) + ')');
    return;
  }
  Logger.log(memberName + ': セクション rows ' + (sectionStart+1) + '-' + (sectionEnd+1) + ' (dateCol=' + dateCol + ', revCol=' + revCol + ')');

  // まず既存の着金額データをクリア（前回sync時の誤データを消す）
  var clearRange = dailySheet.getRange(DAILY_ROW_DATA_START, COL_REVENUE + 1, DAILY_ROW_DATA_END - DAILY_ROW_DATA_START + 1, 1);
  clearRange.clearContent();

  // データ行から日別着金額を読み取り
  var dailyRevenues = {};
  for (var r = sectionStart; r <= sectionEnd && r < allData.length; r++) {
    var dateVal = allData[r][dateCol];
    var revVal = allData[r][revCol];
    if (dateVal instanceof Date && revVal && revVal !== 0) {
      var amount = round1_(parseNum_(revVal));
      if (amount > 0) {
        var day = dateVal.getDate();
        dailyRevenues[day] = (dailyRevenues[day] || 0) + amount;
      }
    }
  }

  // v2日別シートに書き込み
  if (Object.keys(dailyRevenues).length > 0) {
    for (var day in dailyRevenues) {
      var rowNum = DAILY_ROW_DATA_START + (parseInt(day) - 1);
      dailySheet.getRange(rowNum, COL_REVENUE + 1).setValue(dailyRevenues[day]);
    }
    Logger.log(memberName + ': 日別着金(直接) ' + Object.keys(dailyRevenues).length + '日分移行');
  }
}

/**
 * 前月データを月次アーカイブに保存（前月比較用）
 */
function migratePrevMonthToArchive_(ss, settings) {
  var prevMonth = settings.month === 1 ? 12 : settings.month - 1;
  var prevYear = settings.month === 1 ? settings.year - 1 : settings.year;

  var oldSheet = getSheetByMonth_(ss, prevMonth);
  if (!oldSheet) {
    Logger.log(prevMonth + '月の旧シートなし → アーカイブスキップ');
    return;
  }

  var archiveSheet = getArchiveSheet_(ss);
  if (!archiveSheet) {
    Logger.log('アーカイブシートなし → スキップ');
    return;
  }

  // 既にアーカイブ済みかチェック
  var existingData = archiveSheet.getDataRange().getValues();
  for (var i = 1; i < existingData.length; i++) {
    if (parseNum_(existingData[i][0]) === prevYear && parseNum_(existingData[i][1]) === prevMonth) {
      Logger.log(prevYear + '年' + prevMonth + '月は既にアーカイブ済み');
      return;
    }
  }

  var allData = oldSheet.getDataRange().getValues();
  var archiveRows = [];

  for (var m = 0; m < OLD_MEMBER_COLS.length; m++) {
    var col = OLD_MEMBER_COLS[m];
    var name = OLD_MEMBER_NAME_MAP[col];
    if (!name) continue;

    var v2Name = LEGACY_TO_V2_NAME[name] || displayName_(name);
    var closed = parseNum_(allData[OLD_ROW_CLOSED][col]);
    var deals = parseNum_(allData[OLD_ROW_DEALS][col]);
    var fundedDeals = parseNum_(allData[OLD_ROW_FUNDED_DEALS][col]);
    var rateRaw = allData[OLD_ROW_RATE][col];
    var closeRate = typeof rateRaw === 'number'
      ? round1_(rateRaw * 100)
      : (deals > 0 ? round1_((closed / deals) * 100) : 0);

    // v2アーカイブ構造: 17列
    archiveRows.push([
      prevYear,
      prevMonth,
      v2Name,
      fundedDeals,         // 資金有商談数
      closed,              // 資金有成約数（v1の成約数）
      0,                   // 資金なし商談数（v1にデータなし）
      0,                   // 資金なし成約数（v1にデータなし）
      deals,               // 合計商談数
      closeRate,           // 成約率
      round1_(parseNum_(allData[OLD_ROW_SALES][col])),      // 売上
      round1_(parseNum_(allData[OLD_ROW_REVENUE][col])),    // 着金額
      0,                   // CO数（v1にデータなし）
      round1_(parseNum_(allData[OLD_ROW_CO_AMOUNT][col])),  // CO金額
      round1_(parseNum_(allData[OLD_ROW_CREDIT_CARD][col])),// クレカ
      round1_(parseNum_(allData[OLD_ROW_SHINPAN][col])),    // 信販
      String(allData[OLD_ROW_CBS][col] || '-'),              // CBS
      String(allData[OLD_ROW_LIFETY][col] || '-')            // ライフティ
    ]);
  }

  if (archiveRows.length > 0) {
    var lastRow = archiveSheet.getLastRow();
    archiveSheet.getRange(lastRow + 1, 1, archiveRows.length, ARCHIVE_HEADERS.length)
      .setValues(archiveRows);
    Logger.log(prevYear + '年' + prevMonth + '月のデータをアーカイブに保存 (' + archiveRows.length + '件)');
  }
}

/**
 * データが空の日別入力シートを修復（APIからも実行可能）
 */
function repairEmptyDailySheets() {
  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);
  var currentMonth = settings.month;

  var oldSheet = getSheetByMonth_(ss, currentMonth);
  if (!oldSheet) return { error: currentMonth + '月の旧シートなし' };

  var allData = oldSheet.getDataRange().getValues();
  var allFormulas = oldSheet.getDataRange().getFormulas();
  var maxCol = allData[0].length;

  var repaired = [];

  for (var i = 0; i < OLD_MEMBER_COLS.length; i++) {
    var col = OLD_MEMBER_COLS[i];
    var memberName = OLD_MEMBER_NAME_MAP[col];
    if (!memberName) continue;

    var v2Name = LEGACY_TO_V2_NAME[memberName] || memberName;
    var dailySheet = getDailySheet_(ss, v2Name) || getDailySheet_(ss, memberName);
    if (!dailySheet) continue;

    var carryover = dailySheet.getRange(DAILY_ROW_CARRYOVER, 2, 1, DAILY_COL_COUNT - 1).getValues()[0];
    var isEmpty = true;
    for (var c = 0; c < carryover.length; c++) {
      if (carryover[c] !== '' && carryover[c] !== null && carryover[c] !== 0) {
        isEmpty = false;
        break;
      }
    }

    if (!isEmpty) continue;

    migrateMonthlyTotals_(dailySheet, allData, col);
    migrateCBSLifety_(dailySheet, allData, col);
    migrateCreditShinpan_(dailySheet, allData, allFormulas, col);

    SpreadsheetApp.flush();

    migrateDailyData_(dailySheet, oldSheet, allData, col, v2Name, maxCol);

    repaired.push(v2Name);
  }

  if (repaired.length > 0) {
    SpreadsheetApp.flush();
    updateSummary();
  }

  return { repaired: repaired };
}

/**
 * 全日別シートのSUM数式を修復し、診断情報を返す
 * API: ?action=run&fn=fixDailyFormulas
 */
function fixDailyFormulas() {
  var ss = getSpreadsheet_();
  var sheets = ss.getSheets();
  var report = [];

  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name.indexOf(SHEET_DAILY_PREFIX) !== 0) continue;

    var sheet = sheets[i];
    var memberName = name.replace(SHEET_DAILY_PREFIX, '');

    // 行35のSUM数式を再設定
    var totalRow = ['合計'];
    for (var c = 1; c < DAILY_COL_COUNT; c++) {
      var colLetter = columnToLetter_(c + 1);
      totalRow.push('=SUM(' + colLetter + DAILY_ROW_CARRYOVER + ':' + colLetter + DAILY_ROW_DATA_END + ')');
    }
    sheet.getRange(DAILY_ROW_TOTAL, 1, 1, DAILY_COL_COUNT).setValues([totalRow]);

    // 診断: 行15の着金額(D列)の値を確認
    var row15Val = sheet.getRange(15, COL_REVENUE + 1).getValue();
    var row15Display = sheet.getRange(15, COL_REVENUE + 1).getDisplayValue();
    var row15Type = typeof row15Val;

    // 診断: 行35の着金額の数式結果を確認
    SpreadsheetApp.flush();
    var row35Val = sheet.getRange(DAILY_ROW_TOTAL, COL_REVENUE + 1).getValue();
    var row35Formula = sheet.getRange(DAILY_ROW_TOTAL, COL_REVENUE + 1).getFormula();

    report.push({
      member: memberName,
      row15: { value: row15Val, display: row15Display, type: row15Type },
      row35: { value: row35Val, formula: row35Formula }
    });
  }

  // サマリー更新
  SpreadsheetApp.flush();
  updateSummary();

  return report;
}

/**
 * 旧3月ウォーリアーズ数値シートから新日別入力シートへデータ同期
 * ラベル検索で行位置を動的に特定（シート再構成に対応）
 * API: ?action=run&fn=syncOldSheet
 */
function syncFromOldSheet() {
  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);
  var currentMonth = settings.month;

  var oldSheet = getSheetByMonth_(ss, currentMonth);
  if (!oldSheet) {
    return { error: currentMonth + '月の旧シートが見つかりません' };
  }

  // #REF!エラーを自動修正してからデータ読み取り
  var displayVals = oldSheet.getDataRange().getDisplayValues();
  var refFixed = 0;
  for (var ri = 0; ri < displayVals.length; ri++) {
    for (var ci = 0; ci < displayVals[ri].length; ci++) {
      var dv = displayVals[ri][ci];
      if (dv === '#REF!' || dv === '#NAME?' || dv === '#ERROR!' || dv === '#VALUE!') {
        oldSheet.getRange(ri + 1, ci + 1).setValue(0);
        refFixed++;
      }
    }
  }
  if (refFixed > 0) {
    SpreadsheetApp.flush();
    Logger.log('旧シート #REF!修正: ' + refFixed + '件');
  }

  // サマリー行（3-36行）の数式を修正してからデータ読み取り
  try {
    fixOldSheetFormulas();
  } catch (e) {
    Logger.log('旧シート数式修正スキップ: ' + e.message);
  }

  var allData = oldSheet.getDataRange().getValues();
  var allFormulas = oldSheet.getDataRange().getFormulas();
  var maxCol = allData[0].length;

  // === Step 1: ラベルで行インデックスを動的に検索 ===
  var rowMap = {};
  var labelPatterns = {
    'revenue': '合計着金額',
    'closed': '合計成約数',
    'currentRev': '当月着金額',
    'carryoverRev': '前月繰り越し',
    'credit': 'クレジットカード',
    'shinpan': '信販会社',
    'avgPrice': '平均単価',
    'sales': '合計売上',
    'closeRate': '全体の成約率',
    'cbs': 'CBS',
    'lifety': 'ライフティ',
    'deals': '合計商談数',
    'fundedDeals': '資金有りの合計商談数',
    'unfundedDeals': '資金なしの合計商談数',
    'fundedClosed': '資金有成約数',
    'unfundedClosed': '資金なし成約数',
    'coCount': '合計CO数',
    'coAmount': '合計CO金額',
    'ranking': 'ランキング'
  };

  for (var r = 0; r < Math.min(40, allData.length); r++) {
    var label = String(allData[r][0] || '').replace(/\s+/g, '').replace(/[\(\)（）]/g, '');
    if (!label) {
      label = String(allData[r][1] || '').replace(/\s+/g, '').replace(/[\(\)（）]/g, '');
    }
    for (var key in labelPatterns) {
      if (!rowMap[key] && label.indexOf(labelPatterns[key]) !== -1) {
        rowMap[key] = r;
      }
    }
  }

  // === Step 2: メンバー列を名前行から検索 ===
  // 名前行を検索（【xxx】パターンが複数ある行）
  var nameRowIdx = -1;
  for (var r = 0; r < Math.min(10, allData.length); r++) {
    var bracketCount = 0;
    for (var c = 0; c < allData[r].length; c++) {
      var v = String(allData[r][c] || '');
      if (v.indexOf('【') !== -1 && v.indexOf('】') !== -1) bracketCount++;
    }
    if (bracketCount >= 3) {
      nameRowIdx = r;
      break;
    }
  }

  if (nameRowIdx < 0) {
    return { error: '名前行が見つかりません', rowMap: rowMap };
  }

  // メンバー名→列番号のマッピングを構築
  var memberCols = {}; // { v2Name: colIndex }
  for (var c = 0; c < allData[nameRowIdx].length; c++) {
    var rawName = String(allData[nameRowIdx][c] || '').trim();
    if (!rawName) continue;

    // 【】を除去
    var cleanName = rawName.replace(/[【】\[\]]/g, '').trim();
    if (!cleanName) continue;

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
      // 部分一致で検索（「スクリプト通りに営業」→「スクリプト通りに営業するくん」等）
      for (var iconKey in ICON_MAP) {
        if (iconKey.indexOf(normalized) === 0 || normalized.indexOf(iconKey) === 0) {
          v2Name = iconKey;
          break;
        }
      }
    }

    memberCols[v2Name] = c;
  }

  // === Step 2.5: セクション行範囲を動的検出 ===
  var detectedSections = detectMemberSections_(oldSheet);
  var sectionsMap = {};
  for (var ds = 0; ds < detectedSections.length; ds++) {
    sectionsMap[detectedSections[ds].name] = detectedSections[ds];
  }

  // === Step 3: 各メンバーのデータを同期 ===
  var synced = [];

  for (var v2Name in memberCols) {
    var col = memberCols[v2Name];
    var dailySheet = getDailySheet_(ss, v2Name);
    if (!dailySheet) {
      Logger.log('日別入力シートなし: ' + v2Name + ' (col ' + col + ') → スキップ');
      continue;
    }

    // 各行からデータ取得（#REF!エラーは0として扱う）
    var revenue = rowMap.revenue !== undefined ? round1_(parseNum_(allData[rowMap.revenue][col])) : 0;
    var carryoverRev = rowMap.carryoverRev !== undefined ? round1_(parseNum_(allData[rowMap.carryoverRev][col])) : 0;
    var currentRev = rowMap.currentRev !== undefined ? round1_(parseNum_(allData[rowMap.currentRev][col])) : 0;
    var closed = rowMap.closed !== undefined ? parseNum_(allData[rowMap.closed][col]) : 0;
    var sales = rowMap.sales !== undefined ? round1_(parseNum_(allData[rowMap.sales][col])) : 0;
    var coAmount = rowMap.coAmount !== undefined ? round1_(parseNum_(allData[rowMap.coAmount][col])) : 0;
    var coCount = rowMap.coCount !== undefined ? parseNum_(allData[rowMap.coCount][col]) : 0;
    var fundedDeals = rowMap.fundedDeals !== undefined ? parseNum_(allData[rowMap.fundedDeals][col]) : 0;
    var unfundedDeals = rowMap.unfundedDeals !== undefined ? parseNum_(allData[rowMap.unfundedDeals][col]) : 0;
    var fundedClosed = rowMap.fundedClosed !== undefined ? parseNum_(allData[rowMap.fundedClosed][col]) : 0;
    var unfundedClosed = rowMap.unfundedClosed !== undefined ? parseNum_(allData[rowMap.unfundedClosed][col]) : 0;
    var deals = rowMap.deals !== undefined ? parseNum_(allData[rowMap.deals][col]) : 0;
    var creditVal = rowMap.credit !== undefined ? round1_(parseNum_(allData[rowMap.credit][col])) : 0;
    var shinpanVal = rowMap.shinpan !== undefined ? round1_(parseNum_(allData[rowMap.shinpan][col])) : 0;

    // 前月繰越行にセット（前月繰り越し着金額を直接使用）
    var carryoverData = [
      fundedDeals,     // B: 資金有商談数
      fundedClosed > 0 ? fundedClosed : closed,  // C: 資金有成約数
      carryoverRev,    // D: 着金額（前月繰り越し分のみ）
      sales,           // E: 売上
      coCount,         // F: CO数
      coAmount,        // G: CO金額
      0,               // H: 資金有で見送り
      0,               // I: 資金なし見送り数
      unfundedClosed,  // J: 資金なし成約数
      unfundedDeals    // K: 資金なし商談数
    ];
    dailySheet.getRange(DAILY_ROW_CARRYOVER, 2, 1, DAILY_COL_COUNT - 1)
      .setValues([carryoverData]);

    // CBS / ライフティ — 信販会社割合シートから直接取得に移行済み
    // （updateCBSLifetyFromShinpan_() で処理するため旧シートからの読み取りはスキップ）

    // クレカ/信販
    if (creditVal > 0) {
      dailySheet.getRange(DAILY_ROW_CREDIT_TOTAL, 2).setValue(creditVal);
    }
    if (shinpanVal > 0) {
      dailySheet.getRange(DAILY_ROW_SHINPAN_TOTAL, 2).setValue(shinpanVal);
    }

    SpreadsheetApp.flush();

    // 日別着金データを移行（動的検出セクションから読み取り）
    migrateDailyDataDirect_(dailySheet, allData, v2Name, sectionsMap);

    synced.push({
      member: v2Name,
      col: col,
      revenue: revenue,
      deals: deals,
      closed: closed,
      fundedClosed: fundedClosed,
      sales: sales,
      credit: creditVal,
      shinpan: shinpanVal,
      carryoverRev: carryoverRev,
      currentRev: currentRev
    });
  }

  // === ランキングを旧シートにRANK関数で書き戻し ===
  if (rowMap.ranking !== undefined && rowMap.revenue !== undefined && synced.length > 0) {
    var revenueSheetRow = rowMap.revenue + 1; // 1-based sheet row

    // メンバー列のセル参照リストを構築（例: C6,F6,L6,...）
    var refParts = [];
    for (var si = 0; si < synced.length; si++) {
      var colLetter = columnToLetter_(synced[si].col + 1); // 1-based col → A1表記
      refParts.push(colLetter + revenueSheetRow);
    }
    var refArray = '{' + refParts.join(',') + '}';

    // 各メンバーにRANK関数を設定
    for (var si = 0; si < synced.length; si++) {
      var memberColLetter = columnToLetter_(synced[si].col + 1);
      var formula = '=RANK(' + memberColLetter + revenueSheetRow + ',' + refArray + ',0)';
      oldSheet.getRange(rowMap.ranking + 1, synced[si].col + 1).setFormula(formula);
    }
    // ランキング行の背景色グラデーション（1位=明るい金色, 最下位=暗い）
    var rankSheetRow = rowMap.ranking + 1;
    var lastCol = synced[synced.length - 1].col + 1;
    var firstCol = synced[0].col + 1;
    var rankRange = oldSheet.getRange(rankSheetRow, firstCol, 1, lastCol - firstCol + 1);

    // 既存の条件付き書式からランキング行のルールを除去
    var existingRules = oldSheet.getConditionalFormatRules();
    var filteredRules = [];
    for (var ri = 0; ri < existingRules.length; ri++) {
      var ranges = existingRules[ri].getRanges();
      var isRankRow = false;
      for (var rri = 0; rri < ranges.length; rri++) {
        if (ranges[rri].getRow() === rankSheetRow && ranges[rri].getNumRows() === 1) {
          isRankRow = true;
          break;
        }
      }
      if (!isRankRow) filteredRules.push(existingRules[ri]);
    }

    // 個別セル着色: 1-3位=明るい, 下位ほど暗い
    var memberCount = synced.length;
    var rankColors = [
      { max: 1, bg: '#FEF3C7', font: '#000000' },  // 1位: 金色
      { max: 2, bg: '#FFF7ED', font: '#000000' },  // 2位: 薄オレンジ
      { max: 3, bg: '#FFF1E6', font: '#000000' },  // 3位: 薄ピーチ
      { max: 4, bg: '#E2E8F0', font: '#000000' },  // 4位: 薄グレー
      { max: 5, bg: '#CBD5E1', font: '#000000' },  // 5位
      { max: 6, bg: '#94A3B8', font: '#000000' },  // 6位
      { max: 7, bg: '#64748B', font: '#FFFFFF' },  // 7位
      { max: 8, bg: '#475569', font: '#FFFFFF' },  // 8位
      { max: 9, bg: '#334155', font: '#FFFFFF' },  // 9位
      { max: 99, bg: '#1E293B', font: '#FFFFFF' }  // 10位以下: 最暗
    ];
    oldSheet.setConditionalFormatRules(filteredRules);

    for (var si = 0; si < synced.length; si++) {
      var rankCell = oldSheet.getRange(rankSheetRow, synced[si].col + 1);
      var rankVal = parseInt(rankCell.getValue()) || memberCount;
      for (var ci = 0; ci < rankColors.length; ci++) {
        if (rankVal <= rankColors[ci].max) {
          rankCell.setBackground(rankColors[ci].bg);
          rankCell.setFontColor(rankColors[ci].font);
          break;
        }
      }
    }

    // === 38行目以降: 各メンバーセクション内にもランキングを設定 ===
    for (var si = 0; si < synced.length; si++) {
      var section = null;
      for (var mi = 0; mi < MEMBER_SECTIONS.length; mi++) {
        var secName = MEMBER_SECTIONS[mi].name;
        if (secName === synced[si].member ||
            synced[si].member.indexOf(secName) === 0 ||
            secName.indexOf(synced[si].member) === 0) {
          section = MEMBER_SECTIONS[mi];
          break;
        }
      }
      if (!section || section.dataStart < 0) continue;

      // セクションヘッダー付近で「着金額」列を検出
      var sectionRevCol = -1;
      var sectionHeaderIdx = -1;
      var searchStart = Math.max(0, section.dataStart - 4);
      var searchEnd = Math.min(allData.length - 1, section.dataStart);
      for (var r = searchStart; r <= searchEnd; r++) {
        for (var c = 0; c < allData[r].length; c++) {
          if (String(allData[r][c] || '').replace(/\s+/g, '') === '着金額') {
            sectionHeaderIdx = r;
            sectionRevCol = c;
            break;
          }
        }
        if (sectionHeaderIdx >= 0) break;
      }
      if (sectionHeaderIdx < 0) continue;

      // ランク行 = ヘッダーの2行上（例: AをAでやる → row 38付近）
      var sectionRankRow = sectionHeaderIdx - 2; // 0-indexed
      if (sectionRankRow < 0) continue;

      // RANK数式を着金額列+1に設定
      var memberColLetter = columnToLetter_(synced[si].col + 1);
      var formula = '=RANK(' + memberColLetter + revenueSheetRow + ',' + refArray + ',0)';
      oldSheet.getRange(sectionRankRow + 1, sectionRevCol + 2).setFormula(formula);
    }
  }

  // === CBS/ライフティを⚔️信販会社割合シートから取得し旧シートに書き戻し ===
  if ((rowMap.cbs !== undefined || rowMap.lifety !== undefined) && synced.length > 0) {
    var shinpanSheet = ss.getSheetByName(SHEET_SHINPAN);
    if (shinpanSheet) {
      var targetMonth = settings.month + '月';
      var spLastRow = shinpanSheet.getLastRow();
      var spLastCol = shinpanSheet.getLastColumn();
      var scanEnd = Math.min(spLastRow, 300);

      if (scanEnd >= 5) {
        var spData = shinpanSheet.getRange(1, 1, scanEnd, Math.min(spLastCol, 10)).getValues();

        // 対象月のライフティ/CBSセクション開始行を探す
        var lifetyStart = -1, cbsStart = -1;
        for (var r = 0; r < spData.length; r++) {
          var eVal = String(spData[r][4] || '').trim();
          if (eVal === targetMonth + 'ライフティ') lifetyStart = r;
          if (eVal === targetMonth + 'CBS') cbsStart = r;
        }

        // セクションからメンバーデータを読み取る
        var readShinpanSection = function(startRow) {
          var result = {};
          if (startRow < 0) return result;
          for (var r = startRow + 2; r < spData.length; r++) {
            var name = String(spData[r][4] || '').trim();
            if (!name) break;
            if (name.indexOf('月') !== -1) break;
            var approved = parseNum_(spData[r][5]);
            var applied = parseNum_(spData[r][7]);
            var v2Name = resolveRealName_(name);
            if (v2Name) {
              if (!result[v2Name]) result[v2Name] = { approved: 0, applied: 0 };
              result[v2Name].approved += approved;
              result[v2Name].applied += applied;
            }
          }
          return result;
        };

        var cbsShinpan = readShinpanSection(cbsStart);
        var lfShinpan = readShinpanSection(lifetyStart);

        var cbsTotalApproved = 0, cbsTotalApplied = 0;
        var lfTotalApproved = 0, lfTotalApplied = 0;

        for (var si = 0; si < synced.length; si++) {
          var memberCol = synced[si].col + 1; // 1-based sheet column
          var v2Name = synced[si].member;

          // CBS行書き込み
          if (rowMap.cbs !== undefined) {
            var cbsSheetRow = rowMap.cbs + 1;       // summary行 (1-based)
            var cbsDetailRow = rowMap.cbs + 2;       // detail行
            var cbs = cbsShinpan[v2Name] || { approved: 0, applied: 0 };
            var cbsStr = (cbs.approved > 0 || cbs.applied > 0)
              ? cbs.approved + '/' + cbs.applied : '-';
            oldSheet.getRange(cbsSheetRow, memberCol).setValue(cbsStr);
            oldSheet.getRange(cbsDetailRow, memberCol).setValue(cbs.approved);
            oldSheet.getRange(cbsDetailRow, memberCol + 1).setValue('/');
            oldSheet.getRange(cbsDetailRow, memberCol + 2).setValue(cbs.applied);
            cbsTotalApproved += cbs.approved;
            cbsTotalApplied += cbs.applied;
          }

          // ライフティ行書き込み
          if (rowMap.lifety !== undefined) {
            var lfSheetRow = rowMap.lifety + 1;
            var lfDetailRow = rowMap.lifety + 2;
            var lf = lfShinpan[v2Name] || { approved: 0, applied: 0 };
            var lfStr = (lf.approved > 0 || lf.applied > 0)
              ? lf.approved + '/' + lf.applied : '-';
            oldSheet.getRange(lfSheetRow, memberCol).setValue(lfStr);
            oldSheet.getRange(lfDetailRow, memberCol).setValue(lf.approved);
            oldSheet.getRange(lfDetailRow, memberCol + 1).setValue('/');
            oldSheet.getRange(lfDetailRow, memberCol + 2).setValue(lf.applied);
            lfTotalApproved += lf.approved;
            lfTotalApplied += lf.applied;
          }
        }

        // 合計列 (AG = column 33)
        var totalCol = 33;
        if (rowMap.cbs !== undefined) {
          var cbsTotalStr = (cbsTotalApproved > 0 || cbsTotalApplied > 0)
            ? cbsTotalApproved + '/' + cbsTotalApplied : '-';
          oldSheet.getRange(rowMap.cbs + 1, totalCol).setValue(cbsTotalStr);
          oldSheet.getRange(rowMap.cbs + 2, totalCol).setValue(cbsTotalApproved);
          oldSheet.getRange(rowMap.cbs + 2, totalCol + 1).setValue('/');
          oldSheet.getRange(rowMap.cbs + 2, totalCol + 2).setValue(cbsTotalApplied);
        }
        if (rowMap.lifety !== undefined) {
          var lfTotalStr = (lfTotalApproved > 0 || lfTotalApplied > 0)
            ? lfTotalApproved + '/' + lfTotalApplied : '-';
          oldSheet.getRange(rowMap.lifety + 1, totalCol).setValue(lfTotalStr);
          oldSheet.getRange(rowMap.lifety + 2, totalCol).setValue(lfTotalApproved);
          oldSheet.getRange(rowMap.lifety + 2, totalCol + 1).setValue('/');
          oldSheet.getRange(rowMap.lifety + 2, totalCol + 2).setValue(lfTotalApplied);
        }

        Logger.log('CBS/ライフティ書き戻し: CBS=' + cbsTotalApproved + '/' + cbsTotalApplied
          + ', LF=' + lfTotalApproved + '/' + lfTotalApplied);
      }
    }
  }

  // === 名前行(row4)にハイパーリンクを設定（各メンバーの入力欄へジャンプ） ===
  var gid = oldSheet.getSheetId();
  for (var si = 0; si < synced.length; si++) {
    var memberName = synced[si].member;
    var memberCol = synced[si].col + 1;
    var sec = null;
    for (var mi = 0; mi < MEMBER_SECTIONS.length; mi++) {
      if (MEMBER_SECTIONS[mi].name === memberName) { sec = MEMBER_SECTIONS[mi]; break; }
    }
    if (!sec || sec.dataStart < 0) continue;
    var nameCell = oldSheet.getRange(4, memberCol);
    var displayName = nameCell.getValue() || ('【' + memberName + '】');
    var linkUrl = '#gid=' + gid + '&range=A' + sec.dataStart;
    nameCell.setFormula('=HYPERLINK("' + linkUrl + '","' + displayName + '")');
  }

  SpreadsheetApp.flush();
  Logger.log('旧シート同期完了: ' + synced.length + '名');
  return { synced: synced, rowMap: rowMap, nameRow: nameRowIdx, memberCols: memberCols };
}

/**
 * 旧シート同期 + サマリー更新（API・メニュー用）
 */
function syncAndUpdate() {
  var result = syncFromOldSheet();
  updateSummary();
  return result;
}

/**
 * メンバー名リネーム（設定シート + 日別シート + サマリー）
 * API: ?action=run&fn=renameMember&from=ドライ&to=ポジティブ
 */
function renameMember(oldName, newName) {
  if (!oldName || !newName) return { error: 'from/to required' };
  var ss = getSpreadsheet_();
  var result = { renamed: [] };

  // 1. 設定シートのメンバー名・表示名を更新
  var settingsSheet = getSettingsSheet_(ss);
  if (settingsSheet) {
    var lastRow = Math.min(settingsSheet.getLastRow(), SETTINGS_ROW_GLOBAL_START - 2);
    for (var r = SETTINGS_ROW_DATA_START; r <= lastRow; r++) {
      var name = String(settingsSheet.getRange(r, 1).getValue()).trim();
      if (name === oldName) {
        settingsSheet.getRange(r, 1).setValue(newName);
        settingsSheet.getRange(r, 2).setValue(newName);
        result.renamed.push('設定シート row ' + r);
      }
    }
  }

  // 2. 日別入力シートをリネーム
  var oldSheetName = SHEET_DAILY_PREFIX + oldName;
  var newSheetName = SHEET_DAILY_PREFIX + newName;
  var dailySheet = ss.getSheetByName(oldSheetName);
  if (dailySheet) {
    dailySheet.setName(newSheetName);
    result.renamed.push(oldSheetName + ' → ' + newSheetName);
  }

  // 3. サマリーシートのヘッダーを更新
  var summarySheet = ss.getSheetByName(SHEET_SUMMARY);
  if (summarySheet) {
    var headers = summarySheet.getRange(1, 1, 1, summarySheet.getLastColumn()).getValues()[0];
    for (var c = 0; c < headers.length; c++) {
      if (String(headers[c]).trim() === oldName) {
        summarySheet.getRange(1, c + 1).setValue(newName);
        result.renamed.push('サマリー col ' + (c + 1));
      }
    }
  }

  SpreadsheetApp.flush();
  return result;
}

/**
 * 旧シート上のセルを修正（汎用）
 * API: ?action=run&fn=fixCell&gid=XXX&row=YY&col=ZZ&val=VALUE&formula=1
 */
function fixCell(gid, row, col, val, isFormula) {
  var ss = getSpreadsheet_();
  var sheets = ss.getSheets();
  var sheet = null;
  var gidNum = parseInt(gid);
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === gidNum) { sheet = sheets[i]; break; }
  }
  if (!sheet) return { error: 'gid not found' };
  var r = parseInt(row);
  var c = parseInt(col);
  var cell = sheet.getRange(r, c);
  if (isFormula === '1' || isFormula === true) {
    cell.setFormula(val);
  } else {
    var numVal = parseFloat(val);
    cell.setValue(isNaN(numVal) ? val : numVal);
  }
  SpreadsheetApp.flush();
  return { fixed: sheet.getName() + '!' + cell.getA1Notation(), value: cell.getValue() };
}

/**
 * 3月旧シートの集計行数式を修正
 * 各メンバーのセクション（rows 41-71, 77-107, ...）は統一列構造:
 *   col C(3)=資金有商談, F(6)=資金有成約, I(9)=着金額, L(12)=売上,
 *   O(15)=CO数, R(18)=CO金額, U(21)=見送り, X(24)=資金なし見送り,
 *   Y(25)=資金なし成約, Z(26)=資金なし商談
 *
 * セクション→メンバー対応（2月順 + ゴン）:
 *   S1(41-71)=AをAでやる, S2(77-107)=ドライ, S3(113-143)=ヒトコト,
 *   S4(149-179)=ビッグマウス, S5(186-216)=けつだん, S6(223-253)=ぜんぶり,
 *   S7(260-290)=スクリプト, S8(297-327)=ワントーン,
 *   S9(334-364)=トニー, S10(371-401)=ゴン
 *
 * API: ?action=run&fn=fixOldSheetFormulas
 */
function fixOldSheetFormulas() {
  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);
  var oldSheet = getSheetByMonth_(ss, settings.month);
  if (!oldSheet) return { error: 'シートなし' };

  // 動的セクション検出で正しいdataStart/dataEndを取得
  var memberSections = detectMemberSections_(oldSheet);

  // rowMapを構築
  var allData = oldSheet.getRange(1, 1, Math.min(35, oldSheet.getLastRow()), 2).getValues();
  var rowMap = {};
  for (var r = 0; r < allData.length; r++) {
    var label = String(allData[r][0] || '').replace(/\s+/g, '').replace(/[\(\)（）]/g, '');
    if (!label) label = String(allData[r][1] || '').replace(/\s+/g, '').replace(/[\(\)（）]/g, '');
    if (!label) continue;
    if (!rowMap.sales && label.indexOf('合計売上') !== -1) rowMap.sales = r + 1; // 1-indexed
    if (!rowMap.avgPrice && label.indexOf('平均単価') !== -1) rowMap.avgPrice = r + 1;
    if (!rowMap.paymentRate && label.indexOf('着金率') !== -1) rowMap.paymentRate = r + 1;
    if (!rowMap.closeRate && label.indexOf('成約率') !== -1) rowMap.closeRate = r + 1;
    if (!rowMap.deals && label.indexOf('合計商談数') !== -1 && label.indexOf('資金') === -1) rowMap.deals = r + 1;
    if (!rowMap.fundedDeals && label.indexOf('資金有') !== -1 && label.indexOf('商談') !== -1) rowMap.fundedDeals = r + 1;
    if (!rowMap.revenue && label.indexOf('合計着金額') !== -1) rowMap.revenue = r + 1;
    if (!rowMap.closed && label.indexOf('成約数') !== -1) rowMap.closed = r + 1;
    if (!rowMap.fundedClosed && label.indexOf('資金有成約数') !== -1) rowMap.fundedClosed = r + 1;
    if (!rowMap.unfundedClosed && label.indexOf('資金なし成約数') !== -1) rowMap.unfundedClosed = r + 1;
    if (!rowMap.unfundedDeals && label.indexOf('資金なしの合計商談数') !== -1) rowMap.unfundedDeals = r + 1;
    if (!rowMap.coCount && label.indexOf('CO数') !== -1) rowMap.coCount = r + 1;
    if (!rowMap.coAmount && (label.indexOf('CO額') !== -1 || label.indexOf('CO金額') !== -1)) rowMap.coAmount = r + 1;
    if (!rowMap.currentRev && label.indexOf('当月着金額') !== -1) rowMap.currentRev = r + 1;
    if (!rowMap.carryoverRev && label.indexOf('前月繰り越し') !== -1) rowMap.carryoverRev = r + 1;
  }

  // セクション内の列マッピング（全セクション共通）
  var SEC_COL_FUNDED_DEALS   = 3;  // C
  var SEC_COL_FUNDED_CLOSED  = 6;  // F
  var SEC_COL_REVENUE        = 9;  // I
  var SEC_COL_SALES          = 12; // L
  var SEC_COL_CO_COUNT       = 15; // O
  var SEC_COL_CO_AMOUNT      = 18; // R
  var SEC_COL_UNFUNDED_PASSED= 24; // X
  var SEC_COL_UNFUNDED_CLOSED= 25; // Y
  var SEC_COL_UNFUNDED_DEALS = 26; // Z

  var fixed = [];
  var carryoverRow = 36; // 繰り越し売り上げ行

  for (var i = 0; i < memberSections.length; i++) {
    var ms = memberSections[i];
    if (ms.dataStart < 0) continue; // セクションなしのメンバーはスキップ

    var sc = ms.summaryCol;
    var scLetter = columnToLetter_(sc);
    var ds = ms.dataStart;
    var de = ms.dataEnd;
    var sr = ms.sumRow;

    // col letters for section columns
    var salesColLetter = columnToLetter_(SEC_COL_SALES);       // L
    var revColLetter = columnToLetter_(SEC_COL_REVENUE);       // I
    var fundedDealsLetter = columnToLetter_(SEC_COL_FUNDED_DEALS); // C
    var fundedClosedLetter = columnToLetter_(SEC_COL_FUNDED_CLOSED); // F
    var unfundedDealsLetter = columnToLetter_(SEC_COL_UNFUNDED_DEALS); // Z
    var unfundedClosedLetter = columnToLetter_(SEC_COL_UNFUNDED_CLOSED); // Y
    var coCountLetter = columnToLetter_(SEC_COL_CO_COUNT);     // O
    var coAmountLetter = columnToLetter_(SEC_COL_CO_AMOUNT);   // R

    // 行13: 合計売上 = SUM(セクション売上列) + 繰り越し
    if (rowMap.sales) {
      var formula = '=SUM(' + salesColLetter + ds + ':' + salesColLetter + de + ')+' + scLetter + carryoverRow;
      oldSheet.getRange(rowMap.sales, sc).setFormula(formula);
      fixed.push(ms.name + ' sales: ' + formula);
    }

    // 行12: 平均単価 = 合計売上 / 成約数 (if成約数>0)
    if (rowMap.avgPrice && rowMap.sales && rowMap.closed) {
      var formula = '=IF(' + scLetter + rowMap.closed + '=0,"",'+scLetter + rowMap.sales + '/' + scLetter + rowMap.closed + ')';
      oldSheet.getRange(rowMap.avgPrice, sc).setFormula(formula);
    }

    // 行14: 着金率 = 合計着金額 / 合計売上 (if売上>0)
    if (rowMap.paymentRate && rowMap.revenue && rowMap.sales) {
      var formula = '=IF(' + scLetter + rowMap.sales + '=0,0,' + scLetter + rowMap.revenue + '/' + scLetter + rowMap.sales + ')';
      oldSheet.getRange(rowMap.paymentRate, sc).setFormula(formula);
      fixed.push(ms.name + ' paymentRate: ' + formula);
    }

    // 行15: 成約率 = (資金有成約+資金なし成約) / (資金有商談+資金なし商談) (if商談>0)
    if (rowMap.closeRate && rowMap.fundedClosed && rowMap.unfundedClosed && rowMap.fundedDeals && rowMap.unfundedDeals) {
      var numr = '(' + scLetter + rowMap.fundedClosed + '+' + scLetter + rowMap.unfundedClosed + ')';
      var denom = '(' + scLetter + rowMap.fundedDeals + '+' + scLetter + rowMap.unfundedDeals + ')';
      var formula = '=IF(' + denom + '=0,0,' + numr + '/' + denom + ')';
      oldSheet.getRange(rowMap.closeRate, sc).setFormula(formula);
    }

    // 行8: 資金有商談数 = SUM(セクション資金有商談列)
    if (rowMap.fundedDeals) {
      var formula = '=SUM(' + fundedDealsLetter + ds + ':' + fundedDealsLetter + de + ')';
      oldSheet.getRange(rowMap.fundedDeals, sc).setFormula(formula);
    }

    // 行21: 合計商談数 = 資金有商談 + 資金なし商談
    if (rowMap.deals && rowMap.fundedDeals && rowMap.unfundedDeals) {
      var formula = '=' + scLetter + rowMap.fundedDeals + '+' + scLetter + rowMap.unfundedDeals;
      oldSheet.getRange(rowMap.deals, sc).setFormula(formula);
    }

    // 資金有成約数
    if (rowMap.fundedClosed) {
      var formula = '=SUM(' + fundedClosedLetter + ds + ':' + fundedClosedLetter + de + ')';
      oldSheet.getRange(rowMap.fundedClosed, sc).setFormula(formula);
    }
    // 資金なし商談数
    if (rowMap.unfundedDeals) {
      var formula = '=SUM(' + unfundedDealsLetter + ds + ':' + unfundedDealsLetter + de + ')';
      oldSheet.getRange(rowMap.unfundedDeals, sc).setFormula(formula);
    }
    // 資金なし成約数
    if (rowMap.unfundedClosed) {
      var formula = '=SUM(' + unfundedClosedLetter + ds + ':' + unfundedClosedLetter + de + ')';
      oldSheet.getRange(rowMap.unfundedClosed, sc).setFormula(formula);
    }
    // CO数
    if (rowMap.coCount) {
      var formula = '=SUM(' + coCountLetter + ds + ':' + coCountLetter + de + ')';
      oldSheet.getRange(rowMap.coCount, sc).setFormula(formula);
    }
    // CO金額
    if (rowMap.coAmount) {
      var formula = '=SUM(' + coAmountLetter + ds + ':' + coAmountLetter + de + ')';
      oldSheet.getRange(rowMap.coAmount, sc).setFormula(formula);
    }

    // 成約数（row 7）= 資金有成約 + 資金なし成約
    if (rowMap.closed && rowMap.fundedClosed && rowMap.unfundedClosed) {
      var formula = '=' + scLetter + rowMap.fundedClosed + '+' + scLetter + rowMap.unfundedClosed;
      oldSheet.getRange(rowMap.closed, sc).setFormula(formula);
    }

    // 当月着金額 = SUM(セクション着金額列)
    if (rowMap.currentRev) {
      var formula = '=SUM(' + revColLetter + ds + ':' + revColLetter + de + ')';
      oldSheet.getRange(rowMap.currentRev, sc).setFormula(formula);
      fixed.push(ms.name + ' currentRev: ' + formula);
    }

    // 合計着金額 = 当月着金額 + 前月繰り越し - CO額
    if (rowMap.revenue) {
      var formula;
      if (rowMap.currentRev && rowMap.carryoverRev) {
        formula = '=' + scLetter + rowMap.currentRev + '+' + scLetter + rowMap.carryoverRev;
      } else {
        formula = '=SUM(' + revColLetter + ds + ':' + revColLetter + de + ')';
        if (rowMap.carryoverRev) {
          formula += '+' + scLetter + rowMap.carryoverRev;
        }
      }
      // CO額を引く
      if (rowMap.coAmount) {
        formula += '-' + scLetter + rowMap.coAmount;
      }
      oldSheet.getRange(rowMap.revenue, sc).setFormula(formula);
      fixed.push(ms.name + ' revenue: ' + formula);
    }
  }

  // === Row 5 着金ランキング（RANK数式）を修正 ===
  if (rowMap.revenue) {
    // 全メンバーのrevenue行セル参照リスト
    var rankRefs = [];
    for (var ri = 0; ri < memberSections.length; ri++) {
      rankRefs.push(columnToLetter_(memberSections[ri].summaryCol) + rowMap.revenue);
    }
    var rankArray = '{' + rankRefs.join(',') + '}';
    var rankRow = rowMap.revenue - 1; // ランキング行 = 着金額行の1つ上(row 5)
    for (var ri = 0; ri < memberSections.length; ri++) {
      var sc2 = memberSections[ri].summaryCol;
      var cellRef = columnToLetter_(sc2) + rowMap.revenue;
      var formula = '=RANK(' + cellRef + ',' + rankArray + ',0)';
      oldSheet.getRange(rankRow, sc2).setFormula(formula);
    }
    // ランキング行に順位色をつける
    var RANK_COLORS = {
      1:  { bg: '#FEF3C7', fg: '#92400E' },  // 金
      2:  { bg: '#F1F5F9', fg: '#475569' },  // 銀
      3:  { bg: '#FFF7ED', fg: '#9A3412' },  // 銅
      4:  { bg: '#DBEAFE', fg: '#1E40AF' },  // 青
      5:  { bg: '#D1FAE5', fg: '#065F46' },  // 緑
      6:  { bg: '#EDE9FE', fg: '#5B21B6' },  // 紫
      7:  { bg: '#FCE7F3', fg: '#9D174D' },  // ピンク
      8:  { bg: '#FEF9C3', fg: '#854D0E' },  // 黄
      9:  { bg: '#FFE4E6', fg: '#9F1239' },  // 赤
      10: { bg: '#F3F4F6', fg: '#6B7280' },  // グレー
    };
    SpreadsheetApp.flush(); // RANK数式を確定させてから値を取得
    for (var ri = 0; ri < memberSections.length; ri++) {
      var sc3 = memberSections[ri].summaryCol;
      var rankCell = oldSheet.getRange(rankRow, sc3);
      var rankVal = parseInt(rankCell.getValue()) || 0;
      var rc = RANK_COLORS[rankVal] || RANK_COLORS[10];
      rankCell.setBackground(rc.bg)
             .setFontColor(rc.fg)
             .setFontWeight('bold')
             .setFontSize(14)
             .setHorizontalAlignment('center');
    }
    // AG列（合計列）のランキング行
    oldSheet.getRange(rankRow, 33).setBackground('#F9FAFB').setFontColor('#9CA3AF').setFontSize(10);
    fixed.push('Row ' + rankRow + ' (ranking): RANK formulas + colors for all members');
  }

  // === AG列（合計）を修正 ===
  var totalCol = 33; // AG = column 33
  // メンバーのsummaryCol一覧からSUM用の参照を構築
  var memberColRefs = [];
  for (var mi = 0; mi < memberSections.length; mi++) {
    memberColRefs.push(columnToLetter_(memberSections[mi].summaryCol));
  }

  // 単純合計行: SUM of each member's summaryCol
  var simpleRows = ['revenue', 'closed', 'currentRev', 'carryoverRev', 'sales',
                    'fundedDeals', 'fundedClosed', 'unfundedDeals', 'unfundedClosed',
                    'coCount', 'coAmount'];
  for (var si = 0; si < simpleRows.length; si++) {
    var rowNum = rowMap[simpleRows[si]];
    if (!rowNum) continue;
    var parts = [];
    for (var ci = 0; ci < memberColRefs.length; ci++) {
      parts.push(memberColRefs[ci] + rowNum);
    }
    oldSheet.getRange(rowNum, totalCol).setFormula('=' + parts.join('+'));
    fixed.push('AG' + rowNum + ' (' + simpleRows[si] + '): SUM');
  }

  // 合計商談数 (3列パターン: 当月のみ合算)
  if (rowMap.deals) {
    var parts = [];
    for (var ci = 0; ci < memberColRefs.length; ci++) {
      parts.push(memberColRefs[ci] + rowMap.deals);
    }
    oldSheet.getRange(rowMap.deals, totalCol).setFormula('=' + parts.join('+'));
    fixed.push('AG' + rowMap.deals + ' (deals): SUM');
  }

  // 平均単価 = 合計売上 / 合計成約数
  if (rowMap.avgPrice && rowMap.sales && rowMap.closed) {
    var f = '=IF(AG' + rowMap.closed + '=0,0,AG' + rowMap.sales + '/AG' + rowMap.closed + ')';
    oldSheet.getRange(rowMap.avgPrice, totalCol).setFormula(f);
    fixed.push('AG' + rowMap.avgPrice + ' (avgPrice): sales/closed');
  }

  // 着金率 = 合計着金額 / 合計売上
  if (rowMap.paymentRate && rowMap.revenue && rowMap.sales) {
    var f = '=IF(AG' + rowMap.sales + '=0,0,AG' + rowMap.revenue + '/AG' + rowMap.sales + ')';
    oldSheet.getRange(rowMap.paymentRate, totalCol).setFormula(f);
    oldSheet.getRange(rowMap.paymentRate, totalCol).setNumberFormat('0.0%');
    fixed.push('AG' + rowMap.paymentRate + ' (paymentRate): revenue/sales %');
  }

  // 全体の成約率 = 合計成約数 / 合計商談数
  if (rowMap.closeRate && rowMap.closed && rowMap.deals) {
    var f = '=IF(AG' + rowMap.deals + '=0,0,AG' + rowMap.closed + '/AG' + rowMap.deals + ')';
    oldSheet.getRange(rowMap.closeRate, totalCol).setFormula(f);
    oldSheet.getRange(rowMap.closeRate, totalCol).setNumberFormat('0.0%');
    fixed.push('AG' + rowMap.closeRate + ' (closeRate): closed/deals %');
  }

  // 資金ありの成約率
  if (rowMap.fundedDeals && rowMap.fundedClosed) {
    // closeRateの次の行を想定（row 16）
    var fundedRateRow = rowMap.closeRate ? rowMap.closeRate + 1 : 0;
    if (fundedRateRow > 0) {
      var f = '=IF(AG' + rowMap.fundedDeals + '=0,0,AG' + rowMap.fundedClosed + '/AG' + rowMap.fundedDeals + ')';
      oldSheet.getRange(fundedRateRow, totalCol).setFormula(f);
      oldSheet.getRange(fundedRateRow, totalCol).setNumberFormat('0.0%');
      fixed.push('AG' + fundedRateRow + ' (fundedCloseRate): %');
    }
  }

  // === 先月商談数の参照を修正（deals行の summaryCol+1 に前月シートの合計商談数を参照） ===
  if (rowMap.deals) {
    var prevMonth = settings.month === 1 ? 12 : settings.month - 1;
    var prevSheet = getSheetByMonth_(ss, prevMonth);
    if (prevSheet) {
      // 前月シートの名前行と合計商談数行を検出
      var prevScanRows = Math.min(50, prevSheet.getLastRow());
      var prevScanCols = Math.min(50, prevSheet.getLastColumn());
      var prevData = prevSheet.getRange(1, 1, prevScanRows, prevScanCols).getValues();

      // 前月の合計商談数行を検索
      var prevDealsRow = -1;
      for (var pr = 0; pr < prevData.length; pr++) {
        var pLabel = String(prevData[pr][0] || '').replace(/\s+/g, '').replace(/[\(\)（）]/g, '');
        if (!pLabel) pLabel = String(prevData[pr][1] || '').replace(/\s+/g, '').replace(/[\(\)（）]/g, '');
        if (pLabel.indexOf('合計商談数') !== -1 && pLabel.indexOf('資金') === -1) {
          prevDealsRow = pr;
          break;
        }
      }

      // 前月の名前行を検索
      var prevNameRow = -1;
      for (var pr = 0; pr < Math.min(15, prevData.length); pr++) {
        var knownCount = 0;
        for (var pc = 0; pc < prevData[pr].length; pc++) {
          var pv = String(prevData[pr][pc] || '').trim().replace(/[【】\[\]]/g, '');
          if (LEGACY_TO_V2_NAME[pv] || ICON_MAP[pv] || OLD_DISPLAY_TO_V2[pv]) knownCount++;
        }
        if (knownCount >= 3) { prevNameRow = pr; break; }
      }

      if (prevDealsRow >= 0 && prevNameRow >= 0) {
        var prevSheetName = prevSheet.getName();
        // 前月の名前→列マップ
        var prevNameCols = {};
        for (var pc = 0; pc < prevData[prevNameRow].length; pc++) {
          var rawN = String(prevData[prevNameRow][pc] || '').trim().replace(/[【】\[\]]/g, '');
          if (!rawN) continue;
          var normN = NAME_MAP[rawN] || rawN;
          var v2N = LEGACY_TO_V2_NAME[normN] || DISPLAY_NAME_MAP[normN] || OLD_DISPLAY_TO_V2[normN] || normN;
          if (ICON_MAP[v2N]) prevNameCols[v2N] = pc + 1; // 1-indexed
        }

        for (var mi = 0; mi < memberSections.length; mi++) {
          var msi = memberSections[mi];
          var prevCol = prevNameCols[msi.name];
          if (!prevCol) continue;

          var prevColLetter = columnToLetter_(prevCol);
          var prevDealsRowNum = prevDealsRow + 1; // 1-indexed
          // 先月列 = summaryCol + 1
          var formula = "='" + prevSheetName + "'!" + prevColLetter + prevDealsRowNum;
          oldSheet.getRange(rowMap.deals, msi.summaryCol + 1).setFormula(formula);
          fixed.push(msi.name + ' prevDeals: ' + formula);

          // 対比列 = summaryCol + 2
          var dealsLetter = columnToLetter_(msi.summaryCol);
          var prevLetter = columnToLetter_(msi.summaryCol + 1);
          var ratioFormula = '=IF(' + prevLetter + rowMap.deals + '=0,0,' + dealsLetter + rowMap.deals + '/' + prevLetter + rowMap.deals + ')';
          oldSheet.getRange(rowMap.deals, msi.summaryCol + 2).setFormula(ratioFormula);
        }
      }
    }
  }

  SpreadsheetApp.flush();
  return { status: 'fixed', rowMap: rowMap, fixed: fixed };
}

/**
 * CBS/ライフティ周辺のデータをデバッグ出力
 */
function debugCBSLifety() {
  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);
  var oldSheet = getSheetByMonth_(ss, settings.month);
  if (!oldSheet) return { error: 'シートなし' };

  var allData = oldSheet.getDataRange().getValues();
  var result = { rows: {} };

  // rowMap再構築
  var labelPatterns = { 'cbs': 'CBS', 'lifety': 'ライフティ' };
  var rowMap = {};
  for (var r = 0; r < Math.min(40, allData.length); r++) {
    var label = String(allData[r][0] || '').replace(/\s+/g, '');
    if (!label) label = String(allData[r][1] || '').replace(/\s+/g, '');
    for (var key in labelPatterns) {
      if (!rowMap[key] && label.indexOf(labelPatterns[key]) !== -1) {
        rowMap[key] = r;
      }
    }
  }
  result.rowMap = rowMap;

  // ヒトコト=col11, CBS/ライフティ周辺のrow 15-22のデータ
  var memberCols = [2, 5, 11, 14, 17, 20, 23, 26, 29];
  var dumpRows = [];
  for (var r = 14; r < Math.min(25, allData.length); r++) {
    var rowData = { row: r, label: String(allData[r][0] || allData[r][1] || '').substring(0, 30) };
    for (var mi = 0; mi < memberCols.length; mi++) {
      var c = memberCols[mi];
      var v = allData[r][c];
      if (v !== '' && v !== null && v !== undefined) {
        rowData['c' + c] = String(v).substring(0, 30);
      }
      // col+1, col+2 も確認
      for (var offset = 1; offset <= 2; offset++) {
        var vc = allData[r][c + offset];
        if (vc !== '' && vc !== null && vc !== undefined) {
          rowData['c' + (c + offset)] = String(vc).substring(0, 30);
        }
      }
    }
    dumpRows.push(rowData);
  }
  result.rows = dumpRows;

  // 日別入力シートの現在の値も確認
  var dailySheet = getDailySheet_(ss, 'ヒトコト');
  if (dailySheet) {
    result.dailyHitokoto = {
      cbsApproved: String(dailySheet.getRange(DAILY_ROW_CBS_APPROVED, 2).getValue()),
      cbsApplied: String(dailySheet.getRange(DAILY_ROW_CBS_APPLIED, 2).getValue()),
      lfApproved: String(dailySheet.getRange(DAILY_ROW_LF_APPROVED, 2).getValue()),
      lfApplied: String(dailySheet.getRange(DAILY_ROW_LF_APPLIED, 2).getValue())
    };
  }

  // ヒトコト(col11)のlifety行の数式も確認
  var lfRow = (rowMap.lifety || 18) + 1;  // data row
  var formulas = oldSheet.getRange(lfRow + 1, 12, 1, 3).getFormulas();
  var displayVals = oldSheet.getRange(lfRow + 1, 12, 1, 3).getDisplayValues();
  var rawVals = oldSheet.getRange(lfRow + 1, 12, 1, 3).getValues();
  result.hitokotoLifetyRow = {
    sheetRow: lfRow + 1,
    col11_formula: formulas[0][0],
    col12_formula: formulas[0][1],
    col13_formula: formulas[0][2],
    col11_display: displayVals[0][0],
    col12_display: displayVals[0][1],
    col13_display: displayVals[0][2],
    col11_raw: String(rawVals[0][0]),
    col12_raw: String(rawVals[0][1]),
    col13_raw: String(rawVals[0][2])
  };

  // 全メンバーのlifety現在値
  var activeMembers = getActiveMembers_(ss);
  var allLifety = {};
  for (var i = 0; i < activeMembers.length; i++) {
    var m = activeMembers[i];
    var ds = getDailySheet_(ss, m.name);
    if (ds) {
      allLifety[m.displayName] = {
        approved: String(ds.getRange(DAILY_ROW_LF_APPROVED, 2).getValue()),
        applied: String(ds.getRange(DAILY_ROW_LF_APPLIED, 2).getValue())
      };
    }
  }
  result.allMembersLifety = allLifety;

  // 参照先シート「⚔️信販会社割合」の周辺データ確認
  var shinpanSheet = ss.getSheetByName('⚔️信販会社割合');
  if (shinpanSheet) {
    var lastRow = shinpanSheet.getLastRow();
    result.shinpanSheet = { totalRows: lastRow };
  }

  return result;
}

/**
 * 信販会社割合シートのE-J列を広範囲スキャンしてCBS/ライフティの構造を把握
 */
function debugShinpanEJ() {
  var ss = getSpreadsheet_();
  var shinpanSheet = ss.getSheetByName('⚔️信販会社割合');
  if (!shinpanSheet) return { error: 'シートなし' };

  var lastRow = shinpanSheet.getLastRow();
  var lastCol = shinpanSheet.getLastColumn();
  var result = { totalRows: lastRow, totalCols: lastCol, sections: [] };

  // A-J列 (1-10) をスキャン。「3月」「CBS」「ライフティ」を含むセクションを探す
  var settings = getGlobalSettings_(ss);
  var targetMonth = settings.month + '月';
  var scanEnd = Math.min(lastRow, 300);
  var allData = shinpanSheet.getRange(1, 1, scanEnd, Math.min(lastCol, 10)).getValues();

  // キーワードを含む行を抽出
  var keyRows = [];
  for (var r = 0; r < allData.length; r++) {
    var rowText = '';
    for (var c = 0; c < allData[r].length; c++) {
      rowText += String(allData[r][c] || '') + ' ';
    }
    if (rowText.indexOf('3月') !== -1 || rowText.indexOf('CBS') !== -1 ||
        rowText.indexOf('ライフ') !== -1 || rowText.indexOf('承認') !== -1 ||
        rowText.indexOf('否決') !== -1 || rowText.indexOf('成約') !== -1 ||
        rowText.indexOf('月') !== -1 && rowText.indexOf('割合') !== -1) {
      var rowObj = { row: r + 1 };
      for (var c = 0; c < allData[r].length; c++) {
        var v = allData[r][c];
        if (v !== '' && v !== null && v !== undefined) {
          rowObj[String.fromCharCode(65 + c)] = String(v).substring(0, 40);
        }
      }
      keyRows.push(rowObj);
    }
  }
  result.keyRows = keyRows;

  // ヘッダー行（row 1）
  var headerRow = {};
  for (var c = 0; c < allData[0].length; c++) {
    var v = allData[0][c];
    if (v !== '' && v !== null && v !== undefined) {
      headerRow[String.fromCharCode(65 + c)] = String(v).substring(0, 50);
    }
  }
  result.headerRow = headerRow;

  // 対象月のライフティ/CBSセクションを取得
  var lifetyStart = -1, cbsStart = -1;
  for (var r = 0; r < allData.length; r++) {
    var eVal = String(allData[r][4] || '');
    if (eVal === targetMonth + 'ライフティ') lifetyStart = r;
    if (eVal === targetMonth + 'CBS') cbsStart = r;
  }

  result.targetMonth = targetMonth;
  result.lifetyStartRow = lifetyStart + 1;
  result.cbsStartRow = cbsStart + 1;

  // ライフティセクションのデータ行を取得（ヘッダーの次の行からE列が空or次セクションまで）
  var lifetyData = [];
  if (lifetyStart >= 0) {
    for (var r = lifetyStart + 2; r < allData.length; r++) {
      var eName = String(allData[r][4] || '').trim();
      if (!eName) {
        // E列空 but F列にデータがある場合は合計行の可能性
        var fVal = allData[r][5];
        if (fVal !== '' && fVal !== null && fVal !== undefined) {
          lifetyData.push({ row: r + 1, E: '(合計?)', F: String(fVal), G: String(allData[r][6] || ''), H: String(allData[r][7] || '') });
        }
        break;
      }
      if (eName.indexOf('月') !== -1 && eName.indexOf('CBS') !== -1) break;  // 次セクション
      lifetyData.push({
        row: r + 1,
        E: eName,
        F: String(allData[r][5] || ''),
        G: String(allData[r][6] || ''),
        H: String(allData[r][7] || ''),
        I: String(allData[r][8] || ''),
        J: String(allData[r][9] || '')
      });
    }
  }
  result.lifetyData = lifetyData;

  // CBSセクションのデータ行を取得
  var cbsData = [];
  if (cbsStart >= 0) {
    for (var r = cbsStart + 2; r < allData.length; r++) {
      var eName = String(allData[r][4] || '').trim();
      if (!eName) {
        var fVal = allData[r][5];
        if (fVal !== '' && fVal !== null && fVal !== undefined) {
          cbsData.push({ row: r + 1, E: '(合計?)', F: String(fVal), G: String(allData[r][6] || ''), H: String(allData[r][7] || '') });
        }
        break;
      }
      if (eName.indexOf('月') !== -1) break;
      cbsData.push({
        row: r + 1,
        E: eName,
        F: String(allData[r][5] || ''),
        G: String(allData[r][6] || ''),
        H: String(allData[r][7] || ''),
        I: String(allData[r][8] || ''),
        J: String(allData[r][9] || '')
      });
    }
  }
  result.cbsData = cbsData;

  return result;
}

/**
 * 旧シートの構造を診断（どの行・列にデータがあるか調査）
 */
function debugOldSheet() {
  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);
  var oldSheet = getSheetByMonth_(ss, settings.month);
  if (!oldSheet) return { error: 'シートなし' };

  var allData = oldSheet.getDataRange().getValues();
  var result = {
    sheetName: oldSheet.getName(),
    totalRows: allData.length,
    totalCols: allData[0].length
  };

  // 行1-20のB列(index 1)の値を取得（行ラベル把握用）
  var rowLabels = [];
  for (var r = 0; r < Math.min(25, allData.length); r++) {
    var cols = [];
    for (var c = 0; c < Math.min(25, allData[r].length); c++) {
      var v = allData[r][c];
      if (v === '' || v === null || v === undefined) continue;
      cols.push({ col: c, val: String(v).substring(0, 30) });
    }
    if (cols.length > 0) {
      rowLabels.push({ row: r, data: cols });
    }
  }
  result.first25Rows = rowLabels;

  // OLD_MEMBER_COLS の各列で行5(名前行)の値を確認
  var memberCheck = {};
  for (var i = 0; i < OLD_MEMBER_COLS.length; i++) {
    var col = OLD_MEMBER_COLS[i];
    memberCheck['col' + col] = {
      row5: String(allData[4] ? allData[4][col] || '' : ''),
      row6: String(allData[5] ? allData[5][col] || '' : ''),
      row14: String(allData[13] ? allData[13][col] || '' : ''),
      row15: String(allData[14] ? allData[14][col] || '' : '')
    };
  }
  result.memberColCheck = memberCheck;

  // 「着金」「売上」「成約」等のラベルを含む行を検索
  var keyRows = {};
  for (var r = 0; r < Math.min(30, allData.length); r++) {
    for (var c = 0; c < Math.min(5, allData[r].length); c++) {
      var label = String(allData[r][c] || '').replace(/\s+/g, '');
      if (label.indexOf('着金') !== -1 || label.indexOf('売上') !== -1 ||
          label.indexOf('成約') !== -1 || label.indexOf('商談') !== -1 ||
          label.indexOf('ランキング') !== -1 || label.indexOf('CO') !== -1 ||
          label.indexOf('クレカ') !== -1 || label.indexOf('信販') !== -1 ||
          label.indexOf('CBS') !== -1 || label.indexOf('ライフ') !== -1) {
        keyRows['row' + r + '_col' + c] = label;
      }
    }
  }
  result.keyLabels = keyRows;

  return result;
}

/**
 * 指定シートでA-B列固定を阻害する結合セル（B列以前〜C列以降にまたがるもの）を検出
 */
function debugFreezeBlockers() {
  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);
  var oldSheet = getSheetByMonth_(ss, settings.month);
  if (!oldSheet) return { error: 'シートなし' };

  var maxRow = oldSheet.getMaxRows();
  var maxCol = oldSheet.getMaxColumns();

  // シート全体の結合セルを取得
  var allMerges = oldSheet.getRange(1, 1, maxRow, maxCol).getMergedRanges();
  var blockers = [];

  for (var i = 0; i < allMerges.length; i++) {
    var r = allMerges[i];
    var startCol = r.getColumn();     // 1-based
    var endCol = r.getColumn() + r.getNumColumns() - 1;

    // B列(2)以前に始まり、C列(3)以降に広がる結合セル
    if (startCol <= 2 && endCol >= 3) {
      blockers.push({
        range: r.getA1Notation(),
        row: r.getRow(),
        startCol: startCol,
        endCol: endCol,
        value: String(r.getCell(1, 1).getValue() || '').substring(0, 40)
      });
    }
  }

  return {
    sheet: oldSheet.getName(),
    totalMerges: allMerges.length,
    freezeBlockers: blockers
  };
}

/**
 * 旧シートを非表示にする
 */
function hideOldSheets() {
  var ss = getSpreadsheet_();
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert(
    '旧シート非表示',
    '旧ウォーリアーズ数値シートを非表示にします。\nシートは削除されません。続行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  var sheets = ss.getSheets();
  var hidden = [];
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name.indexOf(SHEET_NAME_PATTERN) !== -1) {
      sheets[i].hideSheet();
      hidden.push(name);
    }
  }

  if (hidden.length > 0) {
    ui.alert('完了', '以下のシートを非表示にしました:\n' + hidden.join('\n'), ui.ButtonSet.OK);
  } else {
    ui.alert('該当シートなし', '非表示にするシートが見つかりませんでした。', ui.ButtonSet.OK);
  }
}

/**
 * 非表示にした旧シートを再表示する
 */
function showOldSheets() {
  var ss = getSpreadsheet_();
  var sheets = ss.getSheets();
  var shown = [];
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name.indexOf(SHEET_NAME_PATTERN) !== -1 && sheets[i].isSheetHidden()) {
      sheets[i].showSheet();
      shown.push(name);
    }
  }
  return { shown: shown };
}

// ============================================
// #REF! エラー自動修正
// ============================================

/**
 * 旧シートの #REF! エラーを検出して0に置換
 * API: ?action=run&fn=fixRefErrors
 */
function fixRefErrors() {
  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);
  var oldSheet = getSheetByMonth_(ss, settings.month);
  if (!oldSheet) return { error: '旧シートが見つかりません' };

  var lastRow = oldSheet.getLastRow();
  var lastCol = oldSheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { fixed: [] };

  var displayValues = oldSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var formulas = oldSheet.getRange(1, 1, lastRow, lastCol).getFormulas();

  var fixed = [];
  for (var r = 0; r < lastRow; r++) {
    for (var c = 0; c < lastCol; c++) {
      var dv = displayValues[r][c];
      if (dv === '#REF!' || dv === '#NAME?' || dv === '#ERROR!' || dv === '#VALUE!') {
        oldSheet.getRange(r + 1, c + 1).setValue(0);
        fixed.push({
          row: r + 1,
          col: c + 1,
          error: dv,
          formula: formulas[r][c] || ''
        });
      }
    }
  }

  if (fixed.length > 0) {
    SpreadsheetApp.flush();
  }

  Logger.log('#REF!修正完了: ' + fixed.length + '件');
  return { fixed: fixed, count: fixed.length };
}

// ============================================
// 結合解除 + 列固定
// ============================================

/**
 * 旧シートのB列以前〜C列以降にまたがる結合を解除し、A-B列を固定
 * API: ?action=run&fn=unmergeAndFreeze
 */
function unmergeAndFreeze() {
  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);
  var oldSheet = getSheetByMonth_(ss, settings.month);
  if (!oldSheet) return { error: '旧シートが見つかりません' };

  var maxRow = oldSheet.getMaxRows();
  var maxCol = oldSheet.getMaxColumns();
  var allMerges = oldSheet.getRange(1, 1, maxRow, maxCol).getMergedRanges();

  var unmerged = [];
  for (var i = 0; i < allMerges.length; i++) {
    var range = allMerges[i];
    var startCol = range.getColumn();
    var endCol = startCol + range.getNumColumns() - 1;

    // B列(2)以前に始まり、C列(3)以降に広がる結合セル
    if (startCol <= 2 && endCol >= 3) {
      var notation = range.getA1Notation();
      range.breakApart();
      unmerged.push(notation);
    }
  }

  // A-B列を固定
  oldSheet.setFrozenColumns(2);

  Logger.log('結合解除: ' + unmerged.join(', ') + ' → A-B列固定');
  return { unmerged: unmerged, frozenColumns: 2 };
}

// ============================================
// 旧シート書式改善
// ============================================

/**
 * 旧シートの見やすさを改善する書式設定
 * API: ?action=run&fn=formatOldSheet
 */
function formatOldSheet() {
  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);
  var oldSheet = getSheetByMonth_(ss, settings.month);
  if (!oldSheet) return { error: '旧シートが見つかりません' };

  var allData = oldSheet.getDataRange().getValues();
  var maxCol = allData[0] ? allData[0].length : 0;

  // ラベル行を検出してスタイリング
  var styledRows = [];

  for (var r = 0; r < Math.min(40, allData.length); r++) {
    var label = String(allData[r][0] || '').replace(/\s+/g, '');
    if (!label) {
      label = String(allData[r][1] || '').replace(/\s+/g, '');
    }

    // ランキング行: ダークヘッダー
    if (label.indexOf('ランキング') !== -1) {
      oldSheet.getRange(r + 1, 1, 1, maxCol)
        .setBackground('#1E293B').setFontColor('#F1F5F9').setFontWeight('bold');
      styledRows.push('row' + (r + 1) + ':ランキング');
    }

    // メンバー名前行（【】が3つ以上ある行）: ヘッダー背景
    var bracketCount = 0;
    for (var c = 0; c < allData[r].length; c++) {
      if (String(allData[r][c] || '').indexOf('【') !== -1) bracketCount++;
    }
    if (bracketCount >= 3) {
      oldSheet.getRange(r + 1, 1, 1, maxCol)
        .setBackground('#FEF3C7').setFontWeight('bold').setFontColor('#92400E');
      styledRows.push('row' + (r + 1) + ':名前行');
    }

    // 主要KPI行をハイライト
    if (label.indexOf('合計着金額') !== -1 || label.indexOf('合計売上') !== -1) {
      oldSheet.getRange(r + 1, 1, 1, maxCol)
        .setBackground('#FFF7ED').setFontWeight('bold');
      styledRows.push('row' + (r + 1) + ':' + label);
    }

    // 合計成約数・合計商談数
    if (label.indexOf('合計成約数') !== -1 || label.indexOf('合計商談数') !== -1) {
      oldSheet.getRange(r + 1, 1, 1, maxCol)
        .setBackground('#F0FDF4').setFontWeight('bold');
      styledRows.push('row' + (r + 1) + ':' + label);
    }

    // A列ラベルを太字に
    if (label) {
      oldSheet.getRange(r + 1, 1).setFontWeight('bold');
      if (allData[r][1] && String(allData[r][1] || '').trim()) {
        oldSheet.getRange(r + 1, 2).setFontWeight('bold');
      }
    }
  }

  // A列の列幅を広めに
  oldSheet.setColumnWidth(1, 60);
  oldSheet.setColumnWidth(2, 200);

  Logger.log('旧シート書式改善完了: ' + styledRows.length + '行');
  return { styledRows: styledRows };
}

// ============================================
// ゴン メンバー追加
// ============================================

/**
 * ゴンをアクティブメンバーとして追加（設定シート + 日別入力シート作成）
 * API: ?action=run&fn=addGonMember
 */
function addGonMember() {
  var ss = getSpreadsheet_();
  var settingsSheet = getSettingsSheet_(ss);
  if (!settingsSheet) return { error: '設定シートが見つかりません' };

  // ゴンが既に設定シートにいるかチェック
  var members = getMembersFromSettings_(ss);
  for (var i = 0; i < members.length; i++) {
    if (members[i].name === 'ゴン') {
      // 既にいる場合は日別シートだけ作成
      var existingDaily = getDailySheet_(ss, 'ゴン');
      if (!existingDaily) {
        var settings = getGlobalSettings_(ss);
        createDailySheet_(ss, 'ゴン', settings.year, settings.month);
        var dailySheet = getDailySheet_(ss, 'ゴン');
        if (dailySheet) formatDailySheet_(dailySheet);
        return { status: 'daily_sheet_created', message: 'ゴンの日別入力シートを作成しました' };
      }
      return { status: 'already_exists', message: 'ゴンは既に登録されています' };
    }
  }

  // アクティブメンバーの最後の行を探す（退職済の前に挿入）
  var lastRow = settingsSheet.getLastRow();
  var data = settingsSheet.getRange(SETTINGS_ROW_DATA_START, 1, lastRow - 1, SETTINGS_MEMBER_COL_COUNT).getValues();
  var insertIdx = -1;

  for (var j = 0; j < data.length; j++) {
    var status = String(data[j][3] || '').trim();
    if (status === '退職済') {
      insertIdx = j;
      break;
    }
  }

  // 退職済が見つからなければ末尾に追加
  var writeRow;
  if (insertIdx >= 0) {
    // 退職済行の前に挿入（行を挿入する）
    writeRow = SETTINGS_ROW_DATA_START + insertIdx;
    settingsSheet.insertRowBefore(writeRow);
  } else {
    writeRow = lastRow + 1;
  }

  settingsSheet.getRange(writeRow, 1, 1, SETTINGS_MEMBER_COL_COUNT).setValues([
    ['ゴン', 'ゴン', 3, 'アクティブ', '#F97316']
  ]);

  // ステータスのドロップダウンを適用
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['アクティブ', '退職済'], true)
    .build();
  settingsSheet.getRange(writeRow, SETTINGS_COL_STATUS).setDataValidation(statusRule);

  // 日別入力シートを作成
  var settings = getGlobalSettings_(ss);
  createDailySheet_(ss, 'ゴン', settings.year, settings.month);
  var dailySheet = getDailySheet_(ss, 'ゴン');
  if (dailySheet) formatDailySheet_(dailySheet);

  Logger.log('ゴンをメンバーに追加しました');
  return { status: 'added', message: 'ゴンをアクティブメンバーとして追加し、日別入力シートを作成しました' };
}

/**
 * トニーをメンバーに追加
 * API: ?action=run&fn=addTonyMember
 */
function addTonyMember() {
  var ss = getSpreadsheet_();
  var settingsSheet = getSettingsSheet_(ss);
  if (!settingsSheet) return { error: '設定シートが見つかりません' };

  var members = getMembersFromSettings_(ss);
  for (var i = 0; i < members.length; i++) {
    if (members[i].name === 'トニー') {
      var existingDaily = getDailySheet_(ss, 'トニー');
      if (!existingDaily) {
        var settings = getGlobalSettings_(ss);
        createDailySheet_(ss, 'トニー', settings.year, settings.month);
        var dailySheet = getDailySheet_(ss, 'トニー');
        if (dailySheet) formatDailySheet_(dailySheet);
        return { status: 'daily_sheet_created' };
      }
      return { status: 'already_exists' };
    }
  }

  var lastRow = settingsSheet.getLastRow();
  var data = settingsSheet.getRange(SETTINGS_ROW_DATA_START, 1, lastRow - 1, SETTINGS_MEMBER_COL_COUNT).getValues();
  var insertIdx = -1;
  for (var j = 0; j < data.length; j++) {
    var status = String(data[j][3] || '').trim();
    if (status === '退職済') { insertIdx = j; break; }
  }

  var writeRow;
  if (insertIdx >= 0) {
    writeRow = SETTINGS_ROW_DATA_START + insertIdx;
    settingsSheet.insertRowBefore(writeRow);
  } else {
    writeRow = lastRow + 1;
  }

  settingsSheet.getRange(writeRow, 1, 1, SETTINGS_MEMBER_COL_COUNT).setValues([
    ['トニー', 'トニー', 10, 'アクティブ', '#8B5CF6']
  ]);

  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['アクティブ', '退職済'], true)
    .build();
  settingsSheet.getRange(writeRow, SETTINGS_COL_STATUS).setDataValidation(statusRule);

  var settings = getGlobalSettings_(ss);
  createDailySheet_(ss, 'トニー', settings.year, settings.month);
  var dailySheet = getDailySheet_(ss, 'トニー');
  if (dailySheet) formatDailySheet_(dailySheet);

  return { status: 'added' };
}

/**
 * 旧シートのメンバー名行(row3)に画像を挿入
 * API: ?action=run&fn=insertMemberImages
 */
function insertMemberImages(targetNames) {
  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);
  var oldSheet = getSheetByMonth_(ss, settings.month);
  if (!oldSheet) return { error: '旧シートが見つかりません' };

  // デフォルト: ゴン・トニーのみ
  var targets = targetNames || ['ゴン', 'トニー'];

  var inserted = [];
  for (var i = 0; i < MEMBER_SECTIONS.length; i++) {
    var sec = MEMBER_SECTIONS[i];
    var found = false;
    for (var t = 0; t < targets.length; t++) {
      if (sec.name === targets[t]) { found = true; break; }
    }
    if (!found) continue;

    var url = iconUrl_(sec.name);
    if (!url) { inserted.push(sec.name + ' → URL未設定'); continue; }

    var col = sec.summaryCol; // 1-indexed
    try {
      // =IMAGE() 数式でrow3に画像を表示
      oldSheet.getRange(3, col).setFormula('=IMAGE("' + url + '", 1)');
      inserted.push(sec.name + ' → col' + col + ' row3 OK');

      // セクションヘッダーにも画像を挿入（dataStartの3行上）
      if (sec.dataStart > 0) {
        var sectionImageRow = sec.dataStart - 3;
        oldSheet.getRange(sectionImageRow, 2).setFormula('=IMAGE("' + url + '", 1)');
        inserted.push(sec.name + ' → col2 row' + sectionImageRow + ' (section) OK');
      }
    } catch (e) {
      inserted.push(sec.name + ' ERROR: ' + e.message);
    }
  }
  return { status: 'done', inserted: inserted };
}

/**
 * デバッグ: 指定gidのシートから行データを取得
 * API: ?action=run&fn=debugSheetByGid&gid=xxx&rows=1,3,4,8
 */
function debugSheetByGid(gid, rows, showFormulas) {
  var ss = getSpreadsheet_();
  var sheets = ss.getSheets();
  var sheet = null;
  var gidNum = parseInt(gid);
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === gidNum) { sheet = sheets[i]; break; }
  }
  if (!sheet) return { error: 'gid ' + gid + ' not found' };

  var rowNums = (rows || '1,3,4,8').split(',');
  var maxCol = sheet.getLastColumn();
  var result = { sheetName: sheet.getName(), data: {} };

  for (var r = 0; r < rowNums.length; r++) {
    var rowNum = parseInt(rowNums[r].trim());
    if (rowNum < 1) continue;
    var range = sheet.getRange(rowNum, 1, 1, maxCol);
    var vals = range.getValues()[0];
    var formulas = showFormulas ? range.getFormulas()[0] : [];
    var rowData = {};
    for (var c = 0; c < vals.length; c++) {
      if (vals[c] !== '' && vals[c] !== null && vals[c] !== undefined) {
        var colLetter = '';
        var tmp = c + 1;
        while (tmp > 0) { colLetter = String.fromCharCode(((tmp - 1) % 26) + 65) + colLetter; tmp = Math.floor((tmp - 1) / 26); }
        var entry = String(vals[c]);
        if (formulas[c]) entry += ' [' + formulas[c] + ']';
        rowData[colLetter] = entry;
      }
    }
    result.data['row' + rowNum] = rowData;
  }
  return result;
}

/**
 * 旧シートの4行目名前セルにハイパーリンクを設定
 * タップで各メンバーの個別入力欄（38行目以降）にジャンプ
 * API: ?action=run&fn=setNameLinks
 */
function setNameLinks() {
  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);
  var oldSheet = getSheetByMonth_(ss, settings.month);
  if (!oldSheet) return { error: '旧シートが見つかりません' };

  var gid = oldSheet.getSheetId();
  var sheetUrl = ss.getUrl();
  var linked = [];

  for (var i = 0; i < MEMBER_SECTIONS.length; i++) {
    var sec = MEMBER_SECTIONS[i];
    if (sec.dataStart < 0) {
      linked.push(sec.name + ' → セクションなし(skip)');
      continue;
    }

    var col = sec.summaryCol; // 1-indexed
    var cell = oldSheet.getRange(4, col);
    var name = cell.getValue() || sec.name;

    // シート内リンク: セクションの先頭行へ
    var targetCell = 'A' + sec.dataStart;
    var linkUrl = '#gid=' + gid + '&range=' + targetCell;
    cell.setFormula('=HYPERLINK("' + linkUrl + '","' + name + '")');
    linked.push(sec.name + ' → row' + sec.dataStart);
  }
  return { status: 'done', linked: linked };
}

/**
 * 全シートの名前とgidを一覧表示（デバッグ用）
 * API: ?action=run&fn=listSheets
 */
function listSheets() {
  var ss = getSpreadsheet_();
  var sheets = ss.getSheets();
  var list = [];
  for (var i = 0; i < sheets.length; i++) {
    list.push({
      name: sheets[i].getName(),
      gid: sheets[i].getSheetId(),
      hidden: sheets[i].isSheetHidden()
    });
  }
  return list;
}

/**
 * Row6のCO引き数式とRow5のランキング数式を高速修復（トリガー用）
 * v200のfixOldSheetFormulasがCO引きなしで上書きするため、1分トリガーで修復する
 */
function quickFixFormulas() {
  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);
  var oldSheet = getSheetByMonth_(ss, settings.month);
  if (!oldSheet) return;

  // Row6 C列の数式を確認 — CO引き（マイナス記号）があれば修復不要
  var testFormula = oldSheet.getRange('C6').getFormula();
  if (testFormula && testFormula.indexOf('-') !== -1) return;

  // 修復が必要 → fixOldSheetFormulas(HEAD版 = CO引き + 全員ランキング)を実行
  fixOldSheetFormulas();
}

/**
 * Row 5（着金ランキング）にランキング順の色をつける
 * Apps Scriptエディタから実行する
 */
function colorRankRow() {
  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);
  var oldSheet = getSheetByMonth_(ss, settings.month);
  if (!oldSheet) return 'sheet not found';

  var memberSections = detectMemberSections_(ss, oldSheet, settings);

  // 行検出（detectRowMap_の代替）
  var allData = oldSheet.getDataRange().getValues();
  var revenueRow = 0;
  for (var r = 0; r < allData.length; r++) {
    var label = String(allData[r][0] || '').replace(/\s+/g, '').replace(/[\(\)（）]/g, '');
    if (!label) label = String(allData[r][1] || '').replace(/\s+/g, '').replace(/[\(\)（）]/g, '');
    if (label.indexOf('合計着金額') !== -1) { revenueRow = r + 1; break; }
  }
  if (!revenueRow) return 'revenue row not found';
  var rankRow = revenueRow - 1; // ランキング行 = 着金額行の1つ上

  // ランキング色マップ（1位=金, 2位=銀, 3位=銅, 4位以降=順位色）
  var RANK_COLORS = {
    1:  { bg: '#FEF3C7', fg: '#92400E' },  // 金
    2:  { bg: '#F1F5F9', fg: '#475569' },  // 銀
    3:  { bg: '#FFF7ED', fg: '#9A3412' },  // 銅
    4:  { bg: '#DBEAFE', fg: '#1E40AF' },  // 青
    5:  { bg: '#D1FAE5', fg: '#065F46' },  // 緑
    6:  { bg: '#EDE9FE', fg: '#5B21B6' },  // 紫
    7:  { bg: '#FCE7F3', fg: '#9D174D' },  // ピンク
    8:  { bg: '#FEF9C3', fg: '#854D0E' },  // 黄
    9:  { bg: '#FFE4E6', fg: '#9F1239' },  // 赤
    10: { bg: '#F3F4F6', fg: '#6B7280' },  // グレー
  };

  var colored = [];
  for (var i = 0; i < memberSections.length; i++) {
    var sc = memberSections[i].summaryCol;
    var cell = oldSheet.getRange(rankRow, sc);
    var val = cell.getValue();
    var rank = parseInt(val) || 0;
    var color = RANK_COLORS[rank] || RANK_COLORS[10];
    cell.setBackground(color.bg)
        .setFontColor(color.fg)
        .setFontWeight('bold')
        .setFontSize(14)
        .setHorizontalAlignment('center');
    colored.push(memberSections[i].name + '=' + rank + '位');
  }

  // AG列（合計列）
  var agCell = oldSheet.getRange(rankRow, 33);
  agCell.setBackground('#F9FAFB').setFontColor('#9CA3AF').setFontSize(10);

  SpreadsheetApp.flush();
  return { colored: colored };
}

/**
 * RANK数式だけを直接修正（検出ロジック不要）
 * 10メンバーの列C,F,I,L,O,R,U,X,AA,AD のrow5にRANK数式を設定
 */
function fixRankOnly() {
  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);
  var oldSheet = getSheetByMonth_(ss, settings.month);
  if (!oldSheet) return { error: 'sheet not found' };

  var memberCols = [3, 6, 9, 12, 15, 18, 21, 24, 27, 30]; // C,F,I,L,O,R,U,X,AA,AD
  var rankRow = 5;
  var revRow = 6;

  // RANK参照配列を構築
  var refs = [];
  for (var i = 0; i < memberCols.length; i++) {
    refs.push(columnToLetter_(memberCols[i]) + revRow);
  }
  var rankArray = '{' + refs.join(',') + '}';

  var results = [];
  for (var i = 0; i < memberCols.length; i++) {
    var col = memberCols[i];
    var cellRef = columnToLetter_(col) + revRow;
    var formula = '=RANK(' + cellRef + ',' + rankArray + ',0)';
    oldSheet.getRange(rankRow, col).setFormula(formula);
    results.push(columnToLetter_(col) + rankRow + ': ' + formula);
  }

  // RANK色つけ
  var RANK_COLORS = {
    1: { bg: '#FEF3C7', fg: '#92400E' },
    2: { bg: '#F1F5F9', fg: '#475569' },
    3: { bg: '#FFF7ED', fg: '#9A3412' },
    4: { bg: '#DBEAFE', fg: '#1E40AF' },
    5: { bg: '#D1FAE5', fg: '#065F46' },
    6: { bg: '#EDE9FE', fg: '#5B21B6' },
    7: { bg: '#FCE7F3', fg: '#9D174D' },
    8: { bg: '#FEF9C3', fg: '#854D0E' },
    9: { bg: '#FFE4E6', fg: '#9F1239' },
    10: { bg: '#F3F4F6', fg: '#6B7280' }
  };
  SpreadsheetApp.flush();
  for (var i = 0; i < memberCols.length; i++) {
    var cell = oldSheet.getRange(rankRow, memberCols[i]);
    var rankVal = parseInt(cell.getValue()) || 0;
    var rc = RANK_COLORS[rankVal] || RANK_COLORS[10];
    cell.setBackground(rc.bg).setFontColor(rc.fg).setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  }

  // 確認用: 書き込んだシートの情報
  var verification = {};
  for (var i = 0; i < memberCols.length; i++) {
    var cell = oldSheet.getRange(rankRow, memberCols[i]);
    verification[columnToLetter_(memberCols[i]) + rankRow] = {
      formula: cell.getFormula(),
      value: cell.getValue()
    };
  }

  return { fixed: results, sheetName: oldSheet.getName(), sheetId: oldSheet.getSheetId(), verification: verification };
}

/**
 * quickFixFormulasの1分トリガーをセットアップ
 * Apps Scriptエディタから1回だけ実行する
 */
function setupQuickFixTrigger() {
  // 既存のquickFixFormulasトリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'quickFixFormulas') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // 1分ごとのトリガーを作成
  ScriptApp.newTrigger('quickFixFormulas')
    .timeBased()
    .everyMinutes(1)
    .create();
  return 'quickFixFormulas trigger created (every 1 min)';
}

/**
 * 不要トリガー一括削除（quickFixFormulas + updateSummary）
 * Apps Scriptエディタから1回だけ実行する
 */
function deleteWriteTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var deleted = [];
  var kept = [];
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'quickFixFormulas' || fn === 'updateSummary') {
      ScriptApp.deleteTrigger(triggers[i]);
      deleted.push(fn);
    } else {
      kept.push(fn);
    }
  }
  Logger.log('削除: ' + JSON.stringify(deleted));
  Logger.log('残存: ' + JSON.stringify(kept));
  return { deleted: deleted, kept: kept };
}

/**
 * 全トリガー一覧を返す（確認用）
 */
function listAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var list = [];
  for (var i = 0; i < triggers.length; i++) {
    list.push({
      fn: triggers[i].getHandlerFunction(),
      type: triggers[i].getEventType().toString(),
      source: triggers[i].getTriggerSource().toString()
    });
  }
  Logger.log(JSON.stringify(list, null, 2));
  return list;
}
