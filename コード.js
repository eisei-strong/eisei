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
// ゴール設定 読み込み
// ============================================
function getGoalSettings_(ss) {
  var sheet = ss.getSheetByName('ゴール設定');
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var result = {};
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][0]);
    var teamGoal = data[i][1] || 0;
    var memberKgi = {};
    try { memberKgi = JSON.parse(data[i][2] || '{}'); } catch (e) {}
    result[key] = { teamGoal: teamGoal, memberKgi: memberKgi };
  }
  return result;
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

    if (action === 'saveGoals') {
      var ss = getSpreadsheet_();
      var sheet = ss.getSheetByName('ゴール設定');
      if (!sheet) {
        sheet = ss.insertSheet('ゴール設定');
        sheet.getRange(1, 1, 1, 3).setValues([['月キー', 'チーム目標', '個人目標JSON']]);
      }
      var key = payload.key; // "2026_3" 等
      var teamGoal = payload.teamGoal || 0;
      var memberKgi = payload.memberKgi || {};
      // 既存行を検索
      var data = sheet.getDataRange().getValues();
      var found = -1;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(key)) { found = i + 1; break; }
      }
      if (found > 0) {
        sheet.getRange(found, 2, 1, 2).setValues([[teamGoal, JSON.stringify(memberKgi)]]);
      } else {
        sheet.appendRow([key, teamGoal, JSON.stringify(memberKgi)]);
      }
      return ContentService.createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'getGoals') {
      var ss = getSpreadsheet_();
      var result = getGoalSettings_(ss);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

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

    if (action === 'postRegister') {
      var prResult = postAppRegister_(payload.id, payload.password);
      return ContentService.createTextOutput(JSON.stringify(prResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'postLogin') {
      var plResult = postAppLogin_(payload.id, payload.password);
      return ContentService.createTextOutput(JSON.stringify(plResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'postSave') {
      var psResult = postAppSave_(payload.token, payload.value, payload.col);
      return ContentService.createTextOutput(JSON.stringify(psResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'postLoginStreak') {
      var plsResult = postAppLoginStreak_(payload.token);
      return ContentService.createTextOutput(JSON.stringify(plsResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'postSetGoal') {
      var psgResult = postAppSetGoal_(payload.token, payload.goal);
      return ContentService.createTextOutput(JSON.stringify(psgResult))
        .setMimeType(ContentService.MimeType.JSON);
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

  // デバッグ: 旧シート行6-30のデータ＋数式を全表示
  if (params.action === 'verify') {
    try {
      var ss = getSpreadsheet_();
      var settings = getGlobalSettings_(ss);
      var oldSheet = getSheetByMonth_(ss, settings.month);
      if (!oldSheet) return ContentService.createTextOutput(JSON.stringify({error:'no sheet'})).setMimeType(ContentService.MimeType.JSON);

      var vals = oldSheet.getRange(1, 1, 35, 35).getValues();
      var fmls = oldSheet.getRange(1, 1, 35, 35).getFormulas();

      // メンバー名行（Row 4-5あたり）を検出
      var memberSections = detectMemberSections_(oldSheet);
      var sectionInfo = [];
      for (var si = 0; si < memberSections.length; si++) {
        sectionInfo.push({name: memberSections[si].name, col: memberSections[si].summaryCol, ds: memberSections[si].dataStart, de: memberSections[si].dataEnd});
      }

      // 行1-35の各メンバー列のデータと数式
      var rows = [];
      for (var r = 0; r < 35; r++) {
        var rowData = {row: r+1, B: String(vals[r][1] || '')};
        for (var mi = 0; mi < memberSections.length; mi++) {
          var c = memberSections[mi].summaryCol - 1; // 0-based
          var val = vals[r][c];
          var fml = fmls[r][c];
          rowData[memberSections[mi].name] = {v: val, f: fml || ''};
        }
        // AG列(33)
        var agVal = vals[r][32];
        var agFml = fmls[r][32];
        rowData['AG'] = {v: agVal, f: agFml || ''};
        rows.push(rowData);
      }

      return ContentService.createTextOutput(JSON.stringify({sections: sectionInfo, rows: rows}))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message, stack: err.stack }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // 旧シート数式修正 + B列ラベル復元 + 同期
  if (params.action === 'fixAndSync') {
    try {
      var fixResult = fixOldSheetFormulas();
      syncFromOldSheet();
      updateSummary();
      return ContentService.createTextOutput(JSON.stringify({ fix: fixResult, status: 'synced' }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message, stack: err.stack }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // 同期+キャッシュクリア+データ返却（ダッシュボードの同期ボタン用）
  if (params.action === 'syncAndFetch') {
    try {
      syncFromOldSheet();
      updateSummary();
      clearDashboardCache_();
      var syncResult = getDashboardData(
        params.month ? parseInt(params.month) : undefined,
        params.year ? parseInt(params.year) : undefined,
        true
      );
      return ContentService.createTextOutput(JSON.stringify(syncResult))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (params.action === 'debugMember') {
    try {
      var targetD = params.d || '伊東';
      var m = params.month ? parseInt(params.month) : undefined;
      var y = params.year ? parseInt(params.year) : undefined;
      var result = debugMemberRows(targetD, m, y);
      return ContentService.createTextOutput(JSON.stringify(result, null, 2))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (params.action === 'createAprilHope') {
    try {
      var result = createAprilHopeSheet();
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (params.action === 'designSamples') {
    try {
      var result = createDesignSamples();
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (params.action === 'addMembers') {
    try {
      var result = addMissingMembersToSheets(params.sheet || null);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (params.action === 'extractNames') {
    try {
      var result = extractCustomerNames();
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (params.action === 'hideOldRows') {
    try {
      var result = autoHideNonCurrentMonth();
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (params.action === 'showAllRows') {
    try {
      var result = showAllMasterRows();
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (params.action === 'renameMembers') {
    try {
      var renames = {
        'AをAでやる': '意思決定',
        'ビッグマウス': 'ありのまま',
        'ワントーン': 'スマイル',
        'ドライ': 'ポジティブ'
      };
      var ss = getSpreadsheet_();
      var sheet = ss.getSheetByName(SHEET_SETTINGS);
      var lastRow = Math.min(sheet.getLastRow(), SETTINGS_ROW_GLOBAL_START - 2);
      var changed = [];
      for (var r = SETTINGS_ROW_DATA_START; r <= lastRow; r++) {
        var nameVal = String(sheet.getRange(r, SETTINGS_COL_NAME).getValue()).trim();
        var dispVal = String(sheet.getRange(r, SETTINGS_COL_DISPLAY_NAME).getValue()).trim();
        var newName = renames[nameVal] || renames[dispVal];
        if (newName) {
          sheet.getRange(r, SETTINGS_COL_NAME).setValue(newName);
          sheet.getRange(r, SETTINGS_COL_DISPLAY_NAME).setValue(newName);
          changed.push(nameVal + ' → ' + newName);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ ok: true, changed: changed }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (params.action === 'bottleneck') {
    try {
      var result = sendBottleneckNotification();
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (params.action === 'api') {
    try {
      var skipCache = params.fresh === '1';
      var result;
      if (params.type === 'months') {
        result = getAvailableMonths();
      } else if (params.month && params.year) {
        result = getDashboardData(parseInt(params.month), parseInt(params.year), skipCache);
      } else {
        result = getDashboardData(undefined, undefined, skipCache);
      }
      // ゴール設定を付与（キャッシュ外で常に最新）
      if (params.type !== 'months') {
        try {
          var gss = getSpreadsheet_();
          result.goalSettings = getGoalSettings_(gss);
        } catch (ge) { result.goalSettings = {}; }
        // 共創ポイント + メッセージ
        try {
          result.chatworkPoints = getChatworkPoints_();
          result.chatworkMessages = getChatworkMessages_();
        } catch (cpe) { result.chatworkPoints = {}; result.chatworkMessages = []; }
      }
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ===== 経費ダッシュボード API =====
  if (params.action === 'expense') {
    try {
      var expResult;
      var expType = params.type || '';
      if (expType === 'months') {
        expResult = expGetMonths_();
      } else if (expType === 'trends') {
        expResult = expGetTrends_();
      } else if (expType === 'detail') {
        expResult = expGetDetail_(parseInt(params.year), parseInt(params.month), params.cat || '');
      } else {
        var em = params.month ? parseInt(params.month) : 0;
        var ey = params.year ? parseInt(params.year) : 0;
        expResult = expGetData_(ey, em);
      }
      return ContentService.createTextOutput(JSON.stringify(expResult))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ===== 投稿本数入力アプリ API =====
  if (params.action === 'postCheckId') {
    try {
      var pcResult = postAppCheckId_(params.id);
      return ContentService.createTextOutput(JSON.stringify(pcResult))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (params.action === 'postGet') {
    try {
      var pgResult = postAppGet_(params.token, params.year, params.month);
      return ContentService.createTextOutput(JSON.stringify(pgResult))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (params.action === 'postRegister') {
    try {
      var prResult = postAppRegister_(params.id, params.password);
      return ContentService.createTextOutput(JSON.stringify(prResult))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (params.action === 'postLogin') {
    try {
      var plResult = postAppLogin_(params.id, params.password);
      return ContentService.createTextOutput(JSON.stringify(plResult))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (params.action === 'postSave') {
    try {
      var psResult = postAppSave_(params.token, params.value, params.col, params.year, params.month);
      return ContentService.createTextOutput(JSON.stringify(psResult))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (params.action === 'postResetPassword') {
    try {
      var prpResult = postAppResetPassword_(params.id, params.email);
      return ContentService.createTextOutput(JSON.stringify(prpResult))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (params.action === 'postLoginStreak') {
    try {
      var plsResult = postAppLoginStreak_(params.token);
      return ContentService.createTextOutput(JSON.stringify(plsResult))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (params.action === 'postRanking') {
    try {
      var prkResult = postAppRanking_();
      return ContentService.createTextOutput(JSON.stringify(prkResult))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (params.action === 'postSetGoal') {
    try {
      var psgResult = postAppSetGoal_(params.token, params.goal);
      return ContentService.createTextOutput(JSON.stringify(psgResult))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (params.action === 'postGetHope') {
    try {
      var pghResult = postAppGetHope_(params.token, params.year, params.month);
      return ContentService.createTextOutput(JSON.stringify(pghResult))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ===== 万バズ台本 API =====
  if (params.action === 'manbazu') {
    try {
      var mbResult = manbazuGetData_();
      return ContentService.createTextOutput(JSON.stringify(mbResult))
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
      else if (fn === 'generateKanonReport') fnResult = generateKanonReport();
      else if (fn === 'kanonMembers') fnResult = getKanonSalesMembers();
      else if (fn === 'generateKanonMTGDoc') fnResult = generateKanonMTGDoc();
      else if (fn === 'inspectExtSheet') fnResult = inspectExtSheet_(params.ssid, params.gid, params.startRow, params.endRow);
      else if (fn === 'repairSMCMaster') fnResult = repairSMCMaster();
      else if (fn === 'setupFormSync') fnResult = setupFormSyncTrigger();
      else if (fn === 'syncDebugLog') fnResult = getSyncDebugLog();
      else if (fn === 'debugFormLatest') fnResult = debugFormLatest();
      else if (fn === 'debugGhostRows') fnResult = debugGhostRows();
      else if (fn === 'deleteGhostRows') fnResult = deleteGhostRows();
      else if (fn === 'deleteDeadTriggers') fnResult = deleteDeadTriggers();
      else if (fn === 'setupFormSubmitCleanup') fnResult = setupFormSubmitCleanup();
      else if (fn === 'debugAllTriggersOnSMC') fnResult = debugAllTriggersOnSMC();
      else if (fn === 'debugRowAP') fnResult = debugRowAP(params.row);
      else if (fn === 'listMasterNames') fnResult = listMasterNames(params.month, params.year);
      else if (fn === 'debugMemberRevenue') fnResult = debugMemberRevenue(params.member, params.month, params.year);
      else if (fn === 'auditDuplicates') fnResult = auditDuplicates(params.month, params.year);
      else if (fn === 'deduplicateSeiyaku') fnResult = deduplicateSeiyaku(params.month, params.year, params.dry);
      else if (fn === 'clearAPDupeAmts') fnResult = clearAPDuplicateAmounts(params.month, params.year, params.dry);
      else if (fn === 'dumpRows') fnResult = dumpRows(params.rows);
      else if (fn === 'clearAPForNonSeiyaku') fnResult = clearAPForNonSeiyaku(params.dry);
      else if (fn === 'debugMasterFormulas') fnResult = debugMasterFormulas();
      else if (fn === 'checkFormTail') fnResult = checkFormTail();
      else if (fn === 'backfill') fnResult = { error: '復旧モード: backfillは禁止されています' };
      else if (fn === 'dryRunSync') fnResult = dryRunSyncRecentFormToMaster();
      else if (fn === 'dryRunRepair') fnResult = repairDryRunSMCMaster();
      else if (fn === 'upsertByCDE') fnResult = syncRecentFormToMaster_upsertByCDE();
      else if (fn === 'dryRunUpsert') fnResult = dryRunSyncUpsertByCDE();
      else if (fn === 'syncV2') fnResult = syncFormToMaster_v2();
      else if (fn === 'dryRunV2') fnResult = dryRunSyncFormToMaster_v2();
      else if (fn === 'diagPX') fnResult = diagnosePXProtection();
      else if (fn === 'diagFormLink') fnResult = diagFormLink();
      else if (fn === 'sortMaster') fnResult = sortMasterByDate();
      else if (fn === 'removeOrangeFormat') fnResult = removeOrangeConditionalFormat();
      else if (fn === 'setLFormat') fnResult = setLColumnFormatting();
      else if (fn === 'setLValidation') fnResult = setLColumnValidation();
      else if (fn === 'fixCX') {
        var smc2 = SpreadsheetApp.openById(SMC_SS_ID);
        var m2 = smc2.getSheetByName(SMC_MASTER_SHEET);
        fnResult = { fixed: fixCXOverQ_(m2) };
        SpreadsheetApp.flush();
      }
      else if (fn === 'flagRows') fnResult = flagRows(params.rows);
      else if (fn === 'auditCandidates') fnResult = auditDuplicateCandidates(params.month, params.year, params.members, params.dry);
      else if (fn === 'auditAll') fnResult = auditAllMembers(params.month, params.year);
      else if (fn === 'auditCXvsQ') fnResult = auditCXvsQ(params.month, params.year);
      else if (fn === 'deleteRow') {
        fnResult = { error: '復旧モード: 物理削除は禁止されています' };
      }
      else if (fn === 'contToLost') fnResult = contToLost(params.dry);
      else if (fn === 'setupDailyContToLost') fnResult = setupDailyContToLost();
      else if (fn === 'replaceLStatus') fnResult = replaceLStatus(params.dry);
      else if (fn === 'exportRevenue') fnResult = exportRevenueSheet(parseInt(params.month) || undefined, parseInt(params.year) || undefined);
      else if (fn === 'calcRevenue') fnResult = calcRevenueFromMaster(parseInt(params.month) || undefined, parseInt(params.year) || undefined);
      else if (fn === 'calcRevenueDebug') fnResult = calcRevenueDebug(parseInt(params.month) || undefined, parseInt(params.year) || undefined);
      else if (fn === 'switchToV2') fnResult = switchToV2Trigger();
      else if (fn === 'listTriggers') fnResult = listTriggers();
      else if (fn === 'refreshDropdowns') fnResult = refreshSearchDropdowns();
      else if (fn === 'setDValidation') fnResult = setDColumnValidation();
      else if (fn === 'auditRules') fnResult = auditSheetRules();
      else if (fn === 'normalizeD') fnResult = normalizeDColumn(params.dry);
      else if (fn === 'debugMasterMember') {
        var memberName = params.name || '';
        var smc = SpreadsheetApp.openById('1KxHeLmrpdaw1IUhBaQ46UWSHu-8SCRZqcrHOE2hMwDo');
        var ms = smc.getSheetByName('👑商談マスターデータ');
        var lastRow = ms.getLastRow();
        var vals = ms.getRange(2, 1, lastRow - 1, 102).getValues();
        var tz = Session.getScriptTimeZone();
        var rows = [];
        for (var di = 0; di < vals.length; di++) {
          var d = String(vals[di][3] || '').trim();
          if (d !== memberName) continue;
          var c = vals[di][2];
          if (!rev_isTargetMonth_(c, 2026, 3, tz)) continue;
          var l = String(vals[di][11] || '').trim();
          var q = vals[di][16];
          var cx = vals[di][101];
          var r_name = String(vals[di][17] || '');
          rows.push({row: di+2, D: d, L: l, Q: q, CX: cx, R: r_name});
        }
        fnResult = {member: memberName, count: rows.length, rows: rows};
      }
      else if (fn === 'debugDaily') {
        var memberName = params.name || '';
        var ss = getSpreadsheet_();
        var sheet = getDailySheet_(ss, memberName);
        if (!sheet) { fnResult = {error: 'sheet not found: ' + memberName}; }
        else {
          var data = sheet.getRange(1, 1, 45, 11).getValues();
          var rows = [];
          for (var dr = 0; dr < data.length; dr++) {
            rows.push({row: dr+1, A: String(data[dr][0]), B: data[dr][1], C: data[dr][2], D: data[dr][3], E: data[dr][4], F: data[dr][5], G: data[dr][6]});
          }
          fnResult = {sheet: sheet.getName(), rows: rows};
        }
      }
      else if (fn === 'testSort') fnResult = testSortOnCopy();
      else if (fn === 'applySort') fnResult = applySortFromTest();
      else if (fn === 'setupGuardianTasksSheet') fnResult = (function() {
        var ss = SpreadsheetApp.openById('1k_x3aNRTbojmhJZGMS6JGNiTNJLQR4sD5zyJCBh1YqY');
        var sheet = ss.getSheetByName('ガーディアン担当業務');
        if (sheet) return { status: 'already exists' };
        sheet = ss.insertSheet('ガーディアン担当業務');
        var rows = [
          ['名前', '業務内容'],
          ['星野', 'SDzoom受講生お知らせ・登録 / 炭治郎部屋進捗管理 / 1:1チャット作成・案内 / ID付与・リンク発行 / 数値報告 / 万バズ依頼対応 / アマギフ送付・郵送確認'],
          ['ふうか', '平日日中プッシュ割り振り / 流入経路数値記入 / 優先順位スプレ修正・改善'],
          ['まりん', '添削確認 / 成長シェア確認報告'],
          ['スズカ', '数値報告 / プッシュ報告入力 / 契約書アドレス確認 / 流入経路確認 / 台本・万バズリサーチ'],
          ['ゆいな', '添削時間確認報告 / 平日夜土日祝プッシュ割り振り（シフト）'],
          ['ココ', '台本・万バズリサーチ / シルバー・ゴールド外注部屋管理 / ID付与・リンク発行 / シルバー初投稿お祝いチャット'],
          ['あき', '承諾書・契約書郵送 / 代理店契約書業務 / 土日祝: 数値報告・プッシュ入力'],
          ['レイ', '信販・決済会社対応 / CO・中途解約対応 / ログイン認証 / 入金確認 / 報酬計算 / 契約書作成・修正 / 入電対応'],
          ['あやの', 'シルバー初投稿お祝いチャット（引継ぎ中）/ 万バズリサーチ（ふうか連携中）']
        ];
        sheet.getRange(1, 1, rows.length, 2).setValues(rows);
        sheet.setColumnWidth(1, 120);
        sheet.setColumnWidth(2, 600);
        sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#F5EDE4');
        return { status: 'created', rows: rows.length - 1 };
      })();
      else if (fn === 'testGuardianNotify') fnResult = (function() {
        // PropertiesServiceのトークンとハードコードトークン両方試す
        var token = getChatworkToken_();
        var fallbackToken = CW_API_TOKEN;
        var hasToken = !!token;
        var tokenLen = token ? token.length : 0;
        var memberNames = getGuardianMemberNames_();
        var sendResult = null;
        if (hasToken && memberNames.length > 0) {
          var body = '[info][title]🔧 ガーディアン動的メンバー確認テスト[/title]'
            + '現在の対象メンバー（' + memberNames.length + '名）:\n\n';
          for (var ti = 0; ti < memberNames.length; ti++) {
            body += '✅ ' + memberNames[ti] + '\n';
          }
          body += '\n除外リスト: ' + GUARDIAN_EXCLUDE.join(', ') + '\n';
          body += '\nこのメッセージはテストです。[/info]';
          var url = 'https://api.chatwork.com/v2/rooms/' + GUARDIAN_ROOM_ID + '/messages';
          var res = UrlFetchApp.fetch(url, {
            method: 'post',
            headers: { 'X-ChatWorkToken': token },
            payload: { body: body },
            muteHttpExceptions: true
          });
          sendResult = { code: res.getResponseCode(), body: res.getContentText().substring(0, 200) };
        }
        // 両方のトークンでルームメンバーAPI確認
        var roomUrl = 'https://api.chatwork.com/v2/rooms/' + GUARDIAN_ROOM_ID + '/members';
        var roomRes1 = UrlFetchApp.fetch(roomUrl, { method: 'get', headers: { 'X-ChatWorkToken': token }, muteHttpExceptions: true });
        var roomRes2 = UrlFetchApp.fetch(roomUrl, { method: 'get', headers: { 'X-ChatWorkToken': fallbackToken }, muteHttpExceptions: true });
        return {
          hasToken: hasToken, tokenLen: tokenLen,
          propsToken: { code: roomRes1.getResponseCode(), body: roomRes1.getContentText().substring(0, 200) },
          hardcodeToken: { code: roomRes2.getResponseCode(), body: roomRes2.getContentText().substring(0, 200) },
          memberCount: memberNames.length, members: memberNames, sendResult: sendResult
        };
      })();
      else if (fn === 'pauseFormSync') fnResult = pauseFormSync();
      else if (fn === 'resumeFormSync') fnResult = resumeFormSync();
      else if (fn === 'formSyncStatus') fnResult = getFormSyncStatus();
      else if (fn === 'debugCalendar') {
        // 指定IDのカレンダー + 全アクセス可能カレンダー一覧
        var cal = CalendarApp.getCalendarById('allforone.namaka@gmail.com');
        var calFound = cal ? true : false;
        var evts = cal ? cal.getEvents(new Date(2026,2,1), new Date(2026,3,31,23,59,59)) : [];
        var titles = [];
        for (var ei = 0; ei < Math.min(evts.length, 50); ei++) titles.push(evts[ei].getTitle() + ' | ' + evts[ei].getStartTime());
        // 全カレンダー一覧
        var allCals = CalendarApp.getAllCalendars();
        var calList = [];
        for (var ci = 0; ci < allCals.length; ci++) calList.push({ name: allCals[ci].getName(), id: allCals[ci].getId() });
        fnResult = { calFound: calFound, count: evts.length, titles: titles, calendars: calList };
      }
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

function getDashboardData(targetMonth, targetYear, skipCache) {
  var cache = CacheService.getScriptCache();

  // キャッシュキー構築
  var cacheKey = 'dash_' + (targetYear || 'cur') + '_' + (targetMonth || 'cur');

  // キャッシュ確認（skipCache=trueの場合スキップ）
  if (!skipCache) {
    var cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch (e) { /* キャッシュ破損→再取得 */ }
    }
  }

  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);

  var result;

  // 引数あり & 当月以外 → アーカイブから取得
  if (targetMonth && targetYear) {
    targetMonth = parseNum_(targetMonth);
    targetYear = parseNum_(targetYear);
    if (targetMonth !== settings.month || targetYear !== settings.year) {
      result = getDashboardDataFromArchive_(ss, targetYear, targetMonth);
      // アーカイブは長めにキャッシュ（30分）
      try { cache.put(cacheKey, JSON.stringify(result), 1800); } catch (e) { }
      return result;
    }
  }

  // 当月 → 既存のリアルタイム処理
  var settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  if (settingsSheet) {
    var summarySheet = ss.getSheetByName(SHEET_SUMMARY);
    if (summarySheet) {
      var summaryVal = summarySheet.getRange(SM_ROW_KPI_REVENUE, 2).getValue();
      if (summaryVal === '' || summaryVal === null || summaryVal === undefined) {
        try { updateSummary(); } catch (e) { /* フォールバックに進む */ }
      }
    }

    if (isNewStructureReady_(ss)) {
      result = getDashboardData_new_(ss);
      // マスターシートの着金/売上データで上書き
      try {
        var masterData = calcRevenueFromMaster(settings.month, settings.year);
        if (masterData && masterData.members) {
          var masterMap = {};
          for (var mi = 0; mi < masterData.members.length; mi++) {
            masterMap[masterData.members[mi].name] = masterData.members[mi];
          }
          for (var ri = 0; ri < result.members.length; ri++) {
            var m = result.members[ri];
            var md = masterMap[m.name];
            if (md) {
              m.revenue = md.revenue;
              m.sales = md.sales;
              m.deals = md.deals;
              m.closed = md.closed;
              m.closeRate = md.closeRate;
              m.avgPrice = md.avgPrice;
              m.coRevenue = md.coRevenue;
              m.coAmount = md.coAmount;
              m.lost = md.lost;
            }
          }
          // revenueでソートし直し
          result.members.sort(function(a, b) { return b.revenue - a.revenue; });
          for (var ri = 0; ri < result.members.length; ri++) {
            result.members[ri].rank = ri + 1;
            result.members[ri].gapToTop = ri === 0 ? 0 : round1_(result.members[0].revenue - result.members[ri].revenue);
          }
          result.totalRevenue = round1_(result.members.reduce(function(s, m) { return s + (m.revenue || 0); }, 0));
        }
      } catch (e) {
        Logger.log('Master overlay error: ' + e.message);
      }
      // 当月は5分キャッシュ
      try { cache.put(cacheKey, JSON.stringify(result), 300); } catch (e) { }
      return result;
    }
  }

  result = getDashboardData_legacy_(ss);
  try { cache.put(cacheKey, JSON.stringify(result), 90); } catch (e) { }
  return result;
}

/**
 * ダッシュボードキャッシュをクリア
 */
function clearDashboardCache_() {
  var cache = CacheService.getScriptCache();
  cache.removeAll(['dash_cur_cur']);
  // 当月のキャッシュも念のため削除
  var now = new Date();
  var y = now.getFullYear();
  for (var m = 1; m <= 12; m++) {
    cache.remove('dash_' + y + '_' + m);
  }
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
    var prev = prevData[key] || { revenue: 0, deals: 0, closed: 0, closeRate: 0 };
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

// ニックネーム→マスターD列名マップ
var NICK_TO_MASTER_D_ = {
  '意思決定': '阿部',
  'ポジティブ': '伊東',
  'ヒトコト': '久保田',
  'ありのまま': '辻阪',
  'けつだん': '福島',
  'ぜんぶり': '五十嵐',
  'スクリプト通りに営業するくん': '新居',
  'スクリプトくん': '新居',
  'スマイル': '佐々木',
  'ゴン': '大久保',
  'トニー': '矢吹',
  'ガロウ': '大内',
  'リヴァイ': '川合',
  '嬴政': '嬴政'
};

function getDashboardData_new_(ss) {
  var settings = getGlobalSettings_(ss);
  var activeMembers = getActiveMembers_(ss);
  var currentMonth = settings.month;

  var prevMonthData = getPrevMonthFromMaster_(settings);

  // マスターベースで全データ取得
  var masterRevData = calcRevenueFromMaster(settings.month, settings.year);
  var masterByD = {};
  if (masterRevData && masterRevData.members) {
    for (var ri = 0; ri < masterRevData.members.length; ri++) {
      var rm = masterRevData.members[ri];
      masterByD[rm.name] = rm;
    }
  }

  // マスターから失注・継続等のステータス別カウントも取得
  var masterStats = calcMasterStats_(settings.month, settings.year);
  var statsByD = {};
  if (masterStats) {
    for (var si = 0; si < masterStats.length; si++) {
      statsByD[masterStats[si].name] = masterStats[si];
    }
  }

  // カレンダーから休みデータ取得
  var holidayRaw = { counts: {}, byDate: {} };
  try { holidayRaw = getHolidayData_(settings.month, settings.year); } catch(e) {}
  var holidayData = holidayRaw.counts || {};
  var holidayByDate = holidayRaw.byDate || {};

  var members = [];

  for (var i = 0; i < activeMembers.length; i++) {
    var m = activeMembers[i];
    var masterDName = NICK_TO_MASTER_D_[m.name] || NICK_TO_MASTER_D_[m.displayName] || m.name;
    var mm = masterByD[masterDName] || { revenue: 0, deals: 0, closed: 0, closeRate: 0 };
    var ms = statsByD[masterDName] || { deals: 0, closed: 0, conClosed: 0, onClosed: 0, lost: 0, lostCont: 0, cont: 0, sales: 0 };
    var prev = prevMonthData[m.name] || prevMonthData[m.displayName] || { revenue: 0, deals: 0, closed: 0, closeRate: 0 };

    var closeRate = mm.deals > 0 ? round1_((mm.closed / mm.deals) * 100) : 0;
    var avgPrice = mm.closed > 0 ? round1_(mm.revenue / mm.closed) : 0;

    members.push({
      name: m.displayName,
      icon: m.iconUrl,
      deals: mm.deals,
      closed: mm.closed,
      revenue: mm.revenue,
      closeRate: closeRate,
      fundedDeals: mm.deals,
      sales: ms.sales,
      avgPrice: avgPrice,
      coRevenue: ms.conClosed,
      coAmount: 0,
      creditCard: 0,
      shinpan: 0,
      cbs: ms.cbsTotal > 0 ? ms.cbsApproved + '/' + ms.cbsTotal : '-',
      lifety: ms.lfTotal > 0 ? ms.lfApproved + '/' + ms.lfTotal : '-',
      lifetyRate: ms.lfTotal > 0 ? round1_((ms.lfApproved / ms.lfTotal) * 100) : '-',
      cbsRate: ms.cbsTotal > 0 ? round1_((ms.cbsApproved / ms.cbsTotal) * 100) : '-',
      onClosed: ms.onClosed,
      lost: ms.lost,
      lostCont: ms.lostCont,
      cont: ms.cont,
      prevRevenue: prev.revenue,
      diffRevenue: round1_(mm.revenue - prev.revenue),
      prevDeals: prev.deals,
      diffDeals: mm.deals - prev.deals,
      prevClosed: prev.closed,
      diffClosed: mm.closed - prev.closed,
      prevCloseRate: prev.closeRate,
      diffCloseRate: round1_(closeRate - prev.closeRate),
      holidays: holidayData[m.displayName] || holidayData[m.name] || 0
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

  // holidayByDateにアイコン情報を付与
  var holidayByDateWithIcon = {};
  for (var dk in holidayByDate) {
    holidayByDateWithIcon[dk] = [];
    for (var hi = 0; hi < holidayByDate[dk].length; hi++) {
      var h = holidayByDate[dk][hi];
      holidayByDateWithIcon[dk].push({
        name: h.name,
        type: h.type,
        icon: iconUrl_(h.name)
      });
    }
  }

  // 日別プッシュ数（全メンバー合算）
  var dailyPushes = {};
  try {
    dailyPushes = getDailyPushCounts_(ss, activeMembers, settings);
  } catch(e) {}

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
    holidayByDate: holidayByDateWithIcon,
    dailyPushes: dailyPushes,
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
  var prev = prevMonthData[m.name] || prevMonthData[m.displayName] || { revenue: 0, deals: 0, closed: 0, closeRate: 0 };
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
 * 前月データをマスターシートから取得
 * @returns {Object} { v2Name: { revenue, deals, closed, closeRate } }
 */
function getPrevMonthFromMaster_(settings) {
  var prevMonth = settings.month === 1 ? 12 : settings.month - 1;
  var prevYear = settings.month === 1 ? settings.year - 1 : settings.year;

  var data = calcRevenueFromMaster(prevMonth, prevYear);
  if (!data || !data.members) return {};

  var result = {};
  for (var i = 0; i < data.members.length; i++) {
    var m = data.members[i];
    result[m.name] = {
      revenue: m.revenue,
      deals: m.deals,
      closed: m.closed,
      closeRate: m.closeRate
    };
  }
  return result;
}

/**
 * 日別プッシュ数（全メンバー合算）をマスターシートから集計
 * C列=日付, D列=担当者 の行数をカウント
 */
function getDailyPushCounts_(ss, activeMembers, settings) {
  var totals = {};
  var byMember = {};
  var tz = Session.getScriptTimeZone();
  var year = settings.year;
  var month = settings.month;

  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  if (!master) return { totals: totals, byMember: byMember };

  var lastRow = master.getLastRow();
  if (lastRow <= 1) return { totals: totals, byMember: byMember };
  var data = master.getRange(2, 1, lastRow - 1, 12).getValues();

  // メンバー名→表示名マップ
  var displayMap = {};
  for (var i = 0; i < activeMembers.length; i++) {
    var dName = NICK_TO_MASTER_D_[activeMembers[i].name] || NICK_TO_MASTER_D_[activeMembers[i].displayName] || activeMembers[i].name;
    displayMap[dName] = activeMembers[i].displayName;
  }

  // 日付別・担当者別にカウント
  var countMap = {}; // { dateKey: { displayName: count } }
  for (var r = 0; r < data.length; r++) {
    var row = data[r];
    var d = String(row[3] || '').trim();
    if (!d || d === 'テスト') continue;
    var dispName = displayMap[d];
    if (!dispName) continue;

    var c = row[2];
    if (!rev_isTargetMonth_(c, year, month, tz)) continue;

    // C列をDateオブジェクトに変換してdateKeyを生成
    var dateObj;
    if (c instanceof Date) {
      dateObj = c;
    } else {
      var s = String(c).trim();
      if (!s) continue;
      var mt = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
      if (mt) {
        dateObj = new Date(parseInt(mt[1]), parseInt(mt[2]) - 1, parseInt(mt[3]));
      } else {
        dateObj = new Date(s);
        if (isNaN(dateObj.getTime())) continue;
      }
    }
    var dateKey = Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');

    if (!countMap[dateKey]) countMap[dateKey] = {};
    countMap[dateKey][dispName] = (countMap[dateKey][dispName] || 0) + 1;
  }

  // totals / byMember に変換
  for (var dk in countMap) {
    var dayTotal = 0;
    var members = [];
    for (var name in countMap[dk]) {
      var cnt = countMap[dk][name];
      dayTotal += cnt;
      members.push({ name: name, count: cnt });
    }
    totals[dk] = dayTotal;
    members.sort(function(a, b) { return b.count - a.count; });
    byMember[dk] = members;
  }

  return { totals: totals, byMember: byMember };
}

/**
 * 新構造: 着金速報データを日別入力シートから取得
 */
// マスターD列名→表示名の逆引きマップ
var MASTER_D_TO_NICK_ = {};
(function() {
  for (var nick in NICK_TO_MASTER_D_) {
    var d = NICK_TO_MASTER_D_[nick];
    if (!MASTER_D_TO_NICK_[d]) MASTER_D_TO_NICK_[d] = nick;
  }
})();

function getPaymentNews_new_(ss, activeMembers) {
  var settings = getGlobalSettings_(ss);
  var year = settings.year;
  var month = settings.month;
  var tz = 'Asia/Tokyo';

  // アクティブメンバーの表示名・アイコンマップ（D列名→表示名/icon）
  var memberInfo = {};
  for (var mi = 0; mi < activeMembers.length; mi++) {
    var am = activeMembers[mi];
    var dName = NICK_TO_MASTER_D_[am.name] || NICK_TO_MASTER_D_[am.displayName] || am.name;
    memberInfo[dName] = { displayName: am.displayName, icon: am.iconUrl };
  }

  // マスターシートから着金ブロックを読み取り
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  if (!master) return [];
  var lastRow = master.getLastRow();
  if (lastRow <= 1) return [];
  var data = master.getRange(2, 1, lastRow - 1, 88).getValues();

  // 着金ブロック定義: {amt: 金額列(0-indexed), date: 日付列(0-indexed)}
  var blocks = [
    {amt:15, date:2},   // P
    {amt:21, date:2},   // V
    {amt:29, date:2},   // AD
    {amt:41, date:38},  // AP
    {amt:48, date:45},  // AW
    {amt:55, date:52},  // BD
    {amt:66, date:63},  // BO
    {amt:73, date:70},  // BV
    {amt:80, date:77},  // CC
    {amt:87, date:84}   // CJ
  ];

  var news = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var l = String(row[11] || '').trim();
    if (l !== '成約' && l !== '成約➔CO') continue;

    var d = String(row[3] || '').trim();
    if (!d || d === 'テスト') continue;
    var info = memberInfo[d];
    if (!info) continue;

    for (var bi = 0; bi < blocks.length; bi++) {
      var blk = blocks[bi];
      var dateVal = row[blk.date];
      if (!(dateVal instanceof Date)) continue;
      if (dateVal.getFullYear() !== year || (dateVal.getMonth() + 1) !== month) continue;

      var amt = rev_parseNum_(row[blk.amt]);
      if (amt <= 0) continue;

      news.push({
        date: Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd'),
        dateShort: (dateVal.getMonth() + 1) + '/' + dateVal.getDate(),
        name: info.displayName,
        icon: info.icon,
        amount: round1_(amt)
      });
    }
  }

  news.sort(function(a, b) {
    if (a.date > b.date) return -1;
    if (a.date < b.date) return 1;
    return b.amount - a.amount;
  });

  return news;
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

/**
 * 外部スプレッドシートのデータを検査
 */
function inspectExtSheet_(ssid, gid, startRow, endRow) {
  if (!ssid) return { error: 'ssid required' };
  var ss = SpreadsheetApp.openById(ssid);
  var sheet = null;
  if (gid) {
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (String(sheets[i].getSheetId()) === String(gid)) {
        sheet = sheets[i];
        break;
      }
    }
  }
  if (!sheet) sheet = ss.getSheets()[0];

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var sr = startRow ? parseInt(startRow) : 1;
  var er = endRow ? parseInt(endRow) : Math.min(sr + 9, lastRow);
  er = Math.min(er, lastRow);

  var info = {
    name: sheet.getName(),
    sheetId: sheet.getSheetId(),
    totalRows: lastRow,
    totalCols: lastCol,
    allSheets: ss.getSheets().map(function(s) { return { name: s.getName(), gid: s.getSheetId(), rows: s.getLastRow(), cols: s.getLastColumn() }; })
  };

  if (lastRow < sr) return { info: info, data: [], formulas: [] };

  var range = sheet.getRange(sr, 1, er - sr + 1, Math.min(lastCol, 80));
  var vals = range.getValues();
  var fmls = range.getFormulas();

  var rows = [];
  for (var r = 0; r < vals.length; r++) {
    var row = {};
    for (var c = 0; c < vals[r].length; c++) {
      var v = vals[r][c];
      var f = fmls[r][c];
      if (v !== '' && v !== null && v !== undefined || f) {
        var colLetter = '';
        var cn = c + 1;
        while (cn > 0) { colLetter = String.fromCharCode(((cn - 1) % 26) + 65) + colLetter; cn = Math.floor((cn - 1) / 26); }
        row[colLetter] = f ? { v: String(v).substring(0, 80), f: f.substring(0, 120) } : String(v).substring(0, 80);
      }
    }
    rows.push({ row: sr + r, data: row });
  }

  return { info: info, rows: rows };
}
