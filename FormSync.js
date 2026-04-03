// ============================================
// FormSync.js — SMCマスター ← フォーム回答 同期・修復
//
// ■ 運用ルール
//   - syncRecentFormToMaster: 通常運用（5分トリガー）。自動で直近30件を同期
//   - repairSMCMaster: 手動修復専用。自動運用しない。API手動実行のみ
//
// ■ 重複判定キー（全関数共通）
//   formatDate(A列, tz, 'yyyy/MM/dd HH:mm:ss') + '|' + R列（本名）
//   ※ A列のみだと同秒に別人の商談がある場合に欠損するため
//
// ■ getLastRow() を使わない理由
//   GASの getLastRow() は空フォーマット行も含むため実データ行数とずれる
//   代わりに formSheet.getRange('A:A').getValues().filter(String).length を使用
//
// ■ FORM_SYNC_PAUSED フラグ
//   ScriptProperties の FORM_SYNC_PAUSED=true で sync/repair を即停止
//   API: pauseFormSync / resumeFormSync / formSyncStatus
// ============================================

var SMC_SS_ID = '1KxHeLmrpdaw1IUhBaQ46UWSHu-8SCRZqcrHOE2hMwDo';
var SMC_MASTER_SHEET = '👑商談マスターデータ';
var SMC_FORM_SHEET = 'フォーム回答シート';

// フォーム→マスター カラムマッピング (0-indexed)
// マスターにY(24)=ライフティ、Z(25)=CBSが手動挿入されて2列ずれ
// フォームにAC(28)都道府県〜AJ(35)継続分着金の追加列あり
var FORM_TO_MASTER_MAP = {};
// A〜X: 1:1対応
for (var _fi = 0; _fi <= 23; _fi++) {
  FORM_TO_MASTER_MAP[_fi] = _fi;
}
// Y以降: マスターにライフティ(24)・CBS(25)が挿入 → +2ずれ
FORM_TO_MASTER_MAP[24] = 26; // フォームY(支払方法③)     → マスターAA(26)
FORM_TO_MASTER_MAP[25] = 27; // フォームZ(代理店名)       → マスターAB(27)
FORM_TO_MASTER_MAP[26] = 29; // フォームAA(支払③金額)     → マスターAD(29)
FORM_TO_MASTER_MAP[27] = 30; // フォームAB(郵送確認)      → マスターAE(30)
// フォームAC(28)都道府県〜AJ(35)継続分着金 → マスターに該当列なし（スキップ）
FORM_TO_MASTER_MAP[34] = 27; // フォームAI(代理店Tiktokその他) → マスターAB(27) ※Z(25)と同じ列
FORM_TO_MASTER_MAP[36] = 32; // フォームAK(着金率)        → マスターAG(32)
FORM_TO_MASTER_MAP[37] = 33; // フォームAL(ステータス)      → マスターAH(33)
FORM_TO_MASTER_MAP[38] = 34; // フォームAM(契約アドレス)    → マスターAI(34)
FORM_TO_MASTER_MAP[39] = 35; // フォームAN(契約日)        → マスターAJ(35)

// ============================================
// 5分トリガー同期（通常運用）
// ============================================

/**
 * 5分トリガー用: フォーム直近30件をマスターに同期
 * 判定キー: formatDate(A) + '|' + R列
 */
function syncRecentFormToMaster() {
  if (PropertiesService.getScriptProperties().getProperty('FORM_SYNC_PAUSED') === 'true') return;
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var masterSheet = smc.getSheetByName(SMC_MASTER_SHEET);
  var formSheet = smc.getSheetByName(SMC_FORM_SHEET);
  if (!masterSheet || !formSheet) return;

  // getLastRow()は空フォーマット行も含むので、A列の最終データ行を使う
  var formColA = formSheet.getRange('A:A').getValues();
  var formLastRow = formColA.filter(String).length;
  var masterLastRow = masterSheet.getLastRow();
  if (formLastRow <= 1 || masterLastRow <= 1) return;

  // 直近30件のフォーム回答
  var lookback = Math.min(30, formLastRow - 1);
  var formStart = formLastRow - lookback + 1;
  var formData = formSheet.getRange(formStart, 1, lookback, 41).getValues();

  // マスター全件のA列+R列で重複判定（formatDate + 本名）
  var tz = Session.getScriptTimeZone();
  var masterAR = masterSheet.getRange(2, 1, masterLastRow - 1, 18).getValues();
  var existingKeys = {};
  for (var i = 0; i < masterAR.length; i++) {
    var ts = masterAR[i][0];
    var mR = String(masterAR[i][17] || '').trim();
    var mKey;
    if (ts instanceof Date) {
      mKey = Utilities.formatDate(ts, tz, 'yyyy/MM/dd HH:mm:ss');
    } else if (ts) {
      mKey = String(ts).trim();
    } else {
      continue;
    }
    existingKeys[mKey + '|' + mR] = true;
  }

  var added = 0;

  for (var fi = 0; fi < formData.length; fi++) {
    var formRow = formData[fi];
    var formTs = formRow[0];
    if (!formTs) continue;

    var formA;
    if (formTs instanceof Date) {
      formA = Utilities.formatDate(formTs, tz, 'yyyy/MM/dd HH:mm:ss');
    } else {
      formA = String(formTs).trim();
    }
    var formR = String(formRow[17] || '').trim();
    var formKey = formA + '|' + formR;

    if (existingKeys[formKey]) continue;

    var newRow = new Array(41);
    for (var x = 0; x < 41; x++) newRow[x] = '';
    for (var fc in FORM_TO_MASTER_MAP) {
      var mc = FORM_TO_MASTER_MAP[fc];
      var idx = parseInt(fc);
      if (idx < formRow.length && formRow[idx] !== '' && formRow[idx] !== null) {
        newRow[mc] = formRow[idx];
      }
    }

    // Dateをフォーマット文字列に変換してから書き込み
    if (newRow[0] instanceof Date) {
      newRow[0] = Utilities.formatDate(newRow[0], tz, 'yyyy/MM/dd HH:mm:ss');
    }
    if (newRow[2] instanceof Date) {
      newRow[2] = Utilities.formatDate(newRow[2], tz, 'yyyy/MM/dd');
    }

    masterSheet.appendRow(newRow);
    existingKeys[formKey] = true;
    added++;
  }

  if (added > 0) {
    SpreadsheetApp.flush();
  }
}

// ============================================
// サブマスターシート作成（数式参照方式）
// ============================================

/**
 * サブマスターシートを作成し、マスターを数式で参照する
 * 手動で1回だけ実行すればOK。以降はリアルタイム自動反映。
 */
function createSubMasterSheet() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var subSheet = smc.getSheetByName('サブマスター');
  if (!subSheet) {
    subSheet = smc.insertSheet('サブマスター');
  }
  subSheet.clearContents();
  subSheet.getRange('A1').setFormula("={'👑商談マスターデータ'!A:CX}");
  SpreadsheetApp.flush();
}

/**
 * マスターA1ヘッダー修正（1回実行用）
 */
function fixMasterA1() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var masterSheet = smc.getSheetByName(SMC_MASTER_SHEET);
  masterSheet.getRange('A1').setValue('タイムスタンプ');
  SpreadsheetApp.flush();
}

// ============================================
// CDEキーベース upsert 同期（検証用・本番未使用）
// ============================================

/**
 * 同一 C,D,E の行は append せず update する
 * - 新規なら append
 * - 既存なら必要項目だけ上書き
 * - 空欄で既存値は潰さない
 * - L列は「成約」系を優先
 */
function syncRecentFormToMaster_upsertByCDE() {
  if (PropertiesService.getScriptProperties().getProperty('FORM_SYNC_PAUSED') === 'true') {
    return { paused: true };
  }

  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var masterSheet = smc.getSheetByName(SMC_MASTER_SHEET);
  var formSheet = smc.getSheetByName(SMC_FORM_SHEET);
  if (!masterSheet || !formSheet) return { error: 'sheet not found' };

  var tz = Session.getScriptTimeZone();

  var formColA = formSheet.getRange('A:A').getValues();
  var formLastRow = formColA.filter(String).length;
  var masterLastRow = masterSheet.getLastRow();
  if (formLastRow <= 1) return { error: 'no form data' };

  var lookback = Math.min(30, formLastRow - 1);
  var formStart = formLastRow - lookback + 1;
  var formData = formSheet.getRange(formStart, 1, lookback, 41).getValues();

  var masterData = [];
  if (masterLastRow > 1) {
    masterData = masterSheet.getRange(2, 1, masterLastRow - 1, 41).getValues();
  }

  var existingMap = {};
  for (var i = 0; i < masterData.length; i++) {
    var mRow = masterData[i];
    var cVal = normalizeKeyPart_(mRow[2]);
    var dVal = normalizeKeyPart_(mRow[3]);
    var eVal = normalizeKeyPart_(mRow[4]);
    if (!cVal && !dVal && !eVal) continue;

    var key = buildCDEKey_(cVal, dVal, eVal);
    if (!existingMap[key]) {
      existingMap[key] = { rowNumber: i + 2, values: mRow };
    } else {
      var currentBest = existingMap[key].values;
      if (compareMasterRowsForWinner_(mRow, currentBest, tz) > 0) {
        existingMap[key] = { rowNumber: i + 2, values: mRow };
      }
    }
  }

  var appended = 0;
  var updated = 0;
  var skipped = 0;
  var touchedRows = [];

  for (var fi = 0; fi < formData.length; fi++) {
    var formRow = formData[fi];
    if (!formRow[0]) continue;

    var newRow = buildMasterRowFromFormRow_(formRow, tz);
    var cKey = normalizeKeyPart_(newRow[2]);
    var dKey = normalizeKeyPart_(newRow[3]);
    var eKey = normalizeKeyPart_(newRow[4]);
    if (!cKey && !dKey && !eKey) { skipped++; continue; }

    var key = buildCDEKey_(cKey, dKey, eKey);

    if (!existingMap[key]) {
      masterSheet.appendRow(newRow);
      appended++;
      var newMasterRowNumber = masterSheet.getLastRow();
      existingMap[key] = { rowNumber: newMasterRowNumber, values: newRow.slice() };
      touchedRows.push({ type: 'append', row: newMasterRowNumber, key: key });
    } else {
      var existing = existingMap[key];
      var merged = mergeMasterRows_(existing.values, newRow, tz);
      if (!rowsEqual_(existing.values, merged)) {
        masterSheet.getRange(existing.rowNumber, 1, 1, 41).setValues([merged]);
        updated++;
        existingMap[key] = { rowNumber: existing.rowNumber, values: merged };
        touchedRows.push({ type: 'update', row: existing.rowNumber, key: key });
      } else {
        skipped++;
      }
    }
  }

  SpreadsheetApp.flush();
  return {
    ok: true, lookback: lookback, formStart: formStart,
    appended: appended, updated: updated, skipped: skipped,
    touchedRows: touchedRows.slice(0, 20)
  };
}

// --- ヘルパー関数（_サフィックスで名前衝突回避）---

function buildMasterRowFromFormRow_(formRow, tz) {
  var newRow = new Array(41);
  for (var i = 0; i < 41; i++) newRow[i] = '';
  for (var fc in FORM_TO_MASTER_MAP) {
    var mc = FORM_TO_MASTER_MAP[fc];
    var idx = parseInt(fc, 10);
    if (idx < formRow.length && formRow[idx] !== '' && formRow[idx] !== null) {
      newRow[mc] = formRow[idx];
    }
  }
  if (newRow[0] instanceof Date) {
    newRow[0] = Utilities.formatDate(newRow[0], tz, 'yyyy/MM/dd HH:mm:ss');
  }
  if (newRow[2] instanceof Date) {
    newRow[2] = Utilities.formatDate(newRow[2], tz, 'yyyy/MM/dd');
  }
  return newRow;
}

function mergeMasterRows_(oldRow, newRow, tz) {
  var merged = oldRow.slice();
  for (var i = 0; i < merged.length; i++) {
    var oldVal = oldRow[i];
    var newVal = newRow[i];
    if (i === 11) { merged[i] = chooseBetterStatus_(oldVal, newVal); continue; }
    if (i === 0) { merged[i] = chooseLaterTimestamp_(oldVal, newVal, tz); continue; }
    if (isBlank_(oldVal) && !isBlank_(newVal)) { merged[i] = newVal; continue; }
    if (!isBlank_(oldVal) && !isBlank_(newVal)) {
      // P〜X列(15-23): マスター側の手動修正を保護（既存値を優先）
      if (i >= 15 && i <= 23) { continue; }
      if (i === 24 || i === 25 || i === 26) {
        merged[i] = String(newVal).length > String(oldVal).length ? newVal : oldVal;
      } else {
        merged[i] = newVal;
      }
    }
  }
  return merged;
}

function compareMasterRowsForWinner_(rowA, rowB, tz) {
  var aSeiyaku = isSeiyakuStatus_(String(rowA[11] || ''));
  var bSeiyaku = isSeiyakuStatus_(String(rowB[11] || ''));
  if (aSeiyaku && !bSeiyaku) return 1;
  if (!aSeiyaku && bSeiyaku) return -1;
  var aTs = parseTimestamp_(rowA[0], tz);
  var bTs = parseTimestamp_(rowB[0], tz);
  if (aTs && bTs) {
    if (aTs.getTime() > bTs.getTime()) return 1;
    if (aTs.getTime() < bTs.getTime()) return -1;
  }
  return 0;
}

function chooseBetterStatus_(oldVal, newVal) {
  var oldStr = String(oldVal || '');
  var newStr = String(newVal || '');
  if (isBlank_(oldStr) && !isBlank_(newStr)) return newVal;
  if (!isBlank_(oldStr) && isBlank_(newStr)) return oldVal;
  var oldRank = statusRank_(oldStr);
  var newRank = statusRank_(newStr);
  // 低い状態への巻き戻し禁止
  if (newRank < oldRank) return oldVal;
  // 同格以上なら新しい方を採用
  return newVal;
}

function chooseLaterTimestamp_(oldVal, newVal, tz) {
  var oldTs = parseTimestamp_(oldVal, tz);
  var newTs = parseTimestamp_(newVal, tz);
  if (!oldTs && newTs) return newVal;
  if (oldTs && !newTs) return oldVal;
  if (!oldTs && !newTs) return !isBlank_(newVal) ? newVal : oldVal;
  return newTs.getTime() >= oldTs.getTime() ? newVal : oldVal;
}

function buildCDEKey_(c, d, e) { return [c, d, e].join('|'); }
function normalizeKeyPart_(v) { return String(v == null ? '' : v).trim(); }
function isBlank_(v) { return v === '' || v === null || typeof v === 'undefined'; }
function isSeiyakuStatus_(v) { var s = String(v || ''); return s.indexOf('成約') !== -1 || s.indexOf('CO') !== -1; }

/** L列の優先度 (高い数値 = 強い状態。低い方への巻き戻し禁止) */
function statusRank_(v) {
  var s = String(v || '').trim();
  if (!s) return 0;
  if (s.indexOf('成約') !== -1 && s.indexOf('CO') !== -1) return 4; // 成約➔CO
  if (s.indexOf('成約') !== -1) return 3;
  if (s.indexOf('継続') !== -1) return 2;
  if (s.indexOf('顧客情報') !== -1) return 1;
  if (s.indexOf('失注') !== -1) return 1;
  return 1; // その他
}

function parseTimestamp_(v, tz) {
  if (!v) return null;
  if (v instanceof Date) return v;
  var s = String(v).trim();
  if (!s) return null;
  var d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  var m = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})(?: (\d{1,2}):(\d{1,2}):(\d{1,2}))?$/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0));
}

function rowsEqual_(rowA, rowB) {
  if (rowA.length !== rowB.length) return false;
  for (var i = 0; i < rowA.length; i++) {
    if (String(rowA[i] || '') !== String(rowB[i] || '')) return false;
  }
  return true;
}

/**
 * 業務差分だけを見る比較。形式差分（toString vs formatDate）は無視。
 * 差分がある列名の配列を返す。空なら業務差分なし。
 */
function businessDiffCols_(oldRow, newRow, tz) {
  var diffs = [];
  for (var i = 0; i < Math.max(oldRow.length, newRow.length); i++) {
    var ov = oldRow[i] !== undefined ? oldRow[i] : '';
    var nv = newRow[i] !== undefined ? newRow[i] : '';
    if (String(ov || '') === String(nv || '')) continue;

    // A列(0): timestamp として同値なら無視
    if (i === 0) {
      var oTs = parseTimestamp_(ov, tz);
      var nTs = parseTimestamp_(nv, tz);
      if (oTs && nTs && Math.abs(oTs.getTime() - nTs.getTime()) < 2000) continue; // 2秒以内
      if (!oTs && !nTs) continue;
    }

    // C列(2): date として同値なら無視
    if (i === 2) {
      var oD = parseTimestamp_(ov, tz);
      var nD = parseTimestamp_(nv, tz);
      if (oD && nD && oD.toDateString() === nD.toDateString()) continue;
    }

    // それ以外: 空→空でないなら差分、空でない→空でないなら文字列比較
    diffs.push(i);
  }
  return diffs;
}

// ============================================
// 手動修復専用（repairSMCMaster）
// ※ 自動運用しない。API手動実行のみ。
// ============================================

/**
 * SMCマスターシートを手動修復（API: action=run&fn=repairSMCMaster）
 * Phase 1: ゴースト行削除（A空かつD空）
 * Phase 2: 重複行削除（同一formatDate(A)+R列）
 * Phase 3: フォーム直近200件から不足エントリを追加
 * 判定キー: formatDate(A) + '|' + R列
 */
function repairSMCMaster() {
  if (PropertiesService.getScriptProperties().getProperty('FORM_SYNC_PAUSED') === 'true') {
    return { paused: true };
  }
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var masterSheet = smc.getSheetByName(SMC_MASTER_SHEET);
  var formSheet = smc.getSheetByName(SMC_FORM_SHEET);
  if (!masterSheet || !formSheet) return { error: 'シートが見つかりません' };

  var results = { ghostsDeleted: 0, duplicatesDeleted: 0, added: 0, details: [] };

  // ======== Phase 1 & 2: ゴースト行・重複行の削除 ========
  var masterLastRow = masterSheet.getLastRow();
  if (masterLastRow <= 1) return results;

  var masterData = masterSheet.getRange(2, 1, masterLastRow - 1, 41).getValues();
  var seenTimes = {};      // getTime() → 最初の行index
  var rowsToDelete = [];   // 削除対象の行番号（1-indexed）

  for (var i = 0; i < masterData.length; i++) {
    var row = masterData[i];
    var ts = row[0];
    var d = String(row[3] || '').trim();
    var actualRow = i + 2;

    // Phase 1: A列空かつD列空 → ゴースト行
    if (!ts && !d) {
      var hasAnyData = false;
      for (var ci = 7; ci < 18; ci++) {
        if (String(row[ci] || '').trim()) { hasAnyData = true; break; }
      }
      if (hasAnyData) {
        rowsToDelete.push(actualRow);
        results.ghostsDeleted++;
        results.details.push('ゴースト削除: row ' + actualRow);
      }
      continue;
    }

    if (!ts) continue;

    // Phase 2: タイムスタンプ重複チェック（getTime数値）
    var timeNum;
    if (ts instanceof Date) {
      timeNum = ts.getTime();
    } else {
      var dd = new Date(ts);
      if (isNaN(dd.getTime())) continue;
      timeNum = dd.getTime();
    }

    if (seenTimes[timeNum] !== undefined) {
      rowsToDelete.push(actualRow);
      results.duplicatesDeleted++;
      results.details.push('重複削除: row ' + actualRow + ' (ts=' + timeNum + ')');
    } else {
      seenTimes[timeNum] = i;
    }
  }

  // 下から削除（行番号ずれ防止）
  if (rowsToDelete.length > 0) {
    rowsToDelete.sort(function(a, b) { return b - a; });
    for (var di = 0; di < rowsToDelete.length; di++) {
      masterSheet.deleteRow(rowsToDelete[di]);
    }
    results.details.push('合計 ' + rowsToDelete.length + ' 行削除');
    SpreadsheetApp.flush();
  }

  // ======== Phase 3: フォームから不足エントリを追加 ========
  var tz2 = Session.getScriptTimeZone();
  masterLastRow = masterSheet.getLastRow();
  var existingKeys = {};
  if (masterLastRow > 1) {
    var masterAR2 = masterSheet.getRange(2, 1, masterLastRow - 1, 18).getValues();
    for (var mi = 0; mi < masterAR2.length; mi++) {
      var mTs = masterAR2[mi][0];
      var mR2 = String(masterAR2[mi][17] || '').trim();
      var mK;
      if (mTs instanceof Date) {
        mK = Utilities.formatDate(mTs, tz2, 'yyyy/MM/dd HH:mm:ss');
      } else if (mTs) {
        mK = String(mTs).trim();
      } else {
        continue;
      }
      existingKeys[mK + '|' + mR2] = true;
    }
  }

  var formColA = formSheet.getRange('A:A').getValues();
  var formLastRow = formColA.filter(String).length;
  if (formLastRow <= 1) {
    SpreadsheetApp.flush();
    return results;
  }

  var fLookback = Math.min(200, formLastRow - 1);
  var formStartRow = formLastRow - fLookback + 1;
  var formData = formSheet.getRange(formStartRow, 1, fLookback, 41).getValues();

  for (var fi = 0; fi < formData.length; fi++) {
    var fRow = formData[fi];
    var fTs = fRow[0];
    if (!fTs) continue;

    var fA;
    if (fTs instanceof Date) {
      fA = Utilities.formatDate(fTs, tz2, 'yyyy/MM/dd HH:mm:ss');
    } else {
      fA = String(fTs).trim();
    }
    var fR = String(fRow[17] || '').trim();
    var fKey = fA + '|' + fR;

    if (existingKeys[fKey]) continue;

    var newRow = new Array(41);
    for (var x = 0; x < 41; x++) newRow[x] = '';
    for (var fCol in FORM_TO_MASTER_MAP) {
      var mCol = FORM_TO_MASTER_MAP[fCol];
      var idx = parseInt(fCol);
      if (idx < fRow.length && fRow[idx] !== '' && fRow[idx] !== null) {
        newRow[mCol] = fRow[idx];
      }
    }
    if (newRow[0] instanceof Date) {
      newRow[0] = Utilities.formatDate(newRow[0], tz2, 'yyyy/MM/dd HH:mm:ss');
    }
    if (newRow[2] instanceof Date) {
      newRow[2] = Utilities.formatDate(newRow[2], 'Asia/Tokyo', 'yyyy/MM/dd');
    }
    masterSheet.appendRow(newRow);
    existingKeys[fKey] = true;
    results.added++;
    var fName = String(fRow[3] || '') + '/' + String(fRow[4] || '');
    results.details.push('追加: Form row ' + (formStartRow + fi) + ' (' + fName + ')');
  }

  // ソート
  var newLastRow = masterSheet.getLastRow();
  if (newLastRow > 2) {
    masterSheet.getRange(2, 1, newLastRow - 1, masterSheet.getLastColumn()).sort({ column: 1, ascending: true });
  }

  ensureConditionalFormatting_(masterSheet);
  SpreadsheetApp.flush();
  return results;
}

/**
 * 条件付き書式（CO→黄、成約→赤）
 */
function ensureConditionalFormatting_(sheet) {
  var rules = sheet.getConditionalFormatRules();
  if (rules.length >= 2) return;

  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());

  var newRules = [];
  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=FIND("CO",$L2)')
    .setBackground('#FFFF99')
    .setRanges([range])
    .build());
  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=FIND("成約",$L2)')
    .setBackground('#FF9999')
    .setRanges([range])
    .build());

  sheet.setConditionalFormatRules(newRules);
}

/**
 * dry-run: A+R主キー + CDE副キーのupsertロジック。書き込みしない。
 * ① A+R一致 → update
 * ② A+R不一致 → CDE一致（正規化後） → update
 * ③ 両方なし → append
 */
function dryRunSyncUpsertByCDE() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var masterSheet = smc.getSheetByName(SMC_MASTER_SHEET);
  var formSheet = smc.getSheetByName(SMC_FORM_SHEET);
  if (!masterSheet || !formSheet) return { error: 'sheet not found' };

  var tz = Session.getScriptTimeZone();

  var formColA = formSheet.getRange('A:A').getValues();
  var formLastRow = formColA.filter(String).length;
  var masterLastRow = masterSheet.getLastRow();
  if (formLastRow <= 1) return { error: 'no form data' };

  var lookback = Math.min(30, formLastRow - 1);
  var formStart = formLastRow - lookback + 1;
  var formData = formSheet.getRange(formStart, 1, lookback, 41).getValues();

  var masterData = [];
  if (masterLastRow > 1) {
    masterData = masterSheet.getRange(2, 1, masterLastRow - 1, 41).getValues();
  }

  // 主キーMap: A+R
  var primaryMap = {};
  // 副キーMap: normalizedC + normalizedD + E
  var secondaryMap = {};

  for (var i = 0; i < masterData.length; i++) {
    var mRow = masterData[i];
    var mA = formatA_(mRow[0], tz);
    var mR = normalizeKeyPart_(mRow[17]);
    var mC = normalizeC_(mRow[2], tz);
    var mD = normalizeD_(mRow[3]);
    var mE = normalizeKeyPart_(mRow[4]);

    if (mA && mR) {
      var pk = mA + '|' + mR;
      if (!primaryMap[pk]) {
        primaryMap[pk] = { rowNumber: i + 2, values: mRow };
      }
    }
    if (mC && mD && mE) {
      var sk = mC + '|' + mD + '|' + mE;
      if (!secondaryMap[sk]) {
        secondaryMap[sk] = { rowNumber: i + 2, values: mRow };
      } else {
        if (compareMasterRowsForWinner_(mRow, secondaryMap[sk].values, tz) > 0) {
          secondaryMap[sk] = { rowNumber: i + 2, values: mRow };
        }
      }
    }
  }

  var wouldAppend = [];
  var wouldUpdateAR = [];
  var wouldUpdateCDE = [];
  var skipCount = 0;
  var formatOnlySkip = 0;
  var lReversalBlocked = [];

  for (var fi = 0; fi < formData.length; fi++) {
    var formRow = formData[fi];
    if (!formRow[0]) continue;

    var newRow = buildMasterRowFromFormRow_(formRow, tz);
    var fA = formatA_(newRow[0], tz);
    var fR = normalizeKeyPart_(newRow[17]);
    var fC = normalizeC_(newRow[2], tz);
    var fD = normalizeD_(newRow[3]);
    var fE = normalizeKeyPart_(newRow[4]);

    // ① A+R一致
    var pk = fA + '|' + fR;
    if (fA && fR && primaryMap[pk]) {
      var existing = primaryMap[pk];
      var merged = mergeMasterRows_(existing.values, newRow, tz);

      // L列逆方向更新チェック
      var oldL = String(existing.values[11] || '');
      var newL = String(newRow[11] || '');
      if (statusRank_(newL) < statusRank_(oldL)) {
        lReversalBlocked.push({
          masterRow: existing.rowNumber, formRow: formStart + fi,
          oldL: oldL.substring(0, 20), newL: newL.substring(0, 20),
          oldRank: statusRank_(oldL), newRank: statusRank_(newL)
        });
      }

      // 業務差分チェック（形式差分を除外）
      var bizDiffs = businessDiffCols_(existing.values, merged, tz);
      if (bizDiffs.length === 0) {
        formatOnlySkip++;
        continue;
      }

      var diffs = buildDiffList_(existing.values, merged);
      wouldUpdateAR.push({
        matchType: 'A+R',
        masterRow: existing.rowNumber, formRow: formStart + fi,
        a: fA, c: fC, d: String(newRow[3] || ''), e: fE, l: String(merged[11] || '').substring(0,20), r: fR,
        key: pk, diffs: diffs, bizDiffCols: bizDiffs
      });
      continue;
    }

    // ② CDE一致（正規化後）
    var sk = fC + '|' + fD + '|' + fE;
    if (fC && fD && fE && secondaryMap[sk]) {
      var existing2 = secondaryMap[sk];
      var merged2 = mergeMasterRows_(existing2.values, newRow, tz);

      var bizDiffs2 = businessDiffCols_(existing2.values, merged2, tz);
      if (bizDiffs2.length === 0) {
        formatOnlySkip++;
        continue;
      }

      var diffs2 = buildDiffList_(existing2.values, merged2);
      wouldUpdateCDE.push({
        matchType: 'CDE',
        masterRow: existing2.rowNumber, formRow: formStart + fi,
        a: fA, c: fC, d: String(newRow[3] || ''), e: fE, l: String(merged2[11] || '').substring(0,20), r: fR,
        key: sk, diffs: diffs2, bizDiffCols: bizDiffs2
      });
      continue;
    }

    // ③ 両方なし → append
    wouldAppend.push({
      formRow: formStart + fi,
      a: fA, c: fC, d: String(newRow[3] || ''), e: fE,
      l: String(newRow[11] || '').substring(0,20), r: fR,
      primaryKey: pk, secondaryKey: sk
    });
  } // end for loop

  return {
    formStart: formStart, lookback: lookback,
      wouldAppendCount: wouldAppend.length,
      wouldUpdateARCount: wouldUpdateAR.length,
      wouldUpdateCDECount: wouldUpdateCDE.length,
      skipCount: skipCount,
      formatOnlySkip: formatOnlySkip,
      lReversalBlocked: lReversalBlocked,
      wouldAppend: wouldAppend,
      wouldUpdateAR: wouldUpdateAR,
      wouldUpdateCDE: wouldUpdateCDE
    };
} // end dryRunSyncUpsertByCDE

/** A列をformatDate文字列に正規化 */
function formatA_(v, tz) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy/MM/dd HH:mm:ss');
  return String(v).trim();
}

/** C列を yyyy/MM/dd に正規化 */
function normalizeC_(v, tz) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy/MM/dd');
  var s = String(v).trim();
  var m = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})/);
  if (m) return m[1] + '/' + ('0' + m[2]).slice(-2) + '/' + ('0' + m[3]).slice(-2);
  return s;
}

/** D列を正規化: 括弧内除去 + trim */
function normalizeD_(v) {
  var s = String(v == null ? '' : v).trim();
  return s.replace(/[（(].+?[）)]/g, '').trim();
}

/** 差分列リスト生成 */
function buildDiffList_(oldRow, newRow) {
  var diffs = [];
  var cols = {A:0, C:2, D:3, E:4, L:11, R:17};
  for (var name in cols) {
    var ci = cols[name];
    var ov = String(oldRow[ci] || '').substring(0, 15);
    var nv = String(newRow[ci] || '').substring(0, 15);
    if (ov !== nv) diffs.push(name + ':' + ov + '→' + nv);
  }
  return diffs;
}

/**
 * デバッグログを取得
 */
function getSyncDebugLog() {
  return PropertiesService.getScriptProperties().getProperty('syncDebugLast') || 'no log';
}

/**
 * Step2デバッグ: フォーム最新1件の読み取り値とnewRow生成結果をログに出す
 * appendはしない。読み取りとマッピングだけ。
 */
function debugFormLatest() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var form = smc.getSheetByName(SMC_FORM_SHEET);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var tz = Session.getScriptTimeZone();

  // フォーム最新1件を取得
  var formColA = form.getRange('A:A').getValues();
  var formLastRow = formColA.filter(String).length;
  var formRow = form.getRange(formLastRow, 1, 1, 41).getValues()[0];

  var log = [];
  log.push('=== debugFormLatest ===');
  log.push('timestamp: ' + new Date().toISOString());
  log.push('formLastRow: ' + formLastRow);
  log.push('');

  // フォームA列の生値
  var formTs = formRow[0];
  log.push('--- フォーム最新1件 A列 ---');
  log.push('formTs raw: ' + String(formTs));
  log.push('formTs typeof: ' + typeof formTs);
  log.push('formTs instanceof Date: ' + (formTs instanceof Date));
  if (formTs instanceof Date) {
    log.push('formTs.getTime(): ' + formTs.getTime());
    log.push('formatDate: ' + Utilities.formatDate(formTs, tz, 'yyyy/MM/dd HH:mm:ss'));
  }
  log.push('');

  // 比較キー生成（v206と同じロジック）
  var formKey;
  if (formTs instanceof Date) {
    formKey = Utilities.formatDate(formTs, tz, 'yyyy/MM/dd HH:mm:ss');
  } else {
    formKey = String(formTs).trim();
  }
  log.push('比較キー: ' + formKey);
  log.push('');

  // FORM_TO_MASTER_MAPでnewRow生成
  var newRow = new Array(41);
  for (var x = 0; x < 41; x++) newRow[x] = '';
  for (var fc in FORM_TO_MASTER_MAP) {
    var mc = FORM_TO_MASTER_MAP[fc];
    var idx = parseInt(fc);
    if (idx < formRow.length && formRow[idx] !== '' && formRow[idx] !== null) {
      newRow[mc] = formRow[idx];
    }
  }

  log.push('--- newRow主要列（append前） ---');
  log.push('newRow[0] (A): ' + String(newRow[0]) + ' | type=' + typeof newRow[0] + ' | isDate=' + (newRow[0] instanceof Date));
  log.push('newRow[3] (D): ' + String(newRow[3]));
  log.push('newRow[4] (E): ' + String(newRow[4]));
  log.push('newRow[11](L): ' + String(newRow[11]));
  log.push('newRow[17](R): ' + String(newRow[17]));
  log.push('');

  // formatDate変換後
  if (newRow[0] instanceof Date) {
    var formatted = Utilities.formatDate(newRow[0], tz, 'yyyy/MM/dd HH:mm:ss');
    log.push('newRow[0] formatDate後: ' + formatted);
  } else {
    log.push('newRow[0] はDateではない: ' + String(newRow[0]));
  }
  log.push('');

  // マスター側の比較キーサンプル（末尾3件）
  var masterLast = master.getLastRow();
  var masterSample = master.getRange(masterLast - 2, 1, 3, 1).getValues();
  log.push('--- マスター末尾3件の比較キー ---');
  for (var mi = 0; mi < masterSample.length; mi++) {
    var mts = masterSample[mi][0];
    var mKey;
    if (mts instanceof Date) {
      mKey = Utilities.formatDate(mts, tz, 'yyyy/MM/dd HH:mm:ss');
    } else if (mts) {
      mKey = String(mts).trim();
    } else {
      mKey = '(empty)';
    }
    log.push('Master ' + (masterLast - 2 + mi) + ': raw=' + String(mts) + ' | isDate=' + (mts instanceof Date) + ' | key=' + mKey);
  }

  // existingKeysにformKeyが存在するか
  var masterColA = master.getRange(2, 1, masterLast - 1, 1).getValues();
  var existingKeys = {};
  for (var i = 0; i < masterColA.length; i++) {
    var ts = masterColA[i][0];
    if (ts instanceof Date) {
      existingKeys[Utilities.formatDate(ts, tz, 'yyyy/MM/dd HH:mm:ss')] = true;
    } else if (ts) {
      existingKeys[String(ts).trim()] = true;
    }
  }
  log.push('');
  log.push('existingKeys数: ' + Object.keys(existingKeys).length);
  log.push('formKey in existingKeys: ' + (existingKeys[formKey] ? 'YES (skip)' : 'NO (would add)'));

  // マスター修正プレビューに書き出し
  var logSheet = smc.getSheetByName('マスター修正プレビュー');
  if (logSheet) {
    logSheet.clear();
    for (var li = 0; li < log.length; li++) {
      logSheet.getRange(li + 1, 1).setValue(log[li]);
    }
  }

  return log.join('\n');
}

/**
 * dry-run: repairSMCMaster Phase3と同じロジックで追加対象を出す。appendしない。deleteしない。
 */
function repairDryRunSMCMaster() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var masterSheet = smc.getSheetByName(SMC_MASTER_SHEET);
  var formSheet = smc.getSheetByName(SMC_FORM_SHEET);
  if (!masterSheet || !formSheet) return { error: 'sheet not found' };

  var tz = Session.getScriptTimeZone();
  var masterLastRow = masterSheet.getLastRow();

  // マスターキーセット構築（Phase3と同じロジック）
  var existingKeys = {};
  if (masterLastRow > 1) {
    var masterAR = masterSheet.getRange(2, 1, masterLastRow - 1, 18).getValues();
    for (var mi = 0; mi < masterAR.length; mi++) {
      var mTs = masterAR[mi][0];
      var mR = String(masterAR[mi][17] || '').trim();
      var mK;
      if (mTs instanceof Date) {
        mK = Utilities.formatDate(mTs, tz, 'yyyy/MM/dd HH:mm:ss');
      } else if (mTs) {
        mK = String(mTs).trim();
      } else {
        continue;
      }
      existingKeys[mK + '|' + mR] = true;
    }
  }

  // フォーム直近200件（Phase3と同じ範囲）
  var formColA = formSheet.getRange('A:A').getValues();
  var formLastRow = formColA.filter(String).length;
  if (formLastRow <= 1) return { wouldAddCount: 0, wouldAdd: [], duplicateKeysInBatch: [] };

  var fLookback = Math.min(200, formLastRow - 1);
  var formStartRow = formLastRow - fLookback + 1;
  var formData = formSheet.getRange(formStartRow, 1, fLookback, 41).getValues();

  var wouldAdd = [];
  var seenKeys = {};

  for (var fi = 0; fi < formData.length; fi++) {
    var fRow = formData[fi];
    var fTs = fRow[0];
    if (!fTs) continue;

    var fA;
    if (fTs instanceof Date) {
      fA = Utilities.formatDate(fTs, tz, 'yyyy/MM/dd HH:mm:ss');
    } else {
      fA = String(fTs).trim();
    }
    var fR = String(fRow[17] || '').trim();
    var fKey = fA + '|' + fR;

    if (existingKeys[fKey]) continue;

    var d = String(fRow[3] || '');
    var e = String(fRow[4] || '').substring(0, 20);
    var l = String(fRow[11] || '').substring(0, 20);

    wouldAdd.push({
      formRow: formStartRow + fi,
      a: fA, d: d, e: e, l: l, r: fR, key: fKey
    });

    if (seenKeys[fKey]) {
      seenKeys[fKey].push(formStartRow + fi);
    } else {
      seenKeys[fKey] = [formStartRow + fi];
    }
  }

  var dupKeys = [];
  for (var k in seenKeys) {
    if (seenKeys[k].length > 1) {
      dupKeys.push({ key: k, rows: seenKeys[k] });
    }
  }

  return {
    masterKeys: Object.keys(existingKeys).length,
    formRange: formStartRow + '-' + formLastRow,
    wouldAddCount: wouldAdd.length,
    wouldAdd: wouldAdd,
    duplicateKeysInBatch: dupKeys
  };
}

/**
 * dry-run: syncRecentFormToMasterと同じロジックで追加対象を出す。appendしない。
 */
function dryRunSyncRecentFormToMaster() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var masterSheet = smc.getSheetByName(SMC_MASTER_SHEET);
  var formSheet = smc.getSheetByName(SMC_FORM_SHEET);
  if (!masterSheet || !formSheet) return { error: 'sheet not found' };

  var tz = Session.getScriptTimeZone();

  var formColA = formSheet.getRange('A:A').getValues();
  var formLastRow = formColA.filter(String).length;
  var masterLastRow = masterSheet.getLastRow();
  if (formLastRow <= 1 || masterLastRow <= 1) return { error: 'no data' };

  var lookback = Math.min(30, formLastRow - 1);
  var formStart = formLastRow - lookback + 1;
  var formData = formSheet.getRange(formStart, 1, lookback, 41).getValues();

  var masterAR3 = masterSheet.getRange(2, 1, masterLastRow - 1, 18).getValues();
  var existingKeys = {};
  for (var i = 0; i < masterAR3.length; i++) {
    var ts = masterAR3[i][0];
    var mR3 = String(masterAR3[i][17] || '').trim();
    if (ts instanceof Date) {
      existingKeys[Utilities.formatDate(ts, tz, 'yyyy/MM/dd HH:mm:ss') + '|' + mR3] = true;
    } else if (ts) {
      existingKeys[String(ts).trim() + '|' + mR3] = true;
    }
  }

  var wouldAdd = [];
  var seenKeys = {};

  for (var fi = 0; fi < formData.length; fi++) {
    var formRow = formData[fi];
    var formTs = formRow[0];
    if (!formTs) continue;

    var formA;
    if (formTs instanceof Date) {
      formA = Utilities.formatDate(formTs, tz, 'yyyy/MM/dd HH:mm:ss');
    } else {
      formA = String(formTs).trim();
    }
    var formR = String(formRow[17] || '').trim();
    var formKey = formA + '|' + formR;

    if (existingKeys[formKey]) continue;

    var d = String(formRow[3] || '');
    var e = String(formRow[4] || '');
    var l = String(formRow[11] || '');
    var r = String(formRow[17] || '');

    wouldAdd.push({
      formRow: formStart + fi,
      a: formKey,
      d: d,
      e: e.substring(0, 20),
      l: l.substring(0, 20),
      r: r,
      key: formKey
    });

    if (seenKeys[formKey]) {
      seenKeys[formKey].push(formStart + fi);
    } else {
      seenKeys[formKey] = [formStart + fi];
    }
  }

  var dupKeys = [];
  for (var k in seenKeys) {
    if (seenKeys[k].length > 1) {
      dupKeys.push({ key: k, rows: seenKeys[k] });
    }
  }

  return {
    formLastRow: formLastRow,
    formStart: formStart,
    masterKeys: Object.keys(existingKeys).length,
    wouldAddCount: wouldAdd.length,
    wouldAdd: wouldAdd,
    duplicateKeysInBatch: dupKeys
  };
}

/**
 * 停止フラグ操作
 */
function pauseFormSync() {
  PropertiesService.getScriptProperties().setProperty('FORM_SYNC_PAUSED', 'true');
  return { paused: true };
}

function resumeFormSync() {
  PropertiesService.getScriptProperties().setProperty('FORM_SYNC_PAUSED', 'false');
  return { paused: false };
}

function getFormSyncStatus() {
  var paused = PropertiesService.getScriptProperties().getProperty('FORM_SYNC_PAUSED') === 'true';
  return { paused: paused };
}

/**
 * A列空のゴースト行を一括削除
 */
/**
 * コードに存在しない死んだトリガーを削除
 */
function deleteDeadTriggers() {
  var dead = ['extractCOAndRemind', 'createNewMonthSheet'];
  var triggers = ScriptApp.getProjectTriggers();
  var deleted = [];
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (dead.indexOf(fn) !== -1) {
      ScriptApp.deleteTrigger(triggers[i]);
      deleted.push(fn);
    }
  }
  return { deleted: deleted, remaining: ScriptApp.getProjectTriggers().map(function(t) { return t.getHandlerFunction(); }) };
}

/**
 * onFormSubmitトリガー用: フォーム送信直後にA空ゴースト行を即削除
 */
function onFormSubmitCleanup(e) {
  // 犯人トリガーがゴースト行を作り終わるのを待つ
  Utilities.sleep(10000);
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  if (!master) return;
  var lastRow = master.getLastRow();
  if (lastRow <= 1) return;
  // A,C,D,E列すべてにデータがある行は保護（削除しない）
  var vals = master.getRange(2, 1, lastRow - 1, 5).getValues();
  var checkCols = [0, 2, 3, 4]; // A,C,D,E
  var delRows = [];
  for (var i = 0; i < vals.length; i++) {
    var allFilled = true;
    for (var ci = 0; ci < checkCols.length; ci++) {
      if (!vals[i][checkCols[ci]] || String(vals[i][checkCols[ci]]).trim() === '') { allFilled = false; break; }
    }
    if (!allFilled) delRows.push(i + 2);
  }
  if (delRows.length === 0) return;
  delRows.sort(function(a, b) { return b - a; });
  for (var d = 0; d < delRows.length; d++) {
    master.deleteRow(delRows[d]);
  }
  SpreadsheetApp.flush();
}

/**
 * onFormSubmitCleanupのトリガーをセットアップ
 */
/**
 * フォーム回答シートの物理構造を調査（getLastRow vs A列非空数）
 */
function checkFormTail() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var form = smc.getSheetByName(SMC_FORM_SHEET);
  var physLastRow = form.getLastRow();
  var aVals = form.getRange('A:A').getValues();
  var aCount = aVals.filter(String).length;
  var lastARow = 0;
  for (var j = aVals.length - 1; j >= 0; j--) {
    if (aVals[j][0] && String(aVals[j][0]).trim()) { lastARow = j + 1; break; }
  }
  var result = { physLastRow: physLastRow, aCount: aCount, lastARow: lastARow, gap: physLastRow - lastARow };
  // 境界(lastARow-2 ~ lastARow+20)の各行: A,C,D,H,L,R,AH,AL,AP + 値有無
  var bs = Math.max(2, lastARow - 2);
  var bd = form.getRange(bs, 1, Math.min(25, physLastRow - bs + 1), 42).getValues();
  result.boundary = [];
  for (var i = 0; i < bd.length; i++) {
    var r = bd[i];
    var a = String(r[0]||'').trim(), c = String(r[2]||'').trim(), d = String(r[3]||'').trim();
    var h = String(r[7]||'').trim(), l = String(r[11]||'').trim(), rr = String(r[17]||'').trim();
    var ah = String(r[33]||'').trim(), al = String(r[37]||'').trim(), ap = String(r[41]||'').trim();
    var filled = [];
    for (var ci = 0; ci < 42; ci++) { if (r[ci] !== '' && r[ci] !== null && String(r[ci]).trim()) filled.push(ci); }
    result.boundary.push({
      row: bs+i, a: a.substring(0,20)||'空', c: c.substring(0,12)||'空',
      d: d.substring(0,8)||'空', h: h.substring(0,10)||'空', l: l.substring(0,8)||'空',
      r: rr.substring(0,12)||'空', ah: ah.substring(0,12)||'空', al: al.substring(0,6)||'空',
      ap: ap.substring(0,5)||'空', filledCols: filled
    });
  }
  return result;
}

function setupFormSubmitCleanup() {
  var ss = SpreadsheetApp.openById(SMC_SS_ID);
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onFormSubmitCleanup') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('onFormSubmitCleanup')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
  return { status: 'trigger_created', fn: 'onFormSubmitCleanup', event: 'onFormSubmit' };
}

/**
 * SMCスプレッドシートに対する全ユーザーのトリガーを調査
 */
function debugAllTriggersOnSMC() {
  var ss = SpreadsheetApp.openById(SMC_SS_ID);
  // getUserTriggers: 現ユーザーがこのスプレッドシートに持つ全トリガー（全プロジェクト横断）
  var userTriggers = ScriptApp.getUserTriggers(ss);
  var result = [];
  for (var i = 0; i < userTriggers.length; i++) {
    var t = userTriggers[i];
    result.push({
      fn: t.getHandlerFunction(),
      type: String(t.getEventType()),
      source: String(t.getTriggerSource()),
      id: t.getUniqueId()
    });
  }
  // このプロジェクトのトリガーも追加
  var projTriggers = ScriptApp.getProjectTriggers();
  var projResult = [];
  for (var j = 0; j < projTriggers.length; j++) {
    var pt = projTriggers[j];
    projResult.push({
      fn: pt.getHandlerFunction(),
      type: String(pt.getEventType()),
      source: String(pt.getTriggerSource()),
      id: pt.getUniqueId()
    });
  }
  return { userTriggersOnSMC: result, projectTriggers: projResult };
}

/**
 * マスターシート内の数式を全検索
 */
/**
 * 指定行のAP(42)以降の値と数式を返す（デバッグ用）
 */
function debugRowAP(rowNum) {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var r = parseInt(rowNum) || 2449;
  var lastCol = master.getLastColumn();
  if (lastCol < 42) return { error: 'lastCol=' + lastCol };
  var vals = master.getRange(r, 1, 1, lastCol).getValues()[0];
  var formulas = master.getRange(r, 1, 1, lastCol).getFormulas()[0];
  var result = { row: r, lastCol: lastCol, L_status: String(vals[11] || ''), D_name: String(vals[3] || '') };
  var cols = [];
  for (var c = 41; c < lastCol; c++) {
    var v = vals[c], f = formulas[c];
    if (v !== '' && v !== null && v !== undefined || f) {
      cols.push({ col: c + 1, colLetter: colToLetter_(c), val: String(v), formula: f || '' });
    }
  }
  result.apOnward = cols;
  // 問題行の近傍5行もチェック
  var nearby = [];
  var startR = Math.max(2, r - 2);
  var endR = Math.min(master.getLastRow(), r + 2);
  var nVals = master.getRange(startR, 1, endR - startR + 1, lastCol).getValues();
  var nForms = master.getRange(startR, 1, endR - startR + 1, lastCol).getFormulas();
  for (var ri = 0; ri < nVals.length; ri++) {
    var nr = { row: startR + ri, L: String(nVals[ri][11] || ''), D: String(nVals[ri][3] || '') };
    var nc = [];
    for (var ci = 41; ci < lastCol; ci++) {
      if (nVals[ri][ci] !== '' && nVals[ri][ci] !== null || nForms[ri][ci]) {
        nc.push({ col: ci + 1, letter: colToLetter_(ci), val: String(nVals[ri][ci]), formula: nForms[ri][ci] || '' });
      }
    }
    nr.apCols = nc;
    nearby.push(nr);
  }
  result.nearby = nearby;
  return result;
}
/**
 * 指定メンバーのブロック別着金内訳を返す
 */
/**
 * マスターD列のユニーク名一覧を返す（対象月のみ）
 */
function listMasterNames(targetMonth, targetYear) {
  var now = new Date();
  var year = parseInt(targetYear) || now.getFullYear();
  var month = parseInt(targetMonth) || (now.getMonth() + 1);
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var lastRow = master.getLastRow();
  if (lastRow <= 1) return { error: 'no data' };
  var data = master.getRange(2, 1, lastRow - 1, 12).getValues();
  var tz = Session.getScriptTimeZone();
  var names = {};
  for (var i = 0; i < data.length; i++) {
    var c = data[i][2];
    if (!rev_isTargetMonth_(c, year, month, tz)) continue;
    var d = String(data[i][3] || '').trim();
    var l = String(data[i][11] || '').trim();
    if (!d) continue;
    if (!names[d]) names[d] = { count: 0, seiyaku: 0 };
    names[d].count++;
    if (l === '成約' || l === '成約➔CO') names[d].seiyaku++;
  }
  var list = [];
  for (var n in names) list.push({ name: n, total: names[n].count, seiyaku: names[n].seiyaku });
  list.sort(function(a, b) { return b.total - a.total; });
  return { month: month, year: year, names: list };
}

function debugMemberRevenue(memberName, targetMonth, targetYear) {
  var now = new Date();
  var year = parseInt(targetYear) || now.getFullYear();
  var month = parseInt(targetMonth) || (now.getMonth() + 1);
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var lastRow = master.getLastRow();
  if (lastRow <= 1) return { error: 'no data' };
  var data = master.getRange(2, 1, lastRow - 1, 101).getValues();
  var tz = Session.getScriptTimeZone();
  var blocks = [
    {name:'P(15)', amt:15, date:2},
    {name:'V(21)', amt:21, date:2},
    {name:'AD(29)', amt:29, date:2},
    {name:'AP(41)', amt:41, date:38},
    {name:'AW(48)', amt:48, date:45},
    {name:'BD(55)', amt:55, date:52},
    {name:'BO(66)', amt:66, date:63},
    {name:'BV(73)', amt:73, date:70},
    {name:'CC(80)', amt:80, date:77},
    {name:'CJ(87)', amt:87, date:84}
  ];
  var result = { member: memberName, year: year, month: month, blocks: [], total: 0, rows: [] };
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var l = String(row[11] || '').trim();
    if (l !== '成約' && l !== '成約➔CO') continue;
    var d = String(row[3] || '').trim();
    if (d !== memberName) continue;
    for (var bi = 0; bi < blocks.length; bi++) {
      var dateVal = row[blocks[bi].date];
      if (!rev_isTargetMonth_(dateVal, year, month, tz)) continue;
      var amt = rev_parseNum_(row[blocks[bi].amt]);
      if (amt === 0) continue;
      result.rows.push({ row: i+2, block: blocks[bi].name, amt: amt, L: l });
      result.total = rev_round1_(result.total + amt);
    }
  }
  // ブロック集計
  var bTotals = {};
  for (var ri = 0; ri < result.rows.length; ri++) {
    var bn = result.rows[ri].block;
    if (!bTotals[bn]) bTotals[bn] = 0;
    bTotals[bn] = rev_round1_(bTotals[bn] + result.rows[ri].amt);
  }
  for (var bk in bTotals) result.blocks.push({ block: bk, total: bTotals[bk] });
  return result;
}

/**
 * 成約行のAP+とP/V/ADの二重計上チェック + 同一顧客の重複行チェック
 */
function auditDuplicates(targetMonth, targetYear) {
  var now = new Date();
  var year = parseInt(targetYear) || now.getFullYear();
  var month = parseInt(targetMonth) || (now.getMonth() + 1);
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var lastRow = master.getLastRow();
  if (lastRow <= 1) return { error: 'no data' };
  var data = master.getRange(2, 1, lastRow - 1, 101).getValues();
  var tz = Session.getScriptTimeZone();

  // === 1. AP+とP/Vの金額一致チェック（二重計上疑い） ===
  var apDupes = [];
  // P=15,V=21,AD=29 vs AP=41,AW=48,BD=55
  var pvCols = [15, 21, 29];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var l = String(row[11] || '').trim();
    if (l !== '成約' && l !== '成約➔CO') continue;
    var c = row[2];
    if (!rev_isTargetMonth_(c, year, month, tz)) continue;

    var pvAmts = {};
    for (var pi = 0; pi < pvCols.length; pi++) {
      var pv = rev_parseNum_(row[pvCols[pi]]);
      if (pv > 0) pvAmts[pv] = pvCols[pi];
    }

    // AP以降で同額があるか
    var apCols = [41, 48, 55, 66, 73, 80, 87];
    for (var ai = 0; ai < apCols.length; ai++) {
      var apAmt = rev_parseNum_(row[apCols[ai]]);
      if (apAmt > 0 && pvAmts[apAmt]) {
        apDupes.push({
          row: i + 2,
          D: String(row[3] || ''),
          L: l,
          pvCol: pvCols.indexOf(pvAmts[apAmt]) > -1 ? colToLetter_(pvAmts[apAmt]) : '?',
          apCol: colToLetter_(apCols[ai]),
          amt: apAmt
        });
      }
    }
  }

  // === 2. 同一顧客の重複行チェック（D+メール or D+金額が同じ） ===
  var seiyakuRows = [];
  for (var j = 0; j < data.length; j++) {
    var row2 = data[j];
    var l2 = String(row2[11] || '').trim();
    if (l2 !== '成約' && l2 !== '成約➔CO') continue;
    var c2 = row2[2];
    if (!rev_isTargetMonth_(c2, year, month, tz)) continue;
    var cDate2 = row2[2];
    var cStr2 = (cDate2 instanceof Date) ? Utilities.formatDate(cDate2, tz, 'yyyy/MM/dd') : String(cDate2 || '').trim();
    seiyakuRows.push({
      idx: j,
      row: j + 2,
      D: String(row2[3] || '').trim(),
      E: String(row2[4] || '').trim(),
      C: cStr2,
      email: String(row2[7] || '').trim(),
      P: rev_parseNum_(row2[15]),
      V: rev_parseNum_(row2[21]),
      L: l2,
      A: row2[0]
    });
  }

  var rowDupes = [];
  for (var a = 0; a < seiyakuRows.length; a++) {
    for (var b = a + 1; b < seiyakuRows.length; b++) {
      var ra = seiyakuRows[a], rb = seiyakuRows[b];
      // D+E+C 全一致で重複判定
      if (ra.D !== rb.D || ra.E !== rb.E || ra.C !== rb.C) continue;
      rowDupes.push({
        rowA: ra.row, rowB: rb.row,
        D: ra.D, E: ra.E, C: ra.C,
        La: ra.L, Lb: rb.L
      });
    }
  }

  return {
    apDuplicates: { count: apDupes.length, rows: apDupes.slice(0, 50) },
    rowDuplicates: { count: rowDupes.length, rows: rowDupes.slice(0, 50) }
  };
}

/**
 * 重複成約行を削除（各顧客1行だけ残す）
 * 残す基準: 成約➔CO優先 → AP+データ量多い → 行番号大きい(新しい)
 */
function deduplicateSeiyaku(targetMonth, targetYear, dryRun) {
  var dry = (dryRun === true || dryRun === 'true');
  var now = new Date();
  var year = parseInt(targetYear) || now.getFullYear();
  var month = parseInt(targetMonth) || (now.getMonth() + 1);
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var lastRow = master.getLastRow();
  var lastCol = master.getLastColumn();
  if (lastRow <= 1) return { error: 'no data' };
  var data = master.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var tz = Session.getScriptTimeZone();

  // 成約行を集める
  var seiyaku = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var l = String(row[11] || '').trim();
    if (l !== '成約' && l !== '成約➔CO') continue;
    var c = row[2];
    if (!rev_isTargetMonth_(c, year, month, tz)) continue;
    var d = String(row[3] || '').trim();
    var e = String(row[4] || '').trim();
    var cDate = row[2];
    var cStr = (cDate instanceof Date) ? Utilities.formatDate(cDate, tz, 'yyyy/MM/dd') : String(cDate || '').trim();
    // AP+データ数カウント
    var apCount = 0;
    for (var ac = 41; ac < lastCol; ac++) {
      if (row[ac] !== '' && row[ac] !== null && row[ac] !== undefined && String(row[ac]).trim() !== '' && String(row[ac]) !== '0') apCount++;
    }
    seiyaku.push({ idx: i, row: i + 2, D: d, E: e, C: cStr, L: l, P: rev_parseNum_(row[15]), V: rev_parseNum_(row[21]), apCount: apCount });
  }

  // D+E+C（担当+顧客名+日付）でグループ化
  var groups = {};
  for (var si = 0; si < seiyaku.length; si++) {
    var s = seiyaku[si];
    if (!s.E) continue;
    var key = s.D + '|' + s.E + '|' + s.C;
    if (!groups[key]) groups[key] = [];
    groups[key].push(s);
  }

  // 2行以上あるグループから削除対象を選定
  var toDelete = [];
  var kept = [];
  for (var gk in groups) {
    var g = groups[gk];
    if (g.length < 2) continue;
    // ソート: 成約➔CO優先 → apCount多い → row大きい
    g.sort(function(a, b) {
      var aRank = (a.L === '成約➔CO') ? 2 : 1;
      var bRank = (b.L === '成約➔CO') ? 2 : 1;
      if (bRank !== aRank) return bRank - aRank;
      if (b.apCount !== a.apCount) return b.apCount - a.apCount;
      return b.row - a.row;
    });
    // 先頭を残し、残りを削除
    kept.push({ row: g[0].row, D: g[0].D, E: g[0].E, C: g[0].C, L: g[0].L, apCount: g[0].apCount });
    for (var gi = 1; gi < g.length; gi++) {
      toDelete.push({ row: g[gi].row, D: g[gi].D, E: g[gi].E, C: g[gi].C, L: g[gi].L });
    }
  }

  // 削除実行（下から）
  if (!dry && toDelete.length > 0) {
    var delRows = toDelete.map(function(t) { return t.row; });
    delRows.sort(function(a, b) { return b - a; });
    for (var di = 0; di < delRows.length; di++) {
      master.deleteRow(delRows[di]);
    }
    SpreadsheetApp.flush();
  }

  return {
    dryRun: dry,
    duplicateGroups: Object.keys(groups).filter(function(k) { return groups[k].length > 1; }).length,
    deleted: toDelete.length,
    deletedRows: toDelete,
    keptRows: kept
  };
}

/**
 * 指定行の全データをヘッダー付きで返す
 */
function dumpRows(rowList) {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var lastCol = master.getLastColumn();
  var headers = master.getRange(1, 1, 1, lastCol).getValues()[0];
  var rows = String(rowList).split(',');
  var result = { headers: [], data: [] };
  // 主要列のみ返す
  var keyCols = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40];
  for (var hi = 0; hi < keyCols.length; hi++) {
    result.headers.push(colToLetter_(keyCols[hi]) + '(' + String(headers[keyCols[hi]] || '') + ')');
  }
  for (var ri = 0; ri < rows.length; ri++) {
    var r = parseInt(rows[ri]);
    if (!r || r < 2) continue;
    var vals = master.getRange(r, 1, 1, lastCol).getValues()[0];
    var rowData = { row: r };
    for (var ci = 0; ci < keyCols.length; ci++) {
      var v = vals[keyCols[ci]];
      rowData[colToLetter_(keyCols[ci])] = (v instanceof Date) ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm') : String(v || '');
    }
    // AP以降も非空のみ
    var ap = {};
    for (var ac = 41; ac < lastCol; ac++) {
      if (vals[ac] !== '' && vals[ac] !== null && vals[ac] !== undefined && String(vals[ac]).trim() !== '' && String(vals[ac]) !== '0') {
        ap[colToLetter_(ac)] = (vals[ac] instanceof Date) ? Utilities.formatDate(vals[ac], Session.getScriptTimeZone(), 'yyyy/MM/dd') : String(vals[ac]);
      }
    }
    rowData.AP_plus = ap;
    result.data.push(rowData);
  }
  return result;
}

/**
 * 同一顧客(E列)の別行でP/VとAP+に同額がある場合にAP+をクリア（二重計上除去）
 */
function clearAPDuplicateAmounts(targetMonth, targetYear, dryRun) {
  var dry = (dryRun === true || dryRun === 'true');
  var now = new Date();
  var year = parseInt(targetYear) || now.getFullYear();
  var month = parseInt(targetMonth) || (now.getMonth() + 1);
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var lastRow = master.getLastRow();
  var lastCol = master.getLastColumn();
  if (lastRow <= 1) return { error: 'no data' };
  var data = master.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var tz = Session.getScriptTimeZone();

  var pvCols = [15, 21, 29]; // P, V, AD
  var apAmtCols = [41, 48, 55, 66, 73, 80, 87]; // AP, AW, BD, BO, BV, CC, CJ

  // 成約行を顧客名(E列)でグループ化
  var byCustomer = {};
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var l = String(row[11] || '').trim();
    if (l !== '成約' && l !== '成約➔CO') continue;
    var c = row[2];
    if (!rev_isTargetMonth_(c, year, month, tz)) continue;
    var e = String(row[4] || '').trim();
    if (!e) continue;
    if (!byCustomer[e]) byCustomer[e] = [];
    byCustomer[e].push({ idx: i, row: i + 2, data: row, D: String(row[3] || '') });
  }

  var cleared = [];
  for (var cust in byCustomer) {
    var rows = byCustomer[cust];
    if (rows.length < 2) continue;

    // この顧客の全行のP/V/AD金額を集める
    var pvAmts = {};
    for (var ri = 0; ri < rows.length; ri++) {
      for (var pi = 0; pi < pvCols.length; pi++) {
        var pv = rev_parseNum_(rows[ri].data[pvCols[pi]]);
        if (pv > 0) pvAmts[pv] = { row: rows[ri].row, col: colToLetter_(pvCols[pi]) };
      }
    }

    // 別の行のAP+で同額があればクリア
    for (var ri2 = 0; ri2 < rows.length; ri2++) {
      for (var ai = 0; ai < apAmtCols.length; ai++) {
        var apAmt = rev_parseNum_(rows[ri2].data[apAmtCols[ai]]);
        if (apAmt > 0 && pvAmts[apAmt] && pvAmts[apAmt].row !== rows[ri2].row) {
          cleared.push({
            apRow: rows[ri2].row,
            pvRow: pvAmts[apAmt].row,
            D: rows[ri2].D,
            E: cust,
            pvCol: pvAmts[apAmt].col,
            apCol: colToLetter_(apAmtCols[ai]),
            amt: apAmt
          });
          if (!dry) {
            master.getRange(rows[ri2].row, apAmtCols[ai] + 1).setValue('');
          }
        }
      }
    }
  }
  if (!dry && cleared.length > 0) SpreadsheetApp.flush();
  return { dryRun: dry, cleared: cleared.length, rows: cleared };
}

/**
 * L列「顧客情報に記入」を「継続」に一括変更
 */
/**
 * CX > Q の行を全メンバーで検出
 */
/**
 * 全メンバーの詳細チェック（着金/売上/成約/失注/CO/CO額）
 */
function auditAllMembers(targetMonth, targetYear) {
  var now = new Date();
  var year = parseInt(targetYear) || now.getFullYear();
  var month = parseInt(targetMonth) || (now.getMonth() + 1);
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var lastRow = master.getLastRow();
  if (lastRow <= 1) return { error: 'no data' };
  var data = master.getRange(2, 1, lastRow - 1, 102).getValues();
  var tz = Session.getScriptTimeZone();
  var CX = 101, Q = 16;

  var byPerson = {};
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var c = row[2];
    if (!rev_isTargetMonth_(c, year, month, tz)) continue;
    var d = String(row[3] || '').trim();
    if (!d || d === 'テスト') continue;
    var l = String(row[11] || '').trim();

    if (!byPerson[d]) byPerson[d] = { revenue: 0, sales: 0, deals: 0, closed: 0, lost: 0, lostCont: 0, cont: 0, co: 0, coAmt: 0, rows: [] };
    var p = byPerson[d];
    p.deals++;

    if (l === '成約') {
      p.closed++;
      p.revenue += rev_parseNum_(row[CX]);
      p.sales += rev_parseNum_(row[Q]);
    } else if (l === '成約➔CO') {
      p.co++;
      var coQ = rev_parseNum_(row[Q]);
      p.coAmt += coQ;
      p.revenue += rev_parseNum_(row[CX]);
      p.sales += coQ;
    } else if (l === '失注') {
      p.lost++;
    } else if (l.indexOf('失注') !== -1 && l.indexOf('継続') !== -1) {
      p.lostCont++;
    } else if (l.indexOf('継続') !== -1) {
      p.cont++;
    }

    p.rows.push({ row: i + 2, E: String(row[4] || ''), L: l, Q: rev_parseNum_(row[Q]), CX: rev_parseNum_(row[CX]) });
  }

  var members = [];
  for (var name in byPerson) {
    var m = byPerson[name];
    m.name = name;
    m.revenue = rev_round1_(m.revenue);
    m.sales = rev_round1_(m.sales);
    m.coAmt = rev_round1_(m.coAmt);
    members.push(m);
  }
  members.sort(function(a, b) { return b.revenue - a.revenue; });
  return { month: month, year: year, members: members };
}

/**
 * 復旧モード: 重複候補監査
 * R列優先、なければE列で同一人物判定
 * AZ列(51)に削除候補フラグと理由を書き込む
 * members: 対象メンバー名の配列（省略時は全員）
 */
function auditDuplicateCandidates(targetMonth, targetYear, memberFilter, dryRun) {
  var dry = (dryRun === true || dryRun === 'true');
  var now = new Date();
  var year = parseInt(targetYear) || now.getFullYear();
  var month = parseInt(targetMonth) || (now.getMonth() + 1);
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var lastRow = master.getLastRow();
  if (lastRow <= 1) return { error: 'no data' };
  var data = master.getRange(2, 1, lastRow - 1, 102).getValues();
  var tz = Session.getScriptTimeZone();
  var AZ_COL = 51; // AZ列(0-indexed)
  var filterMembers = memberFilter ? String(memberFilter).split(',') : null;
  var fiveDaysMs = 5 * 24 * 60 * 60 * 1000;

  // 対象月の行を集める
  var rows = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var c = row[2];
    if (!rev_isTargetMonth_(c, year, month, tz)) continue;
    var d = String(row[3] || '').trim();
    if (!d || d === 'テスト') continue;
    if (filterMembers && filterMembers.indexOf(d) === -1) continue;
    var r = String(row[17] || '').trim(); // R列(本名)
    var e = String(row[4] || '').trim(); // E列
    var personKey = r || e; // R優先、なければE
    if (!personKey) continue;
    var l = String(row[11] || '').trim();
    rows.push({
      idx: i,
      row: i + 2,
      D: d,
      R: r,
      E: e,
      personKey: personKey,
      L: l,
      Q: rev_parseNum_(row[16]),
      CX: rev_parseNum_(row[101]),
      C: c,
      cStr: (c instanceof Date) ? Utilities.formatDate(c, tz, 'yyyy/MM/dd') : String(c || '')
    });
  }

  // D + personKey でグループ化
  var groups = {};
  for (var ri = 0; ri < rows.length; ri++) {
    var r2 = rows[ri];
    // personKeyの正規化（スペース除去）
    var normKey = r2.D + '|' + r2.personKey.replace(/\s+/g, '');
    if (!groups[normKey]) groups[normKey] = [];
    groups[normKey].push(r2);
  }

  // 2行以上あるグループを判定
  var candidates = [];
  for (var gk in groups) {
    var g = groups[gk];
    if (g.length < 2) continue;

    // ステータス分類
    var seiyaku = [], seiyakuCO = [], shicchu = [], keizoku = [];
    for (var gi = 0; gi < g.length; gi++) {
      var item = g[gi];
      if (item.L === '成約') seiyaku.push(item);
      else if (item.L === '成約➔CO') seiyakuCO.push(item);
      else if (item.L === '失注') shicchu.push(item);
      else if (item.L === '継続') keizoku.push(item);
    }

    var hasSeiyaku = seiyaku.length > 0 || seiyakuCO.length > 0;

    // パターン判定
    if (hasSeiyaku && shicchu.length > 0) {
      // 成約/成約➔CO + 失注 → 失注を削除候補
      for (var si = 0; si < shicchu.length; si++) {
        var reason = '成約あり(' + (seiyaku.length > 0 ? 'Row' + seiyaku[0].row : 'CORow' + seiyakuCO[0].row) + ')のため失注を削除候補';
        candidates.push({ row: shicchu[si].row, D: shicchu[si].D, personKey: shicchu[si].personKey, E: shicchu[si].E, R: shicchu[si].R, L: shicchu[si].L, reason: reason });
      }
    }

    if (!hasSeiyaku && shicchu.length > 1) {
      // 失注 + 失注 → 1件残して他を削除候補
      for (var si2 = 1; si2 < shicchu.length; si2++) {
        candidates.push({ row: shicchu[si2].row, D: shicchu[si2].D, personKey: shicchu[si2].personKey, E: shicchu[si2].E, R: shicchu[si2].R, L: shicchu[si2].L, reason: '失注重複(Row' + shicchu[0].row + 'を残す)' });
      }
    }

    if (keizoku.length > 0 && shicchu.length > 0) {
      // 継続 + 失注 → 5日超なら継続を失注統合候補
      for (var ki = 0; ki < keizoku.length; ki++) {
        var kc = keizoku[ki].C;
        if (kc instanceof Date && (now.getTime() - kc.getTime()) >= fiveDaysMs) {
          candidates.push({ row: keizoku[ki].row, D: keizoku[ki].D, personKey: keizoku[ki].personKey, E: keizoku[ki].E, R: keizoku[ki].R, L: keizoku[ki].L, reason: '継続5日超+失注あり→失注統合候補' });
        }
      }
    }
  }

  // AZ列にフラグ書き込み（dry=falseの場合）
  if (!dry && candidates.length > 0) {
    for (var ci = 0; ci < candidates.length; ci++) {
      master.getRange(candidates[ci].row, AZ_COL + 1).setValue('削除候補: ' + candidates[ci].reason);
    }
    SpreadsheetApp.flush();
  }

  // メンバー別サマリー
  var byMember = {};
  for (var ci2 = 0; ci2 < candidates.length; ci2++) {
    var c2 = candidates[ci2];
    if (!byMember[c2.D]) byMember[c2.D] = [];
    byMember[c2.D].push(c2);
  }

  return {
    dryRun: dry,
    totalCandidates: candidates.length,
    byMember: byMember,
    candidates: candidates
  };
}

/**
 * 復旧モード: 個別行にAZフラグを書き込む
 * rows: "行番号:理由,行番号:理由,..." 形式
 */
function flagRows(rowsParam) {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var AZ_COL = 52; // AZ列(1-indexed)
  var pairs = String(rowsParam).split(',');
  var flagged = [];
  for (var i = 0; i < pairs.length; i++) {
    var parts = pairs[i].split(':');
    var row = parseInt(parts[0]);
    var reason = parts.slice(1).join(':') || '削除候補';
    if (row > 1) {
      master.getRange(row, AZ_COL).setValue('削除候補: ' + reason);
      flagged.push({ row: row, reason: reason });
    }
  }
  SpreadsheetApp.flush();
  return { flagged: flagged.length, rows: flagged };
}

function auditCXvsQ(targetMonth, targetYear) {
  var now = new Date();
  var year = parseInt(targetYear) || now.getFullYear();
  var month = parseInt(targetMonth) || (now.getMonth() + 1);
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var lastRow = master.getLastRow();
  if (lastRow <= 1) return { error: 'no data' };
  var data = master.getRange(2, 1, lastRow - 1, 102).getValues();
  var tz = Session.getScriptTimeZone();
  var CX = 101;
  var Q = 16;
  var issues = [];
  var byPerson = {};

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var l = String(row[11] || '').trim();
    if (l !== '成約' && l !== '成約➔CO') continue;
    var c = row[2];
    if (!rev_isTargetMonth_(c, year, month, tz)) continue;
    var d = String(row[3] || '').trim();
    var q = rev_parseNum_(row[Q]);
    var cx = rev_parseNum_(row[CX]);
    if (cx > q && q > 0) {
      var diff = rev_round1_(cx - q);
      issues.push({ row: i + 2, D: d, E: String(row[4] || ''), Q: q, CX: cx, diff: diff });
      if (!byPerson[d]) byPerson[d] = { count: 0, totalDiff: 0 };
      byPerson[d].count++;
      byPerson[d].totalDiff = rev_round1_(byPerson[d].totalDiff + diff);
    }
  }
  var summary = [];
  for (var name in byPerson) summary.push({ name: name, count: byPerson[name].count, totalDiff: byPerson[name].totalDiff });
  summary.sort(function(a, b) { return b.totalDiff - a.totalDiff; });
  return { total: issues.length, summary: summary, rows: issues };
}

/**
 * マスターをC列（日付）の時系列順にソート
 */
/**
 * CX > Q の行でCXをQに合わせる
 */
function fixCXOverQ_(master) {
  var lastRow = master.getLastRow();
  if (lastRow <= 1) return 0;
  var data = master.getRange(2, 1, lastRow - 1, 102).getValues();
  var Q = 16;
  var CX = 101;
  var fixed = 0;
  for (var i = 0; i < data.length; i++) {
    var l = String(data[i][11] || '').trim();
    if (l !== '成約' && l !== '成約➔CO') continue;
    var q = rev_parseNum_(data[i][Q]);
    var cx = rev_parseNum_(data[i][CX]);
    if (q > 0 && cx > q) {
      master.getRange(i + 2, CX + 1).setValue(q);
      fixed++;
    }
  }
  return fixed;
}

/**
 * マスターのL列にプルダウン（入力規則）を設定
 */
/**
 * マスターL列の条件付き書式でオレンジ背景を削除
 */
/**
 * L列の条件付き書式を設定（失注=紫文字）
 */
function setLColumnFormatting() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var lastRow = master.getLastRow();
  if (lastRow <= 1) return { error: 'no data' };

  var range = master.getRange(2, 12, lastRow - 1, 1);

  // 既存のL列条件付き書式を削除
  var rules = master.getConditionalFormatRules();
  var kept = [];
  for (var i = 0; i < rules.length; i++) {
    var ranges = rules[i].getRanges();
    var isLCol = false;
    for (var ri = 0; ri < ranges.length; ri++) {
      if (ranges[ri].getColumn() === 12) { isLCol = true; break; }
    }
    if (!isLCol) kept.push(rules[i]);
  }

  // 失注=紫文字
  var shicchu = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('失注')
    .setFontColor('#9900ff')
    .setRanges([range])
    .build();

  kept.push(shicchu);
  master.setConditionalFormatRules(kept);

  return { set: true, rules: kept.length };
}

function removeOrangeConditionalFormat() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var rules = master.getConditionalFormatRules();
  var kept = [];
  var removed = [];
  for (var i = 0; i < rules.length; i++) {
    var rule = rules[i];
    var bg = rule.getBooleanCondition();
    if (!bg) { kept.push(rule); continue; }
    var ranges = rule.getRanges();
    var isLCol = false;
    for (var ri = 0; ri < ranges.length; ri++) {
      if (ranges[ri].getColumn() === 12) { isLCol = true; break; }
    }
    if (!isLCol) { kept.push(rule); continue; }
    // 背景色がオレンジ系かチェック
    var bgColor = bg.getBackground();
    if (bgColor && (bgColor === '#ff9900' || bgColor === '#f6b26b' || bgColor === '#e69138' || bgColor === '#ff6d01' || bgColor === '#fce5cd' || bgColor.indexOf('f') === 1)) {
      removed.push({ index: i, color: bgColor });
    } else {
      kept.push(rule);
    }
  }
  if (removed.length > 0) {
    master.setConditionalFormatRules(kept);
  }
  return { total: rules.length, removed: removed.length, kept: kept.length, details: removed };
}

function setLColumnValidation() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var lastRow = master.getLastRow();
  if (lastRow <= 1) return { error: 'no data' };
  var values = ['成約', '成約➔CO', '失注', '継続', '継続2→成約', '継続3→成約', '継続4→成約'];
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(true)
    .build();
  master.getRange(2, 12, lastRow - 1, 1).setDataValidation(rule);
  return { set: lastRow - 1, values: values };
}

function sortMasterByDate() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var lastRow = master.getLastRow();
  var lastCol = master.getLastColumn();
  if (lastRow <= 1) return { error: 'no data' };
  // ヘッダー除いてC列(3列目)で昇順ソート
  var range = master.getRange(2, 1, lastRow - 1, lastCol);
  range.sort({ column: 3, ascending: true });
  SpreadsheetApp.flush();
  return { sorted: lastRow - 1, byColumn: 'C(日付)', order: '昇順' };
}

/**
 * 継続で5日以上経過してる行を失注に変更
 */
/**
 * 内部用: 継続で5日以上経過を失注に変更
 */
/**
 * 継続→失注チェックのトリガーをセットアップ（1日おき）
 */
function setupDailyContToLost() {
  // 既存トリガー削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'dailyContToLost') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('dailyContToLost')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();
  return { status: 'created', fn: 'dailyContToLost', schedule: '毎日6時' };
}

/**
 * 1日おきトリガー用: 継続5日超→失注 + C列ソート
 */
function dailyContToLost() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  if (!master) return;
  var fixed = contToLost_(master);
  // C列ソート（全列カバーでAP-CW列ズレ防止）
  var lastRow = master.getLastRow();
  var maxCol = master.getMaxColumns();
  if (lastRow > 2) {
    master.getRange(2, 1, lastRow - 1, maxCol).sort({ column: 3, ascending: true });
  }
  SpreadsheetApp.flush();
}

function contToLost_(masterSheet) {
  var lastRow = masterSheet.getLastRow();
  if (lastRow <= 1) return 0;
  var data = masterSheet.getRange(2, 1, lastRow - 1, 12).getValues();
  var now = new Date();
  var fiveDaysMs = 5 * 24 * 60 * 60 * 1000;
  var count = 0;
  for (var i = 0; i < data.length; i++) {
    var l = String(data[i][11] || '').trim();
    if (l !== '継続') continue;
    var c = data[i][2];
    if (!(c instanceof Date)) continue;
    if (now.getTime() - c.getTime() >= fiveDaysMs) {
      masterSheet.getRange(i + 2, 12).setValue('失注');
      count++;
    }
  }
  return count;
}

function contToLost(dryRun) {
  var dry = (dryRun === true || dryRun === 'true');
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  if (!smc) return { error: 'smc null', id: SMC_SS_ID };
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  if (!master) return { error: 'master null', sheet: SMC_MASTER_SHEET, sheets: smc.getSheets().map(function(s){return s.getName();}) };
  var lastRow = master.getLastRow();
  if (lastRow <= 1) return { error: 'no data' };
  var data = master.getRange(2, 1, lastRow - 1, 12).getValues();
  var now = new Date();
  var fiveDaysMs = 5 * 24 * 60 * 60 * 1000;
  var changed = [];

  for (var i = 0; i < data.length; i++) {
    var l = String(data[i][11] || '').trim();
    if (l !== '継続') continue;
    var c = data[i][2];
    if (!(c instanceof Date)) continue;
    var elapsed = now.getTime() - c.getTime();
    if (elapsed >= fiveDaysMs) {
      var tz = Session.getScriptTimeZone();
      changed.push({
        row: i + 2,
        D: String(data[i][3] || ''),
        E: String(data[i][4] || ''),
        C: Utilities.formatDate(c, tz, 'yyyy/MM/dd'),
        days: Math.floor(elapsed / (24 * 60 * 60 * 1000))
      });
      if (!dry) master.getRange(i + 2, 12).setValue('失注');
    }
  }
  if (!dry && changed.length > 0) SpreadsheetApp.flush();
  return { dryRun: dry, changed: changed.length, rows: changed.slice(0, 50) };
}

function replaceLStatus(dryRun) {
  var dry = (dryRun === true || dryRun === 'true');
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var lastRow = master.getLastRow();
  if (lastRow <= 1) return { error: 'no data' };
  var lVals = master.getRange(2, 12, lastRow - 1, 1).getValues();
  var changed = [];
  for (var i = 0; i < lVals.length; i++) {
    if (String(lVals[i][0]).trim() === '顧客情報に記入') {
      changed.push(i + 2);
      if (!dry) master.getRange(i + 2, 12).setValue('継続');
    }
  }
  if (!dry && changed.length > 0) SpreadsheetApp.flush();
  return { dryRun: dry, changed: changed.length, rows: changed.slice(0, 30) };
}

function colToLetter_(i) {
  var s = '';
  while (i >= 0) { s = String.fromCharCode(65 + (i % 26)) + s; i = Math.floor(i / 26) - 1; }
  return s;
}

function debugMasterFormulas() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var lastRow = master.getLastRow();
  var lastCol = master.getLastColumn();
  var formulas = master.getRange(1, 1, lastRow, lastCol).getFormulas();
  var found = [];
  for (var r = 0; r < formulas.length; r++) {
    for (var c = 0; c < formulas[r].length; c++) {
      if (formulas[r][c]) {
        found.push({ row: r + 1, col: c + 1, formula: formulas[r][c].substring(0, 100) });
      }
    }
  }
  // フォーム回答シートのリンク確認
  var sheets = smc.getSheets();
  var sheetInfo = [];
  for (var s = 0; s < sheets.length; s++) {
    var sh = sheets[s];
    var url = '';
    try { url = sh.getFormUrl() || ''; } catch(e) {}
    if (url) sheetInfo.push({ name: sh.getName(), formUrl: url });
  }
  // 全シートの数式数も集計
  var sheetFormulaCounts = [];
  for (var si = 0; si < sheets.length; si++) {
    var sh2 = sheets[si];
    try {
      var lr = sh2.getLastRow();
      var lc = sh2.getLastColumn();
      if (lr > 0 && lc > 0) {
        var fs = sh2.getRange(1, 1, Math.min(lr, 5), Math.min(lc, 50)).getFormulas();
        var cnt = 0;
        for (var fr = 0; fr < fs.length; fr++) {
          for (var fc = 0; fc < fs[fr].length; fc++) {
            if (fs[fr][fc]) cnt++;
          }
        }
        if (cnt > 0) sheetFormulaCounts.push({ name: sh2.getName(), formulasIn5Rows: cnt });
      }
    } catch(e2) {}
  }
  return { formulaCount: found.length, formulas: found.slice(0, 30), linkedForms: sheetInfo, sheetFormulaCounts: sheetFormulaCounts };
}

function deleteGhostRows() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var lastRow = master.getLastRow();
  if (lastRow <= 1) return { deleted: 0 };
  // A,C,D,E列すべてにデータがある行は保護（削除しない）
  var data = master.getRange(2, 1, lastRow - 1, 5).getValues();
  var checkCols = [0, 2, 3, 4]; // A,C,D,E
  var delRows = [];
  for (var i = 0; i < data.length; i++) {
    var allFilled = true;
    for (var ci = 0; ci < checkCols.length; ci++) {
      if (!data[i][checkCols[ci]] || String(data[i][checkCols[ci]]).trim() === '') { allFilled = false; break; }
    }
    if (!allFilled) delRows.push(i + 2);
  }
  delRows.sort(function(a, b) { return b - a; });
  for (var d = 0; d < delRows.length; d++) {
    master.deleteRow(delRows[d]);
  }
  SpreadsheetApp.flush();
  return { deleted: delRows.length, rows: delRows.slice(0, 20) };
}

/**
 * 観測専用: ゴースト行の正体を特定する
 * append/delete/修正なし。読み取りと照合のみ。
 */
function debugGhostRows() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var lastRow = master.getLastRow();
  if (lastRow <= 1) return 'no data';

  // 全行読み取り（A,H,I,J,K,L,R = col 1,8,9,10,11,12,18）
  var allData = master.getRange(2, 1, lastRow - 1, 18).getValues();

  // 既存行（A列あり）とゴースト行（A列空）に分離
  var normals = [];  // {row, a, h, i, j, k, l, r}
  var ghosts = [];
  for (var idx = 0; idx < allData.length; idx++) {
    var row = allData[idx];
    var a = String(row[0] || '').trim();
    var h = String(row[7] || '').trim();
    var ii = String(row[8] || '').trim();
    var j = String(row[9] || '').trim();
    var k = String(row[10] || '').trim();
    var l = String(row[11] || '').trim();
    var r = String(row[17] || '').trim();
    var entry = { row: idx + 2, a: a, h: h, i: ii, j: j, k: k, l: l, r: r };

    if (a) {
      normals.push(entry);
    } else if (h || ii || j || k || l || r) {
      ghosts.push(entry);
    }
  }

  var log = [];
  log.push('=== debugGhostRows ===');
  log.push('timestamp: ' + new Date().toISOString());
  log.push('totalRows: ' + allData.length);
  log.push('normalRows: ' + normals.length);
  log.push('ghostRows: ' + ghosts.length);
  log.push('');

  // 各ゴースト行について既存行と照合
  for (var gi = 0; gi < ghosts.length; gi++) {
    var g = ghosts[gi];
    log.push('--- Ghost #' + (gi + 1) + ' (Row ' + g.row + ') ---');
    log.push('A=' + (g.a || '(empty)') + ' H=' + g.h + ' I=' + g.i + ' J=' + g.j + ' K=' + g.k + ' L=' + g.l + ' R=' + g.r);

    var matches = [];
    for (var ni = 0; ni < normals.length; ni++) {
      var n = normals[ni];
      if (n.h === g.h && n.i === g.i && n.j === g.j && n.k === g.k && n.l === g.l && n.r === g.r) {
        matches.push(n);
      }
    }

    if (matches.length > 0) {
      log.push('MATCH: ' + matches.length + '件');
      for (var mi = 0; mi < matches.length; mi++) {
        log.push('  → Row ' + matches[mi].row + ' A=' + matches[mi].a);
      }
    } else {
      log.push('NO MATCH');
    }
    log.push('');
  }

  // サマリー
  var matchCount = 0;
  var noMatchCount = 0;
  for (var gi2 = 0; gi2 < ghosts.length; gi2++) {
    var g2 = ghosts[gi2];
    var found = false;
    for (var ni2 = 0; ni2 < normals.length; ni2++) {
      var n2 = normals[ni2];
      if (n2.h === g2.h && n2.i === g2.i && n2.j === g2.j && n2.k === g2.k && n2.l === g2.l && n2.r === g2.r) {
        found = true; break;
      }
    }
    if (found) matchCount++; else noMatchCount++;
  }
  log.push('=== SUMMARY ===');
  log.push('ゴースト行数: ' + ghosts.length);
  log.push('既存行コピーと一致: ' + matchCount);
  log.push('一致なし: ' + noMatchCount);
  log.push('ゴースト行は末尾連続か: ' + (ghosts.length > 0 ? (ghosts[ghosts.length - 1].row === lastRow ? 'YES' : 'NO') : 'N/A'));
  if (ghosts.length > 0) {
    log.push('ゴースト行範囲: Row ' + ghosts[0].row + ' ~ Row ' + ghosts[ghosts.length - 1].row);
  }

  // マスター修正プレビューに書き出し
  var logSheet = smc.getSheetByName('マスター修正プレビュー');
  if (logSheet) {
    logSheet.clear();
    for (var li = 0; li < log.length; li++) {
      logSheet.getRange(li + 1, 1).setValue(log[li]);
    }
  }

  return log.join('\n');
}

/**
 * 1回限りのバックフィル: フォーム直近200件のうちマスターに無い行を追加
 * 判定キー: A列表示値 + R列（本名）
 * トリガーから呼ばない。API手動実行のみ。
 */
function oneTimeBackfillMissingRows() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var masterSheet = smc.getSheetByName(SMC_MASTER_SHEET);
  var formSheet = smc.getSheetByName(SMC_FORM_SHEET);
  if (!masterSheet || !formSheet) return { error: 'sheet not found' };

  var tz = Session.getScriptTimeZone();

  // マスターキーセット構築
  var masterLastRow = masterSheet.getLastRow();
  var masterData = masterSheet.getRange(2, 1, masterLastRow - 1, 18).getValues();
  var masterKeys = {};
  for (var i = 0; i < masterData.length; i++) {
    var mTs = masterData[i][0];
    var mR = String(masterData[i][17] || '').trim();
    var mKey;
    if (mTs instanceof Date) {
      mKey = Utilities.formatDate(mTs, tz, 'yyyy/MM/dd HH:mm:ss');
    } else if (mTs) {
      mKey = String(mTs).trim();
    } else {
      continue;
    }
    masterKeys[mKey + '|' + mR] = true;
  }

  // フォーム直近200件
  var formColA = formSheet.getRange('A:A').getValues();
  var formLastRow = formColA.filter(String).length;
  var lookback = Math.min(200, formLastRow - 1);
  var formStart = formLastRow - lookback + 1;
  var formData = formSheet.getRange(formStart, 1, lookback, 41).getValues();

  var added = 0;
  var targetCount = 0;

  for (var fi = 0; fi < formData.length; fi++) {
    var formRow = formData[fi];
    var formTs = formRow[0];
    if (!formTs) continue;

    var formKey;
    if (formTs instanceof Date) {
      formKey = Utilities.formatDate(formTs, tz, 'yyyy/MM/dd HH:mm:ss');
    } else {
      formKey = String(formTs).trim();
    }
    var formR = String(formRow[17] || '').trim();
    var fullKey = formKey + '|' + formR;

    if (masterKeys[fullKey]) continue;
    targetCount++;

    var newRow = new Array(41);
    for (var x = 0; x < 41; x++) newRow[x] = '';
    for (var fc in FORM_TO_MASTER_MAP) {
      var mc = FORM_TO_MASTER_MAP[fc];
      var idx = parseInt(fc);
      if (idx < formRow.length && formRow[idx] !== '' && formRow[idx] !== null) {
        newRow[mc] = formRow[idx];
      }
    }
    if (newRow[0] instanceof Date) {
      newRow[0] = Utilities.formatDate(newRow[0], tz, 'yyyy/MM/dd HH:mm:ss');
    }
    if (newRow[2] instanceof Date) {
      newRow[2] = Utilities.formatDate(newRow[2], tz, 'yyyy/MM/dd');
    }

    masterSheet.appendRow(newRow);
    masterKeys[fullKey] = true;
    added++;
  }

  SpreadsheetApp.flush();

  // 追加後の検証
  var newLastRow = masterSheet.getLastRow();
  var newData = masterSheet.getRange(2, 1, newLastRow - 1, 18).getValues();
  var ghostCount = 0;
  var seenKeys = {};
  var dupCount = 0;
  for (var vi = 0; vi < newData.length; vi++) {
    var vTs = newData[vi][0];
    var vR = String(newData[vi][17] || '').trim();
    if (!vTs) {
      var hasData = false;
      for (var ci = 7; ci <= 17; ci++) {
        if (String(newData[vi][ci] || '').trim()) { hasData = true; break; }
      }
      if (hasData) ghostCount++;
      continue;
    }
    var vKey;
    if (vTs instanceof Date) {
      vKey = Utilities.formatDate(vTs, tz, 'yyyy/MM/dd HH:mm:ss');
    } else {
      vKey = String(vTs).trim();
    }
    var vFullKey = vKey + '|' + vR;
    if (seenKeys[vFullKey]) dupCount++;
    else seenKeys[vFullKey] = true;
  }

  // CX > Q の場合、CXをQに合わせる
  var cxFixed = fixCXOverQ_(masterSheet);

  // C列で時系列ソート
  var sortLR = masterSheet.getLastRow();
  var sortLC = masterSheet.getLastColumn();
  if (sortLR > 2) {
    masterSheet.getRange(2, 1, sortLR - 1, sortLC).sort({ column: 3, ascending: true });
  }

  SpreadsheetApp.flush();

  return {
    targetCount: targetCount,
    added: added,
    totalRows: newLastRow - 1,
    ghostCount: ghostCount,
    dupCount: dupCount,
    cxFixed: cxFixed
  };
}

// ============================================
// トリガー管理
// ============================================

/**
 * SMCフォーム同期トリガーをセットアップ（5分おき）
 */
function setupFormSyncTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'syncRecentFormToMaster') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('syncRecentFormToMaster')
    .timeBased()
    .everyMinutes(5)
    .create();
  return { status: 'trigger_created', interval: '5min' };
}

/**
 * 旧syncトリガーを削除し、v2を5分トリガーに切り替える
 */
function switchToV2Trigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var deletedOld = 0;
  var deletedV2 = 0;
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'syncRecentFormToMaster') {
      ScriptApp.deleteTrigger(triggers[i]);
      deletedOld++;
    }
    if (fn === 'syncFormToMaster_v2') {
      ScriptApp.deleteTrigger(triggers[i]);
      deletedV2++;
    }
  }
  ScriptApp.newTrigger('syncFormToMaster_v2')
    .timeBased()
    .everyMinutes(5)
    .create();

  // 切り替え後のトリガー一覧
  var after = ScriptApp.getProjectTriggers();
  var list = [];
  for (var j = 0; j < after.length; j++) {
    list.push(after[j].getHandlerFunction());
  }
  return {
    deletedOldSync: deletedOld,
    deletedV2Dup: deletedV2,
    createdV2: true,
    currentTriggers: list
  };
}

/**
 * 現在のトリガー一覧を返す
 */
function listTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var list = [];
  for (var i = 0; i < triggers.length; i++) {
    list.push(triggers[i].getHandlerFunction());
  }
  return list;
}

// ============================================
// v2: フォーム→マスター upsert（A+Rキー）
// ============================================

var V2_COL_MAP_ = {};
// A-X (0-23): 1:1（V/W/X含む）
for (var _v2i = 0; _v2i <= 23; _v2i++) { V2_COL_MAP_[_v2i] = _v2i; }
// Y以降: マスターにライフティ(24)・CBS(25)が挿入 → +2ずれ
V2_COL_MAP_[24] = 26; // Form Y(支払方法③)     → Master AA(26)
V2_COL_MAP_[25] = 27; // Form Z(代理店名)       → Master AB(27)
V2_COL_MAP_[26] = 29; // Form AA(支払③金額)     → Master AD(29)
V2_COL_MAP_[27] = 30; // Form AB(郵送確認)      → Master AE(30)
// Form AC(28)〜AH(33) → マスターに該当列なし（スキップ）
V2_COL_MAP_[34] = 27; // Form AI(代理店TikTokその他) → Master AB(27)
V2_COL_MAP_[36] = 32; // Form AK(着金率)        → Master AG(32)
V2_COL_MAP_[37] = 33; // Form AL(ステータス)      → Master AH(33)
V2_COL_MAP_[38] = 34; // Form AM(契約アドレス)    → Master AI(34)
V2_COL_MAP_[39] = 35; // Form AN(契約日)        → Master AJ(35)

// 商談時間1~5: フォーム末尾5列 → マスター DA~DE(col 105-109)
var V2_SHODAN_TIME_FORM_START_ = 41;
var V2_SHODAN_TIME_MASTER_COL_ = 105; // DA = 105列目
var V2_SHODAN_TIME_COUNT_ = 5;

/**
 * フォーム直近30件をA+Rキーでマスターにupsert。別関数。本番未使用。
 */
function syncFormToMaster_v2() {
  if (PropertiesService.getScriptProperties().getProperty('FORM_SYNC_PAUSED') === 'true') {
    return { paused: true };
  }
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var form = smc.getSheetByName(SMC_FORM_SHEET);
  if (!master || !form) return { error: 'sheet not found' };
  var tz = Session.getScriptTimeZone();

  // フォーム直近30件
  var fColA = form.getRange('A:A').getValues();
  var fLast = fColA.filter(String).length;
  if (fLast <= 1) return { error: 'no form data' };
  var lb = Math.min(30, fLast - 1);
  var fStart = fLast - lb + 1;
  var fData = form.getRange(fStart, 1, lb, 41).getValues();

  // マスター全件
  var mLast = master.getLastRow();
  var mData = mLast > 1 ? master.getRange(2, 1, mLast - 1, 41).getValues() : [];

  // A+Rキーマップ
  var keyMap = {}; // key → {row, vals}
  // R列(本名)の存在マップ（重複転記防止）
  var rKeyMap = {}; // R列本名 → {row, vals}
  for (var mi = 0; mi < mData.length; mi++) {
    var k = v2Key_(mData[mi], tz);
    if (k) keyMap[k] = { row: mi + 2, vals: mData[mi] };
    // R列+D列でキー（同一本名でも担当が違えば別行）
    var rName = String(mData[mi][17] || '').trim().replace(/\s+/g, '');
    var dName = String(mData[mi][3] || '').trim();
    if (rName) rKeyMap[rName + '|' + dName] = { row: mi + 2, vals: mData[mi] };
  }

  var appended = 0, updated = 0, deleted = 0, skipped = 0, rDupSkipped = 0;

  for (var fi = 0; fi < fData.length; fi++) {
    var fRow = fData[fi];
    if (!fRow[0]) continue;
    var newRow = v2BuildRow_(fRow, tz);
    var fk = v2Key_(newRow, tz);
    if (!fk) { skipped++; continue; }

    // R列+D列重複チェック: 同じ本名+同じ担当がマスターに既にあればappendせずmerge
    // ※同じ本名でもD列(担当)が違えば別行としてappend
    var fRName = String(newRow[17] || '').trim().replace(/\s+/g, '');
    var fDName = String(newRow[3] || '').trim();
    var rdKey = fRName + '|' + fDName;
    if (!keyMap[fk] && fRName && rKeyMap[rdKey]) {
      // A+Rキーは新規だが、R列+D列が既存 → 既存行にmerge
      var exR = rKeyMap[rdKey];
      var mergedR = v2Merge_(exR.vals, newRow, tz);
      if (!v2Equal_(exR.vals, mergedR)) {
        master.getRange(exR.row, 1, 1, 41).setValues([mergedR]);
        rKeyMap[rdKey] = { row: exR.row, vals: mergedR };
        updated++;
      } else {
        rDupSkipped++;
      }
      continue;
    }

    if (!keyMap[fk]) {
      // 新規（R列+D列も新規）
      master.appendRow(newRow);
      var newRowNum = master.getLastRow();
      // CX列(102列目)にSUM数式をセット
      master.getRange(newRowNum, 102).setFormula('=SUM(P' + newRowNum + ',V' + newRowNum + ',AD' + newRowNum + ',AP' + newRowNum + ',AW' + newRowNum + ',BD' + newRowNum + ',BO' + newRowNum + ',BV' + newRowNum + ',CC' + newRowNum + ',CJ' + newRowNum + ')');
      keyMap[fk] = { row: newRowNum, vals: newRow };
      if (fRName) rKeyMap[rdKey] = { row: newRowNum, vals: newRow };
      appended++;
    } else {
      // 既存 → merge
      var ex = keyMap[fk];
      var merged = v2Merge_(ex.vals, newRow, tz);
      if (!v2Equal_(ex.vals, merged)) {
        master.getRange(ex.row, 1, 1, 41).setValues([merged]);
        keyMap[fk] = { row: ex.row, vals: merged };
        updated++;
      } else {
        skipped++;
      }
    }
  }

  // マスター内: A空ゴースト行削除 + 完全重複削除
  var seen = {};
  var delRows = [];
  var mData2 = master.getLastRow() > 1 ? master.getRange(2, 1, master.getLastRow() - 1, 41).getValues() : [];
  for (var di = 0; di < mData2.length; di++) {
    var aVal = mData2[di][0];
    // A列空 = ゴースト行 → 即削除対象
    if (!aVal || String(aVal).trim() === '') {
      delRows.push(di + 2);
      continue;
    }
    var dk = v2Key_(mData2[di], tz);
    if (!dk) continue;
    if (seen[dk] !== undefined) {
      var prevIdx = seen[dk];
      var prevRow = mData2[prevIdx];
      if (v2Equal_(prevRow, mData2[di])) {
        delRows.push(di + 2);
      }
    } else {
      seen[dk] = di;
    }
  }
  delRows.sort(function(a, b) { return b - a; });
  for (var dri = 0; dri < delRows.length; dri++) {
    master.deleteRow(delRows[dri]);
    deleted++;
  }

  // CX > Q の場合、CXをQに合わせる
  var cxFixed = fixCXOverQ_(master);

  // C列で時系列ソート
  var sortLastRow = master.getLastRow();
  var sortLastCol = master.getLastColumn();
  if (sortLastRow > 2) {
    master.getRange(2, 1, sortLastRow - 1, sortLastCol).sort({ column: 3, ascending: true });
  }

  // CX列(102)が空の行にSUM数式を補完
  var cxFixed2 = 0;
  var finalLastRow = master.getLastRow();
  if (finalLastRow > 1) {
    var cxVals = master.getRange(2, 102, finalLastRow - 1, 1).getValues();
    for (var ci = 0; ci < cxVals.length; ci++) {
      if (cxVals[ci][0] === '' || cxVals[ci][0] === null) {
        var rn = ci + 2;
        master.getRange(rn, 102).setFormula('=SUM(P' + rn + ',V' + rn + ',AD' + rn + ',AP' + rn + ',AW' + rn + ',BD' + rn + ',BO' + rn + ',BV' + rn + ',CC' + rn + ',CJ' + rn + ')');
        cxFixed2++;
      }
    }
  }

  SpreadsheetApp.flush();
  return { appended: appended, updated: updated, deleted: deleted, skipped: skipped, rDupSkipped: rDupSkipped, cxFixed: cxFixed, cxFormulasAdded: cxFixed2 };
}

/**
 * dry-run版。append/update/deleteしない。
 */
function dryRunSyncFormToMaster_v2() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var form = smc.getSheetByName(SMC_FORM_SHEET);
  if (!master || !form) return { error: 'sheet not found' };
  var tz = Session.getScriptTimeZone();

  var fColA = form.getRange('A:A').getValues();
  var fLast = fColA.filter(String).length;
  if (fLast <= 1) return { error: 'no form data' };
  var lb = Math.min(30, fLast - 1);
  var fStart = fLast - lb + 1;
  var fData = form.getRange(fStart, 1, lb, 41).getValues();

  var mLast = master.getLastRow();
  var mData = mLast > 1 ? master.getRange(2, 1, mLast - 1, 41).getValues() : [];

  var keyMap = {};
  for (var mi = 0; mi < mData.length; mi++) {
    var k = v2Key_(mData[mi], tz);
    if (k) keyMap[k] = { row: mi + 2, vals: mData[mi] };
  }

  var wouldAppend = [], wouldUpdate = [], skipList = [], dColSkips = [];

  for (var fi = 0; fi < fData.length; fi++) {
    var fRow = fData[fi];
    if (!fRow[0]) continue;
    var newRow = v2BuildRow_(fRow, tz);
    var fk = v2Key_(newRow, tz);
    if (!fk) { skipList.push({ formRow: fStart + fi, reason: 'no key' }); continue; }

    var a = String(newRow[0] || '').substring(0, 22);
    var l = String(newRow[11] || '').substring(0, 20);
    var r = String(newRow[17] || '');

    if (!keyMap[fk]) {
      wouldAppend.push({ formRow: fStart + fi, a: a, l: l, r: r, reason: 'new' });
    } else {
      var ex = keyMap[fk];

      // D列スキップ追跡
      var exD = String(ex.vals[3] || '');
      var nwD = String(newRow[3] || '');
      if (exD && nwD && exD !== nwD) {
        dColSkips.push({
          masterRow: ex.row, formRow: fStart + fi,
          existingD: exD, newD: nwD,
          reason: '既存値優先'
        });
      }

      var merged = v2Merge_(ex.vals, newRow, tz);
      if (!v2Equal_(ex.vals, merged)) {
        var diffs = [];
        for (var ci = 0; ci < 41; ci++) {
          if (String(ex.vals[ci] || '') !== String(merged[ci] || '')) diffs.push(ci);
        }
        wouldUpdate.push({
          masterRow: ex.row, formRow: fStart + fi,
          a: a, l: String(merged[11] || '').substring(0, 20), r: r,
          reason: 'diff cols: ' + diffs.join(','),
          oldL: String(ex.vals[11] || '').substring(0, 20),
          newL: l
        });
      } else {
        skipList.push({ formRow: fStart + fi, reason: 'identical' });
      }
    }
  }

  // マスター内完全重複チェック
  var seen = {};
  var wouldDelete = [];
  for (var di = 0; di < mData.length; di++) {
    var dk = v2Key_(mData[di], tz);
    if (!dk) continue;
    if (seen[dk] !== undefined) {
      var prevRow = mData[seen[dk]];
      if (v2Equal_(prevRow, mData[di])) {
        wouldDelete.push({
          masterRow: di + 2,
          a: String(mData[di][0] || '').substring(0, 22),
          l: String(mData[di][11] || '').substring(0, 20),
          r: String(mData[di][17] || ''),
          reason: 'exact dup of row ' + (seen[dk] + 2)
        });
      }
    } else {
      seen[dk] = di;
    }
  }

  return {
    formStart: fStart, lookback: lb,
    wouldAppendCount: wouldAppend.length,
    wouldUpdateCount: wouldUpdate.length,
    wouldDeleteCount: wouldDelete.length,
    skipCount: skipList.length,
    dColSkipCount: dColSkips.length,
    wouldAppend: wouldAppend,
    wouldUpdate: wouldUpdate,
    wouldDelete: wouldDelete,
    dColSkips: dColSkips,
    skip: skipList.slice(0, 10)
  };
}

/**
 * マスターシートにGoogleフォームが直接リンクされていないか診断
 */
function diagFormLink() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var formSheet = smc.getSheetByName(SMC_FORM_SHEET);
  var allSheets = smc.getSheets();

  var result = { sheets: [] };

  for (var i = 0; i < allSheets.length; i++) {
    var s = allSheets[i];
    var info = { name: s.getName(), hasForm: false };
    try {
      var url = s.getFormUrl();
      if (url) {
        info.hasForm = true;
        info.formUrl = url;
      }
    } catch (e) {
      info.formError = e.message;
    }
    result.sheets.push(info);
  }

  // マスターシートの最終行付近を確認（フォーム自動挿入の痕跡）
  var mLast = master.getLastRow();
  if (mLast > 1) {
    var tail = master.getRange(Math.max(2, mLast - 4), 1, Math.min(5, mLast - 1), 24).getValues();
    result.masterTail = [];
    for (var ti = 0; ti < tail.length; ti++) {
      var row = tail[ti];
      var aEmpty = !row[0] || String(row[0]).trim() === '';
      var pxHasVal = false;
      for (var ci = 15; ci <= 23; ci++) {
        if (row[ci] !== '' && row[ci] !== null && String(row[ci]).trim() !== '') {
          pxHasVal = true;
          break;
        }
      }
      result.masterTail.push({
        row: Math.max(2, mLast - 4) + ti,
        A: String(row[0] || '').substring(0, 20) || '空',
        D: String(row[3] || '').substring(0, 10) || '空',
        R: String(row[17] || '').substring(0, 15) || '空',
        aEmpty: aEmpty,
        pxHasVal: pxHasVal
      });
    }
  }

  return result;
}

/**
 * P〜X列(15-23)の保護状態を診断: マスター vs フォーム の差分を報告
 */
function diagnosePXProtection() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var form = smc.getSheetByName(SMC_FORM_SHEET);
  if (!master || !form) return { error: 'sheet not found' };
  var tz = Session.getScriptTimeZone();

  var fColA = form.getRange('A:A').getValues();
  var fLast = fColA.filter(String).length;
  var lb = Math.min(30, fLast - 1);
  var fStart = fLast - lb + 1;
  var fData = form.getRange(fStart, 1, lb, 41).getValues();

  var mLast = master.getLastRow();
  var mData = mLast > 1 ? master.getRange(2, 1, mLast - 1, 41).getValues() : [];

  var keyMap = {};
  var rKeyMap = {};
  for (var mi = 0; mi < mData.length; mi++) {
    var k = v2Key_(mData[mi], tz);
    if (k) keyMap[k] = { row: mi + 2, vals: mData[mi] };
    var rName = String(mData[mi][17] || '').trim().replace(/\s+/g, '');
    var dName = String(mData[mi][3] || '').trim();
    if (rName) rKeyMap[rName + '|' + dName] = { row: mi + 2, vals: mData[mi] };
  }

  var colNames = ['P','Q','R','S','T','U','V','W','X'];
  var results = [];

  for (var fi = 0; fi < fData.length; fi++) {
    var fRow = fData[fi];
    if (!fRow[0]) continue;
    var newRow = v2BuildRow_(fRow, tz);
    var fk = v2Key_(newRow, tz);
    if (!fk) continue;

    var fRName = String(newRow[17] || '').trim().replace(/\s+/g, '');
    var fDName = String(newRow[3] || '').trim();
    var rdKey = fRName + '|' + fDName;

    var match = null;
    var matchType = '';
    if (keyMap[fk]) {
      match = keyMap[fk];
      matchType = 'A+R';
    } else if (fRName && rKeyMap[rdKey]) {
      match = rKeyMap[rdKey];
      matchType = 'R+D';
    }

    if (!match) continue;

    // P〜X (15-23) 比較
    var pxDiffs = [];
    for (var ci = 15; ci <= 23; ci++) {
      var mv = match.vals[ci];
      var fv = newRow[ci];
      var ms = String(mv === null || mv === undefined ? '' : mv);
      var fs = String(fv === null || fv === undefined ? '' : fv);
      var mEmpty = (mv === '' || mv === null || typeof mv === 'undefined');
      var merged = v2Merge_(match.vals, newRow, tz);
      var mergedV = String(merged[ci] === null || merged[ci] === undefined ? '' : merged[ci]);
      if (ms !== fs) {
        pxDiffs.push({
          col: colNames[ci - 15],
          idx: ci,
          master: ms.substring(0, 30),
          form: fs.substring(0, 30),
          masterEmpty: mEmpty,
          mergeResult: mergedV.substring(0, 30),
          protected: (ms === mergedV)
        });
      }
    }

    if (pxDiffs.length > 0) {
      results.push({
        formRow: fStart + fi,
        masterRow: match.row,
        matchType: matchType,
        R: String(newRow[17] || '').substring(0, 15),
        D: String(newRow[3] || '').substring(0, 10),
        pxDiffs: pxDiffs
      });
    }
  }

  return {
    totalFormRows: lb,
    totalMatched: results.length + '件にP~X差分あり',
    details: results.slice(0, 20)
  };
}

// --- v2 ヘルパー ---

function v2Key_(row, tz) {
  var a = row[0];
  var r = String(row[17] || '').trim();
  if (!a || !r) return '';
  var aStr;
  if (a instanceof Date) {
    aStr = Utilities.formatDate(a, tz, 'yyyy/MM/dd HH:mm:ss');
  } else {
    // 文字列でもDateに変換してformatDateで正規化（ゼロパディング統一）
    var parsed = new Date(a);
    if (!isNaN(parsed.getTime())) {
      aStr = Utilities.formatDate(parsed, tz, 'yyyy/MM/dd HH:mm:ss');
    } else {
      aStr = String(a).trim();
    }
  }
  if (!aStr) return '';
  return aStr + '|' + r;
}

function v2BuildRow_(fRow, tz) {
  var out = new Array(41);
  for (var i = 0; i < 41; i++) out[i] = '';
  for (var fc in V2_COL_MAP_) {
    var mc = V2_COL_MAP_[fc];
    var idx = parseInt(fc, 10);
    if (idx < fRow.length && fRow[idx] !== '' && fRow[idx] !== null) {
      out[mc] = fRow[idx];
    }
  }
  if (out[0] instanceof Date) out[0] = Utilities.formatDate(out[0], tz, 'yyyy/MM/dd HH:mm:ss');
  if (out[2] instanceof Date) out[2] = Utilities.formatDate(out[2], tz, 'yyyy/MM/dd');
  // AE(30)郵送確認: マスター側に入力規則あり、許可値以外は空にして書込みエラーを防ぐ
  if (out[30] && V2_AE_ALLOWED_.indexOf(String(out[30]).trim()) === -1) out[30] = '';
  return out;
}

function v2Merge_(old, nw, tz) {
  var m = old.slice();
  for (var i = 0; i < 41; i++) {
    var ov = old[i], nv = nw[i];
    if (i === 11) { m[i] = v2BetterL_(ov, nv); continue; }
    if (i === 0) { m[i] = v2LaterA_(ov, nv, tz); continue; }
    if (i === 3) { continue; } // D列は既存値を常に優先（入力規則保護）
    // P〜X列(15-23): マスター側で手動修正した値を保護（既存値があればスキップ）
    if (i >= 15 && i <= 23 && !(ov === '' || ov === null || typeof ov === 'undefined')) { continue; }
    // AE(30)郵送確認: 許可値以外なら既存値を維持
    if (i === 30 && nv && V2_AE_ALLOWED_.indexOf(String(nv).trim()) === -1) { continue; }
    var ob = (ov === '' || ov === null || typeof ov === 'undefined');
    var nb = (nv === '' || nv === null || typeof nv === 'undefined');
    if (ob && !nb) { m[i] = nv; continue; }
    if (!ob && !nb) { m[i] = nv; continue; }
  }
  return m;
}

// マスターL列の入力規則で許可されている値
var V2_L_ALLOWED_ = ['成約', '成約➔CO', '顧客情報に記入', '失注', '継続2→成約', '継続3→成約', '継続4→成約'];
// マスターAE列(30)郵送確認の入力規則で許可されている値
var V2_AE_ALLOWED_ = ['郵送同意を得た', '確認中', '郵送局留め', '自宅以外郵送', 'オンクラス'];

function v2BetterL_(ov, nv) {
  var os = String(ov || ''), ns = String(nv || '');
  if (!os.trim()) {
    // 既存が空→新規値が規則内なら採用
    return v2LAllowed_(ns) ? nv : ov;
  }
  if (!ns.trim()) return ov;
  var or_ = v2Lrank_(os), nr = v2Lrank_(ns);
  if (nr < or_) return ov; // 巻き戻し禁止
  // 新しい値が入力規則外なら既存値を維持
  if (!v2LAllowed_(ns)) return ov;
  return nv;
}

function v2LAllowed_(s) {
  for (var i = 0; i < V2_L_ALLOWED_.length; i++) {
    if (s === V2_L_ALLOWED_[i]) return true;
  }
  return false;
}

/**
 * 非成約行のAP(42)以降を一括クリア
 * L列が「成約」「成約➔CO」を含まない行のみ対象
 */
function clearAPForNonSeiyaku(dryRun) {
  var dry = (dryRun === true || dryRun === 'true');
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var lastRow = master.getLastRow();
  var lastCol = master.getLastColumn();
  if (lastRow <= 1 || lastCol < 42) return { error: 'no data or no AP+ cols', lastCol: lastCol };

  // 全データ読み込み（L列=11, AP=42列目以降）
  var allData = master.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var cleared = [];

  for (var i = 0; i < allData.length; i++) {
    var row = allData[i];
    var l = String(row[11] || '').trim();
    // 成約系はスキップ
    if (l.indexOf('成約') !== -1) continue;

    // AP(col42=index41)以降にデータがあるかチェック
    var hasAPData = false;
    for (var c = 41; c < lastCol; c++) {
      if (row[c] !== '' && row[c] !== null && row[c] !== undefined && String(row[c]).trim() !== '' && String(row[c]) !== '0') {
        hasAPData = true;
        break;
      }
    }
    if (!hasAPData) continue;

    var rowNum = i + 2;
    var apCount = lastCol - 41;
    if (!dry) {
      // AP以降をクリア
      var blank = [];
      for (var bc = 0; bc < apCount; bc++) blank.push('');
      master.getRange(rowNum, 42, 1, apCount).setValues([blank]);
    }
    cleared.push({ row: rowNum, L: l, D: String(row[3] || '') });
  }

  if (!dry && cleared.length > 0) SpreadsheetApp.flush();
  return { dryRun: dry, cleared: cleared.length, rows: cleared.slice(0, 30) };
}

function v2Lrank_(s) {
  if (s.indexOf('成約') !== -1 && s.indexOf('CO') !== -1) return 4;
  if (s.indexOf('成約') !== -1) return 3;
  if (s.indexOf('継続') !== -1) return 2;
  return 1;
}

function v2LaterA_(ov, nv, tz) {
  var od = (ov instanceof Date) ? ov : (ov ? new Date(ov) : null);
  var nd = (nv instanceof Date) ? nv : (nv ? new Date(nv) : null);
  if (!od || isNaN(od.getTime())) return nv || ov;
  if (!nd || isNaN(nd.getTime())) return ov;
  var pick = nd.getTime() >= od.getTime() ? nv : ov;
  if (pick instanceof Date) return Utilities.formatDate(pick, tz, 'yyyy/MM/dd HH:mm:ss');
  return pick;
}

function v2Equal_(a, b) {
  for (var i = 0; i < 41; i++) {
    var av = a[i], bv = b[i];
    var as = String(av || ''), bs = String(bv || '');
    if (as === bs) continue;
    // Date⇔String: getTime比較（同じ瞬間なら同値）
    var ad = (av instanceof Date) ? av : (av ? new Date(av) : null);
    var bd = (bv instanceof Date) ? bv : (bv ? new Date(bv) : null);
    if (ad && bd && !isNaN(ad.getTime()) && !isNaN(bd.getTime()) && ad.getTime() === bd.getTime()) continue;
    return false;
  }
  return true;
}

// ============================================
// マスターベース着金計算
// ============================================

/**
 * SMCマスターから月別着金額を計算する
 * @param {number} targetMonth - 対象月 (1-12)。省略時は当月
 * @param {number} targetYear  - 対象年。省略時は当年
 */
/**
 * マスターからステータス別カウント+売上を取得
 */
function calcMasterStats_(targetMonth, targetYear) {
  var now = new Date();
  var year = targetYear || now.getFullYear();
  var month = targetMonth || (now.getMonth() + 1);
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  if (!master) return [];
  var lastRow = master.getLastRow();
  if (lastRow <= 1) return [];
  var data = master.getRange(2, 1, lastRow - 1, 26).getValues();
  var tz = Session.getScriptTimeZone();
  var byPerson = {};

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var c = row[2];
    if (!rev_isTargetMonth_(c, year, month, tz)) continue;
    var d = String(row[3] || '').trim();
    if (!d || d === 'テスト') continue;
    var l = String(row[11] || '').trim();

    if (!byPerson[d]) byPerson[d] = { name: d, deals: 0, closed: 0, conClosed: 0, onClosed: 0, lost: 0, lostCont: 0, cont: 0, sales: 0, lfApproved: 0, lfTotal: 0, cbsApproved: 0, cbsTotal: 0 };

    // Y列(index 24) ライフティ承認状況
    var lfStatus = String(row[24] || '').trim();
    if (lfStatus === '承認') { byPerson[d].lfApproved++; byPerson[d].lfTotal++; }
    else if (lfStatus === '非承認') { byPerson[d].lfTotal++; }

    // Z列(index 25) CBS承認状況
    var cbsStatus = String(row[25] || '').trim();
    if (cbsStatus === '承認') { byPerson[d].cbsApproved++; byPerson[d].cbsTotal++; }
    else if (cbsStatus === '非承認') { byPerson[d].cbsTotal++; }
    byPerson[d].deals++;

    var mCol = String(row[12] || '').trim(); // M列: オンクラス判定
    if (l === '成約') {
      byPerson[d].closed++;
      byPerson[d].sales += rev_parseNum_(row[16]);
      if (mCol === 'オンクラス') {
        byPerson[d].onClosed++;
      } else {
        byPerson[d].conClosed++;
      }
    } else if (l === '成約➔CO') {
      byPerson[d].closed++;
      byPerson[d].sales += rev_parseNum_(row[16]);
    } else if (l.indexOf('継続') !== -1 && l.indexOf('失注') !== -1) {
      byPerson[d].lostCont++;
    } else if (l === '失注') {
      byPerson[d].lost++;
    } else if (l.indexOf('継続') !== -1) {
      byPerson[d].cont++;
    }
  }

  var result = [];
  for (var name in byPerson) result.push(byPerson[name]);
  return result;
}

/**
 * カレンダーから🟢イベントを取得し、メンバー別休み日数を返す
 */
function getHolidayData_(targetMonth, targetYear) {
  var CALENDAR_ID = 'allforone.namaka@gmail.com';
  // カレンダー名 → v2表示名マッピング
  var CAL_NAME_MAP = {
    // 阿部＝意思決定
    '李信': '意思決定', 'AをAで': '意思決定', '意思決定': '意思決定', '童信': '意思決定', '阿部': '意思決定',
    // 伊東＝ポジティブ
    'ポジティブ': 'ポジティブ', '勝': 'ポジティブ', '勝友美': 'ポジティブ', '伊東': 'ポジティブ', 'ドライ': 'ポジティブ',
    // 新居＝1日1more（旧スクリプトくん）
    'セナ': '1日1more', 'せな': '1日1more', 'スクリプト': '1日1more', '新居': '1日1more', 'スクリプトくん': '1日1more', '1日1more': '1日1more',
    // 久保田＝ヒトコト
    'ヒトコト': 'ヒトコト', '流川': 'ヒトコト', '久保田': 'ヒトコト',
    // 五十嵐＝ぜんぶり
    '本田圭佑': 'ぜんぶり', 'ぜんぶり': 'ぜんぶり', '五十嵐': 'ぜんぶり',
    // 大久保＝言い切り（旧ゴン）
    'ゴン': '言い切り', '大久保': '言い切り', '言い切り': '言い切り',
    // 矢吹＝週1休みくん（旧トニー）
    'トニー': '週1休みくん', '矢吹': '週1休みくん', '週1休みくん': '週1休みくん',
    // 福島＝けつだん
    'ヒカル': 'けつだん', 'けつだん': 'けつだん', '福島': 'けつだん',
    // 辻阪＝ありのまま
    '桓騎': 'ありのまま', 'ありのまま': 'ありのまま', '辻阪': 'ありのまま',
    // 佐々木＝スマイル
    '大飛': 'スマイル', 'スマイル': 'スマイル', '佐々木': 'スマイル',
    // 吉崎＝ゴジータ
    '吉崎': 'ゴジータ', 'ゴジータ': 'ゴジータ',
    // L
    'L': 'L',
    // 荒木＝悟空
    '荒木': '悟空', '悟空': '悟空',
    // やまと
    'こうつさ': 'やまと', 'やまと': 'やまと',
    // 夜神月
    '夜神月': '夜神月',
    // スキップ対象
    '押切': null, 'TAKUYA∞': null, '介入': null, '一歩': null, '龍馬': null, 'サキヨミ': null,
  };
  var now = new Date();
  var year = targetYear || now.getFullYear();
  var month = targetMonth || (now.getMonth() + 1);
  var startDate = new Date(year, month - 1, 1);
  var endDate = new Date(year, month, 0, 23, 59, 59);

  var cal = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!cal) return { counts: {}, byDate: {} };

  var events = cal.getEvents(startDate, endDate);
  var byMember = {};   // { name: { dateKey: 'full'|'half' } }
  var byDate = {};     // { dateKey: [ { name, type } ] }

  for (var i = 0; i < events.length; i++) {
    var title = events[i].getTitle();
    if (title.indexOf('🟢') === -1) continue;

    // 🟢を除去し、名前部分を取得（付随テキストも除去）
    var cleaned = title.replace(/🟢/g, '').replace(/[\s　]+/g, ' ').trim();
    // スペース以降を除去して名前だけ取得、末尾の時間表記も除去
    var namePart = cleaned.split(' ')[0].split('　')[0].split('（')[0].split('~')[0].split('〜')[0];
    namePart = namePart.replace(/\d+時.*$/, '').trim();
    if (!namePart) continue;

    // マッピングで変換
    var v2Name = CAL_NAME_MAP[namePart];
    if (v2Name === null) continue; // 明示的にスキップ
    if (v2Name === undefined) {
      // マッピングにない場合はそのまま使う
      v2Name = namePart;
    }

    var isAllDay = events[i].isAllDayEvent();
    var htype = isAllDay ? 'full' : 'half';

    var eventDate = events[i].getStartTime();
    var dateKey = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    if (!byMember[v2Name]) byMember[v2Name] = {};
    // 同日に終日と半休がある場合、終日を優先
    if (!byMember[v2Name][dateKey] || htype === 'full') {
      byMember[v2Name][dateKey] = htype;
    }

    if (!byDate[dateKey]) byDate[dateKey] = [];
    byDate[dateKey].push({ name: v2Name, type: htype });
  }

  // byDate の重複を除去（同日同名は終日優先）
  for (var dk in byDate) {
    var seen = {};
    var deduped = [];
    for (var di = 0; di < byDate[dk].length; di++) {
      var entry = byDate[dk][di];
      if (!seen[entry.name] || entry.type === 'full') {
        seen[entry.name] = entry;
      }
    }
    for (var sn in seen) deduped.push(seen[sn]);
    byDate[dk] = deduped;
  }

  // メンバー別の日数カウント（終日=1, 半休=0.5）
  var counts = {};
  for (var name in byMember) {
    var total = 0;
    for (var d in byMember[name]) {
      total += byMember[name][d] === 'full' ? 1 : 0.5;
    }
    counts[name] = total;
  }
  return { counts: counts, byDate: byDate };
}

/** テスト用: カレンダーイベント一覧確認 */
function testHolidayData() {
  var data = getHolidayData_();
  Logger.log('parsed: ' + JSON.stringify(data));
}

// ============================================
// 休日データ → スプレッドシート書き出し + PHPサーバー送信
// clasp push のみで反映（デプロイ不要）
// ============================================

/**
 * 休日データをスプレッドシートに書き出し、PHPサーバーにJSON送信
 * time-basedトリガーで1時間ごとに実行
 */
function writeHolidayDataToSheet() {
  var now = new Date();
  var month = now.getMonth() + 1;
  var year = now.getFullYear();

  // 休日データ取得（CalendarApp使用）
  var data = getHolidayData_(month, year);
  if (!data || !data.byDate) {
    Logger.log('休日データなし');
    return;
  }

  // --- 1. スプレッドシートに書き出し（目視確認用） ---
  var MASTER_SS_ID = '1KxHeLmrpdaw1IUhBaQ46UWSHu-8SCRZqcrHOE2hMwDo';
  var ss = SpreadsheetApp.openById(MASTER_SS_ID);
  var sheetName = '休日データ';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  // ヘッダー + データ行を構築
  var rows = [['date', 'name', 'type']];
  var dateKeys = Object.keys(data.byDate).sort();
  for (var di = 0; di < dateKeys.length; di++) {
    var dk = dateKeys[di];
    for (var i = 0; i < data.byDate[dk].length; i++) {
      var h = data.byDate[dk][i];
      rows.push([dk, h.name, h.type]);
    }
  }

  // クリアして書き込み
  if (sheet.getLastRow() > 0) {
    sheet.clearContents();
  }
  if (rows.length > 0) {
    sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  }

  Logger.log('休日データシート更新: ' + (rows.length - 1) + '行, GID=' + sheet.getSheetId());

  // --- 2. PHPサーバーにJSON送信 ---
  var payload = {
    action: 'updateHoliday',
    secret: 'gas_holiday_push_2026',
    month: month,
    year: year,
    byDate: data.byDate,
    counts: data.counts
  };

  try {
    var response = UrlFetchApp.fetch('https://giver.work/sales-dashboard/api-proxy.php', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    Logger.log('PHP push: ' + response.getContentText());
  } catch (e) {
    Logger.log('PHP push error: ' + e.message);
  }
}

/**
 * 休日データ送信の定期トリガーを設定（1時間ごと）
 * Apps Scriptエディタで手動実行してください
 */
function installHolidayPushTrigger() {
  // 既存トリガー削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'writeHolidayDataToSheet') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // 1時間ごとトリガー
  ScriptApp.newTrigger('writeHolidayDataToSheet')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('Holiday push trigger installed (every 1 hour)');
}

function debugCalendarEvents() {
  var cal = CalendarApp.getCalendarById('allforone.namaka@gmail.com');
  if (!cal) { Logger.log('Calendar not found'); return; }
  var now = new Date();
  var start = new Date(now.getFullYear(), now.getMonth(), 1);
  var end = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59);
  var events = cal.getEvents(start, end);
  for (var i = 0; i < events.length; i++) {
    var e = events[i];
    Logger.log(e.getTitle() + ' | ' + e.getStartTime() + ' | allDay=' + e.isAllDayEvent());
  }
  Logger.log('Total events: ' + events.length);
}

/**
 * 全メンバーの10ブロック着金をスプシに書き出す
 */
function exportRevenueSheet(targetMonth, targetYear) {
  var now = new Date();
  var year = parseInt(targetYear) || now.getFullYear();
  var month = parseInt(targetMonth) || (now.getMonth() + 1);
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  var lastRow = master.getLastRow();
  if (lastRow <= 1) return '';
  var data = master.getRange(2, 1, lastRow - 1, 101).getValues();
  var tz = Session.getScriptTimeZone();

  var blocks = [
    {name:'P(15)', amt:15, date:2},
    {name:'V(21)', amt:21, date:2},
    {name:'AD(29)', amt:29, date:2},
    {name:'AP(41)', amt:41, date:38},
    {name:'AW(48)', amt:48, date:45},
    {name:'BD(55)', amt:55, date:52},
    {name:'BO(66)', amt:66, date:63},
    {name:'BV(73)', amt:73, date:70},
    {name:'CC(80)', amt:80, date:77},
    {name:'CJ(87)', amt:87, date:84}
  ];

  var sheetName = year + '年' + month + '月_着金チェック';
  var outSheet = smc.getSheetByName(sheetName);
  if (outSheet) smc.deleteSheet(outSheet);
  outSheet = smc.insertSheet(sheetName);

  var headerRow = ['Row', 'タイムスタンプ', '日付', '担当', '顧客名', 'ステータス', '売上(Q)'];
  for (var hi = 0; hi < blocks.length; hi++) headerRow.push(blocks[hi].name);
  headerRow.push('着金合計');

  var rows = [headerRow];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var l = String(row[11] || '').trim();
    if (l !== '成約' && l !== '成約➔CO') continue;
    var c = row[2];
    if (!rev_isTargetMonth_(c, year, month, tz)) continue;

    var a = row[0];
    var aStr = (a instanceof Date) ? Utilities.formatDate(a, tz, 'yyyy/MM/dd HH:mm') : String(a || '');
    var cStr = (c instanceof Date) ? Utilities.formatDate(c, tz, 'yyyy/MM/dd') : String(c || '');
    var d = String(row[3] || '');
    var e = String(row[4] || '');
    var q = rev_parseNum_(row[16]);

    var outRow = [i + 2, aStr, cStr, d, e, l, q];
    var rowTotal = 0;
    for (var bi = 0; bi < blocks.length; bi++) {
      var dateVal = row[blocks[bi].date];
      var amt = 0;
      if (rev_isTargetMonth_(dateVal, year, month, tz)) {
        amt = rev_parseNum_(row[blocks[bi].amt]);
      }
      outRow.push(amt);
      rowTotal += amt;
    }
    outRow.push(rev_round1_(rowTotal));
    rows.push(outRow);
  }

  outSheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  outSheet.setFrozenRows(1);
  outSheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');

  return { sheet: sheetName, rows: rows.length - 1 };
}

function calcRevenueFromMaster(targetMonth, targetYear) {
  var now = new Date();
  var year = targetYear || now.getFullYear();
  var month = targetMonth || (now.getMonth() + 1);

  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  if (!master) return { error: 'master not found' };

  var lastRow = master.getLastRow();
  if (lastRow <= 1) return { error: 'no data' };
  // CX列=102列目まで読む
  var data = master.getRange(2, 1, lastRow - 1, 102).getValues();
  var tz = Session.getScriptTimeZone();

  var CX_COL = 101; // CX列(0-indexed) = 着金
  var Q_COL = 16;   // Q列(0-indexed) = 売上
  var byPerson = {};
  var grandTotal = 0;
  var grandSales = 0;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var d = String(row[3] || '').trim();
    if (!d || d === 'テスト') continue;
    var c = row[2];
    if (!rev_isTargetMonth_(c, year, month, tz)) continue;

    // 本名→v2名変換
    var v2Name = REAL_NAME_TO_V2[d] || d;

    if (!byPerson[v2Name]) byPerson[v2Name] = { revenue: 0, sales: 0, deals: 0, closed: 0, coCount: 0, coAmount: 0, lost: 0 };
    byPerson[v2Name].deals++;

    var l = String(row[11] || '').trim();
    var qAmt = rev_parseNum_(row[Q_COL]);
    if (l === '成約') {
      byPerson[v2Name].closed++;
      var cxAmt = rev_parseNum_(row[CX_COL]);
      byPerson[v2Name].revenue = rev_round1_(byPerson[v2Name].revenue + cxAmt);
      byPerson[v2Name].sales = rev_round1_(byPerson[v2Name].sales + qAmt);
      grandTotal = rev_round1_(grandTotal + cxAmt);
      grandSales = rev_round1_(grandSales + qAmt);
    } else if (l === '成約➔CO') {
      byPerson[v2Name].closed++;
      byPerson[v2Name].coCount++;
      byPerson[v2Name].coAmount = rev_round1_(byPerson[v2Name].coAmount + qAmt);
      byPerson[v2Name].sales = rev_round1_(byPerson[v2Name].sales + qAmt);
      grandSales = rev_round1_(grandSales + qAmt);
    } else if (l === '見送り' || l === '失注') {
      byPerson[v2Name].lost++;
    }
  }

  // メンバー配列に変換
  var members = [];
  for (var name in byPerson) {
    var p = byPerson[name];
    var total = p.deals || 0;
    var closed = p.closed || 0;
    members.push({
      name: name,
      icon: ICON_MAP[name] || '',
      revenue: p.revenue,
      sales: p.sales,
      deals: total,
      closed: closed,
      closeRate: total > 0 ? rev_round1_((closed / total) * 100) : 0,
      avgPrice: closed > 0 ? rev_round1_(p.revenue / closed) : 0,
      coRevenue: p.coCount,
      coAmount: p.coAmount,
      lost: p.lost
    });
  }
  members.sort(function(a, b) { return b.revenue - a.revenue; });
  for (var ri = 0; ri < members.length; ri++) {
    members[ri].rank = ri + 1;
    members[ri].gapToTop = ri === 0 ? 0 : rev_round1_(members[0].revenue - members[ri].revenue);
  }

  return {
    year: year,
    month: month,
    grandTotal: grandTotal,
    grandSales: grandSales,
    totalRevenue: grandTotal,
    members: members,
    _debug: byPerson
  };
}

/**
 * 指定担当者のダッシュボード値 vs マスター生データを比較（デバッグ用）
 * ?action=debugMember&d=伊東  （全員: d=all）
 */
function debugMemberRows(targetD, targetMonth, targetYear) {
  var now = new Date();
  var year = targetYear || now.getFullYear();
  var month = targetMonth || (now.getMonth() + 1);

  // ダッシュボードの計算値を取得
  var dashData = calcRevenueFromMaster(month, year);
  var dashByName = {};
  if (dashData && dashData.members) {
    for (var di = 0; di < dashData.members.length; di++) {
      var dm = dashData.members[di];
      dashByName[dm.name] = dm;
    }
  }

  // マスターの生データを取得
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  if (!master) return { error: 'master not found' };

  var lastRow = master.getLastRow();
  if (lastRow <= 1) return { error: 'no data' };
  var data = master.getRange(2, 1, lastRow - 1, 102).getValues();
  var tz = Session.getScriptTimeZone();

  // 全担当者 or 指定担当者
  var isAll = (!targetD || targetD === 'all');
  var result = {};

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var d = String(row[3] || '').trim();
    if (!d || d === 'テスト') continue;
    if (!isAll && d !== targetD) continue;
    var c = row[2];
    if (!rev_isTargetMonth_(c, year, month, tz)) continue;

    if (!result[d]) result[d] = { masterRows: [], masterCXTotal: 0, masterQTotal: 0 };

    var e = String(row[4] || '').trim();
    var l = String(row[11] || '').trim();
    var cx = rev_parseNum_(row[101]);
    var q = rev_parseNum_(row[16]);

    if (l === '成約') {
      result[d].masterCXTotal = rev_round1_(result[d].masterCXTotal + cx);
      result[d].masterQTotal = rev_round1_(result[d].masterQTotal + q);
    }

    result[d].masterRows.push({
      row: i + 2,
      E: e,
      L: l,
      Q: q,
      CX: cx
    });
  }

  // ダッシュボード値と比較
  var output = [];
  var targets = isAll ? Object.keys(result) : [targetD];
  for (var ti = 0; ti < targets.length; ti++) {
    var name = targets[ti];
    var v2 = REAL_NAME_TO_V2[name] || name;
    var dashM = dashByName[v2] || { revenue: 0, sales: 0, deals: 0, closed: 0 };
    var raw = result[name] || { masterRows: [], masterCXTotal: 0, masterQTotal: 0 };

    output.push({
      D: name,
      v2Name: v2,
      dashboard: { revenue: dashM.revenue, sales: dashM.sales, deals: dashM.deals, closed: dashM.closed },
      master: { cxTotal: raw.masterCXTotal, qTotal: raw.masterQTotal, rowCount: raw.masterRows.length },
      diff: { revenue: rev_round1_(dashM.revenue - raw.masterCXTotal), sales: rev_round1_(dashM.sales - raw.masterQTotal) },
      rows: raw.masterRows
    });
  }

  return {
    month: year + '/' + month,
    members: output
  };
}

/**
 * ブロック別の着金内訳を返す（デバッグ用）
 */
function calcRevenueDebug(targetMonth, targetYear) {
  var now = new Date();
  var year = targetYear || now.getFullYear();
  var month = targetMonth || (now.getMonth() + 1);

  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  if (!master) return { error: 'master not found' };

  var lastRow = master.getLastRow();
  if (lastRow <= 1) return { error: 'no data' };
  var data = master.getRange(2, 1, lastRow - 1, 101).getValues();
  var tz = Session.getScriptTimeZone();

  var blockDefs = [
    { name: 'P(15)', amtCol: 15, dateCol: 2, dateName: 'C列' },
    { name: 'V(21)', amtCol: 21, dateCol: 2, dateName: 'C列' },
    { name: 'AD(29)', amtCol: 29, dateCol: 2, dateName: 'C列' },
    { name: 'AP(41)', amtCol: 41, dateCol: 38, dateName: 'AM列' },
    { name: 'AW(48)', amtCol: 48, dateCol: 45, dateName: 'AT列' },
    { name: 'BD(55)', amtCol: 55, dateCol: 52, dateName: 'BA列' },
    { name: 'BO(66)', amtCol: 66, dateCol: 63, dateName: 'BL列' },
    { name: 'BV(73)', amtCol: 73, dateCol: 70, dateName: 'BS列' },
    { name: 'CC(80)', amtCol: 80, dateCol: 77, dateName: 'BZ列' },
    { name: 'CJ(87)', amtCol: 87, dateCol: 84, dateName: 'CG列' }
  ];

  var blockTotals = [];
  var grand = 0;

  for (var bi = 0; bi < blockDefs.length; bi++) {
    var bd = blockDefs[bi];
    var total = 0;
    var count = 0;
    var samples = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var _l = String(row[11] || '').trim();
      if (_l !== '成約' && _l !== '成約➔CO') continue;
      if (String(row[3] || '').trim() === 'テスト') continue;

      var dateVal = row[bd.dateCol];
      if (!rev_isTargetMonth_(dateVal, year, month, tz)) continue;

      var amt = rev_parseNum_(row[bd.amtCol]);
      if (amt === 0) continue;

      total = rev_round1_(total + amt);
      count++;
      if (samples.length < 5) {
        samples.push({
          row: i + 2,
          d: String(row[3] || ''),
          amt: amt,
          dateRaw: String(dateVal || '').substring(0, 30)
        });
      }
    }

    grand = rev_round1_(grand + total);
    blockTotals.push({
      block: bd.name,
      dateCol: bd.dateName,
      total: total,
      count: count,
      samples: samples
    });
  }

  return { year: year, month: month, grandTotal: grand, blocks: blockTotals };
}

function rev_isTargetMonth_(val, year, month, tz) {
  if (!val) return false;
  var d;
  if (val instanceof Date) {
    d = val;
  } else {
    var s = String(val).trim();
    if (!s) return false;
    var m = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})/);
    if (m) {
      return parseInt(m[1]) === year && parseInt(m[2]) === month;
    }
    d = new Date(s);
    if (isNaN(d.getTime())) return false;
  }
  return d.getFullYear() === year && (d.getMonth() + 1) === month;
}

function rev_parseNum_(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  var s = String(val).replace(/[,¥￥円万\s%％]/g, '');
  return Number(s) || 0;
}

function rev_round1_(v) {
  return Math.round(v * 10) / 10;
}

// ============================================
// テストソート: 別タブにコピーしてC列ソートを検証
// ============================================
function testSortOnCopy() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  if (!master) return { error: 'master not found' };

  // 既存テストタブがあれば削除
  var existing = smc.getSheetByName('_sort_test');
  if (existing) smc.deleteSheet(existing);

  // マスターをコピー
  var copy = master.copyTo(smc);
  copy.setName('_sort_test');

  var lastRow = copy.getLastRow();
  var lastCol = copy.getLastColumn();
  if (lastRow <= 1) return { error: 'no data' };

  // ソート前のサンプル（末尾5行）
  var beforeRows = copy.getRange(lastRow - 4, 1, 5, Math.min(lastCol, 50)).getValues();
  var beforeSample = [];
  for (var i = 0; i < beforeRows.length; i++) {
    beforeSample.push({ row: lastRow - 4 + i, D: beforeRows[i][3], R: beforeRows[i][17], AR: beforeRows[i][43], AS: beforeRows[i][44] });
  }

  // 全列ソート（getMaxColumnsで全列カバー、AP-CW列のズレ防止）
  var maxCol = copy.getMaxColumns();
  copy.getRange(2, 1, lastRow - 1, maxCol).sort({ column: 3, ascending: true });
  SpreadsheetApp.flush();

  // ソート後: AR列とR列の整合チェック（ランダム10行）
  var afterData = copy.getRange(2, 1, lastRow - 1, Math.min(lastCol, 50)).getValues();
  var mismatches = 0;
  var samples = [];
  // 末尾5行のサンプル
  for (var j = afterData.length - 5; j < afterData.length; j++) {
    if (j < 0) continue;
    samples.push({ row: j + 2, D: afterData[j][3], R: afterData[j][17], AR: afterData[j][43], AS: afterData[j][44] });
  }

  return {
    status: 'ok',
    totalRows: lastRow - 1,
    totalCols: lastCol,
    sortedTab: '_sort_test',
    beforeSample: beforeSample,
    afterSample: samples
  };
}

function applySortFromTest() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  if (!master) return { error: 'master not found' };
  var lastRow = master.getLastRow();
  if (lastRow <= 1) return { error: 'no data' };

  var maxCol = master.getMaxColumns();
  master.getRange(2, 1, lastRow - 1, maxCol).sort({ column: 3, ascending: true });
  SpreadsheetApp.flush();

  // テストタブ削除
  var test = smc.getSheetByName('_sort_test');
  if (test) smc.deleteSheet(test);

  return { sorted: lastRow - 1, cols: lastCol };
}

// ============================================
// マスターD列にプルダウン設定（姓のみ）
// ============================================
function setDColumnValidation() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  if (!master) return { error: 'master not found' };

  // マスターD列から姓のみ（4文字以下）をユニーク取得
  var lastRow = master.getLastRow();
  if (lastRow <= 1) return { error: 'no data' };
  var dVals = master.getRange(2, 4, lastRow - 1, 1).getValues();
  var nameSet = {};
  for (var i = 0; i < dVals.length; i++) {
    var n = String(dVals[i][0] || '').trim();
    if (n && n !== 'テスト' && n.length <= 4) nameSet[n] = true;
  }
  var names = Object.keys(nameSet).sort();

  // D列全体（D2:D10000）にプルダウン設定（警告のみ、入力拒否しない）
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(names, true)
    .setAllowInvalid(true)
    .build();
  master.getRange('D2:D10000').setDataValidation(rule);

  // フォーム回答シートのD列にも同じプルダウンを設定
  var form = smc.getSheetByName(SMC_FORM_SHEET);
  if (form) {
    form.getRange('D2:D10000').setDataValidation(rule);
  }

  // D列: バッジ風（濃い背景色+白文字+太字）
  // 部分一致のため、長い名前を先に配置（堀琉→堀、荒木一→その他、大久保→大内 等）
  var memberColors = [
    ['五十嵐',  '#0D47A1'],
    ['久保田',  '#4A148C'],
    ['佐々木',  '#004D40'],
    ['大久保',  '#1B5E20'],
    ['荒木一',  '#C62828'],
    ['堀琉',    '#00838F'],
    ['ドラゴン', '#EF6C00'],
    ['ヒカル',  '#00695C'],
    ['カルマ',  '#F9A825'],
    ['やまと',  '#4E342E'],
    ['阿部',    '#B71C1C'],
    ['伊東',    '#880E4F'],
    ['辻阪',    '#1A237E'],
    ['新居',    '#006064'],
    ['矢吹',    '#33691E'],
    ['福島',    '#827717'],
    ['嬴政',    '#F57F17'],
    ['森本',    '#E65100'],
    ['柴崎',    '#BF360C'],
    ['川合',    '#3E2723'],
    ['前村',    '#37474F'],
    ['堀',      '#01579B'],
    ['大内',    '#2E7D32'],
    ['鬼道',    '#6A1B9A'],
    ['紅',      '#AD1457'],
    ['ハル',    '#1565C0'],
  ];

  // L列: ステータス背景色（順序重要: 長いテキストを先に）
  var lColors = [
    ['成約➔CO',       '#FFCC80'],
    ['成約➔キャンセル', '#E0E0E0'],
    ['成約➔失注',     '#E0E0E0'],
    ['成約',          '#EF9A9A'],
    ['失注',          '#E0E0E0'],
    ['継続',          '#FFF59D'],
    ['顧客情報に記入', '#A5D6A7'],
  ];

  // 両シートに一括適用
  var sheets = [master];
  if (form) sheets.push(form);

  for (var si = 0; si < sheets.length; si++) {
    var sh = sheets[si];
    var dRange = sh.getRange('D2:D10000');
    var lRange = sh.getRange('L2:L10000');

    // 既存のD列・L列・行全体ルールを除去、その他は保持
    var existing = sh.getConditionalFormatRules();
    var otherRules = [];
    for (var ri = 0; ri < existing.length; ri++) {
      var ranges = existing[ri].getRanges();
      var isManaged = false;
      for (var rri = 0; rri < ranges.length; rri++) {
        var col = ranges[rri].getColumn();
        var ncols = ranges[rri].getNumColumns();
        if ((col === 4 || col === 12) && ncols === 1) isManaged = true;
        if (col === 1 && ncols > 10) isManaged = true; // 行全体ルール
      }
      if (!isManaged) otherRules.push(existing[ri]);
    }

    // D列バッジを最優先（リスト先頭）に配置 → 行全体の背景色に勝つ
    // 部分一致で「佐々木心雪：やまと」→佐々木、「矢吹友一：トニー」→矢吹 等に対応
    var dRules = [];
    for (var mi = 0; mi < memberColors.length; mi++) {
      var mc = memberColors[mi];
      dRules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains(mc[0])
        .setBackground(mc[1])
        .setFontColor('#FFFFFF')
        .setBold(true)
        .setRanges([dRange])
        .build());
    }

    // L列ステータス
    var lRules = [];
    for (var li = 0; li < lColors.length; li++) {
      var lc = lColors[li];
      lRules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains(lc[0])
        .setBackground(lc[1])
        .setRanges([lRange])
        .build());
    }

    // 両シートに行全体の色分けを追加
    var rowRules = [];
    var rowRange = (sh.getName() === SMC_FORM_SHEET)
      ? sh.getRange('A2:AP10000')
      : sh.getRange('A2:DI10000');
    // CO含む → 行全体黄色（成約より先に判定）
    rowRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=FIND("CO",$L2)')
      .setBackground('#FFFF99')
      .setRanges([rowRange])
      .build());
    // 成約（継続2→成約、継続3→成約、継続4→成約 含む）→ 行全体赤
    rowRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=FIND("成約",$L2)')
      .setBackground('#FF9999')
      .setRanges([rowRange])
      .build());

    // D列バッジ → L列ステータス → 行全体ルール → その他既存ルール の順で適用
    sh.setConditionalFormatRules(dRules.concat(lRules).concat(rowRules).concat(otherRules));
  }

  return { members: names.length, list: names, colorsSet: memberColors.length, statusColorsSet: lColors.length };
}

// ============================================
// フォーム回答シート D列 複合名→姓のみ 一括変換
// ============================================
function normalizeDColumn(dryRun) {
  var dry = (dryRun === true || dryRun === 'true' || dryRun === '1');
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var form = smc.getSheetByName(SMC_FORM_SHEET);
  if (!form) return { error: 'フォーム回答シートが見つかりません' };

  var replaceMap = {
    '佐々木心雪：やまと': '佐々木',
    '阿保悠斗：ハル': 'ハル',
    '堀康平：鬼道': '鬼道',
    '荒木一：ドラゴン': 'ドラゴン',
    '大久保友佑悟：ゴン': '大久保',
    '矢吹友一：トニー': '矢吹',
    '迫田健助：紅': '紅',
    '堀琉乃介：北条': '堀琉',
    '森本➔阿部': '阿部'
  };

  var lastRow = form.getLastRow();
  if (lastRow <= 1) return { error: 'データなし' };
  var dVals = form.getRange(2, 4, lastRow - 1, 1).getValues();
  var changes = [];

  for (var i = 0; i < dVals.length; i++) {
    var v = String(dVals[i][0] || '').trim();
    if (replaceMap[v]) {
      changes.push({ row: i + 2, from: v, to: replaceMap[v] });
      if (!dry) dVals[i][0] = replaceMap[v];
    }
  }

  if (!dry && changes.length > 0) {
    form.getRange(2, 4, lastRow - 1, 1).setValues(dVals);
    SpreadsheetApp.flush();
  }

  return { dryRun: dry, totalChanges: changes.length, changes: changes.slice(0, 20) };
}

// ============================================
// _dd シート更新: マスターからユニーク商談者・代理店名を取得
// ============================================
// ============================================
// 監査: 条件付き書式・入力規則・CX数式の状態一覧
// ============================================
function auditSheetRules() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var result = {};
  var targets = [SMC_MASTER_SHEET, SMC_FORM_SHEET];

  for (var t = 0; t < targets.length; t++) {
    var sh = smc.getSheetByName(targets[t]);
    if (!sh) continue;
    var info = { name: targets[t], lastRow: sh.getLastRow(), lastCol: sh.getLastColumn() };

    // 条件付き書式ルール一覧
    var cfRules = sh.getConditionalFormatRules();
    info.cfRulesTotal = cfRules.length;
    info.cfRules = [];
    for (var i = 0; i < cfRules.length; i++) {
      var r = cfRules[i];
      var ranges = r.getRanges();
      var rangeStrs = [];
      for (var j = 0; j < ranges.length; j++) rangeStrs.push(ranges[j].getA1Notation());
      var bc = r.getBooleanCondition();
      var ri = { idx: i, ranges: rangeStrs };
      if (bc) {
        ri.type = String(bc.getCriteriaType());
        ri.values = bc.getCriteriaValues();
        ri.bg = bc.getBackground();
        ri.fc = bc.getFontColor();
        ri.bold = bc.getBold();
      }
      info.cfRules.push(ri);
    }

    // D列(4列目) 入力規則
    var dV2 = sh.getRange(2, 4).getDataValidation();
    info.dValidationRow2 = dV2 ? { type: String(dV2.getCriteriaType()), count: dV2.getCriteriaValues().length, allowInvalid: dV2.getAllowInvalid() } : null;
    var dVLast = sh.getRange(sh.getLastRow(), 4).getDataValidation();
    info.dValidationLastRow = dVLast ? 'あり' : 'なし';
    var dVNew = sh.getRange(Math.min(sh.getLastRow() + 1, 10000), 4).getDataValidation();
    info.dValidationNewRow = dVNew ? 'あり' : 'なし';

    // CX列(102) マスターのみ
    if (targets[t] === SMC_MASTER_SHEET) {
      info.cxRow2 = sh.getRange(2, 102).getFormula() || '(値のみ: ' + sh.getRange(2, 102).getValue() + ')';
      info.cxLastRow = sh.getRange(sh.getLastRow(), 102).getFormula() || '(値のみ: ' + sh.getRange(sh.getLastRow(), 102).getValue() + ')';
      // CX空の行数カウント
      var cxVals = sh.getRange(2, 102, sh.getLastRow() - 1, 1).getFormulas();
      var emptyCount = 0;
      for (var ci = 0; ci < cxVals.length; ci++) {
        if (!cxVals[ci][0]) emptyCount++;
      }
      info.cxEmptyFormulaRows = emptyCount;
    }

    result[targets[t]] = info;
  }
  return result;
}

function refreshSearchDropdowns() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var form = smc.getSheetByName(SMC_FORM_SHEET);
  var dd = smc.getSheetByName('_dd');
  if (!form || !dd) return { error: 'sheet not found' };

  // マスターD列(商談者名・姓のみ)とAB列(代理店名)のユニーク値
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  if (!master) return { error: 'master not found' };
  var mLast = master.getLastRow();
  var dSet = {};
  var abSet = {};
  if (mLast > 1) {
    var mD = master.getRange(2, 4, mLast - 1, 1).getValues();
    for (var i = 0; i < mD.length; i++) {
      var d = String(mD[i][0] || '').trim();
      if (d) dSet[d] = true;
    }
    var mAB = master.getRange(2, 28, mLast - 1, 1).getValues();
    for (var j = 0; j < mAB.length; j++) {
      var v = String(mAB[j][0] || '').trim();
      if (v) abSet[v] = true;
    }
  }

  var dNames = Object.keys(dSet).sort();
  var abNames = Object.keys(abSet).sort();

  // _dd シートのB列(商談者)とC列(代理店名)を更新
  // A列(流入区分)は触らない
  var maxRows = Math.max(dNames.length, abNames.length, dd.getLastRow());

  // B列クリア → 書込み
  if (dd.getLastRow() > 0) {
    dd.getRange(1, 2, dd.getLastRow(), 1).clearContent();
    dd.getRange(1, 3, dd.getLastRow(), 1).clearContent();
  }
  if (dNames.length > 0) {
    var bData = dNames.map(function(n) { return [n]; });
    dd.getRange(1, 2, bData.length, 1).setValues(bData);
  }
  if (abNames.length > 0) {
    var cData = abNames.map(function(n) { return [n]; });
    dd.getRange(1, 3, cData.length, 1).setValues(cData);
  }

  SpreadsheetApp.flush();
  return { members: dNames.length, agents: abNames.length, memberList: dNames, agentSample: abNames.slice(0, 20) };
}

// ============================================
// マスターシート 当月以外非表示
// ============================================

/**
 * 当月以外の行を非表示にする（C列の日付で判定）
 * 連続範囲をバッチ処理で高速化
 * 「全表示」フラグが24時間以内ならスキップ
 */
function autoHideNonCurrentMonth() {
  var props = PropertiesService.getScriptProperties();
  var showUntil = props.getProperty('master_show_all_until');
  if (showUntil && new Date().getTime() < parseInt(showUntil)) {
    return { skipped: true, reason: 'show_all active until ' + new Date(parseInt(showUntil)).toISOString() };
  }

  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var sheet = smc.getSheetByName(SMC_MASTER_SHEET);
  if (!sheet) return { error: 'sheet not found' };

  var now = new Date();
  var year = now.getFullYear();
  var month = now.getMonth() + 1;
  var tz = Session.getScriptTimeZone();

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { error: 'no data' };

  // まず全行を表示してからまとめて非表示にする（バッチ処理）
  sheet.showRows(2, lastRow - 1);

  var dates = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
  var hideRanges = []; // [{start, count}]
  var hideStart = -1;
  var hidden = 0, shown = 0;

  for (var i = 0; i < dates.length; i++) {
    var c = dates[i][0];
    var isTarget = rev_isTargetMonth_(c, year, month, tz);

    if (!isTarget) {
      if (hideStart < 0) hideStart = i + 2;
      hidden++;
    } else {
      if (hideStart >= 0) {
        hideRanges.push({ start: hideStart, count: (i + 2) - hideStart });
        hideStart = -1;
      }
      shown++;
    }
  }
  if (hideStart >= 0) {
    hideRanges.push({ start: hideStart, count: (dates.length + 2) - hideStart });
  }

  // バッチで非表示
  for (var r = 0; r < hideRanges.length; r++) {
    sheet.hideRows(hideRanges[r].start, hideRanges[r].count);
  }

  return { hidden: hidden, shown: shown, ranges: hideRanges.length, month: year + '/' + month };
}

/**
 * 全行を表示して24時間フラグを立てる
 */
function showAllMasterRows() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var sheet = smc.getSheetByName(SMC_MASTER_SHEET);
  if (!sheet) return { error: 'sheet not found' };

  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.showRows(2, lastRow - 1);
  }

  // 24時間フラグ
  var until = new Date().getTime() + 24 * 60 * 60 * 1000;
  PropertiesService.getScriptProperties().setProperty('master_show_all_until', String(until));

  return { ok: true, showUntil: new Date(until).toISOString(), rowsShown: lastRow - 1 };
}

/**
 * autoHideのトリガーを設定（1時間おき）
 */
function setupAutoHideTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'autoHideNonCurrentMonth') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('autoHideNonCurrentMonth')
    .timeBased()
    .everyHours(1)
    .create();
  return { ok: true, trigger: 'autoHideNonCurrentMonth every 1h' };
}

// ============================================
// マスター → 顧客名抽出シート
// ============================================

var EXTRACT_SHEET_NAME = '顧客名抽出（自動）';

/**
 * マスターのB(受講生ID), E(LINE名), R(本名), L(ステータス)を
 * 新規シートに抽出する。成約のみ、CO除外。
 * 5分トリガーで自動更新。
 */
function extractCustomerNames() {
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  if (!master) return { error: 'master not found' };

  var lastRow = master.getLastRow();
  if (lastRow <= 1) return { error: 'no data' };

  var data = master.getRange(2, 1, lastRow - 1, 18).getValues(); // A-R列

  // 成約のみ抽出、CO/キャンセル/失注除外
  var rows = [];
  var seen = {}; // E列(LINE名)で重複除去
  for (var i = 0; i < data.length; i++) {
    var status = String(data[i][11] || '').trim(); // L列
    // 成約のみ（成約➔CO, 成約➔キャンセル, 成約➔失注は除外）
    if (status !== '成約') continue;

    var id = String(data[i][1] || '').trim();   // B列: 受講生ID
    var line = String(data[i][4] || '').trim();  // E列: LINE名
    var real = String(data[i][17] || '').trim(); // R列: 本名

    if (!line) continue;

    // 同じLINE名は最新（後の行）を優先
    seen[line] = { id: id, line: line, real: real };
  }

  // シートに書き出し
  var extract = smc.getSheetByName(EXTRACT_SHEET_NAME);
  if (!extract) {
    extract = smc.insertSheet(EXTRACT_SHEET_NAME);
  }

  // ヘッダー
  var header = [['受講生ID', 'LINE名', '本名', '未登録シート', 'ステータス']];
  var output = [];
  for (var key in seen) {
    var r = seen[key];
    output.push([r.id, r.line, r.real, '', '']);
  }

  // LINE名でソート
  output.sort(function(a, b) { return a[1].localeCompare(b[1]); });

  extract.clear();

  // ルール説明（行1〜6）
  var rules = [
    ['【このシートについて】自動生成シートです。手動編集しないでください。'],
    ['■ データ元：👑商談マスターデータ + 投稿/ホープ/プッシュの3シート'],
    ['■ 白 = 成約済み＆3シート全てに登録済み（問題なし）'],
    ['■ オレンジ = 成約済みだが一部シートに未登録（D列に未登録シート名あり→登録してください）'],
    ['■ 黄色 = 成約済みだがどのシートにも未登録（D列参照→全シートに登録してください）'],
    ['■ グレー = 成約ではないのにシートに登録されている（D列にステータスあり→確認してください）'],
  ];
  extract.getRange(1, 1, rules.length, 1).setValues(rules);
  extract.getRange(1, 1, 1, 1).setFontWeight('bold').setFontSize(11).setFontColor('#B71C1C');
  extract.getRange(2, 1, rules.length - 1, 1).setFontSize(10).setFontColor('#333333');
  extract.getRange(3, 1, 1, 1).setBackground(null);       // 白
  extract.getRange(4, 1, 1, 1).setBackground('#FFE0B2');   // オレンジ
  extract.getRange(5, 1, 1, 1).setBackground('#FFF9C4');   // 黄色
  extract.getRange(6, 1, 1, 1).setBackground('#E0E0E0');   // グレー

  // ヘッダー（行7）
  var headerRow = 7;
  extract.getRange(headerRow, 1, 1, 5).setValues(header);
  extract.getRange(headerRow, 1, 1, 5).setFontWeight('bold').setBackground('#F1F5F9');
  extract.setFrozenRows(headerRow);

  // データ（行8〜）
  if (output.length > 0) {
    extract.getRange(headerRow + 1, 1, output.length, 5).setValues(output);
  }

  // 3シートとの照合で背景色を付ける
  highlightMissingCustomers_(smc, extract, output, data);

  return { ok: true, rows: output.length };
}

/**
 * 抽出シートの顧客を3シート（投稿/ホープ/プッシュ）と照合し背景色をつける
 * 黄色: どのシートにもいない
 * オレンジ: 一部のシートにいるが全部ではない
 */
function highlightMissingCustomers_(smc, extract, output, data) {
  var LP_SS_ID = '1LP_eye2PMswK1OuGfCpzALJkiRE7gsAvrdFjrj5zoik';
  var lp = SpreadsheetApp.openById(LP_SS_ID);

  // 各シートからID一覧 + 名前→ID逆引きを取得
  var nameToId = {}; // 名前 → ID（3シート統合）
  function getIdsFromSheet(sheetGid, idCol, nameCol) {
    var sheets = lp.getSheets();
    var sheet = null;
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() === sheetGid) { sheet = sheets[i]; break; }
    }
    if (!sheet) return {};
    var sheetData = sheet.getDataRange().getValues();
    var ids = {};
    for (var r = 0; r < sheetData.length; r++) {
      var id = String(sheetData[r][idCol] || '').trim();
      var name = String(sheetData[r][nameCol] || '').trim();
      if (id && /^\d{3,}$/.test(id)) {
        ids[id] = name || true;
        // 名前→ID逆引き（ビジネスネーム部分で照合）
        if (name) {
          var bizName = name.replace(/[（(][^）)]*[）)]/g, '').replace(/(メイン|サブ\d*|DFアカウント|DF|YT|IG)$/g, '').trim();
          if (bizName && !nameToId[bizName]) nameToId[bizName] = id;
          if (!nameToId[name]) nameToId[name] = id;
        }
      }
    }
    return ids;
  }

  var postIds = getIdsFromSheet(655724533, 12, 15);   // 投稿: ID=M(12), 名前=P(15)
  var hopeIds = getIdsFromSheet(2140456237, 0, 5);    // ホープ: ID=A(0), 名前=F(5)
  var pushIds = getIdsFromSheet(1948536159, 1, 6);    // プッシュ: ID=B(1), 名前=G(6)

  // IDがない成約者を3シートから補完
  for (var fi = 0; fi < output.length; fi++) {
    if (!output[fi][0] && output[fi][1]) {
      var lineName = output[fi][1];
      var foundId = nameToId[lineName];
      // LINE名で見つからなければ部分一致
      if (!foundId) {
        for (var nk in nameToId) {
          if (nk.indexOf(lineName) >= 0 || lineName.indexOf(nk) >= 0) {
            foundId = nameToId[nk];
            break;
          }
        }
      }
      if (foundId) {
        output[fi][0] = foundId;
      } else {
        output[fi][4] = 'IDなし（3シートに存在しない）';
      }
    }
  }

  // 背景色リセット
  if (output.length > 0) {
    extract.getRange(2, 1, output.length, 5).setBackground(null);
  }

  // マスターの全IDとステータス・名前を取得（成約以外も含む）
  var allMasterById = {}; // id → {status, line, real}
  for (var ai = 0; ai < data.length; ai++) {
    var aId = String(data[ai][1] || '').trim();
    var aStatus = String(data[ai][11] || '').trim();
    var aLine = String(data[ai][4] || '').trim();
    var aReal = String(data[ai][17] || '').trim();
    if (aId && /^\d{3,}$/.test(aId)) {
      allMasterById[aId] = { status: aStatus, line: aLine, real: aReal };
    }
  }

  // 3シートにいるが成約でないIDを検出
  var allSheetIds = {};
  for (var pid in postIds) allSheetIds[pid] = true;
  for (var hid in hopeIds) allSheetIds[hid] = true;
  for (var uid in pushIds) allSheetIds[uid] = true;

  // 成約IDセット
  var seiyakuIds = {};
  for (var oi = 0; oi < output.length; oi++) {
    if (output[oi][0]) seiyakuIds[output[oi][0]] = true;
  }

  // 非成約だがシートにいる人を追加
  var nonSeiyaku = [];
  for (var sid in allSheetIds) {
    if (!seiyakuIds[sid]) {
      var info = allMasterById[sid] || { status: 'マスターに存在しない', line: '', real: '' };
      // マスターに名前がなければ3シートから取得
      var lineName = info.line;
      if (!lineName) {
        lineName = (typeof hopeIds[sid] === 'string' ? hopeIds[sid] : '') ||
                   (typeof pushIds[sid] === 'string' ? pushIds[sid] : '') ||
                   (typeof postIds[sid] === 'string' ? postIds[sid] : '');
      }
      nonSeiyaku.push([sid, lineName, info.real, 'シートにいるが' + info.status, '', 'gray']);
    }
  }

  // 非成約をoutputに追加
  for (var ni = 0; ni < nonSeiyaku.length; ni++) {
    output.push([nonSeiyaku[ni][0], nonSeiyaku[ni][1], nonSeiyaku[ni][2], nonSeiyaku[ni][3], nonSeiyaku[ni][4]]);
  }

  // シート再書き込み（ルール行は親関数で書き込み済み → データ行のみ上書き）
  var dataStartRow = 8; // ルール6行 + ヘッダー1行 の次
  if (output.length > 0) {
    extract.getRange(dataStartRow, 1, output.length, 5).setValues(output);
  }

  var nonSeiyakuStart = output.length - nonSeiyaku.length;

  // バッチ処理: 背景色とD列を一括設定
  var bgColors = [];
  for (var i = 0; i < output.length; i++) {
    var id = output[i][0];
    var isGray = (i >= nonSeiyakuStart);

    if (isGray) {
      bgColors.push(['#E0E0E0', '#E0E0E0', '#E0E0E0', '#E0E0E0', '#E0E0E0']);
      continue;
    }

    var inPost = !!postIds[id];
    var inHope = !!hopeIds[id];
    var inPush = !!pushIds[id];
    var count = (inPost ? 1 : 0) + (inHope ? 1 : 0) + (inPush ? 1 : 0);

    var missing = [];
    if (!inPost) missing.push('投稿');
    if (!inHope) missing.push('ホープ');
    if (!inPush) missing.push('プッシュ');
    output[i][3] = missing.join('・');

    if (count === 0) {
      bgColors.push(['#FFF9C4', '#FFF9C4', '#FFF9C4', '#FFF9C4', '#FFF9C4']);
    } else if (count < 3) {
      bgColors.push(['#FFE0B2', '#FFE0B2', '#FFE0B2', '#FFE0B2', '#FFE0B2']);
    } else {
      bgColors.push([null, null, null, null, null]);
    }
  }

  // 一括書き込み（背景色のみ。データは上で書き込み済み）
  if (output.length > 0) {
    extract.getRange(dataStartRow, 1, output.length, 5).setValues(output);
    extract.getRange(dataStartRow, 1, output.length, 5).setBackgrounds(bgColors);
  }
}

// ============================================
// 3シート参考デザイン作成
// ============================================

/**
 * 4月ホープ数シートを新規スプレッドシートに作成
 * マスターの成約者名を全員入れて、参考デザインを適用
 */
function createAprilHopeSheet() {
  // マスターから成約者を取得
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  if (!master) return { error: 'master not found' };

  var lastRow = master.getLastRow();
  var data = master.getRange(2, 1, lastRow - 1, 18).getValues();

  // 3月ホープシートのID一覧を取得
  var LP_SS_ID = '1LP_eye2PMswK1OuGfCpzALJkiRE7gsAvrdFjrj5zoik';
  var lp = SpreadsheetApp.openById(LP_SS_ID);
  var hopeSheet = null;
  var sheets = lp.getSheets();
  for (var si = 0; si < sheets.length; si++) {
    if (sheets[si].getSheetId() === 2140456237) { hopeSheet = sheets[si]; break; }
  }
  var hopeIds = {};
  if (hopeSheet) {
    var hopeLastRow = Math.min(hopeSheet.getLastRow(), 912);
    var hopeData = hopeSheet.getRange(1, 1, hopeLastRow, 1).getValues();
    for (var hi = 0; hi < hopeData.length; hi++) {
      var hid = String(hopeData[hi][0] || '').trim();
      if (hid && /^\d{3,}$/.test(hid)) hopeIds[hid] = true;
    }
  }

  var seen = {};
  var members = [];
  for (var i = 0; i < data.length; i++) {
    var status = String(data[i][11] || '').trim();
    if (status !== '成約') continue;
    var id = String(data[i][1] || '').trim();
    var line = String(data[i][4] || '').trim();
    var real = String(data[i][17] || '').trim();
    if (!line) continue;
    // 3月ホープシートにいる人だけ
    if (!hopeIds[id]) continue;
    var key = id || line;
    if (seen[key]) continue;
    seen[key] = true;
    var display = real ? line + '（' + real + '）' : line;
    members.push([id, '', '', '', '', display]);
  }

  // 新規スプレッドシート作成
  var ss = SpreadsheetApp.create('4月ホープ数');
  var sheet = ss.getActiveSheet();
  sheet.setName('ホープ数');

  // 4月の日付列（4/1〜4/30）
  var dateCols = [];
  for (var d = 1; d <= 30; d++) dateCols.push('4/' + d);

  var h1 = ['ID', 'ティア', 'チーム', '担当', '契約日', '名前', '合計'];
  h1 = h1.concat(dateCols);
  var totalCols = h1.length;

  // ヘッダー
  sheet.getRange(1, 1, 1, totalCols).setValues([h1]);
  sheet.getRange(1, 1, 1, totalCols)
    .setFontWeight('bold').setBackground('#1E3A5F').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center').setFontSize(10);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(6);

  // メンバーデータ（日付列は❌で初期化）
  var rows = [];
  for (var mi = 0; mi < members.length; mi++) {
    var row = members[mi].slice();
    // 合計列 = COUNTIF数式（❌以外をカウント）
    var rowNum = mi + 2;
    row.push('=COUNTIF(H' + rowNum + ':AK' + rowNum + ',"<>❌")');
    for (var dc = 0; dc < 30; dc++) row.push('❌');
    rows.push(row);
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, totalCols).setValues(rows);
  }

  // プルダウン（日付列: 8列目〜37列目）
  var dropdownRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['❌', '1本', '2本', '3本'], true)
    .setAllowInvalid(false)
    .build();
  if (rows.length > 0) {
    sheet.getRange(2, 8, rows.length, 30).setDataValidation(dropdownRule);
  }

  // 条件付き書式: ❌ → グレー背景
  var dateRange = sheet.getRange(2, 8, Math.max(rows.length, 500), 30);
  var grayRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('❌')
    .setBackground('#E0E0E0')
    .setRanges([dateRange])
    .build();
  sheet.setConditionalFormatRules([grayRule]);

  // 交互背景（名前列まで）
  for (var r = 0; r < rows.length; r++) {
    var bg = (r % 2 === 0) ? '#FFFFFF' : '#F8FAFC';
    sheet.getRange(r + 2, 1, 1, 6).setBackground(bg);
  }

  // 書式
  sheet.getRange(2, 7, rows.length, 31).setHorizontalAlignment('center');
  sheet.getRange(2, 1, rows.length, 1).setHorizontalAlignment('center');
  sheet.getRange(2, 7, rows.length, 1).setFontWeight('bold').setFontColor('#1E40AF');

  // 列幅
  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 45);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 60);
  sheet.setColumnWidth(5, 90);
  sheet.setColumnWidth(6, 180);
  sheet.setColumnWidth(7, 60);
  for (var c = 8; c <= 37; c++) sheet.setColumnWidth(c, 50);

  // 罫線
  sheet.getRange(1, 1, rows.length + 1, totalCols)
    .setBorder(true, true, true, true, true, true, '#E2E8F0', SpreadsheetApp.BorderStyle.SOLID);

  return { ok: true, url: ss.getUrl(), members: rows.length };
}

function createDesignSamples() {
  var LP_SS_ID = '1LP_eye2PMswK1OuGfCpzALJkiRE7gsAvrdFjrj5zoik';
  var lp = SpreadsheetApp.openById(LP_SS_ID);
  var results = {};

  // --- ホープ数 参考デザイン ---
  results.hope = createHopeDesign_(lp);
  // --- プッシュ数 参考デザイン ---
  results.push = createPushDesign_(lp);
  // --- 投稿本数 参考デザイン ---
  results.post = createPostDesign_(lp);

  return results;
}

function createHopeDesign_(lp) {
  var name = '📊 ホープ数（参考デザイン）';
  var existing = lp.getSheetByName(name);
  if (existing) lp.deleteSheet(existing);
  var sheet = lp.insertSheet(name);

  // ヘッダー
  var h1 = ['ID', 'ティア', 'チーム', '担当', '契約日', '名前', '合計', '3/1', '3/2', '3/3', '3/4', '3/5'];
  sheet.getRange(1, 1, 1, h1.length).setValues([h1]);
  sheet.getRange(1, 1, 1, h1.length)
    .setFontWeight('bold').setBackground('#1E3A5F').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center').setFontSize(10);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(6);

  // サンプルデータ（❌=未投稿）
  var samples = [
    ['1001', '👑', 'ほしVIP', 'KEI', '2024/06/01', 'ガガ（ナカムラユミコ）', 16, '2本', '1本', '3本', '❌', '2本'],
    ['1005', '👑', 'ほしVIP', 'KEI', '2024/07/15', 'かっちゃん（神原）', 1041, '3本', '3本', '3本', '3本', '3本'],
    ['1007', '👑', 'ほしVIP', 'KEI', '2024/08/01', 'シャンクス', 532, '1本', '2本', '2本', '3本', '2本'],
    ['', '', '', '', '', '', '', '', '', '', '', ''],
    ['2003', '🟡', 'ほしVIP', 'ほし', '2024/10/01', 'Kanoa(加納永也)', 200, '1本', '❌', '1本', '1本', '1本'],
    ['2006', '🟡', 'ほしVIP', 'ほし', '2024/11/01', 'SHIHO（あやこ）', 150, '❌', '❌', '1本', '❌', '1本'],
  ];
  sheet.getRange(2, 1, samples.length, h1.length).setValues(samples);

  // プルダウン（日付列: 8〜12列目）
  var dropdownRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['❌', '1本', '2本', '3本'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 8, samples.length, 5).setDataValidation(dropdownRule);

  // 条件付き書式: ❌ → グレー背景
  var dateRange = sheet.getRange(2, 8, 500, 5);
  var grayRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('❌')
    .setBackground('#E0E0E0')
    .setRanges([dateRange])
    .build();
  sheet.setConditionalFormatRules([grayRule]);

  // チーム区切り行を薄グレー
  sheet.getRange(5, 1, 1, h1.length).setBackground('#F5F5F5');

  // 交互背景色
  for (var r = 0; r < samples.length; r++) {
    if (samples[r][0] === '') continue;
    var bg = (r % 2 === 0) ? '#FFFFFF' : '#F8FAFC';
    sheet.getRange(r + 2, 1, 1, 6).setBackground(bg);
  }

  // 数値列の書式
  sheet.getRange(2, 7, samples.length, 6).setHorizontalAlignment('center');
  sheet.getRange(2, 1, samples.length, 1).setHorizontalAlignment('center');
  sheet.getRange(2, 2, samples.length, 1).setHorizontalAlignment('center');

  // 列幅
  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 45);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 60);
  sheet.setColumnWidth(5, 90);
  sheet.setColumnWidth(6, 180);
  sheet.setColumnWidth(7, 60);
  for (var c = 8; c <= 12; c++) sheet.setColumnWidth(c, 50);

  // 合計列を太字+色
  sheet.getRange(2, 7, samples.length, 1).setFontWeight('bold').setFontColor('#1E40AF');

  // 罫線
  sheet.getRange(1, 1, samples.length + 1, h1.length)
    .setBorder(true, true, true, true, true, true, '#E2E8F0', SpreadsheetApp.BorderStyle.SOLID);

  // ルール説明
  sheet.getRange(samples.length + 3, 1).setValue('【デザインポイント】');
  sheet.getRange(samples.length + 3, 1).setFontWeight('bold');
  var notes = [
    '✅ ヘッダー: 濃紺背景+白文字で視認性UP',
    '✅ 固定行/列: ヘッダー行+名前列まで固定でスクロールしても迷わない',
    '✅ 交互背景: 白/薄グレーで行の区切りが見やすい',
    '✅ チーム区切り: 空行+薄グレーでグループ分けが明確',
    '✅ 合計列: 太字+青で一目で数値がわかる',
    '✅ ID/ティア: 中央揃えでスッキリ',
    '✅ 列幅: 内容に合わせて最適化',
  ];
  for (var ni = 0; ni < notes.length; ni++) {
    sheet.getRange(samples.length + 4 + ni, 1).setValue(notes[ni]).setFontSize(9).setFontColor('#64748B');
  }

  return { ok: true, sheet: name };
}

function createPushDesign_(lp) {
  var name = '📊 プッシュ数（参考デザイン）';
  var existing = lp.getSheetByName(name);
  if (existing) lp.deleteSheet(existing);
  var sheet = lp.insertSheet(name);

  var h1 = ['', 'ID', 'ティア', 'チーム', '担当', '契約日', '名前', '合計', '3/1', '3/2', '3/3', '3/4', '3/5'];
  sheet.getRange(1, 1, 1, h1.length).setValues([h1]);
  sheet.getRange(1, 1, 1, h1.length)
    .setFontWeight('bold').setBackground('#065F46').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center').setFontSize(10);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(7);

  var samples = [
    ['', '1001', '👑', 'ほしVIP', 'KEI', '2024/06/01', 'ガガ（ナカムラユミコ）', 2, '❌', '❌', '1本', '❌', '1本'],
    ['', '1005', '👑', 'ほしVIP', 'KEI', '2024/07/15', 'かっちゃん（神原）', 40, '3本', '2本', '1本', '2本', '3本'],
    ['', '1007', '👑', 'ほしVIP', 'KEI', '2024/08/01', 'シャンクス', 30, '1本', '2本', '1本', '3本', '2本'],
    ['', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['', '2003', '🟡', 'ほしVIP', 'ほし', '2024/10/01', 'Kanoa(加納永也)', 15, '1本', '❌', '1本', '1本', '2本'],
    ['', '2006', '🟡', 'ほしVIP', 'ほし', '2024/11/01', 'SHIHO（あやこ）', 10, '❌', '1本', '1本', '❌', '1本'],
  ];
  sheet.getRange(2, 1, samples.length, h1.length).setValues(samples);

  // プルダウン（日付列: 9〜13列目）
  var dropdownRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['❌', '1本', '2本', '3本'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 9, samples.length, 5).setDataValidation(dropdownRule);

  // 条件付き書式: ❌ → グレー背景
  var dateRange = sheet.getRange(2, 9, 500, 5);
  var grayRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('❌')
    .setBackground('#E0E0E0')
    .setRanges([dateRange])
    .build();
  sheet.setConditionalFormatRules([grayRule]);

  sheet.getRange(5, 1, 1, h1.length).setBackground('#F5F5F5');
  for (var r = 0; r < samples.length; r++) {
    if (!samples[r][1]) continue;
    var bg = (r % 2 === 0) ? '#FFFFFF' : '#F0FDF4';
    sheet.getRange(r + 2, 1, 1, 7).setBackground(bg);
  }

  sheet.getRange(2, 8, samples.length, 6).setHorizontalAlignment('center');
  sheet.getRange(2, 2, samples.length, 1).setHorizontalAlignment('center');
  sheet.getRange(2, 3, samples.length, 1).setHorizontalAlignment('center');
  sheet.getRange(2, 8, samples.length, 1).setFontWeight('bold').setFontColor('#065F46');

  sheet.setColumnWidth(1, 30);
  sheet.setColumnWidth(2, 60);
  sheet.setColumnWidth(3, 45);
  sheet.setColumnWidth(4, 90);
  sheet.setColumnWidth(5, 60);
  sheet.setColumnWidth(6, 90);
  sheet.setColumnWidth(7, 180);
  sheet.setColumnWidth(8, 60);
  for (var c = 9; c <= 13; c++) sheet.setColumnWidth(c, 50);

  sheet.getRange(1, 1, samples.length + 1, h1.length)
    .setBorder(true, true, true, true, true, true, '#D1FAE5', SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(samples.length + 3, 1).setValue('【デザインポイント】').setFontWeight('bold');
  var notes = [
    '✅ ヘッダー: 深緑背景で投稿シートと区別',
    '✅ 交互背景: 白/薄緑で清潔感',
    '✅ 合計列: 太字+深緑で強調',
    '✅ プルダウン: ❌/1本/2本/3本（❌=グレー背景）',
  ];
  for (var ni = 0; ni < notes.length; ni++) {
    sheet.getRange(samples.length + 4 + ni, 1).setValue(notes[ni]).setFontSize(9).setFontColor('#64748B');
  }

  return { ok: true, sheet: name };
}

function createPostDesign_(lp) {
  var name = '📊 投稿本数（参考デザイン）';
  var existing = lp.getSheetByName(name);
  if (existing) lp.deleteSheet(existing);
  var sheet = lp.insertSheet(name);

  var h1 = ['ID', 'ティア', 'チーム', '担当', '契約日', '名前', 'SNS', '合計', '3/1', '3/2', '3/3', '3/4', '3/5'];
  sheet.getRange(1, 1, 1, h1.length).setValues([h1]);
  sheet.getRange(1, 1, 1, h1.length)
    .setFontWeight('bold').setBackground('#7C2D12').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center').setFontSize(10);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(7);

  // SNS色分け
  var snsBg = { 'TT': '#F0F0F0', 'YT': '#FFF0F0', 'IG': '#F0F0FF' };

  var samples = [
    ['1001', '👑', 'ほしVIP', 'KEI', '2024/06/01', 'ガガ（ナカムラユミコ）', 'TT', 30, '1本', '❌', '2本', '1本', '❌'],
    ['', '', '', '', '', 'ガガ（ナカムラユミコ）', 'YT', 5, '❌', '❌', '1本', '❌', '❌'],
    ['', '', '', '', '', 'ガガ（ナカムラユミコ）', 'IG', 10, '❌', '1本', '❌', '1本', '❌'],
    ['1005', '👑', 'ほしVIP', 'KEI', '2024/07/15', 'かっちゃん（神原）', 'TT', 120, '3本', '2本', '3本', '3本', '3本'],
    ['', '', '', '', '', 'かっちゃん（神原）', 'YT', 20, '1本', '❌', '1本', '1本', '❌'],
    ['', '', '', '', '', 'かっちゃん（神原）', 'IG', 45, '2本', '1本', '2本', '1本', '2本'],
    ['', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['2003', '🟡', 'ほしVIP', 'ほし', '2024/10/01', 'Kanoa(加納永也)', 'TT', 60, '3本', '2本', '1本', '2本', '3本'],
    ['', '', '', '', '', 'Kanoa(加納永也)', 'YT', 8, '❌', '1本', '❌', '❌', '1本'],
    ['', '', '', '', '', 'Kanoa(加納永也)', 'IG', 15, '1本', '❌', '1本', '1本', '❌'],
  ];
  sheet.getRange(2, 1, samples.length, h1.length).setValues(samples);

  // プルダウン設定（日付列: 9列目〜13列目）
  var dropdownRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['❌', '1本', '2本', '3本'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 9, samples.length, 5).setDataValidation(dropdownRule);

  // SNS別の背景色 + ❌セルをグレー
  for (var r = 0; r < samples.length; r++) {
    var sns = samples[r][6];
    if (sns && snsBg[sns]) {
      sheet.getRange(r + 2, 7, 1, 1).setBackground(snsBg[sns]).setFontWeight('bold');
    }
    if (sns === 'TT') {
      sheet.getRange(r + 2, 1, 1, h1.length)
        .setBorder(true, null, null, null, null, null, '#94A3B8', SpreadsheetApp.BorderStyle.SOLID);
    }
    if (samples[r][0] === '' && samples[r][6] === '') {
      sheet.getRange(r + 2, 1, 1, h1.length).setBackground('#F5F5F5');
    }
    // ❌セルをグレー背景
    for (var dc = 8; dc <= 12; dc++) {
      if (samples[r][dc] === '❌') {
        sheet.getRange(r + 2, dc + 1, 1, 1).setBackground('#E0E0E0').setHorizontalAlignment('center');
      } else if (samples[r][dc]) {
        sheet.getRange(r + 2, dc + 1, 1, 1).setBackground(null).setHorizontalAlignment('center');
      }
    }
  }

  sheet.getRange(2, 8, samples.length, 1).setHorizontalAlignment('center');
  sheet.getRange(2, 1, samples.length, 1).setHorizontalAlignment('center');
  sheet.getRange(2, 2, samples.length, 1).setHorizontalAlignment('center');
  sheet.getRange(2, 8, samples.length, 1).setFontWeight('bold').setFontColor('#7C2D12');

  // 条件付き書式: ❌ → グレー背景
  var dateRange = sheet.getRange(2, 9, 500, 5);
  var grayRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('❌')
    .setBackground('#E0E0E0')
    .setRanges([dateRange])
    .build();
  sheet.setConditionalFormatRules([grayRule]);

  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 45);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 60);
  sheet.setColumnWidth(5, 90);
  sheet.setColumnWidth(6, 180);
  sheet.setColumnWidth(7, 45);
  sheet.setColumnWidth(8, 60);
  for (var c = 9; c <= 13; c++) sheet.setColumnWidth(c, 45);

  sheet.getRange(1, 1, samples.length + 1, h1.length)
    .setBorder(true, true, true, true, true, true, '#FECACA', SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(samples.length + 3, 1).setValue('【デザインポイント】').setFontWeight('bold');
  var notes = [
    '✅ ヘッダー: 深赤背景で3シート中もっとも目立つ',
    '✅ SNS列: TT=グレー / YT=薄赤 / IG=薄青 で色分け',
    '✅ TT行に上ボーダー: 1人分の3行グループが明確',
    '✅ ID/ティア/チーム/担当/契約日はTT行のみに記載（YT/IG行は名前+SNSだけ）',
    '✅ 現状の問題: 列が多すぎて見づらい → 不要列を削除 or 非表示に',
    '✅ 提案: A-Lの不要列を削減し、重要情報を左に集約',
  ];
  for (var ni = 0; ni < notes.length; ni++) {
    sheet.getRange(samples.length + 4 + ni, 1).setValue(notes[ni]).setFontSize(9).setFontColor('#64748B');
  }

  return { ok: true, sheet: name };
}

// ============================================
// 成約者を3シート（投稿/ホープ/プッシュ）に自動追加
// ============================================

/**
 * マスターの成約者で3シートにいない人を各シートの末尾に追加
 * 既存データは一切削除・変更しない（追加のみ）
 */
function addMissingMembersToSheets(targetSheet) {
  var LP_SS_ID = '1LP_eye2PMswK1OuGfCpzALJkiRE7gsAvrdFjrj5zoik';

  // マスターから成約者を取得（B=ID, E=LINE名, R=本名）
  var smc = SpreadsheetApp.openById(SMC_SS_ID);
  var master = smc.getSheetByName(SMC_MASTER_SHEET);
  if (!master) return { error: 'master not found' };

  var lastRow = master.getLastRow();
  if (lastRow <= 1) return { error: 'no data' };
  var data = master.getRange(2, 1, lastRow - 1, 18).getValues();

  // 成約者マップ: ID → {line, real}（ID必須）
  var seiyaku = {};
  for (var i = 0; i < data.length; i++) {
    var status = String(data[i][11] || '').trim();
    if (status !== '成約') continue;
    var id = String(data[i][1] || '').trim();
    var line = String(data[i][4] || '').trim();
    if (!id || !line) continue;
    if (!seiyaku[id]) {
      var real = String(data[i][17] || '').trim();
      var displayName = real ? line + '（' + real + '）' : line;
      seiyaku[id] = { line: line, real: real, display: displayName };
    }
  }

  var lp = SpreadsheetApp.openById(LP_SS_ID);
  var results = {};

  // --- ホープシート (gid=2140456237): A=ID, F=名前 ---
  if (!targetSheet || targetSheet === 'hope') try {
  results.hope = addToSheet_(lp, 2140456237, seiyaku, {
    idCol: 0,     // A列
    nameCol: 5,   // F列
    makeRow: function(id, info, totalCols) {
      var row = new Array(totalCols).fill('');
      row[0] = id;
      row[5] = info.display;
      return row;
    }
  });

  } catch(e) { results.hope = { error: e.message }; }

  // --- プッシュシート (gid=1948536159): B=ID, G=名前 ---
  if (!targetSheet || targetSheet === 'push') try {
  results.push = addToSheet_(lp, 1948536159, seiyaku, {
    idCol: 1,     // B列
    nameCol: 6,   // G列
    makeRow: function(id, info, totalCols) {
      var row = new Array(totalCols).fill('');
      row[1] = id;
      row[6] = info.display;
      return row;
    }
  });

  } catch(e) { results.push = { error: e.message }; }

  // --- 投稿シート (gid=655724533): M(12)=ID, P(15)=名前, B(1)=TT/YT/IG ---
  if (!targetSheet || targetSheet === 'post') try {
  results.post = addToPostSheet_(lp, 655724533, seiyaku);
  } catch(e) { results.post = { error: e.message }; }

  return results;
}

/**
 * 指定シートにIDがない成約者を末尾に追加
 */
function addToSheet_(lp, sheetGid, seiyaku, opts) {
  var sheets = lp.getSheets();
  var sheet = null;
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === sheetGid) { sheet = sheets[i]; break; }
  }
  if (!sheet) return { error: 'sheet not found', gid: sheetGid };

  // 既存ID+名前を取得
  var sheetData = sheet.getDataRange().getValues();
  var existingIds = {};
  var existingNames = {};
  for (var r = 0; r < sheetData.length; r++) {
    var id = String(sheetData[r][opts.idCol] || '').trim();
    var name = String(sheetData[r][opts.nameCol] || '').trim();
    if (id) existingIds[id] = true;
    if (name) existingNames[name] = true;
  }

  // 不足者を特定（ID or 名前で照合）
  var toAdd = [];
  for (var id in seiyaku) {
    if (existingIds[id]) continue;
    if (existingNames[seiyaku[id].display]) continue;
    if (existingNames[seiyaku[id].line]) continue;
    toAdd.push(id);
  }

  if (toAdd.length === 0) return { added: 0 };

  // ID列の最後のデータ行を探す（空行を飛ばさない）
  var lastDataRow = 0;
  for (var lr = sheetData.length - 1; lr >= 0; lr--) {
    var lrId = String(sheetData[lr][opts.idCol] || '').trim();
    var lrName = String(sheetData[lr][opts.nameCol] || '').trim();
    if (lrId || lrName) { lastDataRow = lr + 1; break; }
  }
  if (lastDataRow === 0) lastDataRow = sheetData.length;

  // 書き込みは名前列+1列分だけ（軽量化）
  var writeCols = Math.max(opts.idCol, opts.nameCol) + 1;
  var newRows = [];
  for (var ti = 0; ti < toAdd.length; ti++) {
    newRows.push(opts.makeRow(toAdd[ti], seiyaku[toAdd[ti]], writeCols));
  }

  var insertRow = lastDataRow + 1;
  sheet.getRange(insertRow, 1, newRows.length, writeCols).setValues(newRows);
  sheet.getRange(insertRow, 1, newRows.length, writeCols).setBackground('#FFCDD2');

  return { added: toAdd.length, startRow: insertRow };
}

/**
 * 投稿シートに追加（TT/YT/IG の3行セット）
 */
function addToPostSheet_(lp, sheetGid, seiyaku) {
  var sheets = lp.getSheets();
  var sheet = null;
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === sheetGid) { sheet = sheets[i]; break; }
  }
  if (!sheet) return { error: 'sheet not found' };

  // M列(13列目)とP列(16列目)を読んでID+名前で照合（軽量化）
  var lastRow = sheet.getLastRow();
  var checkData = sheet.getRange(1, 13, lastRow, 4).getValues(); // M〜P列
  var existingIds = {};
  var existingNames = {};
  for (var r = 0; r < checkData.length; r++) {
    var id = String(checkData[r][0] || '').trim(); // M列
    var name = String(checkData[r][3] || '').trim(); // P列
    if (id) existingIds[id] = true;
    if (name) existingNames[name] = true;
  }

  var toAdd = [];
  for (var id in seiyaku) {
    if (existingIds[id]) continue;
    if (existingNames[seiyaku[id].display]) continue;
    if (existingNames[seiyaku[id].line]) continue;
    toAdd.push(id);
  }

  if (toAdd.length === 0) return { added: 0 };

  // 16列分だけ書き込み（A〜P）
  var writeCols = 16;
  var newRows = [];
  for (var ti = 0; ti < toAdd.length; ti++) {
    var id = toAdd[ti];
    var info = seiyaku[id];
    var ttRow = new Array(writeCols).fill('');
    ttRow[1] = 'TT'; ttRow[12] = id; ttRow[15] = info.display;
    newRows.push(ttRow);
    var ytRow = new Array(writeCols).fill('');
    ytRow[1] = 'YT'; ytRow[15] = info.display + 'YT';
    newRows.push(ytRow);
    var igRow = new Array(writeCols).fill('');
    igRow[1] = 'IG'; igRow[15] = info.display + 'IG';
    newRows.push(igRow);
  }

  var insertRow = lastRow + 1;
  sheet.getRange(insertRow, 1, newRows.length, writeCols).setValues(newRows);
  sheet.getRange(insertRow, 1, newRows.length, writeCols).setBackground('#FFCDD2');

  return { added: toAdd.length, rows: newRows.length, startRow: insertRow };
}
