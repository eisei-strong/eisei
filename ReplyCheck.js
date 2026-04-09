// ReplyCheck.js - 受講生への返信漏れ検知・講師通知
// Chatwork 1:1チャット（約400室）を監視し、未返信を検出して通知する

// ===== 設定 =====
var RC_TOKEN = PropertiesService.getScriptProperties().getProperty('RC_TOKEN');
var RC_STALE_HOURS = 0;        // 0=全件表示（15h/20h/24hでアイコン変化）
var RC_BATCH_SIZE = 80;        // ディープスキャン1回あたりの最大ルーム数
var RC_SLEEP_MS = 300;         // API呼び出し間隔（ms）
var RC_ROOM_KEYWORD = '1:1';   // 対象ルーム名のキーワード
var RC_NOTIFY_ROOM_ID = 432896678; // 通知先ルームID
var RC_OWNER_NAME = '宮津宏始'; // このトークンの所有者名（ルーム名判定用）
// 講師・運営アカウントID一覧
var RC_STAFF_IDS = [
  7607018,   // レイ
  10144671,  // 嬴政ver2.0
  10945552,  // 嬴政2026
  3602185,   // ゴール
  5264121,   // ハイエフィ(TAKUYA∞)
  10933943,  // 一歩
  10733308,  // ルルーシュ
  7528759,   // さかね(魯迅)
  9855353,   // ハヤ
  11180179,  // 煉獄
  10754430,  // 嬴政ver3.0
  10541954,  // 押切
  9939321,   // 鬼スピ(ゼロワン)
  10659454,  // SMC運営
  10578812,  // エース〔プロ講師〕
  9333978,   // 炭治郎プロ講師
  10572112,  // 緑谷出久【プロ講師】
  11169713   // 代表:ほし（ブレスル）
];
var RC_MSG_LIMIT = 5000;       // 1メッセージあたりの文字数上限（超えたら分割送信）

// ===== Chatwork API ヘルパー =====
function rcFetch_(method, path, payload) {
  var url = 'https://api.chatwork.com/v2' + path;
  var options = {
    method: method,
    headers: { 'X-ChatWorkToken': RC_TOKEN },
    muteHttpExceptions: true
  };
  if (payload) {
    options.contentType = 'application/x-www-form-urlencoded';
    options.payload = payload;
  }
  var res = UrlFetchApp.fetch(url, options);
  var code = res.getResponseCode();
  if (code === 204) return [];
  if (code === 429) {
    Logger.log('Rate limited: ' + res.getHeaders()['x-ratelimit-reset']);
    throw new Error('RC_RATE_LIMITED');
  }
  if (code >= 400) throw new Error('CW API Error ' + code + ': ' + res.getContentText().substring(0, 200));
  var text = res.getContentText();
  return text ? JSON.parse(text) : [];
}

// ===== クイックスキャン（API 2回のみ） =====
// unread_num > 0 のルームを検出（未読＝確実に未返信）
function quickScanUnreplied() {
  var me = rcFetch_('get', '/me');
  var rooms = rcFetch_('get', '/rooms');

  var unreplied = [];
  for (var i = 0; i < rooms.length; i++) {
    var r = rooms[i];
    // 1:1チャットかダイレクトメッセージで未読ありのもの
    var is1on1 = r.type === 'direct' || (r.name && r.name.indexOf(RC_ROOM_KEYWORD) >= 0);
    var isClosed = r.name && r.name.indexOf('❌') >= 0;
    var isOwner = r.name && r.name.indexOf('嬴政') >= 0;
    if (is1on1 && !isClosed && !isOwner && r.unread_num > 0) {
      var elapsedHours = (Date.now() / 1000 - r.last_update_time) / 3600;
      if (elapsedHours >= RC_STALE_HOURS) {
        unreplied.push({
          roomId: r.room_id,
          roomName: r.name || 'DM',
          unreadCount: r.unread_num,
          lastUpdate: r.last_update_time,
          elapsedHours: elapsedHours
        });
      }
    }
  }

  // 24時間に近い順（経過時間が長い順）
  unreplied.sort(function(a, b) { return b.elapsedHours - a.elapsedHours; });

  Logger.log('クイックスキャン完了: ' + unreplied.length + '件の未読1:1チャット');
  return { myId: me.account_id, myName: me.name, unreplied: unreplied };
}

// ===== ディープスキャン（バッチ処理） =====
// 最終メッセージの送信者をチェックし「既読だが未返信」も検出
function deepScanUnreplied() {
  var me = rcFetch_('get', '/me');
  var myId = me.account_id;
  var rooms = rcFetch_('get', '/rooms');

  // 1:1チャットをフィルタ
  var targets = [];
  for (var i = 0; i < rooms.length; i++) {
    var r = rooms[i];
    var is1on1 = r.type === 'direct' || (r.name && r.name.indexOf(RC_ROOM_KEYWORD) >= 0);
    var isClosed = r.name && r.name.indexOf('❌') >= 0;
    var isOwner = r.name && r.name.indexOf('嬴政') >= 0;
    if (is1on1 && !isClosed && !isOwner) targets.push(r);
  }

  // 最近更新された順にソート
  targets.sort(function(a, b) { return b.last_update_time - a.last_update_time; });

  // 前回のスキャン位置を取得（バッチ継続用）
  var props = PropertiesService.getScriptProperties();
  var offset = parseInt(props.getProperty('RC_DEEP_OFFSET') || '0', 10);
  if (offset >= targets.length) offset = 0;

  var unreplied = [];
  var now = Math.floor(Date.now() / 1000);
  var apiCalls = 0;
  var end = Math.min(offset + RC_BATCH_SIZE, targets.length);

  Logger.log('ディープスキャン: ' + offset + '〜' + end + ' / ' + targets.length + '室');

  for (var j = offset; j < end; j++) {
    var room = targets[j];
    try {
      Utilities.sleep(RC_SLEEP_MS);
      var msgs = rcFetch_('get', '/rooms/' + room.room_id + '/messages?force=1');
      apiCalls++;

      if (!msgs || msgs.length === 0) continue;

      // 最後のメッセージを確認
      var last = msgs[msgs.length - 1];
      if (!last.account || RC_STAFF_IDS.indexOf(last.account.account_id) >= 0) continue;

      // 相手が最後に送信 → 未返信
      var elapsed = Math.floor((now - last.send_time) / 3600);
      if (elapsed >= RC_STALE_HOURS) {
        unreplied.push({
          roomId: room.room_id,
          roomName: room.name || last.account.name,
          senderName: last.account.name,
          preview: last.body.substring(0, 80).replace(/\n/g, ' '),
          hoursAgo: elapsed,
          sendTime: last.send_time
        });
      }
    } catch (e) {
      if (e.message === 'RC_RATE_LIMITED') {
        Logger.log('Rate limited at index ' + j + ', saving offset for next run');
        props.setProperty('RC_DEEP_OFFSET', String(j));
        break;
      }
      Logger.log('Error room ' + room.room_id + ': ' + e.message);
    }
  }

  // 次回のオフセットを保存（全完了したらリセット）
  var nextOffset = (end >= targets.length) ? 0 : end;
  props.setProperty('RC_DEEP_OFFSET', String(nextOffset));

  // 古い順にソート（長時間放置が上）
  unreplied.sort(function(a, b) { return b.hoursAgo - a.hoursAgo; });

  Logger.log('ディープスキャン結果: ' + unreplied.length + '件（API ' + apiCalls + '回）');
  return { myId: myId, unreplied: unreplied, scanned: end - offset, total: targets.length, nextOffset: nextOffset };
}

// ===== 統合チェック＆通知 =====
function checkAndNotifyUnreplied() {
  // クイックスキャン（未読ベース、API 2回）
  var quick = quickScanUnreplied();

  if (quick.unreplied.length === 0) {
    Logger.log('未返信なし - 通知スキップ');
    return;
  }

  // 各ルームの受講生名を解決
  var myId = quick.myId;
  resolveStudentNames_(quick.unreplied, myId);

  // ※quickScanで24時間に近い順（経過時間が長い順）にソート済み

  // 全件を分割送信
  var notifyRoom = RC_NOTIFY_ROOM_ID;
  if (!notifyRoom) {
    Logger.log('通知先が見つかりません');
    return;
  }

  var total = quick.unreplied.length;
  var now = new Date();
  var timeStr = Utilities.formatDate(now, 'Asia/Tokyo', 'M/d HH:mm');
  var header = '[info][title]⚠ 返信漏れ検知 ' + timeStr + '（' + total + '件）[/title]';
  var body = header;
  var msgNum = 1;

  for (var i = 0; i < total; i++) {
    var u = quick.unreplied[i];
    var elapsedSec = Math.floor(Date.now() / 1000 - u.lastUpdate);
    var elapsedH = elapsedSec / 3600;
    var icon = urgencyIcon_(elapsedH);
    var elapsed = formatElapsed_(elapsedSec);
    var line = (icon ? icon + ' ' : '') + '▸ ' + u.studentName + '（' + elapsed + '前）\n';
    line += '  → https://www.chatwork.com/#!rid' + u.roomId + '\n\n';

    // メッセージが長くなりすぎたら分割送信
    if (body.length + line.length > RC_MSG_LIMIT) {
      body += '[/info]';
      rcFetch_('post', '/rooms/' + notifyRoom + '/messages', { body: body });
      Utilities.sleep(RC_SLEEP_MS);
      msgNum++;
      body = '[info][title]⚠ 返信漏れ検知（続き ' + msgNum + '）[/title]';
    }
    body += line;
  }

  body += '\n💡 自動チェック（2時間おき）[/info]';
  rcFetch_('post', '/rooms/' + notifyRoom + '/messages', { body: body });
  Logger.log('嬴政チャットに通知送信完了（' + msgNum + '通・' + total + '件）');
}

// ===== ディープスキャン＆通知（既読未返信も検出） =====
function deepCheckAndNotify() {
  var result = deepScanUnreplied();

  if (result.unreplied.length === 0) {
    Logger.log('ディープスキャン: 未返信なし（' + result.scanned + '/' + result.total + '室スキャン済）');
    return;
  }

  var notifyRoom = RC_NOTIFY_ROOM_ID;
  var total = result.unreplied.length;
  var now = new Date();
  var timeStr = Utilities.formatDate(now, 'Asia/Tokyo', 'M/d HH:mm');
  var body = '[info][title]🔍 返信漏れ詳細チェック ' + timeStr + '（' + total + '件）[/title]';
  body += '※ ' + result.scanned + '/' + result.total + '室をスキャン\n\n';
  var msgNum = 1;

  for (var i = 0; i < total; i++) {
    var u = result.unreplied[i];
    var icon = urgencyIcon_(u.hoursAgo);
    var line = (icon ? icon + ' ' : '') + '▸ ' + u.senderName + '（' + u.hoursAgo + '時間前）\n';
    line += '  → https://www.chatwork.com/#!rid' + u.roomId + '\n\n';

    if (body.length + line.length > RC_MSG_LIMIT) {
      body += '[/info]';
      rcFetch_('post', '/rooms/' + notifyRoom + '/messages', { body: body });
      Utilities.sleep(RC_SLEEP_MS);
      msgNum++;
      body = '[info][title]🔍 返信漏れ（続き ' + msgNum + '）[/title]';
    }
    body += line;
  }

  body += '💡 ' + RC_STALE_HOURS + '時間以上返信のないチャットを表示[/info]';
  rcFetch_('post', '/rooms/' + notifyRoom + '/messages', { body: body });
}

// ===== ユーティリティ =====

// ルーム名から受講生名を抽出（オーナー名以外の場合のみ有効）
function extractStudentName_(roomName) {
  if (!roomName) return '';
  var cleaned = roomName.replace(/✅/g, '').replace(/1:1チャ.*/g, '').replace(/[【】]/g, '').trim();
  // オーナー名と同じ or 空なら取得不可
  if (!cleaned || cleaned === RC_OWNER_NAME) return '';
  return cleaned;
}

// 未返信リストの受講生名を解決（ルーム名→メンバー取得フォールバック）
function resolveStudentNames_(unrepliedList, myId) {
  // キャッシュからメンバー情報を読み込み
  var cache = CacheService.getScriptCache();
  var cacheKey = 'RC_MEMBER_CACHE';
  var memberCache = {};
  try {
    var cached = cache.get(cacheKey);
    if (cached) memberCache = JSON.parse(cached);
  } catch (e) {}

  var needsFetch = [];
  for (var i = 0; i < unrepliedList.length; i++) {
    var u = unrepliedList[i];
    // まずルーム名から試す
    var nameFromRoom = extractStudentName_(u.roomName);
    if (nameFromRoom) {
      u.studentName = nameFromRoom;
    } else if (memberCache[u.roomId]) {
      // キャッシュにある
      u.studentName = memberCache[u.roomId];
    } else {
      // API取得が必要
      needsFetch.push(u);
    }
  }

  // メンバーAPIで名前取得（必要なルームのみ）
  for (var j = 0; j < needsFetch.length; j++) {
    var u = needsFetch[j];
    try {
      Utilities.sleep(RC_SLEEP_MS);
      var members = rcFetch_('get', '/rooms/' + u.roomId + '/members');
      var name = '不明';
      for (var k = 0; k < members.length; k++) {
        if (members[k].account_id !== myId) {
          name = members[k].name;
          break;
        }
      }
      u.studentName = name;
      memberCache[u.roomId] = name;
    } catch (e) {
      u.studentName = 'ルーム' + u.roomId;
      if (e.message === 'RC_RATE_LIMITED') break;
    }
  }

  // キャッシュ保存（6時間）
  try {
    cache.put(cacheKey, JSON.stringify(memberCache), 21600);
  } catch (e) {}
}

function formatElapsed_(seconds) {
  if (seconds < 3600) return Math.floor(seconds / 60) + '分';
  if (seconds < 86400) return Math.floor(seconds / 3600) + '時間';
  return Math.floor(seconds / 86400) + '日';
}

// 経過時間に応じた緊急度アイコン
function urgencyIcon_(hours) {
  if (hours >= 24) return '☠️☠️';
  if (hours >= 20) return '⚠️⚠️';
  if (hours >= 15) return '⚠️';
  return '';
}


// ===== トリガー管理 =====
// クイックチェックを2時間おきに実行
function setupReplyCheckTrigger() {
  removeReplyCheckTrigger();
  ScriptApp.newTrigger('checkAndNotifyUnreplied')
    .timeBased()
    .everyHours(2)
    .create();
  Logger.log('返信チェックトリガーを設定（2時間おき）');
}

// ディープスキャンを1日2回（9時・18時）実行
function setupDeepCheckTrigger() {
  removeDeepCheckTrigger();
  ScriptApp.newTrigger('deepCheckAndNotify')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();
  ScriptApp.newTrigger('deepCheckAndNotify')
    .timeBased()
    .atHour(18)
    .everyDays(1)
    .create();
  Logger.log('ディープチェックトリガーを設定（9時・18時）');
}

function removeReplyCheckTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkAndNotifyUnreplied') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function removeDeepCheckTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'deepCheckAndNotify') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// ===== テスト用 =====
function testQuickScan() {
  var result = quickScanUnreplied();
  Logger.log('未読1:1チャット: ' + result.unreplied.length + '件');
  result.unreplied.forEach(function(u) {
    Logger.log('  ' + u.roomName + ' (未読' + u.unreadCount + ') rid:' + u.roomId);
  });
}

function testDeepScan() {
  var result = deepScanUnreplied();
  Logger.log(result.unreplied.length + '件の未返信（' + result.scanned + '/' + result.total + '室）');
  result.unreplied.forEach(function(u) {
    Logger.log('  ' + u.senderName + '（' + u.hoursAgo + 'h前）: ' + u.preview);
  });
}
