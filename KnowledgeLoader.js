// ============================================
// KnowledgeLoader.js — Drive から Knowledge MD を読み込み
// ============================================
// 質問内容から該当タグを判定し、 該当MDをDriveから取得して Claude API 用に整形

/**
 * 質問から該当する上位タグを判定
 * @param {string} question - ユーザー質問本文
 * @param {number} topN - 取得する上位タグ数（デフォルト 2）
 * @returns {Array<string>} - タグファイルパス（例: 'sales/反論処理.md'）の配列
 */
function detectRelevantTags_(question, topN) {
  topN = topN || 2;
  var scores = {};

  for (var tag in TAG_KEYWORDS) {
    var keywords = TAG_KEYWORDS[tag];
    var score = 0;
    for (var i = 0; i < keywords.length; i++) {
      if (question.indexOf(keywords[i]) !== -1) {
        score++;
      }
    }
    if (score > 0) {
      scores[tag] = score;
    }
  }

  // スコア順でソート、 上位 topN を返す
  var entries = Object.keys(scores).map(function (k) { return [k, scores[k]]; });
  entries.sort(function (a, b) { return b[1] - a[1]; });
  var top = entries.slice(0, topN).map(function (e) { return e[0]; });

  // 1個もマッチしなかったら 「ヒアリング」 「冒頭・フック」 をデフォルトに
  if (top.length === 0) {
    top = ['sales/ヒアリング.md', 'writing/冒頭・フック.md'];
  }

  return top;
}

/**
 * Drive から指定タグの MD を取得して結合
 * @param {Array<string>} tags - タグファイルパス配列
 * @returns {string} - 全MD を結合した文字列
 */
function loadKnowledgeFromDrive_(tags) {
  var folder;
  try {
    folder = DriveApp.getFolderById(SHACHO_FF_DRIVE_FOLDER_ID);
  } catch (e) {
    logShachoBotError_('Drive フォルダ取得失敗: ' + e.message);
    return '';
  }

  var combined = '';
  for (var i = 0; i < tags.length; i++) {
    var tag = tags[i];
    var fileName = tag.split('/').pop(); // 例: sales/反論処理.md → 反論処理.md
    try {
      var files = folder.getFilesByName(fileName);
      if (files.hasNext()) {
        var file = files.next();
        var content = file.getBlob().getDataAsString('UTF-8');
        combined += '\n\n===== ' + tag + ' =====\n' + content;
      } else {
        logShachoBotActivity_('Drive ファイル未発見: ' + fileName);
      }
    } catch (e) {
      logShachoBotError_('Drive 読み込みエラー (' + fileName + '): ' + e.message);
    }
  }

  return combined;
}

/**
 * 質問本文を受けて、 関連 Knowledge MD を返す
 * @param {string} question - ユーザー質問
 * @returns {Object} - { tags: [...], content: '...' }
 */
function getRelevantKnowledge_(question) {
  var tags = detectRelevantTags_(question, 2);
  var content = loadKnowledgeFromDrive_(tags);
  return { tags: tags, content: content };
}
