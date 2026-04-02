// ===== 万バズ台本 API =====
// Google Docs内の👑マーク付き台本数をカウントして返す

var MANBAZU_DOCS = [
  { id: '10dA-C8iulP7H1tBseKj3qqTsxpDQDHq1Oxo25n8rLAA', name: 'マネタイズ台本' },
  { id: '1vpmfl5-0xzoB41uFxnzMbnTJQypfsS9eIvP1bDdmh5g', name: 'TikTok攻略台本' },
  { id: '1UCB3AFRAQ1AE684xgEAC40qzzHMGBGfJ7cw_pyBrIq4', name: 'IG攻略台本' }
];

function manbazuGetData_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('manbazu_data');
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  var results = [];
  var totalCrown = 0;

  for (var i = 0; i < MANBAZU_DOCS.length; i++) {
    var info = MANBAZU_DOCS[i];
    try {
      var doc = DocumentApp.openById(info.id);
      var text = doc.getBody().getText();
      var crowns = (text.match(/\u{1F451}/gu) || []).length;
      totalCrown += crowns;
      results.push({ name: info.name, count: crowns, ok: true });
    } catch (e) {
      results.push({ name: info.name, count: 0, ok: false });
    }
  }

  var result = {
    docs: results,
    total: totalCrown,
    at: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')
  };

  try {
    var j = JSON.stringify(result);
    if (j.length < 100000) cache.put('manbazu_data', j, 300);
  } catch (e) {}

  return result;
}
