/**
 * かのんさん MTG用 課題点ドキュメント生成（2026年3月10日）
 * GASスクリプトエディタで実行 → Googleドキュメントが自動生成される
 */
function generateKanonMTGDoc() {
  var doc = DocumentApp.create('【MTG資料】かのんさん 課題点まとめ（2026年3月10日）');
  var body = doc.getBody();

  // ===== スタイル定義 =====
  var titleStyle = {};
  titleStyle[DocumentApp.Attribute.FONT_SIZE] = 20;
  titleStyle[DocumentApp.Attribute.BOLD] = true;
  titleStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#1a73e8';

  var h1Style = {};
  h1Style[DocumentApp.Attribute.FONT_SIZE] = 16;
  h1Style[DocumentApp.Attribute.BOLD] = true;
  h1Style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#202124';

  var h2Style = {};
  h2Style[DocumentApp.Attribute.FONT_SIZE] = 13;
  h2Style[DocumentApp.Attribute.BOLD] = true;
  h2Style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#333333';

  var normalStyle = {};
  normalStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  normalStyle[DocumentApp.Attribute.BOLD] = false;
  normalStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#333333';
  normalStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';

  var boldStyle = {};
  boldStyle[DocumentApp.Attribute.BOLD] = true;

  var redBold = {};
  redBold[DocumentApp.Attribute.FOREGROUND_COLOR] = '#d93025';
  redBold[DocumentApp.Attribute.BOLD] = true;

  var greenBold = {};
  greenBold[DocumentApp.Attribute.FOREGROUND_COLOR] = '#0d652d';
  greenBold[DocumentApp.Attribute.BOLD] = true;

  // ===== タイトル =====
  var title = body.appendParagraph('かのんさん 課題点まとめ');
  title.setHeading(DocumentApp.ParagraphHeading.TITLE);
  title.setAttributes(titleStyle);

  body.appendParagraph('MTG日: 2026年3月10日').setAttributes(normalStyle);
  body.appendParagraph('目的: 現状の課題を整理し、優先的に取り組むべきポイントを議論する').setAttributes(normalStyle);

  body.appendHorizontalRule();

  // ===== 2月実績サマリー =====
  var s0 = body.appendParagraph('2月実績サマリー');
  s0.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  s0.setAttributes(h1Style);

  var summaryTable = body.appendTable([
    ['指標', '数値', 'メモ'],
    ['売上（着金）', '¥3,412,500', ''],
    ['リスト数（LINE登録）', '110人', '広告+オーガニック'],
    ['説明会希望', '24人', 'LINE登録の21.8%'],
    ['セールス実施', '17人', '希望の70.8%（7人脱落）'],
    ['TikTok', '5本 / 23,597再生', '@kanon_dress_diycafe'],
    ['Instagram', '3本 / 11,853再生', '@kanomade'],
    ['SNS合計', '8本 / 35,450再生', ''],
    ['セールス体制', '5人', '振り分け制'],
    ['広告', 'IG / FB運用中', '費用・CPA未確認']
  ]);
  styleTableMTG_(summaryTable);

  body.appendParagraph('');
  var kpi = body.appendParagraph('▶ 売上/セールス件数 = 約¥200,735/件');
  kpi.setAttributes(normalStyle);
  kpi.editAsText().setBold(true);

  body.appendHorizontalRule();

  // ===== ファネル図 =====
  var sf = body.appendParagraph('ファネル全体像');
  sf.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  sf.setAttributes(h1Style);

  body.appendParagraph('SNS再生 35,450回').setAttributes(normalStyle);
  body.appendParagraph('　　↓ CVR 0.31%（※広告込み）').setAttributes(normalStyle);
  body.appendParagraph('LINE登録 110人').setAttributes(normalStyle);
  body.appendParagraph('　　↓ CVR 21.8%').setAttributes(normalStyle);
  body.appendParagraph('説明会希望 24人').setAttributes(normalStyle);

  var dropP = body.appendParagraph('　　↓ CVR 70.8%（★ 7人脱落 = 最大課題）');
  dropP.setAttributes(normalStyle);
  dropP.editAsText().setAttributes(0, dropP.getText().length - 1, redBold);

  body.appendParagraph('セールス実施 17人').setAttributes(normalStyle);
  body.appendParagraph('　　↓').setAttributes(normalStyle);

  var revP = body.appendParagraph('着金 ¥3,412,500');
  revP.setAttributes(normalStyle);
  revP.editAsText().setBold(true);

  body.appendHorizontalRule();

  // ===== 課題点 =====
  var s1 = body.appendParagraph('課題点（優先度順）');
  s1.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  s1.setAttributes(h1Style);

  // 課題①
  var i1 = body.appendParagraph('課題① 希望→実施の脱落率 29%');
  i1.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  i1.setAttributes(h2Style);
  i1.editAsText().setAttributes(0, i1.getText().length - 1, redBold);

  body.appendParagraph('24人が「やりたい」と言ったのに、17人しか実施できていない。7人が途中で消えている。').setAttributes(normalStyle);

  body.appendParagraph('');
  var loss1 = body.appendParagraph('機会損失: ¥200,735 × 7人 ≒ ¥1,405,000/月');
  loss1.setAttributes(normalStyle);
  loss1.editAsText().setAttributes(0, loss1.getText().length - 1, redBold);

  body.appendParagraph('');
  body.appendParagraph('考えられる原因:').setAttributes(boldStyle);
  body.appendListItem('希望を出してから実施まで時間が空きすぎ（熱が冷める）').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);
  body.appendListItem('日程調整のやり取りが面倒で離脱').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);
  body.appendListItem('LINEでのリマインド不足').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);
  body.appendListItem('事前学習動画で「自分には合わない」と判断').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);
  body.appendListItem('セールス担当5人の対応スピード・品質のバラつき').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);

  body.appendParagraph('');
  body.appendParagraph('【MTGで確認したいこと】').setAttributes(boldStyle);
  body.appendListItem('7人が脱落した具体的な理由は把握しているか？').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);
  body.appendListItem('希望〜実施までの平均日数は？').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);
  body.appendListItem('セールス担当ごとの対応状況は？').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);

  // 課題②
  body.appendParagraph('');
  var i2 = body.appendParagraph('課題② セールスチーム5人の成績が見えない');
  i2.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  i2.setAttributes(h2Style);

  body.appendParagraph('5人で振り分けているが、誰が何件担当して何件成約したかが不明。録画もまだ全部シートに貼れていない。').setAttributes(normalStyle);

  body.appendParagraph('');
  body.appendParagraph('【MTGで確認したいこと】').setAttributes(boldStyle);
  body.appendListItem('5人それぞれの担当件数・成約数・着金額は？').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);
  body.appendListItem('録画のシート貼り付けはいつ完了するか？').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);
  body.appendListItem('振り分けの基準は？（ランダム / 順番 / 相性）').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);

  // 課題③
  body.appendParagraph('');
  var i3 = body.appendParagraph('課題③ 投稿量の不足（月8本）');
  i3.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  i3.setAttributes(h2Style);

  var postTable = body.appendTable([
    ['', '現状', '業界水準'],
    ['TikTok', '月5本', '月12〜20本（週3〜5本）'],
    ['Instagram', '月3本', '月8本（週2本）'],
    ['合計', '月8本', '月20〜28本']
  ]);
  styleTableMTG_(postTable);

  body.appendParagraph('');
  body.appendParagraph('再生数自体は悪くない（平均4,431回/本）。量を増やせばリーチは比例して伸びる可能性が高い。').setAttributes(normalStyle);

  body.appendParagraph('');
  body.appendParagraph('【MTGで確認したいこと】').setAttributes(boldStyle);
  body.appendListItem('投稿が少ない理由は？（時間不足 / ネタ不足 / 撮影環境）').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);
  body.appendListItem('投稿頻度を上げるために必要なサポートは？').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);

  // 課題④
  body.appendParagraph('');
  var i4 = body.appendParagraph('課題④ LINE教育→説明会希望の転換率（21.8%）');
  i4.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  i4.setAttributes(h2Style);

  body.appendParagraph('110人のLINE登録者のうち、説明会を希望したのは24人（21.8%）。残り86人はLINE内で止まっている。').setAttributes(normalStyle);

  body.appendParagraph('');
  body.appendParagraph('【MTGで確認したいこと】').setAttributes(boldStyle);
  body.appendListItem('事前学習動画の視聴完了率はどのくらいか？').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);
  body.appendListItem('動画から説明会への誘導の仕方はどうなっているか？').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);
  body.appendListItem('LINE内のメッセージ開封率は？').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);

  // 課題⑤
  body.appendParagraph('');
  var i5 = body.appendParagraph('課題⑤ 広告の費用対効果が不明');
  i5.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  i5.setAttributes(h2Style);

  body.appendParagraph('Instagram / Facebook広告を運用中だが、以下が見えない:').setAttributes(normalStyle);
  body.appendListItem('月間の広告費はいくらか？').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);
  body.appendListItem('広告経由のLINE登録は何人か？（110人中）').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);
  body.appendListItem('CPA（1リスト獲得にいくらかかっているか）').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);
  body.appendListItem('ROAS（広告費に対してどのくらいの売上が出ているか）').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);

  body.appendHorizontalRule();

  // ===== MTGアジェンダ案 =====
  var s2 = body.appendParagraph('MTGアジェンダ案');
  s2.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  s2.setAttributes(h1Style);

  var agendaTable = body.appendTable([
    ['#', '議題', '目的', '時間目安'],
    ['1', '2月実績の振り返り', '数値の共有認識を作る', '5分'],
    ['2', '希望→実施の脱落7人について', '原因特定 → 即アクション決定', '15分'],
    ['3', 'セールスチーム5人の状況確認', '個人別成績 / 録画共有の進捗', '10分'],
    ['4', 'LINE導線の確認', '事前学習動画→説明会の流れ', '10分'],
    ['5', '広告の費用対効果', '数値確認 → 継続/調整の判断', '5分'],
    ['6', '3月のアクションプラン決定', '誰が何をいつまでにやるか', '15分']
  ]);
  styleTableMTG_(agendaTable);
  // ヘッダー行の色
  var agendaHeader = agendaTable.getRow(0);
  for (var i = 0; i < agendaHeader.getNumCells(); i++) {
    agendaHeader.getCell(i).setBackgroundColor('#1a73e8');
    agendaHeader.getCell(i).editAsText().setForegroundColor('#ffffff').setBold(true);
  }

  body.appendHorizontalRule();

  // ===== 参照資料保管庫 =====
  var s3 = body.appendParagraph('参照資料保管庫');
  s3.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  s3.setAttributes(h1Style);

  body.appendParagraph('【営業資料】').setAttributes(boldStyle);
  body.appendListItem('サロン説明会スライド（Canva）\nhttps://www.canva.com/design/DAGkhsDLMBQ/5_HCTUfxQyEN4w9kVMS04w/edit').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);
  body.appendListItem('セールスチーム共有シート\nhttps://docs.google.com/spreadsheets/d/1-QjFKOeNXdy4nLZlX6DjX6ECPWeKq6keZZmg_ej2JpE/edit').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);
  body.appendListItem('⭕️レコーディングシート（セールス録画）\nhttps://docs.google.com/spreadsheets/d/1-QjFKOeNXdy4nLZlX6DjX6ECPWeKq6keZZmg_ej2JpE/edit?gid=378649631#gid=378649631').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);

  body.appendParagraph('').setAttributes(normalStyle);
  body.appendParagraph('【SNS】').setAttributes(boldStyle);
  body.appendListItem('TikTok: @kanon_dress_diycafe\nhttps://www.tiktok.com/@kanon_dress_diycafe').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);
  body.appendListItem('Instagram: @kanomade\nhttps://www.instagram.com/kanomade/').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);

  body.appendParagraph('').setAttributes(normalStyle);
  body.appendParagraph('【LINE関連】').setAttributes(boldStyle);
  body.appendListItem('L Message管理画面\nhttps://step.lme.jp/').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);
  body.appendListItem('LINE公式 事務局招待リンク\nhttps://manager.line.biz/invitation/SNrfC2dNfJXmctLLsumz8FN6XRlNLe').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);

  body.appendParagraph('').setAttributes(normalStyle);
  body.appendParagraph('【講座オープンチャット】').setAttributes(boldStyle);
  body.appendListItem('【1月】コスプレ衣装作家特別講座\nhttps://line.me/ti/g2/aDeV8OzEJ8UU-Jzu-9kUKqQm0JqX1FHqei9T0A').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);

  // ドキュメントを保存
  doc.saveAndClose();

  var url = doc.getUrl();
  Logger.log('MTGドキュメントURL: ' + url);
  return url;
}

/**
 * テーブルスタイル（MTG用）
 */
function styleTableMTG_(table) {
  var header = table.getRow(0);
  for (var i = 0; i < header.getNumCells(); i++) {
    header.getCell(i).setBackgroundColor('#f1f3f4');
    header.getCell(i).editAsText().setBold(true).setFontSize(10);
  }
  for (var r = 1; r < table.getNumRows(); r++) {
    var row = table.getRow(r);
    for (var c = 0; c < row.getNumCells(); c++) {
      row.getCell(c).editAsText().setFontSize(10).setBold(false);
      if (r % 2 === 0) {
        row.getCell(c).setBackgroundColor('#fafafa');
      }
    }
  }
}

/**
 * レコーディングシートのB列（担当者）を取得
 */
function getKanonSalesMembers() {
  var ss = SpreadsheetApp.openById('1-QjFKOeNXdy4nLZlX6DjX6ECPWeKq6keZZmg_ej2JpE');
  var sheets = ss.getSheets();
  var sheet = null;
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === 378649631) {
      sheet = sheets[i];
      break;
    }
  }
  if (!sheet) return {error: 'シートが見つかりません', sheetNames: sheets.map(function(s){return s.getName() + '(gid:' + s.getSheetId() + ')';})};
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(1, 1, lastRow, 3).getValues();
  var members = {};
  var rows = [];
  for (var r = 0; r < data.length; r++) {
    var name = String(data[r][1]).trim();
    rows.push({row: r+1, A: String(data[r][0]).trim(), B: name, C: String(data[r][2]).trim()});
    if (name && name !== '' && name !== '担当者' && name !== 'undefined' && name !== 'B') {
      members[name] = (members[name] || 0) + 1;
    }
  }
  return {sheetName: sheet.getName(), members: members, sampleRows: rows.slice(0, 20)};
}
