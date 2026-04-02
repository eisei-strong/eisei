/**
 * かのんさんボトルネック分析レポートをGoogleドキュメントで生成
 */
function generateKanonReport() {
  var doc = DocumentApp.create('【分析レポート】かのんさん ボトルネック分析 & 改善プラン（2026年3月8日）');
  var body = doc.getBody();

  // スタイル定義
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

  var redStyle = {};
  redStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#d93025';
  redStyle[DocumentApp.Attribute.BOLD] = true;

  var greenStyle = {};
  greenStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#0d652d';
  greenStyle[DocumentApp.Attribute.BOLD] = true;

  var blueStyle = {};
  blueStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#1a73e8';
  blueStyle[DocumentApp.Attribute.BOLD] = true;

  // ===== タイトル =====
  var title = body.appendParagraph('かのんさん ボトルネック分析 & 改善プラン');
  title.setHeading(DocumentApp.ParagraphHeading.TITLE);
  title.setAttributes(titleStyle);

  body.appendParagraph('作成日: 2026年3月8日').setAttributes(normalStyle);
  body.appendParagraph('目的: 着金最大化に向けたボトルネック特定と改善施策の優先順位付け').setAttributes(normalStyle);

  body.appendHorizontalRule();

  // ===== 1. 現状ファネル整理 =====
  var s1 = body.appendParagraph('1. 現状ファネル整理（2月実績）');
  s1.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  s1.setAttributes(h1Style);

  // SNS発信
  var s1sub1 = body.appendParagraph('SNS発信');
  s1sub1.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  s1sub1.setAttributes(h2Style);

  var snsTable = body.appendTable([
    ['プラットフォーム', '投稿本数', '合計再生数', '平均再生数/本'],
    ['TikTok (@kanon_dress_diycafe)', '5本', '23,597回', '4,719回'],
    ['Instagram (@kanomade)', '3本', '11,853回', '3,951回'],
    ['合計', '8本', '35,450回', '4,431回']
  ]);
  styleTable_(snsTable);

  // ファネル数値
  var s1sub2 = body.appendParagraph('セールスファネル');
  s1sub2.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  s1sub2.setAttributes(h2Style);

  var funnelTable = body.appendTable([
    ['ステージ', '人数', '転換率', 'ステータス'],
    ['SNS合計再生', '35,450回', '-', '-'],
    ['LINE登録', '※未計測', '※未計測', '⚠️ 要計測'],
    ['説明会希望', '24人', '※LINE登録数が不明', '⚠️ 要計測'],
    ['セールス実施', '17人', '70.8%（対希望）', '🔴 要改善'],
    ['成約（着金）', '※未計測', '※未計測', '⚠️ 要計測']
  ]);
  styleTable_(funnelTable);

  body.appendParagraph('');
  var funnelNote = body.appendParagraph('▶ 全体CVR（SNS再生→希望）: 約0.07%（35,450再生 → 24人希望）');
  funnelNote.setAttributes(normalStyle);

  body.appendHorizontalRule();

  // ===== 2. ボトルネック特定 =====
  var s2 = body.appendParagraph('2. ボトルネック特定（インパクト順）');
  s2.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  s2.setAttributes(h1Style);

  // ボトルネック①
  var bn1 = body.appendParagraph('🔴 ボトルネック①: 希望→実施の脱落率 29%（最優先）');
  bn1.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  bn1.setAttributes(h2Style);

  var bn1desc = body.appendParagraph('24人が希望を出したにもかかわらず、実施に至ったのは17人。7人（29%）が脱落している。');
  bn1desc.setAttributes(normalStyle);

  var bn1imp = body.appendParagraph('これが最大のボトルネック。「やりたい」と意思表示した温度の高い見込み客の3割を失っている。仮に単価30万円なら、月210万円の機会損失に相当。');
  bn1imp.setAttributes(normalStyle);

  body.appendParagraph('原因仮説:').setAttributes(boldStyle);
  var bn1causes = body.appendListItem('希望〜実施までの間隔が長く、熱が冷めている');
  bn1causes.setGlyphType(DocumentApp.GlyphType.BULLET);
  bn1causes.setAttributes(normalStyle);
  body.appendListItem('LINE上でのリマインド・日程調整が不十分').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);
  body.appendListItem('セールスチームの初動対応速度が遅い').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);
  body.appendListItem('事前学習動画で「自分には合わない」と判断される（逆選別の可能性）').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);
  body.appendListItem('日程調整のやり取りが多く面倒になっている').setGlyphType(DocumentApp.GlyphType.BULLET).setAttributes(normalStyle);

  // ボトルネック②
  var bn2 = body.appendParagraph('🟡 ボトルネック②: 投稿量の絶対的な不足（月8本）');
  bn2.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  bn2.setAttributes(h2Style);

  var bn2desc = body.appendParagraph('TikTok月5本・Instagram月3本の計8本は、業界標準（TikTok週3〜5本、IG週2本）と比較して大幅に少ない。市場にリーチする「打席数」が根本的に不足している。');
  bn2desc.setAttributes(normalStyle);

  // ボトルネック③
  var bn3 = body.appendParagraph('🟡 ボトルネック③: 成約率・着金額が未計測');
  bn3.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  bn3.setAttributes(h2Style);

  var bn3desc = body.appendParagraph('17人実施して何件・いくら着金したかが見えない状態。計測なくして改善なし。特にセールストークの質・クロージング力の評価ができない。');
  bn3desc.setAttributes(normalStyle);

  // ボトルネック④
  var bn4 = body.appendParagraph('🟡 ボトルネック④: LINE登録数が不明（ファネル上流の分断）');
  bn4.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  bn4.setAttributes(h2Style);

  var bn4desc = body.appendParagraph('35,450再生に対して24人の希望という数値はあるが、途中の「LINE登録数」が不明。SNSが弱いのか、LINE内の教育コンテンツが弱いのかの切り分けができていない。');
  bn4desc.setAttributes(normalStyle);

  body.appendHorizontalRule();

  // ===== 3. 改善プラン =====
  var s3 = body.appendParagraph('3. 改善プラン（優先度順）');
  s3.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  s3.setAttributes(h1Style);

  // P1
  var p1 = body.appendParagraph('🔴 P1（今すぐ・最大インパクト）: 希望→実施の歩留まり改善');
  p1.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  p1.setAttributes(h2Style);

  var p1table = body.appendTable([
    ['施策', '具体アクション', '期待効果'],
    ['即レス体制の構築', '希望が出た瞬間から24時間以内に日程確定。48時間超えると成約率が半減するデータあり', '脱落率の大幅削減'],
    ['リマインド自動化', 'LINE公式で前日・当日朝にリマインドメッセージ配信（L Messageのステップ配信活用）', 'ドタキャン防止'],
    ['日程候補の即提示', '「いつがいいですか？」ではなく「この3つからどれが良いですか？」で選択式に', '日程確定率UP'],
    ['特別感の演出', '「〇〇さん専用の時間を確保しました」と伝える', 'キャンセル心理的障壁UP'],
    ['脱落者への再アプローチ', '未実施7人に再度連絡。理由ヒアリング→再日程調整', '即効性あり']
  ]);
  styleTable_(p1table);

  var p1effect = body.appendParagraph('期待効果: 実施率70%→85%で月+3〜4人のセールス機会増');
  p1effect.setAttributes(normalStyle);
  p1effect.editAsText().setAttributes(0, p1effect.getText().length - 1, greenStyle);

  // P2
  var p2 = body.appendParagraph('🟡 P2（1〜2週間以内）: KPI計測基盤の整備');
  p2.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  p2.setAttributes(h2Style);

  var p2table = body.appendTable([
    ['計測項目', '具体アクション'],
    ['KPIシート作成', 'LINE登録数・希望数・実施数・成約数・着金額を週次で記録'],
    ['成約率の可視化', '実施者のうち何人が成約したか、単価はいくらかを必ず記録'],
    ['セールスチーム個人別成績', '誰がどの案件を担当し、結果はどうだったかを追跡'],
    ['LINE登録数の取得', 'L Messageまたは LINE公式のダッシュボードから毎月記録']
  ]);
  styleTable_(p2table);

  // P3
  var p3 = body.appendParagraph('🟢 P3（2〜4週間）: 投稿量の増加');
  p3.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  p3.setAttributes(h2Style);

  var p3table = body.appendTable([
    ['施策', '現状', '目標'],
    ['TikTok投稿数', '月5本', '月15本（週3〜4本）'],
    ['Instagram投稿数', '月3本', '月8本（週2本）'],
    ['コンテンツ再利用', '-', 'TikTok→IGリール転用で効率化'],
    ['バズ動画の横展開', '-', '高パフォーマンス動画の切り口を再利用']
  ]);
  styleTable_(p3table);

  var p3effect = body.appendParagraph('期待効果: 投稿2倍→再生数・LINE登録の増加');
  p3effect.setAttributes(normalStyle);

  // P4
  var p4 = body.appendParagraph('🔵 P4（並行して検証）: LINE導線の最適化');
  p4.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  p4.setAttributes(h2Style);

  var p4table = body.appendTable([
    ['施策', '具体アクション'],
    ['事前学習動画の見直し', '動画が「教育」になりすぎて満足→離脱を起こしていないか確認'],
    ['LINE内CTA強化', '事前学習後に「次のステップはこちら」と明確に説明会誘導'],
    ['セグメント配信', '動画視聴完了者と未完了者で異なるメッセージを送る']
  ]);
  styleTable_(p4table);

  body.appendHorizontalRule();

  // ===== 4. 実行ロードマップ =====
  var s4 = body.appendParagraph('4. 実行ロードマップ');
  s4.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  s4.setAttributes(h1Style);

  var roadmapTable = body.appendTable([
    ['フェーズ', '時期', '施策', '担当'],
    ['短期', '今週〜', '① 希望→実施の歩留まり改善（即レス・リマインド・日程提示）', 'セールスチーム'],
    ['短期', '今週〜', '② 未実施7人への再アプローチ', 'セールスチーム'],
    ['中期', '2〜4週間', '③ KPI計測基盤の整備', 'かのんさん＋事務局'],
    ['中期', '2〜4週間', '④ 投稿量2倍化（TikTok月15本、IG月8本）', 'かのんさん'],
    ['中長期', '1〜2ヶ月', '⑤ LINE導線の教育→セールス誘導の最適化', 'セールスチーム'],
    ['中長期', '1〜2ヶ月', '⑥ セールストーク・スライドの改善（録画分析後）', '顧問＋セールスチーム']
  ]);
  styleTable_(roadmapTable);
  // ヘッダー行の色を変える
  var roadmapHeader = roadmapTable.getRow(0);
  for (var i = 0; i < roadmapHeader.getNumCells(); i++) {
    roadmapHeader.getCell(i).setBackgroundColor('#1a73e8');
    roadmapHeader.getCell(i).editAsText().setForegroundColor('#ffffff').setBold(true);
  }

  body.appendHorizontalRule();

  // ===== 5. 追加で必要な情報 =====
  var s5 = body.appendParagraph('5. 精度向上のために追加で必要な情報');
  s5.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  s5.setAttributes(h1Style);

  body.appendListItem('2月の成約数と着金額（セールス成約率の算出に必須）').setGlyphType(DocumentApp.GlyphType.NUMBER).setAttributes(normalStyle);
  body.appendListItem('LINE登録者数（ファネル上流のCVR算出に必須）').setGlyphType(DocumentApp.GlyphType.NUMBER).setAttributes(normalStyle);
  body.appendListItem('セールス録画の閲覧（トーク内のボトルネック特定）').setGlyphType(DocumentApp.GlyphType.NUMBER).setAttributes(normalStyle);
  body.appendListItem('事前学習動画の視聴完了率（LINE内での離脱ポイント特定）').setGlyphType(DocumentApp.GlyphType.NUMBER).setAttributes(normalStyle);
  body.appendListItem('セールスチーム各メンバーの個人別成績（属人性の確認）').setGlyphType(DocumentApp.GlyphType.NUMBER).setAttributes(normalStyle);

  body.appendParagraph('');
  var footer = body.appendParagraph('これらの情報が揃えば、さらに精度の高い改善プランに落とし込むことが可能です。');
  footer.setAttributes(normalStyle);
  footer.editAsText().setBold(true);

  // ドキュメントを保存
  doc.saveAndClose();

  var url = doc.getUrl();
  Logger.log('レポートURL: ' + url);
  return url;
}

/**
 * テーブルのスタイルを整える
 */
function styleTable_(table) {
  // ヘッダー行
  var header = table.getRow(0);
  for (var i = 0; i < header.getNumCells(); i++) {
    header.getCell(i).setBackgroundColor('#f1f3f4');
    header.getCell(i).editAsText().setBold(true).setFontSize(10);
  }
  // データ行
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
