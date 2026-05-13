/**
 * 営業切り返しマスター(AI) - 学習提出ログ書き込み GAS Web App
 *
 * デプロイ:
 *   1. clasp push
 *   2. デプロイ → ウェブアプリ → 「自分」として実行 / 「全員」がアクセス可能 → デプロイ
 *   3. 表示された URL を Xserver の /home/kodaidai/.kirigaeshi_gas_url に保存
 *
 * 呼び出し元: ai_api.php (category === 'ai_kirigaeshi' のとき自動でPOST)
 *
 * 期待される POST body (JSON):
 *   {
 *     userName: "受講生名",
 *     scriptNo: "No.1",
 *     questionCategory: "料金・コスト",
 *     submissionText: "新人が書いた回答テキスト or Loom transcript",
 *     loomUrl: "https://www.loom.com/share/xxx" (任意),
 *     aiGrade: "S" | "A" | "B" | "C",
 *     aiFeedback: "Claude が返した採点コメント全文",
 *     missingPoints: ["..."],  // 任意
 *     ngPoints: ["..."]        // 任意
 *   }
 */

const SPREADSHEET_ID = '1MqoJiyn8syUQP3SCvdFogqJ_iyLCxIfZl8oEfJwpMk0';
const SHEET_NAME = '学習提出ログ';

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({success: false, error: 'no body'});
    }
    const data = JSON.parse(e.postData.contents);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      return jsonResponse({success: false, error: 'sheet not found: ' + SHEET_NAME});
    }

    sheet.appendRow([
      new Date().toISOString(),
      String(data.userName || '匿名'),
      String(data.scriptNo || ''),
      String(data.questionCategory || ''),
      String(data.submissionText || ''),
      String(data.loomUrl || ''),
      String(data.aiGrade || ''),
      String(data.aiFeedback || ''),
      (data.missingPoints || []).join('\n'),
      (data.ngPoints || []).join('\n'),
    ]);

    return jsonResponse({success: true});
  } catch (err) {
    return jsonResponse({success: false, error: err.toString()});
  }
}

function doGet(e) {
  return jsonResponse({status: 'ok', message: 'KirigaeshiSubmit Web App is alive'});
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
