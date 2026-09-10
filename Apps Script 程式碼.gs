/**
 * 痛點收集表單 — 接收端
 * AI 資訊素養與應用（通識3）W02
 *
 * ── 部署步驟（大約 3 分鐘）──────────────────────────────
 * 1. 到 https://sheets.new 建立一個新的 Google 試算表，命名為「W02 痛點收集」
 * 2. 上方選單「擴充功能 → Apps Script」
 * 3. 把預設的 myFunction 全部刪掉，貼上這整份程式碼，存檔
 * 4. 右上角「部署 → 新增部署作業」
 *      類型：網頁應用程式
 *      執行身分：我
 *      具有存取權的使用者：★ 任何人 ★（這個一定要選，否則學生送不出來）
 * 5. 按「部署」，第一次會要求授權，一路允許
 * 6. 複製產生的網址（結尾是 /exec），貼進 index.html 最上面的 ENDPOINT
 * ────────────────────────────────────────────────────────
 *
 * 測試：部署後在瀏覽器直接打開那個 /exec 網址，
 *       應該會看到「痛點收集表單接收端運作中」，代表部署成功。
 */

const SHEET_NAME = '痛點收集';

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['送出時間', '組別', '填寫人', '痛點項目', '最煩的一項', '項目數']);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
      sheet.setColumnWidth(4, 420);
      sheet.setColumnWidth(5, 260);
    }

    const d = JSON.parse(e.postData.contents);
    const items = (d.items || [])
      .map(function (x) { return String(x).trim(); })
      .filter(function (x) { return x.length > 0; });

    sheet.appendRow([
      new Date(),
      d.group || '',
      d.name || '',
      items.map(function (v, k) { return (k + 1) + '. ' + v; }).join('\n'),
      d.worst || '',
      items.length
    ]);
    sheet.getRange(sheet.getLastRow(), 4).setWrap(true);

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/** 直接用瀏覽器打開網址時顯示，用來確認部署成功 */
function doGet() {
  return ContentService
    .createTextOutput('痛點收集表單接收端運作中 ✅')
    .setMimeType(ContentService.MimeType.TEXT);
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
