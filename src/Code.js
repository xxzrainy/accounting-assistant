/**
 * 記帳小幫手 — 掃描國泰世華消費彙整通知，解析後寫入 Google Sheets
 */

// ─── 設定 ────────────────────────────────────────────
var SPREADSHEET_ID = '1WZ51Dw9-r4yzdCHCyRhr-fDbtpegIoI55M1uG_Zw-kg';
var SHEET_NAME = '消費DB';
var HEADERS = ['ID', '日期', '項次', '金額', '類別', '子類別', '付款方式', '備註'];
var GMAIL_QUERY = 'subject:"國泰世華銀行消費彙整通知" -label:已記帳';
var LABEL_NAME = '已記帳';
var MAX_THREADS = 500;

// ─── 主程式 ──────────────────────────────────────────

function main() {
  var threads = GmailApp.search(GMAIL_QUERY, 0, MAX_THREADS);
  if (threads.length === 0) {
    Logger.log('沒有新的消費通知');
    return 0;
  }

  var label = getOrCreateLabel(LABEL_NAME);
  var totalCount = 0;

  for (var i = 0; i < threads.length; i++) {
    var threadRecords = [];
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var records = parseEmail(messages[j]);
      threadRecords = threadRecords.concat(records);
    }
    if (threadRecords.length > 0) {
      threadRecords.sort(function(a, b) { return a.date.localeCompare(b.date); });
      writeToSheet(threadRecords);
      threads[i].addLabel(label);
      totalCount += threadRecords.length;
    }
  }

  if (totalCount === 0) {
    Logger.log('信件已處理但未解析到交易紀錄');
    return 0;
  }

  Logger.log('完成，共寫入 ' + totalCount + ' 筆');
  return totalCount;
}

// ─── 信件解析 ────────────────────────────────────────

/**
 * 解析單封國泰世華消費彙整通知 HTML 信件
 *
 * 每筆消費是一個 <table class="spend-table">，內含 4 列：
 *   列1 (標題): 卡別 | 行動卡號後4碼 | 授權日期 | 授權時間 | 消費地區
 *   列2 (資料): 正卡 | 7755         | 2026/04/01| 11:47   | TW
 *   列3 (標題): 消費金額(colspan=2) | 商店名稱  | 消費類別 | 備註
 *   列4 (資料): NT$500(colspan=2)   | xxx       | 餐飲    | 註一
 */
function parseEmail(message) {
  var html = message.getBody();
  var records = [];

  // 擷取每個 <table class="spend-table"> 區塊
  var tablePattern = /<table[^>]*class="spend-table"[^>]*>([\s\S]*?)<\/table>/gi;
  var tableMatch;

  while ((tableMatch = tablePattern.exec(html)) !== null) {
    var tableHtml = tableMatch[1];

    // 擷取這個 table 內所有 <tr> 的 <td> 內容
    var rows = [];
    var rowPattern = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
    var rowMatch;
    while ((rowMatch = rowPattern.exec(tableHtml)) !== null) {
      var cells = extractCells(rowMatch[1]);
      if (cells.length > 0) rows.push(cells);
    }

    // 需要至少 4 列（標題+資料+標題+資料）
    if (rows.length < 4) continue;

    // 列2（index 1）: 卡別, 行動卡號後4碼, 授權日期, 授權時間, 消費地區
    var infoRow = rows[1];
    var date = '';
    var time = '';
    var region = '';
    var mobileCard = '';
    if (infoRow.length >= 3) date = infoRow[2];
    if (infoRow.length >= 4) time = infoRow[3];
    if (infoRow.length >= 5) region = infoRow[4];
    if (infoRow.length >= 2) mobileCard = infoRow[1];

    // 列4（index 3）: 消費金額, 商店名稱, 消費類別, 備註
    var detailRow = rows[3];
    var amount = '';
    var merchant = '';
    var category = '';
    var remark = '';
    if (detailRow.length >= 1) {
      var m = detailRow[0].match(/NT\$\s*([\d,]+)/);
      amount = m ? m[1].replace(/,/g, '') : detailRow[0];
    }
    if (detailRow.length >= 2) merchant = detailRow[1];
    if (detailRow.length >= 3) category = detailRow[2];
    if (detailRow.length >= 4) remark = detailRow[3];

    // 組裝備註
    var noteParts = [];
    if (remark && remark !== '註一') noteParts.push(remark);
    if (time) noteParts.push(time);
    if (region && region !== 'TW') noteParts.push('海外-' + region);
    if (mobileCard) noteParts.push('行動卡' + mobileCard);

    records.push({
      date: date,
      item: merchant || '未知',
      amount: amount,
      category: category || '',
      subCategory: '',
      payment: '信用卡',
      note: noteParts.join(' ')
    });
  }

  return records;
}

/**
 * 從 <tr> 內容擷取所有 <td> 的文字
 */
function extractCells(rowHtml) {
  var cells = [];
  var cellPattern = /<td[^>]*>([\s\S]*?)<\/td>/gi;
  var m;
  while ((m = cellPattern.exec(rowHtml)) !== null) {
    cells.push(stripHtml(m[1]).trim());
  }
  return cells;
}

// ─── 試算表寫入 ──────────────────────────────────────

function writeToSheet(records) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#4a86c8')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  var lastRow = sheet.getLastRow();
  // ID 從現有最後一筆的 ID + 1 開始
  var startId = 1;
  if (lastRow > 1) {
    var lastId = sheet.getRange(lastRow, 1).getValue();
    if (typeof lastId === 'number') startId = lastId + 1;
  }

  var rows = records.map(function(r, idx) {
    return [startId + idx, r.date, r.item, r.amount, r.category, r.subCategory, r.payment, r.note];
  });

  sheet.getRange(lastRow + 1, 1, rows.length, HEADERS.length).setValues(rows);
  sheet.getRange(lastRow + 1, 4, rows.length, 1).setNumberFormat('#,##0');

  return rows.length;
}

// ─── 工具函式 ────────────────────────────────────────

function stripHtml(html) {
  return html.replace(/<[^>]*>/g, '').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&');
}

function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
}

function getOrCreateLabel(name) {
  var label = GmailApp.getUserLabelByName(name);
  if (!label) label = GmailApp.createLabel(name);
  return label;
}

// ─── 觸發器管理 ──────────────────────────────────────

function createDailyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'main') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('main').timeBased().everyDays(1).atHour(8).create();
  Logger.log('已建立每日觸發器（每天早上 8 點）');
}

function removeAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  Logger.log('已移除所有觸發器');
}

/**
 * 移除所有消費通知信上的「已記帳」標籤
 */
function removeAllLabels() {
  var label = GmailApp.getUserLabelByName(LABEL_NAME);
  if (!label) {
    Logger.log('標籤「' + LABEL_NAME + '」不存在');
    return;
  }
  var threads = label.getThreads();
  for (var i = 0; i < threads.length; i++) {
    threads[i].removeLabel(label);
  }
  Logger.log('已從 ' + threads.length + ' 個信件移除「' + LABEL_NAME + '」標籤');
}

// ─── Debug ───────────────────────────────────────────

/**
 * 抓一封消費通知的 HTML，用 Logger 輸出
 */
function debugEmailHtml() {
  var threads = GmailApp.search('subject:"國泰世華銀行消費彙整通知"', 0, 1);
  if (threads.length === 0) {
    Logger.log('找不到信件');
    return;
  }
  var html = threads[0].getMessages()[0].getBody();
  // Logger.log 有字數限制，分段輸出
  var chunkSize = 4000;
  for (var i = 0; i < html.length; i += chunkSize) {
    Logger.log(html.substring(i, i + chunkSize));
  }
}

// ─── 試算表選單 ──────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi().createMenu('記帳小幫手')
    .addItem('立即掃描', 'main')
    .addItem('設定每日自動執行', 'createDailyTrigger')
    .addItem('移除自動執行', 'removeAllTriggers')
    .addToUi();
}
