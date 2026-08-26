function doGet(e) {
  return ContentService.createTextOutput("API is running");
}

function doPost(e) {
  try {
    const contents = JSON.parse(e.postData.contents);
    const action = contents.action;
    let result = {};

    if (action === 'checkLogin') {
      result = checkLogin(contents.empId);
    } else if (action === 'searchArticle') {
      result = searchArticle(contents.keyword);
    } else if (action === 'saveIncoming') {
      result = saveIncoming(contents.items, contents.user);
    } else if (action === 'saveOutgoing') {
      result = saveOutgoing(contents.items, contents.user);
    } else if (action === 'getDashboardData') {
      result = getDashboardData();
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Function สำหรับเชื่อม URL Sheet
function getSpreadsheetUrl() {
  return SpreadsheetApp.getActiveSpreadsheet().getUrl();
}

// --- ฟังก์ชันเดิมที่มีอยู่แล้ว นำมาต่อตรงนี้ ---
function checkLogin(empId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Emp');
  if (!sheet) {
    sheet = ss.insertSheet('Emp');
    sheet.appendRow(['รหัสพนักงาน', 'ชื่อ-นามสกุล']);
    sheet.appendRow(['1001', 'พนักงาน คลังสินค้า']);
  }
  const data = sheet.getSheetValues(1, 1, sheet.getLastRow() || 1, sheet.getLastColumn() || 1);
  const searchId = String(empId).trim();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === searchId) {
      return { success: true, name: data[i][1] || searchId };
    }
  }
  return { success: false, message: 'ไม่พบรหัสพนักงานนี้ในระบบ' };
}

function searchArticle(keyword) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Data Aticle ALL SAP');
  if (!sheet) return null;
  const data = sheet.getSheetValues(1, 1, sheet.getLastRow() || 1, sheet.getLastColumn() || 1);
  const kw = String(keyword).trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    const article = String(data[i][0]).trim();
    const desc = String(data[i][1]).trim();
    const sku = String(data[i][2]).trim();
    if (article.toLowerCase() === kw || sku.toLowerCase() === kw || desc.toLowerCase().includes(kw)) {
      return { article: article, description: desc, sku: sku };
    }
  }
  return null;
}

function saveIncoming(items, user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetIn = ss.getSheetByName('บันทึกสินค้าเกิน') || ss.insertSheet('บันทึกสินค้าเกิน');
  let sheetHist = ss.getSheetByName('History') || ss.insertSheet('History');
  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');

  items.forEach(item => {
    sheetIn.appendRow([dateStr, timeStr, item.article, item.desc, item.sku, item.qty, item.note, user]);
    sheetHist.appendRow(['รับเข้า', dateStr, timeStr, item.article, item.desc, item.sku, item.qty, item.note, user]);
  });
  return { success: true, message: 'บันทึกรายการสินค้าเกินเรียบร้อยแล้ว' };
}

function saveOutgoing(items, user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetOut = ss.getSheetByName('รายการตัดสินค้าเกิน') || ss.insertSheet('รายการตัดสินค้าเกิน');
  let sheetHist = ss.getSheetByName('History') || ss.insertSheet('History');
  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');

  items.forEach(item => {
    sheetOut.appendRow([dateStr, timeStr, item.article, item.desc, item.sku, item.qty, item.note, user]);
    sheetHist.appendRow(['ตัดออก', dateStr, timeStr, item.article, item.desc, item.sku, item.qty, item.note, user]);
  });
  return { success: true, message: 'บันทึกรายการตัดเบิกสินค้าเกินเรียบร้อยแล้ว' };
}

function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetIn = ss.getSheetByName('บันทึกสินค้าเกิน');
  const sheetOut = ss.getSheetByName('รายการตัดสินค้าเกิน');
  const sheetHist = ss.getSheetByName('History');
  const stockMap = {};

  if (sheetIn && sheetIn.getLastRow() > 1) {
    const dataIn = sheetIn.getRange(1, 1, sheetIn.getLastRow(), sheetIn.getLastColumn()).getDisplayValues();
    for (let i = 1; i < dataIn.length; i++) {
      const art = String(dataIn[i][2] || '').trim();
      const desc = String(dataIn[i][3] || '').trim();
      const sku = String(dataIn[i][4] || '').trim();
      const qty = parseFloat(dataIn[i][5]) || 0;
      const bill = String(dataIn[i][6] || '').trim();
      if (art) {
        if (!stockMap[art]) stockMap[art] = { article: art, desc: desc, sku: sku, qty: 0, bill: bill };
        stockMap[art].qty += qty;
        if (bill) stockMap[art].bill = bill;
      }
    }
  }

  if (sheetOut && sheetOut.getLastRow() > 1) {
    const dataOut = sheetOut.getRange(1, 1, sheetOut.getLastRow(), sheetOut.getLastColumn()).getDisplayValues();
    for (let i = 1; i < dataOut.length; i++) {
      const art = String(dataOut[i][2] || '').trim();
      const qty = parseFloat(dataOut[i][5]) || 0;
      if (art && stockMap[art]) stockMap[art].qty -= qty;
    }
  }

  const historyList = [];
  if (sheetHist && sheetHist.getLastRow() > 1) {
    const dataHist = sheetHist.getRange(1, 1, sheetHist.getLastRow(), sheetHist.getLastColumn()).getDisplayValues();
    for (let i = dataHist.length - 1; i >= 1; i--) {
      historyList.push({
        type: dataHist[i][0], date: dataHist[i][1], time: dataHist[i][2],
        article: dataHist[i][3], desc: dataHist[i][4], sku: dataHist[i][5],
        qty: dataHist[i][6], note: dataHist[i][7], user: dataHist[i][8]
      });
    }
  }
  return { stock: Object.values(stockMap), history: historyList };
}