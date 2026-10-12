// 泰時刻客製代購：接收網頁表單並寫入 Google Sheet
// 貼到「與試算表綁定」的 Apps Script（試算表 > 擴充功能 > Apps Script）
// 一張訂單有幾個商品就寫幾列，客人資料每列重複，方便篩選與對帳。
// 匯款總金額由這裡重新計算（不採用網頁傳來的數字），只寫在每張訂單的第一列。

var SHEET_NAME = '代購訂單';
var HEADERS = ['時間', '訂單編號', '到貨通知聯繫方式', '姓名(賣貨便核對)', '商品名稱', '數量',
               '單件金額', '小計', '匯款總金額', '匯款後五碼', '出貨平台', '備註',
               '已同意條款', '狀態'];
var CONFIG_SHEET = '設定';
var MAX_ITEMS = 30; // 防濫用的上限，正常訂單遠低於此

// 網頁載入時讀取匯款資料：「設定」分頁第 1 列是標題「銀行」「帳號」，第 2 列是內容
// 只回傳這兩項，分頁裡其他內容不會外流。分頁不存在時自動建立（內容是範例，請自行改掉）。
function doGet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CONFIG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG_SHEET);
    sh.getRange(1, 1, 2, 2).setValues([['銀行', '帳號'], ['範例銀行（代碼 000）', '0000-000-000000']]);
  }
  // 用顯示值讀取，帳號開頭的 0 才不會被吃掉
  var v = sh.getDataRange().getDisplayValues();
  var head = (v[0] || []).map(function (x) { return String(x).trim(); });
  var row = v[1] || [];
  var bank = String(row[head.indexOf('銀行')] || '').trim();
  var account = String(row[head.indexOf('帳號')] || '').trim();
  return json_({ok: true, bank: bank, account: account});
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.website) return json_({ok: true}); // honeypot：機器人填了就直接忽略
    var items = (d.items || []).slice(0, MAX_ITEMS);
    if (!d.contact || !items.length || d.agreed !== true) return json_({ok: false, error: 'missing'});
    if (d.platform === '賣貨便' && !d.name) return json_({ok: false, error: 'missing name'});
    if (!/^\d{5}$/.test(String(d.last5 || ''))) return json_({ok: false, error: 'bad last5'});
    for (var i = 0; i < items.length; i++) {
      if (!items[i].product) return json_({ok: false, error: 'missing product'});
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.setFrozenRows(1);
    }
    var now = new Date();
    var orderId = Utilities.formatDate(now, 'Asia/Taipei', 'yyMMddHHmmss');
    var rows = [];
    var total = 0;
    items.forEach(function (it) {
      var qty = Number(it.qty) || 1;
      var price = Number(it.price) || 0;
      total += qty * price;
      rows.push([
        now, orderId, s_(d.contact), s_(d.name), s_(it.product), qty, price, qty * price, '',
        s_(d.last5), s_(d.platform), s_(d.note), '是', '待確認匯款'
      ]);
    });
    rows[0][8] = total;
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
    return json_({ok: true, orderId: orderId, total: total});
  } catch (err) {
    return json_({ok: false, error: String(err)});
  } finally {
    lock.releaseLock();
  }
}

// 防止試算表公式注入：開頭是 = + - @ 的文字前面加單引號
function s_(v) {
  v = String(v == null ? '' : v).slice(0, 1000);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
