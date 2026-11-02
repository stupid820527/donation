/**
 * 聖義堂記帳 App — Google 試算表同步後端
 *
 * 安裝步驟（只做一次）：
 * 1. 開 https://sheets.google.com 建立一個新的試算表，命名「聖義堂記帳」。
 * 2. 上方選單 擴充功能 → Apps Script。
 * 3. 把編輯器裡原本的內容全部刪掉，貼上這整份程式碼，按存檔。
 * 4. 右上角「部署」→「新增部署作業」→ 類型選「網頁應用程式」。
 *    - 執行身分：我
 *    - 誰可以存取：任何人
 *    - 按「部署」，第一次會要你授權，一路允許。
 * 5. 複製它給你的「網頁應用程式網址」（結尾是 /exec）。
 * 6. 回到記帳 App 的「歷次」分頁 → 雲端同步 → 立即取回。
 *
 * ★ 更新這份程式碼之後，一定要「部署 → 管理部署作業 → 編輯（鉛筆）→
 *   版本：新版本 → 部署」，網址才不變、而且新程式碼才會生效。
 *   改完在 App 按一下任何欄位（觸發同步），分頁就會重新產生。
 */

var SHEET_DATA = '資料';   // 存 JSON 全文（App 讀寫的真正資料）
var SHEET_FEE  = '收費名單';
var SHEET_SPON = '隨喜贊助';
var SHEET_COST = '支出明細';
var SHEET_INS  = '保險名單';
var SHEET_PUDU = '普渡名單';
var SHEET_ROSTER = '名冊';
var SHEET_TEMPLE = '參拜紀錄';

function doGet(e) {
  var out = { ok: true, data: null, updatedAt: 0 };
  try {
    var sh = getSheet_(SHEET_DATA);
    var json = sh.getRange('A2').getValue();
    var ts = sh.getRange('B2').getValue();
    if (json) {
      out.data = JSON.parse(json);
      out.updatedAt = Number(ts) || 0;
    }
  } catch (err) {
    out.ok = false;
    out.error = String(err);
  }
  return reply_(out);
}

function doPost(e) {
  var out = { ok: true };
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var body = JSON.parse(e.postData.contents);
    var data = body.data;
    var ts = Number(body.updatedAt) || Date.now();

    var sh = getSheet_(SHEET_DATA);
    if (sh.getLastRow() < 1) sh.getRange('A1:B1').setValues([['JSON 資料（請勿手動修改）', '更新時間']]);
    sh.getRange('A2').setValue(JSON.stringify(data));
    sh.getRange('B2').setValue(ts);
    sh.getRange('C2').setValue(new Date());

    writeReadable_(data);
    out.updatedAt = ts;
  } catch (err) {
    out.ok = false;
    out.error = String(err);
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
  return reply_(out);
}

/* 把資料同時攤平成人看得懂的分頁，方便直接在試算表上檢視或列印 */
function writeReadable_(d) {
  var trips = d.trips || [];
  var roster = d.roster || [];
  var events = d.puduEvents || (d.pudu ? [{year:'', name:'普渡', rows:d.pudu}] : []);

  var fee = [['出遊', '日期', '姓名', '桌號', '應繳', '狀態']];
  var spon = [['出遊', '日期', '姓名', '隨喜金額']];
  var cost = [['出遊', '日期', '費用項目', '金額']];

  trips.forEach(function (t) {
    (t.fees || []).forEach(function (x) {
      fee.push([t.place || '', t.date || '', x.name || '', x.table || '', numv_(x.due), x.paid ? '已繳' : '未繳']);
    });
    (t.sponsors || []).forEach(function (x) {
      spon.push([t.place || '', t.date || '', x.name || '', numv_(x.amount)]);
    });
    (t.costs || []).forEach(function (c) {
      cost.push([t.place || '', t.date || '', c.name || '', numv_(c.amount)]);
    });
  });

  var ins = [['姓名', '出生年月日', '身分證字號']];
  var ros = [['順序', '姓名', '出生年月日', '身分證字號', '投保資料']];
  roster.forEach(function (r, i) {
    if (r.name && r.birth && r.id) ins.push([r.name, r.birth, r.id]);
    ros.push([i + 1, r.name || '', r.birth || '', r.id || '',
      (r.birth && r.id) ? '齊全' : '缺資料']);
  });

  var pd = [['年份', '場次', '姓名', '箱', '零包', '總包數', '金額', '已收']];
  events.forEach(function (ev) {
    var price = numv_(ev.price) || 1680;
    var bagP = price / 8;
    var tb = 0, tdue = 0, tpaid = 0;
    (ev.rows || []).forEach(function (p) {
      var bags = numv_(p.boxes) * 8 + numv_(p.bags);
      var amt = Math.round(bags * bagP);
      tb += bags; tdue += amt; if (p.paid) tpaid += amt;
      pd.push([ev.year || '', ev.name || '', p.name || '',
        numv_(p.boxes), numv_(p.bags), bags, amt, p.paid ? '付清' : '未收']);
    });
    pd.push([ev.year || '', '合計', (ev.rows || []).length + ' 位',
      Math.floor(tb / 8), tb % 8, tb, tdue, '已收 ' + tpaid + ' / 未收 ' + (tdue - tpaid)]);
  });

  dump_(SHEET_FEE, fee);
  dump_(SHEET_SPON, spon);
  dump_(SHEET_COST, cost);
  dump_(SHEET_INS, ins);
  dump_(SHEET_PUDU, pd);
  dump_(SHEET_ROSTER, ros);
  var tp = [['進香', '日期', '天', '順序', '廟宇', '供品／備註', '住宿']];
  trips.forEach(function (t) {
    (t.days || []).forEach(function (d) {
      (d.temples || []).forEach(function (x, i) {
        tp.push([t.place || '', t.date || '', d.label || '', i + 1, x.name || '', x.note || '', i === 0 ? (d.stay || '') : '']);
      });
    });
  });
  dump_(SHEET_TEMPLE, tp);
}

function dump_(name, rows) {
  var sh = getSheet_(name);
  sh.clear();
  if (!rows.length) return;
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  sh.setFrozenRows(1);
}

function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function numv_(v) {
  var n = parseInt(String(v == null ? '' : v).replace(/[^0-9-]/g, ''), 10);
  return isNaN(n) ? 0 : n;
}

function reply_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
