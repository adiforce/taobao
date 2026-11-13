/**
 * 淘寶代購管理 - Apps Script 後端
 * 貼在「買いもの_googleSheet」的 擴充功能 > Apps Script（容器綁定）
 * 只寫入「輸入欄」，計算欄（手續費、運費分攤、總額…）全部由試算表公式負責。
 */
var SS_ID = '';
var SH = { orders: '訂單', batches: '運費批次', pay: '付款紀錄', settle: '結算', set: '設定' };

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('淘寶代購管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- 工具 ---------- */
function ss_() {
  try { var s = SpreadsheetApp.getActive(); if (s) return s; } catch (e) {}
  return SpreadsheetApp.openById(SS_ID);
}
function sh_(k) {
  var s = ss_().getSheetByName(SH[k]);
  if (!s) throw new Error('找不到分頁：' + SH[k]);
  return s;
}
function tz_() { return ss_().getSpreadsheetTimeZone(); }
function lastRow_(sheet) {            // A 欄最後一筆資料所在列（沒有資料回傳 1）
  var v = sheet.getRange(1, 1, sheet.getMaxRows(), 1).getValues();
  for (var i = v.length - 1; i >= 1; i--) if (v[i][0] !== '') return i + 1;
  return 1;
}
function d2s_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v);
}
function s2d_(s) {
  if (!s) return '';
  var p = String(s).split('-');
  if (p.length !== 3) return s;
  return new Date(+p[0], +p[1] - 1, +p[2], 12, 0, 0);   // 取正午，避免時區位移
}
function num_(v) { return v === '' || v === null || v === undefined ? '' : Number(v); }
function numIn_(v) { return v === '' || v === null || v === undefined || isNaN(Number(v)) ? '' : Number(v); }
function ensureRows_(sheet, row) {
  var need = row - sheet.getMaxRows();
  if (need > 0) sheet.insertRowsAfter(sheet.getMaxRows(), need + 50);
}
function findRow_(sheet, col, key) {
  var last = lastRow_(sheet);
  if (last < 2) return -1;
  var v = sheet.getRange(2, col, last - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]) === String(key)) return i + 2;
  return -1;
}
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------- 讀取 ---------- */
function getAll_() {
  SpreadsheetApp.flush();
  return {
    settings: getSettings_(),
    orders: getOrders_(),
    batches: getBatches_(),
    payments: getPayments_(),
    settle: getSettle_()
  };
}

function getSettings_() {
  var s = sh_('set');
  var b = s.getRange(2, 2, 5, 1).getValues();
  function col(c, n) {
    return s.getRange(2, c, n || 8, 1).getValues().map(function (r) { return String(r[0]); })
      .filter(function (x) { return x !== ''; });
  }
  return { fee: Number(b[0][0]), card: Number(b[1][0]), reward: Number(b[2][0]), payer: String(b[3][0]), rateType: String(b[4][0] || '即期賣出'),
           people: col(4), statuses: col(6), methods: col(8), recipients: col(10, 20), addresses: col(12, 20),
           carriers: (function () { var c = s.getMaxColumns() >= 14 ? col(14, 10) : []; return c.length ? c : ['新竹貨運|https://www.sjlexpress.com/m/search?Num={num}']; })() };
}

function getOrders_() {
  var s = sh_('orders'), last = lastRow_(s);
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, 23).getValues().map(function (r, i) {
    return { id: r[0], date: d2s_(r[1]), name: String(r[2]), track: String(r[3]), weight: num_(r[4]),
      who: String(r[5]), cny: num_(r[6]), rate: num_(r[7]), batch: String(r[8]), status: String(r[9]), note: String(r[10]),
      fee1: num_(r[11]), fee2: num_(r[12]), reward: num_(r[13]), twd: num_(r[14]), share: num_(r[15]),
      ship: num_(r[16]), total: num_(r[17]), dup: r[18] === 'DUP', tb: String(r[19]), r1: r[20] === '' ? null : num_(r[20]), r2: r[21] === '' ? null : num_(r[21]), r3: r[22] === '' ? null : num_(r[22]), row: i + 2 };
  });
}

function getBatches_() {
  var s = sh_('batches'), last = lastRow_(s);
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, 19).getValues().map(function (r, i) {
    return { name: String(r[0]), date: d2s_(r[1]), kg: num_(r[2]), unit: num_(r[3]), cny: num_(r[4]),
      rate: num_(r[5]), link: String(r[6]), note: String(r[7]), twd: num_(r[8]), count: num_(r[9]),
      wsum: num_(r[10]), wdiff: num_(r[11]), itemTwd: num_(r[12]), total: num_(r[13]), addr: String(r[14]), rcp: String(r[15]), r1: r[16] === '' ? null : num_(r[16]), r2: r[17] === '' ? null : num_(r[17]), r3: r[18] === '' ? null : num_(r[18]), row: i + 2 };
  });
}

function getPayments_() {
  var s = sh_('pay'), last = lastRow_(s);
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, 6).getValues().map(function (r, i) {
    return { date: d2s_(r[0]), who: String(r[1]), amount: num_(r[2]), method: String(r[3]),
      batch: String(r[4]), note: String(r[5]), row: i + 2 };
  });
}

function getSettle_() {
  var s = sh_('settle');
  var v = s.getRange(1, 1, 12, 5).getValues();
  var persons = [];
  for (var i = 1; i <= 8; i++) {
    if (v[i][0] === '') continue;
    persons.push({ who: String(v[i][0]), due: num_(v[i][1]), paid: num_(v[i][2]),
      unpaid: v[i][3], pending: num_(v[i][4]) });
  }
  var total = { due: num_(v[10][1]), paid: num_(v[10][2]), unpaid: num_(v[10][3]), pending: num_(v[10][4]) };
  var check = num_(v[11][1]);
  var n = Math.max(0, Math.min(100, s.getMaxRows() - 15));
  var head = s.getRange(15, 1, 1, 10).getValues()[0];
  var rows = n ? s.getRange(16, 1, n, 10).getValues() : [];
  var cols = [];
  for (var c = 1; c <= 8; c++) if (head[c] !== '') cols.push({ idx: c, who: String(head[c]) });
  var matrix = [];
  rows.forEach(function (r) {
    if (r[0] === '') return;
    var cells = {};
    cols.forEach(function (cc) { cells[cc.who] = num_(r[cc.idx]); });
    matrix.push({ batch: String(r[0]), cells: cells, total: num_(r[9]) });
  });
  return { persons: persons, total: total, check: check, cols: cols.map(function (c) { return c.who; }), matrix: matrix };
}

/* ---------- 訂單 ---------- */
function orderRow_(o, id) {
  if (!o.name || !String(o.name).trim()) throw new Error('品名必填');
  if (!o.who) throw new Error('WHO 必填');
  return [id, s2d_(o.date), String(o.name).trim(), String(o.track || '').trim(), numIn_(o.weight),
    o.who, numIn_(o.cny), numIn_(o.rate), o.batch || '', o.status || '', o.note || ''];
}
function writeOrder_(s, row, vals, tb) {
  s.getRange(row, 2).setNumberFormat('yyyy/mm/dd');
  s.getRange(row, 4).setNumberFormat('@');         // 運單號 15~16 碼，必須是文字
  s.getRange(row, 1, 1, 11).setValues([vals]);
  var t = s.getRange(row, 20); t.setNumberFormat('@'); t.setValue(String(tb || '').trim());   // 淘寶訂單號(19碼)也要存文字
}

function snap_() {                               // 目前設定的三個費率；新增時寫進該筆，之後改設定不影響舊資料
  var b = sh_('set').getRange(2, 2, 3, 1).getValues();
  return [Number(b[0][0]) || 0, Number(b[1][0]) || 0, Number(b[2][0]) || 0];
}
function addOrder_(o) {
  return withLock_(function () {
    var s = sh_('orders'), last = lastRow_(s), max = 0;
    if (last >= 2) s.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
      var n = parseInt(String(r[0]).replace(/\D/g, ''), 10); if (n > max) max = n;
    });
    var id = 'O' + ('0000' + (max + 1)).slice(-4);
    var row = last + 1;
    ensureRows_(s, row);
    writeOrder_(s, row, orderRow_(o, id), o.tb);
    s.getRange(row, 21, 1, 3).setValues([snap_()]);
    return id;
  });
}

var LOCKED_ = ['集運回台', '已完成'];     // 這兩個狀態的訂單不能改狀態／集運單；要改只能刪除整張集運單
function lockedMsg_(name, st) { return '「' + name + '」已是「' + st + '」，不能改狀態或集運單；要改請先刪除整張集運單（會回到「在集運倉」）'; }
function updateOrder_(o) {
  return withLock_(function () {
    var s = sh_('orders'), row = findRow_(s, 1, o.id);
    if (row < 0) throw new Error('找不到訂單 ' + o.id + '，請重新整理');
    var cur = s.getRange(row, 3, 1, 8).getValues()[0], curBatch = String(cur[6]), curSt = String(cur[7]);   // C..J → I 批次、J 狀態
    if (LOCKED_.indexOf(curSt) >= 0 && (String(o.status || '') !== curSt || String(o.batch || '') !== curBatch)) throw new Error(lockedMsg_(String(cur[0]), curSt));
    writeOrder_(s, row, orderRow_(o, o.id), o.tb);
    return true;
  });
}

function deleteOrder_(id) {
  return withLock_(function () {
    var s = sh_('orders'), row = findRow_(s, 1, id);
    if (row < 0) throw new Error('找不到訂單 ' + id);
    s.deleteRow(row);
    return true;
  });
}

/* ---------- 運費批次 ---------- */
function batchRow_(b, name) {
  return [name, s2d_(b.date), numIn_(b.kg), numIn_(b.unit), numIn_(b.cny), numIn_(b.rate), b.link || '', b.note || ''];
}
function batchName_(b) {
  if (b.name && String(b.name).trim()) return String(b.name).trim();
  if (!b.date) throw new Error('請填集運單名稱或日期');
  var d = String(b.date).replace(/-/g, ''), k = (b.kg !== '' && b.kg != null && !isNaN(Number(b.kg))) ? '_' + Number(b.kg) + 'KG' : '';
  return d + k;                                    // 例：20261113_34KG
}
function writeBatch_(s, row, vals, addr, rcp) {
  s.getRange(row, 2).setNumberFormat('yyyy/mm/dd');
  s.getRange(row, 1, 1, 8).setValues([vals]);
  s.getRange(row, 15, 1, 2).setValues([[addr || '', rcp || '']]);
}

function addBatch_(b) {
  return withLock_(function () {
    var s = sh_('batches'), name = batchName_(b);
    if (findRow_(s, 1, name) > 0) throw new Error('集運單名稱已存在：' + name);
    var row = lastRow_(s) + 1;
    ensureRows_(s, row);
    writeBatch_(s, row, batchRow_(b, name), b.addr, b.rcp);
    s.getRange(row, 17, 1, 3).setValues([snap_()]);
    return name;
  });
}

function updateBatch_(b) {                 // b.oldName 為原名稱；若改名，連動訂單與付款紀錄
  return withLock_(function () {
    var s = sh_('batches'), row = findRow_(s, 1, b.oldName);
    if (row < 0) throw new Error('找不到集運單 ' + b.oldName + '，請重新整理');
    var name = batchName_(b);
    if (name !== b.oldName) {
      if (findRow_(s, 1, name) > 0) throw new Error('集運單名稱已存在：' + name);
      renameRefs_('orders', 9, b.oldName, name);
      renameRefs_('pay', 5, b.oldName, name);
    }
    writeBatch_(s, row, batchRow_(b, name), b.addr, b.rcp);
    return name;
  });
}
function renameRefs_(key, col, from, to) {
  var s = sh_(key), last = lastRow_(s);
  if (last < 2) return;
  var r = s.getRange(2, col, last - 1, 1), v = r.getValues(), hit = false;
  v.forEach(function (x) { if (String(x[0]) === from) { x[0] = to; hit = true; } });
  if (hit) r.setValues(v);
}

function deleteBatch_(name) {      // 刪集運單：底下訂單放回未集運，「集運回台」「已完成」的狀態都還原成「在集運倉」
  return withLock_(function () {
    var s = sh_('batches'), row = findRow_(s, 1, name);
    if (row < 0) throw new Error('找不到集運單 ' + name);
    var o = sh_('orders'), last = lastRow_(o), freed = 0;
    if (last >= 2) {
      var rng = o.getRange(2, 9, last - 1, 2), v = rng.getValues();            // I 批次、J 狀態
      v.forEach(function (r) {
        if (String(r[0]) === name) { r[0] = ''; if (LOCKED_.indexOf(String(r[1])) >= 0) r[1] = '在集運倉'; freed++; }
      });
      if (freed) rng.setValues(v);
    }
    s.deleteRow(row);
    return freed;
  });
}

/* ---------- 付款紀錄 ---------- */
function addPayment_(p) {
  return withLock_(function () {
    if (!p.who) throw new Error('請選擇付款人');
    var amt = numIn_(p.amount);
    if (amt === '' || amt === 0) throw new Error('金額必填');
    var s = sh_('pay'), row = lastRow_(s) + 1;
    ensureRows_(s, row);
    s.getRange(row, 1).setNumberFormat('yyyy/mm/dd');
    s.getRange(row, 1, 1, 6).setValues([[s2d_(p.date), p.who, amt, p.method || '', p.batch || '', p.note || '']]);
    return true;
  });
}

function deletePayment_(p) {               // 以列號 + 內容雙重確認，避免刪錯列
  return withLock_(function () {
    var s = sh_('pay'), r = s.getRange(p.row, 1, 1, 3).getValues()[0];
    if (String(r[1]) !== p.who || Number(r[2]) !== Number(p.amount)) throw new Error('資料已變動，請重新整理後再刪');
    s.deleteRow(p.row);
    return true;
  });
}

/* ---------- 設定 ---------- */
function saveSettings_(st) {
  return withLock_(function () {
    var s = sh_('set');
    s.getRange(2, 2, 5, 1).setValues([[Number(st.fee)], [Number(st.card)], [Number(st.reward)], [st.payer], [st.rateType || '即期賣出']]);
    function pad(arr) { var a = arr.slice(0, 8); while (a.length < 8) a.push(''); return a.map(function (x) { return [x]; }); }
    function pad20(arr) { var a = (arr || []).slice(0, 20); while (a.length < 20) a.push(''); return a.map(function (x) { return [x]; }); }
    s.getRange(2, 4, 8, 1).setValues(pad(st.people));
    s.getRange(2, 10, 20, 1).setValues(pad20(st.recipients));
    s.getRange(2, 12, 20, 1).setValues(pad20(st.addresses));
    if (s.getMaxColumns() < 14) s.insertColumnsAfter(s.getMaxColumns(), 14 - s.getMaxColumns());
    s.getRange(1, 14).setValue('集運公司|查詢網址');
    var cr = (st.carriers || []).slice(0, 10); while (cr.length < 10) cr.push('');
    s.getRange(2, 14, 10, 1).setValues(cr.map(function (x) { return [x]; }));
    s.getRange(2, 6, 8, 1).setValues(pad(st.statuses));
    s.getRange(2, 8, 8, 1).setValues(pad(st.methods));
    return true;
  });
}

/* ---------- 新增人員 / 批次指派 / 拆單 ---------- */
function addPerson_(name) {
  return withLock_(function () {
    name = String(name || '').trim();
    if (!name) throw new Error('請輸入名稱');
    var st = getSettings_();
    if (st.people.indexOf(name) >= 0) return name;
    if (st.people.length >= 8) throw new Error('人員最多 8 位（結算表只有 8 欄）');
    sh_('set').getRange(2 + st.people.length, 4).setValue(name);
    return name;
  });
}

function bulkAssign_(ids, batch, status) {      // batch: null=不改、''=清空；status: ''=不改
  return withLock_(function () {
    var s = sh_('orders'), last = lastRow_(s);
    if (last < 2) return 0;
    var A = s.getRange(2, 1, last - 1, 3).getValues(), rng = s.getRange(2, 9, last - 1, 2), v = rng.getValues(), n = 0;
    var want = {}; ids.forEach(function (x) { want[x] = 1; });
    var blocked = [];
    for (var k = 0; k < A.length; k++) {                  // 先檢查：已鎖定的訂單不能改狀態／集運單
      if (!want[A[k][0]] || LOCKED_.indexOf(String(v[k][1])) < 0) continue;
      var chB = batch !== null && batch !== undefined && String(batch) !== String(v[k][0]), chS = status && String(status) !== String(v[k][1]);
      if (chB || chS) blocked.push(String(A[k][2]) + '（' + v[k][1] + '）');
    }
    if (blocked.length) throw new Error('這些訂單已鎖定，不能改狀態或集運單：' + blocked.slice(0, 5).join('、') + (blocked.length > 5 ? ' 等 ' + blocked.length + ' 筆' : '') + '。要改請先刪除整張集運單。');
    for (var i = 0; i < A.length; i++) {
      if (!want[A[i][0]]) continue;
      if (batch !== null && batch !== undefined) v[i][0] = batch;
      if (status) v[i][1] = status;
      n++;
    }
    rng.setValues(v);
    return n;
  });
}

function splitOrder_(id, parts) {               // 同一筆淘寶訂單拆給多人：金額、重量各自填，其餘欄位沿用
  return withLock_(function () {
    if (!parts || parts.length < 2) throw new Error('至少要拆成兩份');
    var s = sh_('orders'), row = findRow_(s, 1, id);
    if (row < 0) throw new Error('找不到訂單 ' + id + '，請重新整理');
    var cur = s.getRange(row, 1, 1, 23).getValues()[0];
    var orig = Number(cur[6]) || 0, sum = 0;
    parts.forEach(function (p) {
      if (!p.who) throw new Error('每一份都要選 WHO');
      sum += Number(p.cny) || 0;
    });
    if (Math.abs(sum - orig) > 0.011) throw new Error('拆單金額合計 ¥' + sum.toFixed(2) + ' 與原訂單 ¥' + orig.toFixed(2) + ' 不符');
    var last = lastRow_(s), max = 0;
    s.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
      var n = parseInt(String(r[0]).replace(/\D/g, ''), 10); if (n > max) max = n;
    });
    var tb = String(cur[19] || '');
    var mk = function (id2, p, note) {
      return [id2, cur[1], cur[2], cur[3], numIn_(p.weight), p.who, numIn_(p.cny), cur[7], cur[8], cur[9], note];
    };
    var baseNote = cur[10] ? String(cur[10]) + '；' : '';
    var first = mk(id, parts[0], baseNote + '拆單(共' + parts.length + '份)');
    s.getRange(row, 1, 1, 11).setValues([first]);
    var next = last + 1, ids = [id];
    ensureRows_(s, last + parts.length);
    for (var i = 1; i < parts.length; i++) {
      var nid = 'O' + ('0000' + (++max)).slice(-4);
      writeOrder_(s, next, mk(nid, parts[i], baseNote + '拆自 ' + id), tb);
      s.getRange(next, 21, 1, 3).setValues([[cur[20], cur[21], cur[22]]]);     // 沿用原訂單當時的費率
      s.getRange(next, 2).setValue(cur[1]);
      next++; ids.push(nid);
    }
    return ids;
  });
}

/* ---------- 匯率（臺灣銀行牌告，可指定日期；假日自動往前找最近營業日） ---------- */
var RATE_IDX = { '現金買入': 0, '現金賣出': 1, '即期買入': 2, '即期賣出': 3 };

function fetchBot_(ds) {
  var res = UrlFetchApp.fetch('https://rate.bot.com.tw/xrt/quote/' + ds + '/CNY/spot', { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return null;
  var text = res.getContentText().replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ');
  var re = /(\d{4}\/\d{2}\/\d{2})\s+(\d{2}:\d{2}:\d{2})[^\d]*?\(CNY\)[^\d]*([\d.]+)[^\d]+([\d.]+)[^\d]+([\d.]+)[^\d]+([\d.]+)/g;
  var m, last = null;
  while ((m = re.exec(text)) !== null) last = m;      // 取當天最後一筆（收盤前最新）
  if (!last) return null;
  return { t: last[1] + ' ' + last[2], v: [+last[3], +last[4], +last[5], +last[6]] };
}

function parseQuote_(text) { return text; }   // 保留給測試用

function getRate_(dateStr, type) {
  type = type || getSettings_().rateType || '即期賣出';
  var tz = tz_(), now = new Date(), d = dateStr ? s2d_(dateStr) : now;
  if (!(d instanceof Date) || d > now) d = now;
  var todayS = Utilities.formatDate(now, tz, 'yyyy-MM-dd'), cache = CacheService.getScriptCache();
  for (var back = 0; back <= 7; back++) {
    var ds = Utilities.formatDate(new Date(d.getTime() - back * 86400000), tz, 'yyyy-MM-dd');
    var key = 'bot2:' + ds, q = cache.get(key);
    if (q) q = JSON.parse(q);
    else {
      try { q = fetchBot_(ds); } catch (e) { q = null; }
      if (q) cache.put(key, JSON.stringify(q), ds === todayS ? 1800 : 21600);
    }
    if (!q) continue;
    var rate = type === '中間價' ? (q.v[2] + q.v[3]) / 2 : q.v[RATE_IDX[type] != null ? RATE_IDX[type] : 3];
    return { rate: Math.round(rate * 1000) / 1000, date: ds, time: q.t, source: '臺灣銀行', type: type };
  }
  try {                                               // 後備：只有「當日」中間價
    var r = JSON.parse(UrlFetchApp.fetch('https://open.er-api.com/v6/latest/CNY').getContentText());
    if (r.result === 'success' && r.rates && r.rates.TWD) return { rate: Math.round(r.rates.TWD * 1000) / 1000, date: todayS, time: r.time_last_update_utc, source: 'exchangerate-api(後備)', type: '中間價' };
  } catch (e) {}
  throw new Error('抓不到匯率，請手動輸入');
}

function setWeights_(list) {                    // 集運後才知道重量：只更新 E 欄
  return withLock_(function () {
    var s = sh_('orders'), last = lastRow_(s);
    if (last < 2) return 0;
    var A = s.getRange(2, 1, last - 1, 1).getValues(), rng = s.getRange(2, 5, last - 1, 1), v = rng.getValues(), m = {}, n = 0;
    list.forEach(function (x) { m[x.id] = numIn_(x.weight); });
    for (var i = 0; i < A.length; i++) {
      if (Object.prototype.hasOwnProperty.call(m, A[i][0])) { v[i][0] = m[A[i][0]]; n++; }
    }
    rng.setValues(v);
    return n;
  });
}

/* ---------- 資料入口（誰能開網頁由部署設定「具有存取權的使用者：只有我」控管） ---------- */
var API_ = {
  getAll: getAll_, addOrder: addOrder_, updateOrder: updateOrder_, deleteOrder: deleteOrder_,
  addBatch: addBatch_, updateBatch: updateBatch_, deleteBatch: deleteBatch_,
  addPayment: addPayment_, deletePayment: deletePayment_, saveSettings: saveSettings_,
  addPerson: addPerson_, bulkAssign: bulkAssign_, splitOrder: splitOrder_, setWeights: setWeights_, getRate: getRate_
};

function api(token, fn, args) {                    // 唯一的資料入口（token 已不使用，保留參數是為了相容前端）
  if (!Object.prototype.hasOwnProperty.call(API_, fn)) throw new Error('未知的操作：' + fn);
  return API_[fn].apply(null, args || []);
}
