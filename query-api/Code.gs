/**
 * 集運單查詢 - 公開「唯讀」API（獨立 Apps Script 專案，跟管理後台分開）
 * 只提供讀取，沒有任何寫入功能；姓名在這裡就遮罩（真名不會離開試算表）。
 * 所有請求都要帶通行碼 key＝「設定」分頁收件人欄裡任一支電話（伺服器端比對，網頁本身不含密碼）。
 * 部署：網頁應用程式 → 執行身分「我」、存取權「任何人」
 *
 *   ?action=list            → 集運單清單
 *   ?action=batch&name=XXX  → 單張集運單明細
 *   ?action=people          → 可查詢的人（遮罩後）
 *   ?action=person&who=A**  → 某人的未付款總額與各集運單明細
 */
var SS_ID = '';   // 貼上試算表 ID（網址 /d/ 與 /edit 中間那串）

function ss_() { return SpreadsheetApp.openById(SS_ID); }
function sh_(n) { var s = ss_().getSheetByName(n); if (!s) throw new Error('找不到分頁：' + n); return s; }
function lastRow_(s) {
  var v = s.getRange(1, 1, s.getMaxRows(), 1).getValues();
  for (var i = v.length - 1; i >= 1; i--) if (v[i][0] !== '') return i + 1;
  return 1;
}
function d2s_(v) { return v instanceof Date ? Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd') : (v == null ? '' : String(v)); }
function num_(v) { return v === '' || v == null || isNaN(Number(v)) ? 0 : Number(v); }

// 姓名遮罩：第一個字 + 其餘字數的星號（Amy→A**、Karrie→K*****、JR→J*）
function mask_(w) {
  w = String(w || '').trim();
  if (!w) return '';
  var a = Array.from(w);
  return a[0] + new Array(a.length).join('*');
}

function orders_() {
  var s = sh_('訂單'), last = lastRow_(s);
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, 20).getValues().map(function (r, i) {
    return { name: String(r[2]), track: String(r[3]), weight: r[4] === '' ? '' : num_(r[4]), who: String(r[5]),
      cny: r[6], batch: String(r[8]), status: String(r[9]),
      twd: r[14] === '' ? '' : num_(r[14]), share: r[15] === '' ? '' : num_(r[15]),
      ship: r[16] === '' ? '' : num_(r[16]), total: r[17] === '' ? '' : num_(r[17]), date: d2s_(r[1]), row: i };
  });
}
function batches_() {
  var s = sh_('運費批次'), last = lastRow_(s);
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, 14).getValues().map(function (r, i) {
    return { name: String(r[0]), date: d2s_(r[1]), kg: num_(r[2]), unit: num_(r[3]), cny: num_(r[4]), rate: num_(r[5]),
      link: String(r[6]), twd: num_(r[8]), count: num_(r[9]), total: num_(r[13]), row: i };
  });
}
function payments_() {
  var s = sh_('付款紀錄'), last = lastRow_(s);
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, 6).getValues().map(function (r) { return { who: String(r[1]), amount: num_(r[2]), batch: String(r[4]) }; });
}
function payer_() { return String(sh_('設定').getRange(5, 2).getValue()); }   // 設定 B5＝代墊人
function people_() {
  return sh_('設定').getRange(2, 4, 8, 1).getValues().map(function (r) { return String(r[0]); }).filter(Boolean);
}

function statusOf_(os) {
  if (!os.length) return '';
  var seen = {}, list = [];
  os.forEach(function (o) { if (o.status && !seen[o.status]) { seen[o.status] = 1; list.push(o.status); } });
  return list.length === 1 ? list[0] : (list.indexOf('集運回台') >= 0 ? '集運回台' : list[list.length - 1]);
}

function list_() {
  var os = orders_(), by = {};
  os.forEach(function (o) { if (o.batch) (by[o.batch] = by[o.batch] || []).push(o); });
  var out = batches_().filter(function (b) { return by[b.name]; }).sort(function (a, b) {
    return a.name < b.name ? 1 : a.name > b.name ? -1 : b.row - a.row;      // 依集運單號 新→舊
  }).map(function (b) { return { name: b.name, date: b.date, kg: b.kg, status: statusOf_(by[b.name]), count: by[b.name].length }; });
  return { batches: out };
}

function r2_(x) { return Math.round(x * 100) / 100; }

// 可查詢的人（遮罩後；代墊人不用付款，不列出）
function people2_() {
  var payer = payer_(), ppl = people_(), seen = {}, list = [];
  orders_().forEach(function (o) { if (o.who && o.who !== payer && !seen[o.who]) { seen[o.who] = 1; list.push(o.who); } });
  list.sort(function (a, b) { var i = ppl.indexOf(a), j = ppl.indexOf(b); return (i < 0 ? 99 : i) - (j < 0 ? 99 : j); });
  return { people: list.map(mask_) };
}

// 某人的付款狀況：未付總額（取自「結算」分頁）＋各集運單的應付／已付／明細
function person_(masked) {
  var payer = payer_(), mine = {};
  orders_().forEach(function (o) { if (o.who && o.who !== payer && mask_(o.who) === masked) mine[o.who] = 1; });
  var names = Object.keys(mine);
  if (!names.length) throw new Error('找不到這個人');
  var os = orders_().filter(function (o) { return mine[o.who] && o.batch; });
  var bmap = {}; batches_().forEach(function (b) { bmap[b.name] = b; });
  var pays = payments_().filter(function (p) { return mine[p.who]; }), paidAll = 0, paidBy = {};
  pays.forEach(function (p) { paidAll += p.amount; if (p.batch) paidBy[p.batch] = (paidBy[p.batch] || 0) + p.amount; });
  var g = {}, order = [];
  os.forEach(function (o) { if (!g[o.batch]) { g[o.batch] = []; order.push(o.batch); } g[o.batch].push(o); });
  order.sort(function (a, b) { return a < b ? 1 : a > b ? -1 : 0; });          // 集運單號 新→舊
  var dueAll = 0, inBatchPaid = 0;
  var batches = order.map(function (n) {
    var rows = g[n], due = 0;
    rows.forEach(function (o) { due += num_(o.twd) + num_(o.ship); });
    var paid = paidBy[n] || 0; dueAll += due; inBatchPaid += paid;
    return { name: n, date: bmap[n] ? bmap[n].date : '', status: statusOf_(rows), due: r2_(due), paid: r2_(paid), unpaid: r2_(due - paid),
      rows: rows.map(function (o) { return { name: o.name, track: o.track, weight: o.weight, cny: o.cny === '' ? '' : num_(o.cny), twd: o.twd, ship: o.ship }; }) };
  });
  var unpaid = r2_(dueAll - paidAll);                                              // 預設自行加總
  var st = sh_('結算').getRange(2, 1, 8, 4).getValues();                           // 有「結算」分頁數字就以它為準
  st.forEach(function (r) { if (mine[String(r[0])] && typeof r[3] === 'number') unpaid = r2_(r[3]); });
  return { who: masked, due: r2_(dueAll), paid: r2_(paidAll), unpaid: unpaid, otherPaid: r2_(paidAll - inBatchPaid), batches: batches };
}

function batch_(name) {
  name = String(name || '');
  var b = batches_().filter(function (x) { return x.name === name; })[0];
  if (!b) throw new Error('找不到這張集運單');
  var rows = orders_().filter(function (o) { return o.batch === name; });
  if (!rows.length) throw new Error('這張集運單沒有品項');
  var ppl = people_(), rank = function (w) { var i = ppl.indexOf(w); return i < 0 ? 99 : i; };
  rows.sort(function (a, c) { return rank(a.who) - rank(c.who) || (a.who < c.who ? -1 : a.who > c.who ? 1 : 0) || (a.date < c.date ? -1 : 1) || a.row - c.row; });
  var paid = {};
  payments_().forEach(function (p) { if (p.batch === name) paid[p.who] = (paid[p.who] || 0) + p.amount; });
  var payer = payer_();
  var persons = [], seen = {};
  rows.forEach(function (o) {
    if (!seen[o.who]) { seen[o.who] = { who: mask_(o.who), n: 0, a: 0, s: 0, paid: paid[o.who] || 0, payer: o.who === payer }; persons.push(seen[o.who]); }
    var p = seen[o.who]; p.n++; p.a += num_(o.twd); p.s += num_(o.ship);
  });
  return {
    name: b.name, date: b.date, kg: b.kg, unit: b.unit, rate: b.rate, link: b.link, status: statusOf_(rows),
    rows: rows.map(function (o) {
      return { who: mask_(o.who), name: o.name, track: o.track, weight: o.weight, cny: o.cny === '' ? '' : num_(o.cny),
        twd: o.twd, ship: o.ship };
    }),
    persons: persons
  };
}

/* ---------- 通行碼：用「設定」分頁收件人欄（J 欄，格式「姓名 電話」）裡的任何一支電話 ---------- */
function digits_(t) { return String(t || '').replace(/\D/g, ''); }
function normPhone_(d) { return d.indexOf('886') === 0 ? '0' + d.slice(3) : d; }   // +886912… → 0912…
function phones_() {
  var out = [];
  sh_('設定').getRange(2, 10, 20, 1).getValues().forEach(function (r) {
    var m = String(r[0] || '').match(/\+?\d[\d\- ]{6,}\d/);                       // 抓出像電話的數字串
    if (m) { var d = normPhone_(digits_(m[0])); if (d.length >= 8) out.push(d); }
  });
  return out;
}
function auth_(key) {
  var cache = CacheService.getScriptCache(), bad = Number(cache.get('badtry') || 0);
  if (bad >= 20) throw new Error('LOCK');                                           // 10 分鐘內錯 20 次就暫停
  var d = normPhone_(digits_(key));
  if (d.length >= 8 && phones_().indexOf(d) >= 0) return;
  cache.put('badtry', String(bad + 1), 600);
  throw new Error('AUTH');
}

function doGet(e) {
  var p = (e && e.parameter) || {}, out;
  try {
    auth_(p.key);
    if (p.action === 'batch') out = batch_(p.name);
    else if (p.action === 'people') out = people2_();
    else if (p.action === 'person') out = person_(p.who);
    else out = list_();
    out.ok = true;
  } catch (err) { out = { ok: false, error: String(err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
