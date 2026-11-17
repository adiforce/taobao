/**
 * 集運單查詢 - 公開「唯讀」API（獨立 Apps Script 專案，跟管理後台分開）
 * 只提供讀取，沒有任何寫入功能；姓名在這裡就遮罩（真名不會離開試算表）。
 * 部署：網頁應用程式 → 執行身分「我」、存取權「任何人」
 *
 *   ?action=list            → 集運單清單
 *   ?action=batch&name=XXX  → 單張集運單明細
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
    return a.date < b.date ? 1 : a.date > b.date ? -1 : b.row - a.row;      // 新→舊
  }).map(function (b) { return { name: b.name, date: b.date, kg: b.kg, status: statusOf_(by[b.name]), count: by[b.name].length }; });
  return { batches: out };
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

function doGet(e) {
  var p = (e && e.parameter) || {}, out;
  try {
    if (p.action === 'batch') out = batch_(p.name);
    else out = list_();
    out.ok = true;
  } catch (err) { out = { ok: false, error: String(err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
