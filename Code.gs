/**
 * Budget app — Google Apps Script backend (salary-cycle model).
 * A "cycle" runs from salary day (default 25th) to the day before the next salary.
 * Reads from / writes to the "Budget" spreadsheet this script is bound to.
 */

const PLAN = 'Total Plan';
const MONTH = 'This Month';
const LOAN = 'Home Loan';
const CELL_SALARY = 'A39';
const CELL_CURRENT = 'A41';
const CELL_SAVINGS = 'A43';
const CELL_CARD = 'A45';      // prepaid credit card balance
const CELL_CASH = 'A47';      // cash in hand
const LOG = 'App Log';
const PAID = 'App Paid';
const SETTINGS = 'App Settings';
const HISTORY = 'App History';
const ACCOUNTS = 'App Accounts';   // Name | Balance | Kind (spend/save) | Deduct (1 = ticked items come out of this one) | Order | Color
const FIXED = ['Monthly', 'Quarterly', 'Annually', 'Annually (Hijri)', 'One Time'];
const DEFAULT_PAID = ['Monthly'];           // assumed paid at the start of the cycle unless switched off
const BIG = ['Quarterly', 'Annually', 'Annually (Hijri)', 'One Time'];
// Total Plan layout: M cycle (1st of salary month) | N Opening | O Expenses | P Big items | Q Savings | R Cumulative savings | S spacer | T:U dashboard
const DASH = { labels: 'T2:T16', values: 'U2:U16' };

// The app page is hosted on GitHub Pages and calls this script as a JSON API.
// Every call must carry the passphrase stored in Script Properties under APP_TOKEN.
function doGet() {
  return ContentService.createTextOutput('Budget API').setMimeType(ContentService.MimeType.TEXT);
}
function doPost(e) {
  let out;
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const want = PropertiesService.getScriptProperties().getProperty('APP_TOKEN');
    if (!want || String(body.token || '') !== want) throw new Error('Wrong passphrase');
    const API = { getData, updateBalances, updateAccounts, addAccount, editAccount, deleteAccount, moveAccount, setPaid, addExpense, editExpense, deleteExpense, updateLoan, setSetting, setSpread, monthlyClose, salaryLanded };
    const fn = API[body.fn];
    if (!fn) throw new Error('Unknown function: ' + body.fn);
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try { out = { ok: true, data: fn.apply(null, body.args || []) }; }
    finally { lock.releaseLock(); }
  } catch (err) {
    out = { ok: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- helpers ----------
function ss() { return SpreadsheetApp.getActive(); }
function sh(name) { return ss().getSheetByName(name); }
function dateOnly(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function addMonths(d, n) { return new Date(d.getFullYear(), d.getMonth() + n, d.getDate()); }
function addDays(d, n) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); }
function isDate(v) { return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v); }
function ymd(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function fmt(d, p) { return Utilities.formatDate(d, Session.getScriptTimeZone(), p); }
function num(v) { return typeof v === 'number' ? v : (parseFloat(String(v).replace(/,/g, '')) || 0); }
function r2(v) { return Math.round(num(v) * 100) / 100; }
function serialToDate(s) { return new Date(Date.UTC(1899, 11, 30) + s * 86400000); }
function hijriToSerial(d, m, y) { return Math.floor((11 * y + 3) / 30) + 354 * y + 30 * m - Math.floor((m - 1) / 2) + d - 466964; }
// A cycle is named after its salary month (key 'Myyyy-MM'). Nominally it starts on salaryDay of that month and ends the day
// before the next cycle starts. If a salary lands early, the app records cycleStart:<key> and the cycle starts on that day instead.
function nominalStart(y, mo, sd) { return new Date(y, mo, Math.min(28, Math.max(1, sd))); }
function parseYmd(v) { if (isDate(v)) return dateOnly(v); const m = String(v || '').match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null; }
function keyOf(y, mo) { const d = new Date(y, mo, 1); return 'M' + d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2); }
function startOf(y, mo, settings) { // actual start of the cycle named (y, mo)
  const early = parseYmd(settings['cycleStart:' + keyOf(y, mo)]);
  const nom = nominalStart(y, mo, settings.salaryDay);
  return (early && early < nom) ? early : nom;
}
function cycleFor(d, settings) { // the cycle that contains day d
  let y = d.getFullYear(), mo = d.getMonth();
  if (d < startOf(y, mo, settings)) { mo -= 1; }           // before this month's start -> previous month's cycle
  if (d >= startOf(y, mo + 1, settings)) { mo += 1; }      // next cycle already started early
  const cs = startOf(y, mo, settings);
  const ce = addDays(startOf(y, mo + 1, settings), -1);
  return { cs: cs, ce: ce, key: keyOf(y, mo), month: new Date(y, mo, 1) };
}
function cycleByMonth(m1, settings) { // the cycle for a 1st-of-month date from the sheet
  const y = m1.getFullYear(), mo = m1.getMonth();
  return { cs: startOf(y, mo, settings), ce: addDays(startOf(y, mo + 1, settings), -1), key: keyOf(y, mo), month: new Date(y, mo, 1) };
}
function cycleKey(cs) { return 'M' + cs.getFullYear() + '-' + ('0' + (cs.getMonth() + 1)).slice(-2); }
function cycleLabelOf(c) { return fmt(c.cs, 'd MMM') + ' – ' + fmt(c.ce, 'd MMM'); }
function paidMap(mk) {
  const rows = sh(PAID).getDataRange().getValues().slice(1);
  const o = {};
  rows.forEach(r => { if (String(r[0]) === mk) o[String(r[1])] = { state: String(r[2]), when: isDate(r[3]) ? r[3] : null }; });
  return o;
}

function ensureAppSheets() {
  const c = CacheService.getScriptCache();
  if (c.get('sheetsOk')) return;
  const s = ss();
  if (!s.getSheetByName(LOG)) s.insertSheet(LOG).appendRow(['Timestamp', 'Current', 'Savings', 'Total', 'Expected', 'Delta', 'Card', 'Cash']);
  else if (s.getSheetByName(LOG).getRange('H1').getValue() === '') s.getSheetByName(LOG).getRange('G1:H1').setValues([['Card', 'Cash']]);
  if (!s.getSheetByName(PAID)) { const t = s.insertSheet(PAID); t.appendRow(['MonthKey', 'Item', 'State', 'When']); t.getRange('A:A').setNumberFormat('@'); }
  if (!s.getSheetByName(SETTINGS)) {
    const t = s.insertSheet(SETTINGS);
    t.appendRow(['Key', 'Value']);
    t.appendRow(['salaryDay', 27]);
  }
  if (!s.getSheetByName(HISTORY)) s.insertSheet(HISTORY).appendRow(['Cycle', 'Planned', 'Opening', 'Closing', 'Actual spend', 'Delta vs plan']);
  if (!s.getSheetByName(ACCOUNTS)) createAccountsSheet();
  c.put('sheetsOk', '1', 21600);
}

// ---------- accounts (as many as you like, named as you like) ----------
// First run: seeds the list from the old fixed cells (Current A41, Saving A43, Credit A45, Cash A47).
function createAccountsSheet() {
  const s = ss();
  if (s.getSheetByName(ACCOUNTS)) return;
  const t = s.insertSheet(ACCOUNTS);
  t.appendRow(['Name', 'Balance', 'Kind', 'Deduct', 'Order', 'Color']);
  const mv = s.getSheetByName(MONTH).getRange('A41:A47').getValues();
  t.getRange(2, 1, 4, 6).setValues([
    ['Current', r2(mv[0][0]), 'spend', 1, 1, ''],
    ['Saving', r2(mv[2][0]), 'save', '', 2, ''],
    ['Credit', r2(mv[4][0]), 'spend', '', 3, ''],
    ['Cash', r2(mv[6][0]), 'spend', '', 4, '']
  ]);
  t.getRange('B:B').setNumberFormat('#,##0.00');
  try { s.getSheetByName(PLAN).getRange('T10').setValue('Spendable now (all spending accounts)'); } catch (e) {}
  SpreadsheetApp.flush();
}
function readAccounts() {
  let t = sh(ACCOUNTS);
  if (!t) { createAccountsSheet(); t = sh(ACCOUNTS); }
  const last = t.getLastRow();
  if (last < 2) return [];
  if (t.getRange('E1').getValue() === '') t.getRange('E1:F1').setValues([['Order', 'Color']]);
  return t.getRange(2, 1, last - 1, 6).getValues()
    .map((r, i) => ({ row: i + 2, name: String(r[0]).trim(), balance: r2(r[1]), kind: String(r[2]).toLowerCase() === 'save' ? 'save' : 'spend', deduct: r[3] === true || num(r[3]) === 1, order: r[4] === '' ? i + 2 : num(r[4]), color: String(r[5] || '') }))
    .filter(a => a.name)
    .sort((a, b) => a.order - b.order || a.row - b.row);
}
function deductAccount(accs) { return accs.find(a => a.deduct && a.kind === 'spend') || accs.find(a => a.kind === 'spend') || null; }
function sumKind(accs, kind) { return r2(accs.filter(a => a.kind === kind).reduce((x, a) => x + a.balance, 0)); }
// Keep the old This Month cells as a mirror (A41 = all spending money, A43 = all savings) so anything in the sheet that pointed there still works.
function mirrorAccounts(accs, mv) {
  const sp = sumKind(accs, 'spend'), sv = sumKind(accs, 'save');
  if (!mv || r2(mv[2][0]) !== sp || r2(mv[4][0]) !== sv || mv[6][0] !== '' || mv[8][0] !== '') {
    const m = sh(MONTH);
    m.getRange(CELL_CURRENT).setValue(sp); m.getRange(CELL_SAVINGS).setValue(sv);
    m.getRange(CELL_CARD).setValue(''); m.getRange(CELL_CASH).setValue('');
  }
}
function updateAccounts(list) {
  ensureAppSheets();
  const t = sh(ACCOUNTS) || (createAccountsSheet(), sh(ACCOUNTS));
  const accs = readAccounts();
  (list || []).forEach(u => {
    const a = accs.find(x => (u.row && x.row === Number(u.row)) || (!u.row && x.name === u.name));
    if (a) { a.balance = r2(u.balance); t.getRange(a.row, 2).setValue(a.balance); }
  });
  mirrorAccounts(accs, null);
  SpreadsheetApp.flush();
  const d = getData();
  sh(LOG).appendRow([new Date(), d.spendable, d.savings, d.total, d.expectedSpendable, d.spendDelta, 0, 0]);
  d.log.push({ t: new Date().toISOString(), current: d.spendable, savings: d.savings, total: d.total, expected: d.expectedSpendable, delta: d.spendDelta });
  return d;
}
function addAccount(name, kind, deduct, color) {
  ensureAppSheets();
  name = String(name || '').trim(); if (!name) throw new Error('Name is required');
  const accs = readAccounts();
  if (accs.some(a => a.name.toLowerCase() === name.toLowerCase())) throw new Error('An account called "' + name + '" already exists');
  kind = kind === 'save' ? 'save' : 'spend';
  const t = sh(ACCOUNTS);
  if (deduct && kind === 'spend') accs.forEach(a => { if (a.deduct) t.getRange(a.row, 4).setValue(''); });
  const maxOrder = accs.reduce((m, a) => Math.max(m, a.order), 0);
  t.appendRow([name, 0, kind, (deduct && kind === 'spend') ? 1 : '', maxOrder + 1, String(color || '')]);
  SpreadsheetApp.flush();
  return getData();
}
function editAccount(row, name, kind, deduct, color) {
  ensureAppSheets();
  row = Number(row); name = String(name || '').trim(); if (!name) throw new Error('Name is required');
  const accs = readAccounts(); const a = accs.find(x => x.row === row); if (!a) throw new Error('Account not found');
  if (accs.some(x => x.row !== row && x.name.toLowerCase() === name.toLowerCase())) throw new Error('An account called "' + name + '" already exists');
  kind = kind === 'save' ? 'save' : 'spend';
  const t = sh(ACCOUNTS);
  if (deduct && kind === 'spend') accs.forEach(x => { if (x.deduct && x.row !== row) t.getRange(x.row, 4).setValue(''); });
  t.getRange(row, 1, 1, 6).setValues([[name, a.balance, kind, (deduct && kind === 'spend') ? 1 : '', a.order, color === undefined || color === null ? a.color : String(color)]]);
  SpreadsheetApp.flush();
  return getData();
}
// Move an account one step up (dir = -1) or down (dir = +1) in the list.
function moveAccount(row, dir) {
  ensureAppSheets();
  row = Number(row); dir = Number(dir) < 0 ? -1 : 1;
  const accs = readAccounts();               // already sorted by order
  const i = accs.findIndex(a => a.row === row); if (i < 0) throw new Error('Account not found');
  const j = i + dir; if (j < 0 || j >= accs.length) return getData();
  const t = sh(ACCOUNTS);
  // renumber 1..n so the order column is always clean, then swap i and j
  const tmp = accs[i]; accs[i] = accs[j]; accs[j] = tmp;
  accs.forEach((a, k) => t.getRange(a.row, 5).setValue(k + 1));
  SpreadsheetApp.flush();
  return getData();
}
function deleteAccount(row) {
  ensureAppSheets();
  row = Number(row);
  const accs = readAccounts(); const a = accs.find(x => x.row === row); if (!a) throw new Error('Account not found');
  if (accs.length <= 1) throw new Error('Keep at least one account');
  const t = sh(ACCOUNTS);
  t.deleteRow(row);
  if (a.deduct) { const rest = readAccounts(); const nx = rest.find(x => x.kind === 'spend'); if (nx) t.getRange(nx.row, 4).setValue(1); }
  mirrorAccounts(readAccounts(), null);
  SpreadsheetApp.flush();
  return getData();
}

function getSettings() {
  ensureAppSheets();
  const rows = sh(SETTINGS).getDataRange().getValues().slice(1);
  const o = { salaryDay: 27, spread: {}, perWeek: {} };
  const truthy = v => v === true || v === 1 || String(v).toLowerCase() === 'true';
  rows.forEach(r => { const k = String(r[0] || ''); if (!k) return; if (k.indexOf('spread:') === 0) { if (truthy(r[1])) o.spread[k.slice(7)] = true; } else if (k.indexOf('perWeek:') === 0) { if (truthy(r[1])) o.perWeek[k.slice(8)] = true; } else o[k] = r[1]; });
  o.salaryDay = num(o.salaryDay) || 27;
  return o;
}
function writeSetting(key, value) {
  const t = sh(SETTINGS);
  const rows = t.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) if (String(rows[i][0]) === key) { t.getRange(i + 1, 2).setValue(value); return; }
  t.appendRow([key, value]);
}
function setSpread(item, on) {
  ensureAppSheets();
  writeSetting('spread:' + item, !!on);
  return getData();
}

// ---------- read the expense table ----------
function readItems() {
  const p = sh(PLAN);
  const last = p.getLastRow();
  const vals = p.getRange(2, 1, Math.max(1, last - 1), 11).getValues();
  const items = [];
  vals.forEach((r, i) => {
    if (!r[0]) return;
    items.push({
      row: i + 2,
      item: String(r[0]),
      amount: num(r[1]),
      freq: String(r[2]),
      monthly: num(r[5]),
      annually: num(r[6]),
      add: isDate(r[9]) ? ymd(r[9]) : (r[9] === '' ? null : String(r[9])),
      addRaw: r[9],
      stop: isDate(r[10]) ? ymd(r[10]) : null,
      stopRaw: r[10]
    });
  });
  return items;
}

// Amount due for an item inside the window [ws, we] (both inclusive, date-only). Mirrors the sheet's Expenses formula.
function dueInWindow(it, ws, we) {
  const nd = Math.round((we - ws) / 86400000) + 1;
  const amt = it.amount;
  const k = it.stopRaw;
  if (isDate(k) && dateOnly(k) < ws) return 0;
  const a = it.addRaw;
  const inWin = d => d >= ws && d <= we;
  switch (it.freq) {
    case 'Daily': return amt * nd;
    case 'Weekly': return amt * nd / 7;
    case 'Monthly': return amt;
    case 'Quarterly': {
      if (!isDate(a)) return 0;
      let d = dateOnly(a);
      if (d > we) return 0;
      while (d < ws) d = addMonths(d, 3);
      return inWin(d) ? amt : 0;
    }
    case 'Annually': {
      if (!isDate(a)) { // no date given: due on 1 January
        for (let y = ws.getFullYear(); y <= we.getFullYear(); y++) if (inWin(new Date(y, 0, 1))) return amt;
        return 0;
      }
      const d0 = dateOnly(a);
      if (d0 > we) return 0;
      for (let y = Math.max(d0.getFullYear(), ws.getFullYear()); y <= we.getFullYear(); y++) if (inWin(new Date(y, d0.getMonth(), d0.getDate()))) return amt;
      return 0;
    }
    case 'Annually (Hijri)': {
      const parts = String(a || '').split('/');
      if (parts.length < 2) return 0;
      const d = parseInt(parts[0], 10), m = parseInt(parts[1], 10);
      if (!d || !m) return 0;
      for (let y = 1440; y < 1480; y++) {
        const g = serialToDate(hijriToSerial(d, m, y));
        const gl = new Date(g.getUTCFullYear(), g.getUTCMonth(), g.getUTCDate());
        if (inWin(gl)) return amt;
      }
      return 0;
    }
    case 'One Time':
      if (!isDate(a)) return 0;
      return inWin(dateOnly(a)) ? amt : 0;
    default: return 0;
  }
}
function bigItemsText(items, ws, we) {
  const parts = [];
  items.forEach(it => {
    const d = dueInWindow(it, ws, we);
    if (d > 0 && BIG.indexOf(it.freq) >= 0) parts.push(it.item + ' (' + Math.round(d).toLocaleString('en-US') + ')');
    if (isDate(it.stopRaw)) { const k = dateOnly(it.stopRaw); if (k >= ws && k <= we && it.freq === 'Monthly') parts.push(it.item + ' – last month'); }
  });
  return parts.join(', ');
}

// ---------- main payload ----------
function getData() {
  ensureAppSheets();
  const settings = getSettings();
  const today = dateOnly(new Date());
  const cyc = cycleFor(today, settings);
  const cs = cyc.cs, ce = cyc.ce;
  const nd = Math.round((ce - cs) / 86400000) + 1;
  const elapsed = Math.round((today - cs) / 86400000) + 1;   // today counts as a spending day
  const left = nd - elapsed;
  const items = readItems();
  const m = sh(MONTH);
  const mv = m.getRange('A39:A47').getValues();   // A39 salary; A41/A43 are now mirrors of the accounts list
  const salary = num(mv[0][0]);
  const accounts = readAccounts();
  mirrorAccounts(accounts, mv);
  const spendable = sumKind(accounts, 'spend');
  const savings = sumKind(accounts, 'save');
  const da = deductAccount(accounts);
  const current = da ? da.balance : 0;
  const total = r2(spendable + savings);
  const mk = cyc.key;
  const over = paidMap(mk);

  // --- salary status for this cycle: manual switch, or detected from a balance jump logged in this cycle
  const lg = sh(LOG); const lgLast = lg.getLastRow();
  const logRows = lgLast > 1 ? lg.getRange(Math.max(2, lgLast - 80), 1, Math.min(81, lgLast - 1), 8).getValues().filter(r => isDate(r[0])) : [];
  let salaryJumpDate = null;
  for (let i = 1; i < logRows.length; i++) { const d = logRows[i][0]; if (d >= cs && salary > 0 && num(r3(logRows[i])) - num(r3(logRows[i - 1])) >= 0.8 * salary) { salaryJumpDate = d; break; } }
  const ov = settings['salaryReceived:' + mk];
  const ovStr = ov === undefined || ov === '' ? '' : String(ov).toLowerCase();
  const salaryMode = ovStr === 'true' ? 'manual-yes' : (ovStr === 'false' ? 'manual-no' : (salaryJumpDate ? 'auto' : 'none'));
  const salaryReceived = salaryMode === 'manual-yes' || salaryMode === 'auto';

  // --- opening snapshot: the first time the salary is seen in a cycle, freeze the spendable money and the savings balance
  let opening = settings['opening:' + mk], savingsStart = settings['savingsStart:' + mk];
  if (salaryReceived && (opening === undefined || opening === '')) {
    opening = spendable; savingsStart = savings;
    writeSetting('opening:' + mk, opening); writeSetting('savingsStart:' + mk, savingsStart);
  }
  const hasOpening = !(opening === undefined || opening === '');
  opening = hasOpening ? num(opening) : null;
  savingsStart = hasOpening ? num(savingsStart) : savings;

  // --- items due in this cycle
  let planned = 0, remaining = 0, spentPlan = 0, remainingVar = 0;
  const thisMonth = [];
  items.forEach(it => {
    const due = dueInWindow(it, cs, ce);
    if (due <= 0) return;
    const spread = it.freq === 'Monthly' && !!settings.spread[it.item];
    const isFixed = FIXED.indexOf(it.freq) >= 0 && !spread;
    const o = over[it.item];
    const defPaid = DEFAULT_PAID.indexOf(it.freq) >= 0;
    const paid = isFixed && (o ? o.state === 'paid' : defPaid);
    let rem, done;
    if (isFixed) { rem = paid ? 0 : due; done = paid ? due : 0; }
    else { rem = due * left / nd; done = due - rem; }
    planned += due; remaining += rem; spentPlan += done; if (!isFixed) remainingVar += rem;
    thisMonth.push({ row: it.row, item: it.item, freq: it.freq, due: due, remaining: rem, done: done, paid: paid, fixed: isFixed, spread: spread, big: BIG.indexOf(it.freq) >= 0, defPaid: defPaid, perDay: isFixed ? 0 : due / nd, perWeek: !isFixed && !!settings.perWeek[it.item] });
  });
  thisMonth.sort((a, b) => (a.fixed === b.fixed) ? b.due - a.due : (a.fixed ? -1 : 1)); // switches first, then bars; largest first

  // --- the plan for this cycle
  const openingUsed = hasOpening ? opening : (salaryReceived ? spendable : spendable + salary);
  const savingsPlan = openingUsed - planned;                 // move this to savings when the salary lands (negative = pull back)
  const cumTarget = savingsStart + savingsPlan;              // what the savings account should hold this cycle
  const moved = savings - savingsStart;                      // what has actually been moved so far
  const savingsDelta = savings - cumTarget;                  // + ahead / - still to move
  const expectedSpendable = openingUsed - moved - spentPlan; // what should be left in current+credit+cash today
  const spendDelta = spendable - expectedSpendable;          // + under plan / - over plan
  const closing = spendable - remaining;                     // what will be left at the end of the cycle (move it to savings then)

  // --- every cycle in the sheet (M column): expenses, big items, item lines. Written to O:P only when the plan changed.
  const p = sh(PLAN);
  const months = p.getRange(2, 13, 90, 1).getValues();
  const rowsOut = [], cycles = [];
  months.forEach(r => {
    if (!isDate(r[0])) { rowsOut.push(['', '']); return; }
    const wc = cycleByMonth(r[0], settings);
    let tot = 0; const lines = [];
    items.forEach(it => { const d = dueInWindow(it, wc.cs, wc.ce); if (d > 0) { tot += d; lines.push([it.item, Math.round(d * 100) / 100, BIG.indexOf(it.freq) >= 0 ? 1 : 0]); } });
    lines.sort((a, b) => b[1] - a[1]);
    cycles.push({ key: wc.key, label: fmt(wc.month, 'MMM yyyy'), window: cycleLabelOf(wc), total: tot, lines: lines });
    rowsOut.push([tot, bigItemsText(items, wc.cs, wc.ce)]);
  });
  const sig = JSON.stringify([settings.salaryDay, items.map(it => [it.item, it.amount, it.freq, String(it.addRaw), String(it.stopRaw)]), Object.keys(settings).filter(k => k.indexOf('cycleStart:') === 0).map(k => k + '=' + settings[k])]);
  const cache = CacheService.getScriptCache();
  const sigHash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, sig));
  const planChanged = cache.get('planSig') !== sigHash;
  if (planChanged) { p.getRange(2, 15, rowsOut.length, 2).setValues(rowsOut); cache.put('planSig', sigHash, 21600); }
  const future = cycles.filter(c => c.key > mk).slice(0, 12).map(c => ({ key: c.key, label: c.label, window: c.window, total: c.total, lines: c.lines.map(l => ({ item: l[0], amount: l[1], freq: l[2] ? 'Big' : 'Regular' })) }));

  // --- write the live numbers into the Total Plan dashboard so the sheet tells the same story as the app
  const dashVals = [
    [cycleLabelOf(cyc) + ' (' + fmt(cyc.month, 'MMM yyyy') + ' salary)'],
    [nd], [elapsed], [left],
    [planned], [spentPlan], [remaining], [left > 0 ? remainingVar / left : remainingVar],
    [spendable], [closing],
    [openingUsed], [savingsPlan], [cumTarget], [savings]
  ];
  const dashHash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(dashVals)));
  if (planChanged || cache.get('dashSig') !== dashHash) {
    p.getRange(DASH.values).setValues(dashVals.concat([[new Date()]]));
    SpreadsheetApp.flush();
    cache.put('dashSig', dashHash, 21600);
  }

  // --- read the plan table back (N, Q, R are sheet formulas)
  const fr = p.getRange(2, 13, 90, 6).getValues().filter(r => isDate(r[0]));
  const byKey = {}; cycles.forEach(c => byKey[c.key] = c);
  const forecast = fr.map(r => { const k = cycleKey(r[0]); const c = byKey[k] || {}; return { label: fmt(r[0], 'MMM yy'), key: k, window: c.window || '', opening: num(r[1]), expenses: num(r[2]), big: String(r[3] || ''), savings: num(r[4]), cumulative: num(r[5]), lines: c.lines || [] }; });
  const log = logRows.slice(-60).map(r => ({ t: r[0].toISOString(), current: num(r[1]) + num(r[6]) + num(r[7]), savings: num(r[2]), total: num(r[3]), expected: r[4] === '' ? null : num(r[4]), delta: r[5] === '' ? null : num(r[5]) }));

  // --- loan inputs (cached; refreshed when the loan inputs or the plan change)
  let loan, loanSeries;
  const loanCached = planChanged ? null : cache.get('loan');
  if (loanCached) { const lc = JSON.parse(loanCached); loan = lc.loan; loanSeries = lc.loanSeries; }
  else {
    const l = sh(LOAN);
    const loanVals = l.getRange('A2:B22').getValues();
    loan = loanVals.map(r => ({ label: String(r[0]), value: isDate(r[1]) ? fmt(r[1], 'MMM yyyy') : r[1] }));
    loanSeries = l.getRange(2, 4, 90, 6).getValues().filter(r => isDate(r[0])).map(r => ({ label: fmt(r[0], 'MMM yy'), rent: num(r[1]), buy: num(r[2]), owed: r[3] === '' ? null : num(r[3]), wealth: num(r[4]), diff: num(r[5]) }));
    try { cache.put('loan', JSON.stringify({ loan: loan, loanSeries: loanSeries }), 600); } catch (e) {}
  }

  return {
    today: today.toISOString(), hijri: fmt(today, 'EEEE, d MMMM'),
    cycle: { key: mk, label: cycleLabelOf(cyc), month: fmt(cyc.month, 'MMMM yyyy'), start: ymd(cs), end: ymd(ce), days: nd, elapsed: elapsed, left: left, early: parseYmd(settings['cycleStart:' + mk]) ? true : false },
    salary: salary, current: current, savings: savings, spendable: spendable, total: total,
    accounts: accounts, deductRow: da ? da.row : null,
    salaryDay: settings.salaryDay, salaryReceived: salaryReceived, salaryMode: salaryMode, monthKey: mk,
    opening: openingUsed, hasOpening: hasOpening, savingsStart: savingsStart,
    planned: planned, remaining: remaining, remainingVar: remainingVar, spentPlan: spentPlan, dailyLeft: left > 0 ? remainingVar / left : remainingVar,
    savingsPlan: savingsPlan, cumTarget: cumTarget, moved: moved, savingsDelta: savingsDelta,
    expectedSpendable: expectedSpendable, spendDelta: spendDelta, closing: closing,
    thisMonth: thisMonth, future: future, forecast: forecast, log: log,
    items: items.map(it => ({ row: it.row, item: it.item, amount: it.amount, freq: it.freq, monthly: it.monthly, annually: it.annually, add: it.add, stop: it.stop, spread: it.freq === 'Monthly' && !!settings.spread[it.item], perWeek: !!settings.perWeek[it.item] })),
    loan: loan, loanSeries: loanSeries
  };
}
function r3(r) { return r[3] !== '' ? r[3] : num(r[1]) + num(r[2]); }

// ---------- writes ----------
function updateBalances(current, savings, card, cash) {
  // legacy (fixed four accounts) — maps onto the first spend account / first save account
  const accs = readAccounts();
  const sp = deductAccount(accs), sv = accs.find(a => a.kind === 'save');
  const list = []; if (sp) list.push({ row: sp.row, balance: num(current) + num(card) + num(cash) }); if (sv) list.push({ row: sv.row, balance: savings });
  return updateAccounts(list);
}

// Marking an item paid takes its amount off the deduct account (the spending account marked ● in the app) (you can always correct the balance by hand);
// un-marking it puts back whatever that tick had taken off. Bills that are assumed paid by default never move the balance.
function setPaid(item, paid) {
  ensureAppSheets();
  const t = sh(PAID);
  const mk = cycleFor(dateOnly(new Date()), getSettings()).key;
  if (t.getRange('E1').getValue() === '') t.getRange('E1').setValue('Deducted');
  // (cycle + item resolved below without a full getData)
  const rows = t.getDataRange().getValues();
  let prevState = null, prevDeducted = 0;
  for (let i = rows.length - 1; i >= 1; i--) if (String(rows[i][0]) === mk && String(rows[i][1]) === item) { if (prevState === null) { prevState = String(rows[i][2]); prevDeducted = num(rows[i][4]); } t.deleteRow(i + 1); }
  const settings = getSettings();
  const cyc = cycleFor(dateOnly(new Date()), settings);
  const itm = readItems().find(x => x.item === item);
  const due = itm ? dueInWindow(itm, cyc.cs, cyc.ce) : 0;
  const defPaid = itm && DEFAULT_PAID.indexOf(itm.freq) >= 0 && !settings.spread[item];
  const wasPaid = prevState === null ? !!defPaid : prevState === 'paid';
  let deducted = 0;
  const accs = readAccounts(); const da = deductAccount(accs); const at = sh(ACCOUNTS);
  if (paid && !wasPaid && due > 0 && da) { at.getRange(da.row, 2).setValue(r2(da.balance - due)); deducted = r2(due); }
  else if (!paid && prevDeducted > 0 && da) { at.getRange(da.row, 2).setValue(r2(da.balance + prevDeducted)); deducted = 0; }
  else if (paid) deducted = prevDeducted;
  t.appendRow([mk, item, paid ? 'paid' : 'unpaid', new Date(), deducted]);
  t.getRange('A:A').setNumberFormat('@');
  SpreadsheetApp.flush();
  return getData();
}

function addExpense(o) {
  const p = sh(PLAN);
  const items = readItems();
  const lastRow = items.length ? items[items.length - 1].row : 1;
  const r = lastRow + 1;
  p.getRange(r, 1, 1, 3).setValues([[o.item, num(o.amount), o.freq]]);
  if (o.add) p.getRange(r, 10).setValue(o.freq === 'Annually (Hijri)' ? String(o.add) : new Date(o.add));
  if (o.stop) p.getRange(r, 11).setValue(new Date(o.stop));
  const f = p.getRange(2, 4, 1, 6).getFormulasR1C1();
  p.getRange(r, 4, 1, 6).setFormulasR1C1(f);
  p.getRange(2, 2, 1, 10).copyTo(p.getRange(r, 2, 1, 10), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  return getData();
}

function editExpense(o) {
  const p = sh(PLAN);
  const r = parseInt(o.row, 10);
  p.getRange(r, 1, 1, 3).setValues([[o.item, num(o.amount), o.freq]]);
  p.getRange(r, 10).setValue(o.add ? (o.freq === 'Annually (Hijri)' ? String(o.add) : new Date(o.add)) : '');
  p.getRange(r, 11).setValue(o.stop ? new Date(o.stop) : '');
  return getData();
}

function deleteExpense(row) {
  sh(PLAN).deleteRow(parseInt(row, 10));
  return getData();
}

function updateLoan(o) {
  const l = sh(LOAN);
  if (o.price !== undefined) l.getRange('B2').setValue(num(o.price));
  if (o.loan !== undefined) l.getRange('B4').setValue(num(o.loan));
  if (o.rate !== undefined) l.getRange('B5').setValue(num(o.rate) / 100);
  if (o.years !== undefined) l.getRange('B6').setValue(num(o.years));
  if (o.fees !== undefined) l.getRange('B7').setValue(num(o.fees) / 100);
  if (o.floor !== undefined) l.getRange('B11').setValue(num(o.floor));
  if (o.month) l.getRange('B13').setValue(new Date(o.month));
  if (o.growth !== undefined) l.getRange('B17').setValue(num(o.growth) / 100);
  SpreadsheetApp.flush();
  CacheService.getScriptCache().remove('loan');
  return getData();
}

function setSetting(key, value) {
  ensureAppSheets();
  writeSetting(key, value);
  return getData();
}
// The "Salary received" switch. ON when this cycle has no salary yet -> mark it received (opening gets frozen).
// ON when this cycle already has its salary -> the NEXT salary landed early: start the next cycle today.
// OFF -> un-mark this cycle (and, if it was started early today, remove that early start).
function salaryLanded(on) {
  ensureAppSheets();
  const settings = getSettings();
  const today = dateOnly(new Date());
  const cyc = cycleFor(today, settings);
  const has = settings['opening:' + cyc.key] !== undefined && settings['opening:' + cyc.key] !== '';
  if (on && !has) { writeSetting('salaryReceived:' + cyc.key, 'true'); return getData(); }
  if (on && has) {
    const nk = keyOf(cyc.month.getFullYear(), cyc.month.getMonth() + 1);
    writeSetting('cycleStart:' + nk, ymd(today));
    writeSetting('salaryReceived:' + nk, 'true');
    return getData();
  }
  writeSetting('salaryReceived:' + cyc.key, 'false');
  if (settings['cycleStart:' + cyc.key] && ymd(parseYmd(settings['cycleStart:' + cyc.key])) === ymd(today)) writeSetting('cycleStart:' + cyc.key, '');
  return getData();
}

function monthlyClose() {
  ensureAppSheets();
  const d = getData();
  sh(HISTORY).appendRow([d.cycle.label, d.planned, d.opening, d.spendable, d.opening - d.moved - d.spendable, d.spendDelta]);
  return getData();
}

// ---------- sheet edit trigger: keep Expenses / Big items / dashboard current when the table is edited by hand ----------
function onSheetEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    const name = sheet.getName();
    const col = e.range.getColumn();
    const touchesPlan = (name === PLAN && col <= 11) || (name === MONTH) || (name === SETTINGS);
    if (!touchesPlan) return;
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(5000)) return;
    try { getData(); } finally { lock.releaseLock(); }
  } catch (err) { /* never block the user's edit */ }
}
function installEditTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'onSheetEdit') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onSheetEdit').forSpreadsheet(ss()).onEdit().create();
  return 'installed';
}

// ---------- one-off maintenance (run from the editor) ----------
// Move the salary day to the 27th and record that the Oct 2026 salary landed early on 25 Oct 2026.
function setSalaryDay27() {
  writeSetting('salaryDay', 27);
  writeSetting('cycleStart:M2026-10', '2026-10-25');
  sh(PLAN).getRange('M1').setValue('Salary month (cycle 27th → 26th)');
  getData();
  return 'done';
}
// Rebuilds the forecast block of Total Plan for the salary-cycle model:
//  M salary month (1st of month; the cycle runs salaryDay -> salaryDay-1 of next month)
//  N Opening balance  O Expenses  P Big items  Q Savings this cycle  R Cumulative savings  | S spacer | T:U dashboard
function setupCycleLayout() {
  const p = sh(PLAN);
  const sd = getSettings().salaryDay;
  // drop the old extra columns (S "Move this month", T "Left over") if they are still there
  if (String(p.getRange('T1').getValue()).indexOf('Left over') === 0) p.deleteColumns(19, 2);
  p.getRange('M1:R1').setValues([['Salary month (cycle ' + sd + 'th → ' + (sd - 1) + 'th)', 'Opening balance', 'Expenses', 'Big items', 'Savings this cycle', 'Cumulative savings']]);
  const n = 78; // rows 2..79
  const N = [], Q = [], R = [];
  for (let i = 0; i < n; i++) {
    const r = i + 2;
    N.push(["=IF(M" + r + "=\"\",\"\",IFERROR(VLOOKUP(\"opening:M\"&TEXT(M" + r + ",\"yyyy-MM\"),'App Settings'!$A:$B,2,FALSE),'This Month'!$A$39))"]);
    Q.push(["=IF(M" + r + "=\"\",\"\",N" + r + "-O" + r + ")"]);
    R.push(["=IF(M" + r + "=\"\",\"\",IFERROR(VLOOKUP(\"savingsStart:M\"&TEXT(M" + r + ",\"yyyy-MM\"),'App Settings'!$A:$B,2,FALSE),N(R" + (r - 1) + "))+Q" + r + ")"]);
  }
  p.getRange(2, 14, n, 1).setFormulas(N);
  p.getRange(2, 15, n, 2).clearContent(); // Expenses + Big items are written by the app on every refresh
  p.getRange(2, 17, n, 1).setFormulas(Q);
  p.getRange(2, 18, n, 1).setFormulas(R);
  const sar = '[$SAR] #,##0.00;[Red]-[$SAR] #,##0.00';
  p.getRange(2, 14, n, 2).setNumberFormat(sar);
  p.getRange(2, 17, n, 2).setNumberFormat(sar);
  p.getRange(2, 13, n, 1).setNumberFormat('MMM yyyy');
  p.getRange('M1:R1').setFontWeight('bold');
  p.setColumnWidth(16, 420); p.setColumnWidth(14, 120); p.setColumnWidth(15, 110); p.setColumnWidth(17, 130); p.setColumnWidth(18, 140);
  // dashboard labels (values are written by the app on every refresh)
  p.getRange(DASH.labels).setValues([
    ['Current cycle'], ['Days in cycle'], ['Days elapsed'], ['Days remaining'],
    ['Planned expenses this cycle'], ['Spent so far (from app)'], ['Still to spend (from app)'], ['Daily budget left'],
    ['Spendable now (all spending accounts)'], ['Closing balance at cycle end → move to savings'],
    ['Opening balance (after salary)'], ['Savings this cycle → move to savings on salary day'], ['Savings account should hold'], ['Savings account now'],
    ['Last synced from app']
  ]);
  p.getRange('T2:T16').setFontWeight('bold');
  p.getRange('U2:U16').setNumberFormat('[$SAR] #,##0.00;[Red]-[$SAR] #,##0.00');
  p.getRange('U2:U4').setNumberFormat('0'); p.getRange('U2').setNumberFormat('@');
  p.getRange('U16').setNumberFormat('dd/MM/yyyy HH:mm');
  // Home Loan tab: "savings if renting" now reads the Cumulative savings column (R) instead of the old Balance column (P)
  const l = sh(LOAN);
  const rng = l.getDataRange();
  const fs = rng.getFormulas();
  let changed = false;
  for (let i = 0; i < fs.length; i++) for (let j = 0; j < fs[i].length; j++) {
    const f = fs[i][j];
    if (f && f.indexOf("'Total Plan'!P") >= 0) { fs[i][j] = f.replace(/'Total Plan'!\$?P/g, "'Total Plan'!R"); changed = true; }
  }
  if (changed) { for (let i = 0; i < fs.length; i++) for (let j = 0; j < fs[i].length; j++) if (fs[i][j]) l.getRange(i + 1, j + 1).setFormula(fs[i][j]); }
  // freeze this cycle's opening if the salary is already marked received
  getData();
  return 'done';
}
