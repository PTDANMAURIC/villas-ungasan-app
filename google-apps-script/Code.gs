/**
 * VILLAS UNGASAN — operations app backend (Google Apps Script)
 * ------------------------------------------------------------
 * Stores everything in YOUR Google account:
 *   • Google Sheet "Villas Ungasan – App"  (Users, Expenses, Attendance, Tasks, Summary)
 *   • Drive folder "Villas Ungasan – App"   (Invoices/Villa 131…134, Attendance photos/YYYY-MM)
 *
 * HOW TO INSTALL — see SETUP-GUIDE. Short version:
 *   1. Paste this file into a new Apps Script project (script.google.com).
 *   2. Edit the SETUP block below if needed, then run  setup  once (▶ Run).
 *      The execution log prints everyone's starting PIN.
 *   3. Project Settings → Script properties → add ANTHROPIC_API_KEY (or GEMINI_API_KEY)
 *      so invoices can be read automatically.
 *   4. Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone.
 *
 * To refresh the look of the Google Sheet (Overview tab, column titles, Rupiah amounts): run  formatSheets  once.
 *      Copy the /exec URL into the app's index.html (API_URL).
 */

// ============================== SETUP ==============================
var SETUP = {
  appName: 'Villas Ungasan',
  timezone: 'Asia/Makassar',          // Bali time (WITA, UTC+8)
  currency: 'IDR',
  villas: [
    { id: '131', owner: 'Mauri' },
    { id: '132', owner: 'Dani' },
    { id: '133', owner: 'Dani' },
    { id: '134', owner: 'Ric' }
  ],
  // People created by setup(). Random PINs are generated and printed in the log.
  // role: 'admin' (everything + Team), 'owner' (expenses of ALL villas + workers + tasks; 'villas' = which they own), 'worker'
  people: [
    { name: 'Leon',  role: 'admin', villas: 'all' },
    { name: 'Mauri', role: 'owner', villas: '131' },
    { name: 'Dani',  role: 'owner', villas: '132,133' },
    { name: 'Ric',   role: 'owner', villas: '134' },
    { name: 'Elton', role: 'worker', dailyRate: 150000, mealAllowance: 23000, monthlyBonus: 500000, deductionPct: 4 }, // handyman / maintenance
    { name: 'Elsi',  role: 'worker', dailyRate: 150000, mealAllowance: 23000, monthlyBonus: 0, deductionPct: 4 }       // cleaning & housekeeping
    // More workers: add here before running setup, or later in the app (Team tab, admin only).
  ],
  categories: ['Materials', 'Labour', 'Furniture & equipment', 'Utilities', 'Cleaning', 'Pool & garden',
               'Maintenance & repairs', 'Permits & fees', 'Transport', 'Other']
};

var CLAUDE_MODEL_DEFAULT = 'claude-sonnet-5-5';
var GEMINI_MODEL_DEFAULT = 'gemini-2.5-flash';
var SESSION_DAYS = 60;

var COLS = {
  Users:      ['id', 'name', 'role', 'villas', 'dailyRate', 'active', 'pinHash', 'salt', 'createdAt',
               'mealAllowance', 'monthlyBonus', 'deductionPct'],
  Expenses:   ['id', 'date', 'month', 'villa', 'owner', 'vendor', 'description', 'category', 'amount',
               'currency', 'invoiceNo', 'fileUrl', 'fileId', 'uploadedBy', 'createdAt', 'notes',
               'groupId', 'invoiceTotal', 'split'],
  Attendance: ['id', 'date', 'month', 'userId', 'worker', 'villa', 'checkIn', 'checkInPhotoUrl', 'checkInPhotoId',
               'checkInGPS', 'checkOut', 'checkOutPhotoUrl', 'checkOutPhotoId', 'checkOutGPS', 'hours', 'note'],
  Tasks:      ['id', 'userId', 'worker', 'title', 'notes', 'villa', 'due', 'status', 'createdBy', 'createdAt', 'doneAt']
};
var NUMERIC = { amount: '#,##0.##', invoiceTotal: '#,##0.##', dailyRate: '#,##0', mealAllowance: '#,##0', monthlyBonus: '#,##0',
                deductionPct: '0.##', hours: '0.00' };
var SCHEMA_VERSION = '3';

// ============================== ONE-TIME SETUP ==============================
function setup() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('SHEET_ID')) {
    Logger.log('Already set up. Sheet: ' + SpreadsheetApp.openById(props.getProperty('SHEET_ID')).getUrl());
    Logger.log('To reset a PIN use the Team tab in the app, or run resetPin("Name").');
    return;
  }
  var ss = SpreadsheetApp.create(SETUP.appName + ' – App');
  ss.setSpreadsheetTimeZone(SETUP.timezone);
  Object.keys(COLS).forEach(function (name, i) {
    var sh = i === 0 ? ss.getSheets()[0] : ss.insertSheet(name);
    if (i === 0) sh.setName(name);
    var cols = COLS[name];
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#e8efe9');
    sh.setFrozenRows(1);
    var rows = sh.getMaxRows() - 1;
    sh.getRange(2, 1, rows, cols.length).setNumberFormat('@');      // keep dates/times/ids as plain text
    cols.forEach(function (c, ci) {
      if (NUMERIC[c]) sh.getRange(2, ci + 1, rows, 1).setNumberFormat(NUMERIC[c]);
    });
  });
  var root = DriveApp.createFolder(SETUP.appName + ' – App');
  var inv = root.createFolder('Invoices');
  SETUP.villas.forEach(function (v) { inv.createFolder(villaFolderName_(v)); });
  inv.createFolder('Shared between villas');
  root.createFolder('Attendance photos');
  root.createFolder('_Pending scans');
  DriveApp.getFileById(ss.getId()).moveTo(root);

  props.setProperty('SHEET_ID', ss.getId());
  props.setProperty('ROOT_FOLDER_ID', root.getId());
  props.setProperty('SCHEMA_VERSION', SCHEMA_VERSION);

  var lines = ['', '======== STARTING PINs — hand these out privately ========'];
  SETUP.people.forEach(function (p) {
    var pin = randomPin_();
    createUser_({ name: p.name, role: p.role, villas: p.villas || '', dailyRate: p.dailyRate || '', mealAllowance: p.mealAllowance || '',
                  monthlyBonus: p.monthlyBonus || '', deductionPct: p.deductionPct || '', pin: pin });
    lines.push(pad_(p.name, 12) + ' PIN ' + pin + '   (' + p.role + ')');
  });
  lines.push('==========================================================');
  lines.push('Sheet:  ' + ss.getUrl());
  lines.push('Folder: ' + root.getUrl());
  formatSheets();
  lines.push('Next: add ANTHROPIC_API_KEY or GEMINI_API_KEY in Project Settings → Script properties, then Deploy → Web app.');
  Logger.log(lines.join('\n'));
}

/** Run from the editor if someone forgets their PIN: resetPin('Wayan') */
function resetPin(name) {
  var u = findUserByName_(name);
  if (!u) throw new Error('No user called ' + name);
  var pin = randomPin_();
  var h = hashPin_(pin);
  updateRow_('Users', u._row, { pinHash: h.hash, salt: h.salt });
  Logger.log('New PIN for ' + u.name + ': ' + pin);
}

// ============================== HTTP ENTRY POINTS ==============================
function doGet() {
  return json_({ ok: true, app: SETUP.appName, message: 'API is running. Open the app page, not this URL.' });
}

function doPost(e) {
  var res;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var fn = ACTIONS[req.action];
    if (!fn) throw new Error('Unknown action');
    ensureSchema_();
    res = { ok: true, data: fn(req) };
  } catch (err) {
    res = { ok: false, error: String(err && err.message || err) };
  }
  return json_(res);
}

var ACTIONS = {
  // ---------- auth ----------
  login: function (r) {
    var name = String(r.name || '').trim();
    var key = 'fail_' + name.toLowerCase();
    var cache = CacheService.getScriptCache();
    var fails = Number(cache.get(key) || 0);
    if (fails >= 5) throw new Error('LOCKED');
    var u = findUserByName_(name);
    if (!u || !isTrue_(u.active) || hashPin_(String(r.pin || ''), u.salt).hash !== u.pinHash) {
      cache.put(key, String(fails + 1), 900);
      throw new Error('BAD_LOGIN');
    }
    cache.remove(key);
    purgeSessions_();
    var token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
    PropertiesService.getScriptProperties().setProperty('S_' + token,
      JSON.stringify({ uid: u.id, exp: Date.now() + SESSION_DAYS * 864e5 }));
    return { token: token, user: publicUser_(u), config: clientConfig_() };
  },
  me: function (r) {
    var u = auth_(r);
    return { user: publicUser_(u), config: clientConfig_() };
  },
  logout: function (r) {
    PropertiesService.getScriptProperties().deleteProperty('S_' + r.token);
    return true;
  },
  changePin: function (r) {
    var u = auth_(r);
    if (hashPin_(String(r.oldPin || ''), u.salt).hash !== u.pinHash) throw new Error('BAD_LOGIN');
    if (!/^\d{4,8}$/.test(String(r.newPin || ''))) throw new Error('PIN must be 4–8 digits');
    var h = hashPin_(String(r.newPin));
    updateRow_('Users', u._row, { pinHash: h.hash, salt: h.salt });
    return true;
  },

  // ---------- worker: attendance ----------
  myStatus: function (r) {
    var u = auth_(r);
    var today = today_();
    var month = today.slice(0, 7);
    var mine = readTable_('Attendance').filter(function (a) { return a.userId === u.id; });
    var rec = mine.filter(function (a) { return a.date === today; })[0] || null;
    var days = {};
    mine.forEach(function (a) { if (a.month === month && a.checkIn) days[a.date] = 1; });
    var lastVilla = mine.length ? mine[mine.length - 1].villa : '';
    return { today: today, record: rec ? cleanAtt_(rec) : null, daysThisMonth: Object.keys(days).length, datesThisMonth: Object.keys(days),
             lastVilla: lastVilla, now: nowTime_() };
  },
  checkIn: function (r) {
    var u = auth_(r);
    return withLock_(function () {
      var today = today_();
      var existing = readTable_('Attendance').filter(function (a) { return a.userId === u.id && a.date === today; })[0];
      if (existing) throw new Error('ALREADY_IN');
      if (!r.photo) throw new Error('Photo required');
      var f = savePhoto_(r.photo, attendanceFolder_(today.slice(0, 7)), today + '_' + safe_(u.name) + '_in.jpg');
      var rec = { id: id_(), date: today, month: today.slice(0, 7), userId: u.id, worker: u.name,
                  villa: String(r.villa || ''), checkIn: nowStamp_(), checkInPhotoUrl: f.getUrl(),
                  checkInPhotoId: f.getId(), checkInGPS: gps_(r.gps) };
      appendObj_('Attendance', rec);
      return cleanAtt_(rec);
    });
  },
  checkOut: function (r) {
    var u = auth_(r);
    return withLock_(function () {
      var today = today_();
      var rec = readTable_('Attendance').filter(function (a) { return a.userId === u.id && a.date === today; })[0];
      if (!rec) throw new Error('NOT_IN');
      if (rec.checkOut) throw new Error('ALREADY_OUT');
      if (!r.photo) throw new Error('Photo required');
      var f = savePhoto_(r.photo, attendanceFolder_(rec.month), today + '_' + safe_(u.name) + '_out.jpg');
      var out = nowStamp_();
      var patch = { checkOut: out, checkOutPhotoUrl: f.getUrl(), checkOutPhotoId: f.getId(),
                    checkOutGPS: gps_(r.gps), hours: hoursBetween_(rec.checkIn, out) };
      updateRow_('Attendance', rec._row, patch);
      for (var k in patch) rec[k] = patch[k];
      return cleanAtt_(rec);
    });
  },

  // ---------- tasks ----------
  listTasks: function (r) {
    var u = auth_(r);
    var all = readTable_('Tasks');
    if (u.role === 'worker') all = all.filter(function (t) { return t.userId === u.id; });
    else if (r.userId) all = all.filter(function (t) { return t.userId === r.userId; });
    if (!r.includeDone) {
      var cutoff = Utilities.formatDate(new Date(Date.now() - 7 * 864e5), SETUP.timezone, 'yyyy-MM-dd');
      all = all.filter(function (t) { return t.status !== 'done' || String(t.doneAt).slice(0, 10) >= cutoff; });
    }
    return all.map(strip_).sort(function (a, b) {
      return (a.status === 'done') - (b.status === 'done') || String(a.due || '9999').localeCompare(String(b.due || '9999'));
    });
  },
  saveTask: function (r) {
    var u = auth_(r, ['admin', 'owner']);
    var t = r.task || {};
    if (!t.title) throw new Error('Title required');
    var w = findUserById_(t.userId);
    if (!w) throw new Error('Pick a worker');
    return withLock_(function () {
      if (t.id) {
        var ex = readTable_('Tasks').filter(function (x) { return x.id === t.id; })[0];
        if (!ex) throw new Error('Task not found');
        var patch = { userId: w.id, worker: w.name, title: t.title, notes: t.notes || '', villa: t.villa || '', due: t.due || '' };
        updateRow_('Tasks', ex._row, patch);
        return true;
      }
      appendObj_('Tasks', { id: id_(), userId: w.id, worker: w.name, title: t.title, notes: t.notes || '',
        villa: t.villa || '', due: t.due || '', status: 'open', createdBy: u.name, createdAt: nowStamp_(), doneAt: '' });
      return true;
    });
  },
  setTaskStatus: function (r) {
    var u = auth_(r);
    return withLock_(function () {
      var t = readTable_('Tasks').filter(function (x) { return x.id === r.id; })[0];
      if (!t) throw new Error('Task not found');
      if (u.role === 'worker' && t.userId !== u.id) throw new Error('Not allowed');
      var done = r.status === 'done';
      updateRow_('Tasks', t._row, { status: done ? 'done' : 'open', doneAt: done ? nowStamp_() : '' });
      return true;
    });
  },
  deleteTask: function (r) {
    auth_(r, ['admin', 'owner']);
    return withLock_(function () {
      var t = readTable_('Tasks').filter(function (x) { return x.id === r.id; })[0];
      if (t) sheet_('Tasks').deleteRow(t._row);
      return true;
    });
  },

  // ---------- expenses ----------
  scanInvoice: function (r) {
    var u = auth_(r, ['admin', 'owner']);
    if (!r.file || !r.file.data) throw new Error('No file');
    var pending = childFolder_(rootFolder_(), '_Pending scans');
    var ext = r.file.mime === 'application/pdf' ? '.pdf' : '.jpg';
    var f = pending.createFile(Utilities.newBlob(Utilities.base64Decode(r.file.data), r.file.mime,
                               'scan_' + nowStamp_().replace(/[: ]/g, '-') + '_' + safe_(u.name) + ext));
    var extracted = null, aiError = '';
    try { extracted = readInvoiceWithAI_(r.file.data, r.file.mime); }
    catch (e) { aiError = String(e.message || e); }
    return { fileId: f.getId(), extracted: extracted, aiError: aiError };
  },
  /**
   * expense: { groupId?, date, vendor, description, category, currency, invoiceNo, notes, fileId?,
   *            total, shares: [{villa, amount}, ...] }   (one share = whole invoice to one villa)
   * Each share becomes one row in Expenses (so per-villa totals stay correct); rows of the
   * same invoice share a groupId, the invoice total and the invoice file.
   */
  saveExpense: function (r) {
    var u = auth_(r, ['admin', 'owner']);
    var x = r.expense || {};
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(x.date || ''))) throw new Error('Date must be YYYY-MM-DD');
    var total = parseAmount_(x.total);
    if (!(total > 0)) throw new Error('Amount must be a positive number');
    var shares = (x.shares || []).map(function (sh) { return { villa: villaById_(sh.villa), amount: parseAmount_(sh.amount) }; });
    if (!shares.length) throw new Error('Pick at least one villa');
    var seen = {};
    shares.forEach(function (sh) {
      if (!sh.villa) throw new Error('Unknown villa');
      if (seen[sh.villa.id]) throw new Error('Villa ' + sh.villa.id + ' is listed twice');
      seen[sh.villa.id] = 1;
      if (!(sh.amount > 0)) throw new Error('Each villa\'s share must be a positive amount');
    });
    var sum = shares.reduce(function (a, sh) { return a + sh.amount; }, 0);
    if (Math.abs(sum - total) > 0.005) throw new Error('The shares add up to ' + sum + ' but the invoice total is ' + total);
    var shared = shares.length > 1;
    var splitNote = shared ? 'Shared: ' + shares.map(function (sh) { return sh.villa.id + ' ' + sh.amount; }).join(' / ') : '';
    return withLock_(function () {
      var old = [];
      if (x.groupId) {
        old = readTable_('Expenses').filter(function (e) { return (e.groupId || e.id) === x.groupId; });
        if (!old.length) throw new Error('Expense not found');
      }
      var invoices = childFolder_(rootFolder_(), 'Invoices');
      var folder = shared ? childFolder_(invoices, 'Shared between villas') : childFolder_(invoices, villaFolderName_(shares[0].villa));
      var fname = x.date + '_' + shares.map(function (sh) { return sh.villa.id; }).join('-') + '_' + safe_(x.vendor || 'invoice') + '_' + total;
      var fileId = x.fileId || (old[0] && old[0].fileId) || '';
      var file = null;
      if (fileId) {
        file = DriveApp.getFileById(fileId);
        file.moveTo(folder);
        file.setName(fname + (file.getMimeType() === 'application/pdf' ? '.pdf' : '.jpg'));
      } else if (r.file && r.file.data) {
        file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(r.file.data), r.file.mime,
                                 fname + (r.file.mime === 'application/pdf' ? '.pdf' : '.jpg')));
      }
      // replace the old rows of this invoice (bottom-up so row numbers stay valid)
      old.sort(function (a, b) { return b._row - a._row; }).forEach(function (e) { sheet_('Expenses').deleteRow(e._row); });
      var groupId = x.groupId || id_();
      var uploadedBy = old[0] ? old[0].uploadedBy : u.name;
      var createdAt = old[0] ? old[0].createdAt : nowStamp_();
      shares.forEach(function (sh) {
        appendObj_('Expenses', {
          id: id_(), groupId: groupId, date: x.date, month: x.date.slice(0, 7), villa: sh.villa.id, owner: sh.villa.owner,
          vendor: x.vendor || '', description: x.description || '', category: x.category || 'Other',
          amount: sh.amount, invoiceTotal: total, split: splitNote, currency: x.currency || SETUP.currency,
          invoiceNo: x.invoiceNo || '', notes: x.notes || '', fileUrl: file ? file.getUrl() : '', fileId: file ? file.getId() : '',
          uploadedBy: uploadedBy, createdAt: createdAt });
      });
      return { groupId: groupId };
    });
  },
  listExpenses: function (r) {
    var u = auth_(r, ['admin', 'owner']);
    var rows = readTable_('Expenses').filter(function (e) {
      return canSeeVilla_(u, e.villa) && (!r.month || e.month === r.month) && (!r.villa || e.villa === r.villa);
    }).map(function (e) { var o = strip_(e); o.groupId = o.groupId || o.id; o.invoiceTotal = o.invoiceTotal === '' ? o.amount : o.invoiceTotal; return o; });
    rows.sort(function (a, b) { return String(b.date).localeCompare(String(a.date)) || String(b.createdAt).localeCompare(String(a.createdAt)); });
    return rows;
  },
  getExpenseGroup: function (r) {
    auth_(r, ['admin', 'owner']);
    return readTable_('Expenses').filter(function (e) { return (e.groupId || e.id) === r.groupId; })
      .map(function (e) { var o = strip_(e); o.groupId = o.groupId || o.id; return o; });
  },
  deleteExpense: function (r) {
    var u = auth_(r, ['admin', 'owner']);
    return withLock_(function () {
      var rows = readTable_('Expenses').filter(function (e) { return (e.groupId || e.id) === r.groupId; });
      if (!rows.length) return true;
      if (rows[0].fileId) { try { DriveApp.getFileById(rows[0].fileId).setTrashed(true); } catch (e) {} }
      rows.sort(function (a, b) { return b._row - a._row; }).forEach(function (e) { sheet_('Expenses').deleteRow(e._row); });
      return true;
    });
  },

  // ---------- attendance reports ----------
  monthlyReport: function (r) {
    auth_(r, ['admin', 'owner']);
    var month = r.month || today_().slice(0, 7);
    var users = readTable_('Users');
    var att = readTable_('Attendance').filter(function (a) { return a.month === month; });
    var byUser = {};
    att.forEach(function (a) { (byUser[a.userId] = byUser[a.userId] || []).push(cleanAtt_(a)); });
    var today = today_();
    var out = users.filter(function (u) { return u.role === 'worker' && (isTrue_(u.active) || byUser[u.id]); })
      .map(function (u) {
        var recs = (byUser[u.id] || []).sort(function (a, b) { return a.date.localeCompare(b.date); });
        var days = {}, hours = 0, missing = 0;
        recs.forEach(function (a) {
          if (a.checkIn) days[a.date] = 1;
          hours += Number(a.hours || 0);
          if (a.checkIn && !a.checkOut && a.date < today) missing++;
        });
        var n = Object.keys(days).length;
        var p = payslip_(u, n);
        return { userId: u.id, name: u.name, days: n, hours: Math.round(hours * 100) / 100,
                 missingCheckOut: missing, dailyRate: p.dailyRate, mealAllowance: p.mealAllowance, monthlyBonus: p.monthlyBonus,
                 deductionPct: p.deductionPct, wage: p.wage, meal: p.meal, bonus: p.bonus, gross: p.gross,
                 deduction: p.deduction, net: p.net, pay: p.net, records: recs };
      });
    return { month: month, workers: out };
  },
  addManualDay: function (r) {
    var u = auth_(r, ['admin', 'owner']);
    var w = findUserById_(r.userId);
    if (!w) throw new Error('Worker not found');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(r.date || ''))) throw new Error('Date must be YYYY-MM-DD');
    return withLock_(function () {
      var ex = readTable_('Attendance').filter(function (a) { return a.userId === w.id && a.date === r.date; })[0];
      if (ex) throw new Error('That day is already recorded');
      appendObj_('Attendance', { id: id_(), date: r.date, month: r.date.slice(0, 7), userId: w.id, worker: w.name,
        villa: r.villa || '', checkIn: r.date + ' ' + (r.inTime || '08:00') + ':00',
        checkOut: r.outTime ? r.date + ' ' + r.outTime + ':00' : '',
        hours: r.outTime ? hoursBetween_(r.date + ' ' + (r.inTime || '08:00') + ':00', r.date + ' ' + r.outTime + ':00') : '',
        note: 'Added manually by ' + u.name + (r.note ? ': ' + r.note : '') });
      return true;
    });
  },
  setCheckOut: function (r) {
    var u = auth_(r, ['admin', 'owner']);
    if (!/^\d{2}:\d{2}$/.test(String(r.time || ''))) throw new Error('Time must be HH:MM');
    return withLock_(function () {
      var a = readTable_('Attendance').filter(function (x) { return x.id === r.id; })[0];
      if (!a) throw new Error('Record not found');
      var out = a.date + ' ' + r.time + ':00';
      updateRow_('Attendance', a._row, { checkOut: out, hours: hoursBetween_(a.checkIn, out),
        note: (a.note ? a.note + '; ' : '') + 'Check-out set by ' + u.name });
      return true;
    });
  },
  deleteAttendance: function (r) {
    auth_(r, ['admin']);
    return withLock_(function () {
      var a = readTable_('Attendance').filter(function (x) { return x.id === r.id; })[0];
      if (a) sheet_('Attendance').deleteRow(a._row);
      return true;
    });
  },

  // ---------- photos (served through the API so Drive files stay private) ----------
  getFile: function (r) {
    var u = auth_(r);
    var id = String(r.fileId || '');
    var allowed = false;
    if (u.role === 'worker') {
      allowed = readTable_('Attendance').some(function (a) {
        return a.userId === u.id && (a.checkInPhotoId === id || a.checkOutPhotoId === id); });
    } else {
      allowed = readTable_('Expenses').some(function (e) { return e.fileId === id && canSeeVilla_(u, e.villa); }) ||
                readTable_('Attendance').some(function (a) { return a.checkInPhotoId === id || a.checkOutPhotoId === id; });
    }
    if (!allowed) throw new Error('Not allowed');
    var f = DriveApp.getFileById(id);
    var blob = f.getBlob();
    if (blob.getContentType() === 'application/pdf' || blob.getBytes().length > 6e6) return { url: f.getUrl(), mime: blob.getContentType() };
    return { mime: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) };
  },

  // ---------- team (admin) ----------
  listUsers: function (r) {
    var u = auth_(r, ['admin', 'owner']);
    return readTable_('Users').filter(function (x) { return u.role === 'admin' || x.role === 'worker'; }).map(publicUser_);
  },
  saveUser: function (r) {
    auth_(r, ['admin']);
    var p = r.user || {};
    if (!p.name) throw new Error('Name required');
    if (['admin', 'owner', 'worker'].indexOf(p.role) < 0) throw new Error('Bad role');
    if (p.pin && !/^\d{4,8}$/.test(String(p.pin))) throw new Error('PIN must be 4–8 digits');
    return withLock_(function () {
      var dup = findUserByName_(p.name);
      if (dup && dup.id !== p.id) throw new Error('Someone already has that name');
      if (p.id) {
        var ex = findUserById_(p.id);
        if (!ex) throw new Error('User not found');
        var patch = { name: p.name, role: p.role, villas: p.villas || '', dailyRate: p.dailyRate || '',
                      mealAllowance: p.mealAllowance || '', monthlyBonus: p.monthlyBonus || '', deductionPct: p.deductionPct || '',
                      active: p.active === false ? 'FALSE' : 'TRUE' };
        if (p.pin) { var h = hashPin_(String(p.pin)); patch.pinHash = h.hash; patch.salt = h.salt; }
        updateRow_('Users', ex._row, patch);
        return true;
      }
      if (!p.pin) throw new Error('Set a PIN for the new person');
      createUser_(p);
      return true;
    });
  }
};

// ============================== SHEET LAYOUT (readable for owners) ==============================
/** Column titles people see in row 1. The app finds data by column position, so never move columns. */
var LABELS = {
  Users:      { id: 'ID', name: 'Name', role: 'Role', villas: 'Villas owned', dailyRate: 'Daily wage', active: 'Active',
                pinHash: 'PIN (scrambled)', salt: 'PIN key', createdAt: 'Added on', mealAllowance: 'Daily meal allowance',
                monthlyBonus: 'Monthly bonus', deductionPct: 'Deduction %' },
  Expenses:   { id: 'ID', date: 'Date', month: 'Month', villa: 'Villa', owner: 'Owner', vendor: 'Vendor / shop',
                description: 'What was bought', category: 'Category', amount: 'Villa share', currency: 'Currency',
                invoiceNo: 'Invoice no.', fileUrl: 'Invoice file', fileId: 'File ID', uploadedBy: 'Added by',
                createdAt: 'Added on', notes: 'Notes', groupId: 'Invoice group', invoiceTotal: 'Invoice total', split: 'Split between villas' },
  Attendance: { id: 'ID', date: 'Date', month: 'Month', userId: 'Worker ID', worker: 'Worker', villa: 'Villa',
                checkIn: 'Check-in', checkInPhotoUrl: 'Check-in photo', checkInPhotoId: 'Check-in photo ID', checkInGPS: 'Check-in location',
                checkOut: 'Check-out', checkOutPhotoUrl: 'Check-out photo', checkOutPhotoId: 'Check-out photo ID',
                checkOutGPS: 'Check-out location', hours: 'Hours', note: 'Note' },
  Tasks:      { id: 'ID', userId: 'Worker ID', worker: 'Worker', title: 'Task', notes: 'Notes', villa: 'Villa', due: 'Due',
                status: 'Status', createdBy: 'Given by', createdAt: 'Given on', doneAt: 'Done on' }
};
var HIDDEN = {
  Expenses:   ['id', 'month', 'currency', 'fileId', 'groupId'],
  Attendance: ['id', 'month', 'userId', 'villa', 'checkInPhotoId', 'checkOutPhotoId'],
  Tasks:      ['id', 'userId']
};
var WIDTHS = { date: 95, villa: 60, owner: 80, vendor: 190, description: 260, category: 150, amount: 130, invoiceTotal: 130,
  invoiceNo: 110, fileUrl: 220, uploadedBy: 90, createdAt: 140, notes: 240, split: 260, worker: 100, checkIn: 140, checkOut: 140,
  checkInPhotoUrl: 200, checkOutPhotoUrl: 200, checkInGPS: 170, checkOutGPS: 170, hours: 70, note: 220, title: 260, due: 95,
  status: 80, createdBy: 90, doneAt: 140 };
var DATE_COLS = { Expenses: ['date'], Attendance: ['date'], Tasks: ['due'] };
var MONEY = '"Rp "#,##0';
var VILLA_COLORS = ['#4fb7aa', '#f0b43c', '#d9668a', '#7b8fe0'];
var INK = '#1d2b2d', OCEAN = '#0b4f57', PALE = '#eef3f1';

function label_(sheetName, key) { return (LABELS[sheetName] && LABELS[sheetName][key]) || key; }
function colLetter_(sheetName, key) {
  var n = COLS[sheetName].indexOf(key) + 1, s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/**
 * Makes the Google Sheet easy to read: an Overview tab with totals and a chart, clear column titles,
 * Rupiah amounts (1.250.000), dd/mm/yyyy dates, hidden technical columns, filters and striped rows.
 * Safe to run again at any time — it only changes the look, never the data.
 */
function formatSheets() {
  var ss = ss_();
  ss.setSpreadsheetLocale('en_US');            // write formulas in one known syntax…
  ss.setSpreadsheetTimeZone(SETUP.timezone);

  Object.keys(COLS).forEach(function (name) {
    var sh = ss.getSheetByName(name), cols = COLS[name];
    var maxRows = sh.getMaxRows(), last = sh.getLastRow();
    sh.getRange(1, 1, 1, cols.length).setValues([cols.map(function (c) { return label_(name, c); })])
      .setFontWeight('bold').setBackground(OCEAN).setFontColor('#ffffff').setVerticalAlignment('middle').setWrap(true);
    sh.setRowHeight(1, 36);
    sh.setFrozenRows(1);
    (sh.getBandings() || []).forEach(function (b) { b.remove(); });
    if (sh.getFilter()) sh.getFilter().remove();
    var all = sh.getRange(1, 1, maxRows, cols.length);
    all.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
    sh.getRange(1, 1, maxRows, cols.length).createFilter();
    sh.getRange(2, 1, maxRows - 1, cols.length).setVerticalAlignment('top');
    cols.forEach(function (c, i) {
      if (WIDTHS[c]) sh.setColumnWidth(i + 1, WIDTHS[c]);
      if (c === 'amount' || c === 'invoiceTotal' || c === 'dailyRate' || c === 'mealAllowance' || c === 'monthlyBonus')
        sh.getRange(2, i + 1, maxRows - 1, 1).setNumberFormat(MONEY);
      if (c === 'hours') sh.getRange(2, i + 1, maxRows - 1, 1).setNumberFormat('0.0');
      if (c === 'description' || c === 'notes' || c === 'note' || c === 'split' || c === 'title')
        sh.getRange(2, i + 1, maxRows - 1, 1).setWrap(true);
      if ((HIDDEN[name] || []).indexOf(c) >= 0) sh.hideColumns(i + 1); else sh.showColumns(i + 1);
    });
    // real dates, shown as dd/mm/yyyy (existing text dates are converted)
    (DATE_COLS[name] || []).forEach(function (c) {
      var ci = cols.indexOf(c) + 1;
      var r = sh.getRange(2, ci, maxRows - 1, 1);
      var vals = last > 1 ? sh.getRange(2, ci, last - 1, 1).getValues() : [];
      r.setNumberFormat('dd/mm/yyyy');
      if (vals.length) sh.getRange(2, ci, vals.length, 1).setValues(vals.map(function (v) {
        return [v[0] instanceof Date ? v[0] : String(v[0] || '')]; }));
    });
  });
  var tabColors = { Expenses: '#0e6670', Attendance: '#f0b43c', Tasks: '#7b8fe0' };
  Object.keys(tabColors).forEach(function (n) { ss.getSheetByName(n).setTabColor(tabColors[n]); });
  ss.getSheetByName('Users').hideSheet();
  var old = ss.getSheetByName('Summary'); if (old) ss.deleteSheet(old);

  buildOverview_(ss);
  ss.setSpreadsheetLocale('id_ID');            // …then show numbers as 1.250.000 and dates as dd/mm/yyyy
  Logger.log('Done. Open the Sheet: ' + ss.getUrl());
}

function buildOverview_(ss) {
  var old = ss.getSheetByName('Overview'); if (old) ss.deleteSheet(old);
  var sh = ss.insertSheet('Overview', 0);
  ss.setActiveSheet(sh); ss.moveActiveSheet(1);
  sh.setTabColor(OCEAN);
  sh.setHiddenGridlines(true);
  var V = SETUP.villas, nV = V.length;
  var E = function (key) { return 'Expenses!$' + colLetter_('Expenses', key) + '$2:$' + colLetter_('Expenses', key); };
  var AMT = E('amount'), VIL = E('villa'), MON = E('month'), CAT = E('category'), OWN = E('owner');
  // exact text matching (TO_TEXT) so ids like 131 and months like 2026-11 are never misread as numbers or dates
  var sumWhere = function (conds) { return '=IFERROR(SUM(FILTER(' + AMT + ',' + conds.join(',') + ')),0)'; };
  var isVilla = function (id) { return 'TO_TEXT(' + VIL + ')="' + id + '"'; };
  var isMonth = function (cell) { return 'TO_TEXT(' + MON + ')=TO_TEXT(' + cell + ')'; };
  var isOwner = function (o) { return 'TO_TEXT(' + OWN + ')="' + o + '"'; };
  var lastCol = 2 + nV;                                  // A label, B..(villas), then total
  sh.setColumnWidth(1, 210);
  for (var c = 2; c <= lastCol + 1; c++) sh.setColumnWidth(c, 130);

  var row = 1;
  sh.getRange(row, 1).setValue(SETUP.appName + ' — Overview').setFontSize(20).setFontWeight('bold').setFontColor(OCEAN);
  sh.setRowHeight(row, 40); row++;
  sh.getRange(row, 1).setValue('Updates automatically from the Expenses tab. Amounts in Rupiah. Salaries appear as "Staff salaries" (category Labour).')
    .setFontColor('#5f6f70'); row += 2;

  // month picker
  var MONTH_CELL = 'B' + row;
  sh.getRange(row, 1).setValue('Month shown').setFontWeight('bold');
  sh.getRange(row, 2).setNumberFormat('@').setFormula('=TEXT(TODAY(),"yyyy-mm")').setBackground('#fbecc8').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(row, 3, 1, 3).merge().setValue('← type another month here, e.g. 2026-10').setFontColor('#5f6f70').setFontStyle('italic');
  row += 2;

  var section = function (title) {
    sh.getRange(row, 1, 1, lastCol + 1).merge().setValue(title).setFontSize(13).setFontWeight('bold').setFontColor(INK)
      .setBorder(null, null, true, null, null, null, OCEAN, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sh.setRowHeight(row, 30); row++;
  };
  var header = function (cells) {
    var r = sh.getRange(row, 1, 1, cells.length);
    r.setValues([cells]).setFontWeight('bold').setBackground(PALE).setHorizontalAlignment('center');
    sh.getRange(row, 1).setHorizontalAlignment('left'); row++;
  };
  var villaHeaderColors = function (r0, c0) {
    V.forEach(function (v, i) { sh.getRange(r0, c0 + i).setBackground(VILLA_COLORS[i % 4]).setFontColor(i === 1 ? INK : '#ffffff'); });
  };

  // 1. per villa: selected month and since start
  section('1. Spending per villa');
  header(['Villa', 'Owner', 'Selected month', 'Since the start']);
  var first = row;
  V.forEach(function (v, i) {
    sh.getRange(row, 1).setValue('Villa ' + v.id).setFontWeight('bold');
    sh.getRange(row, 1).setBackground(VILLA_COLORS[i % 4]).setFontColor(i === 1 ? INK : '#ffffff');
    sh.getRange(row, 2).setValue(v.owner);
    sh.getRange(row, 3).setFormula(sumWhere([isVilla(v.id), isMonth(MONTH_CELL)]));
    sh.getRange(row, 4).setFormula(sumWhere([isVilla(v.id)]));
    row++;
  });
  sh.getRange(row, 1).setValue('All villas');
  sh.getRange(row, 3).setFormula('=SUM(C' + first + ':C' + (row - 1) + ')');
  sh.getRange(row, 4).setFormula('=SUM(D' + first + ':D' + (row - 1) + ')');
  sh.getRange(row, 1, 1, 4).setFontWeight('bold').setBorder(true, null, null, null, null, null, INK, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(first, 3, row - first + 1, 2).setNumberFormat(MONEY);
  row += 2;

  // 2. per owner
  section('2. Spending per owner');
  header(['Owner', 'Villas', 'Selected month', 'Since the start']);
  var owners = [];
  V.forEach(function (v) { if (owners.indexOf(v.owner) < 0) owners.push(v.owner); });
  first = row;
  owners.forEach(function (o) {
    var ids = V.filter(function (v) { return v.owner === o; }).map(function (v) { return v.id; });
    sh.getRange(row, 1).setValue(o).setFontWeight('bold');
    sh.getRange(row, 2).setValue(ids.join(' + '));
    sh.getRange(row, 3).setFormula(sumWhere([isOwner(o), isMonth(MONTH_CELL)]));
    sh.getRange(row, 4).setFormula(sumWhere([isOwner(o)]));
    row++;
  });
  sh.getRange(first, 3, row - first, 2).setNumberFormat(MONEY);
  row++;

  // 3. by category for selected month
  section('3. Selected month by category');
  header(['Category'].concat(V.map(function (v) { return 'Villa ' + v.id; })).concat(['Total']));
  villaHeaderColors(row - 1, 2);
  first = row;
  SETUP.categories.forEach(function (cat) {
    sh.getRange(row, 1).setValue(cat);
    V.forEach(function (v, i) {
      sh.getRange(row, 2 + i).setFormula(sumWhere([isVilla(v.id), 'TO_TEXT(' + CAT + ')=$A' + row, isMonth(MONTH_CELL)]));
    });
    sh.getRange(row, 2 + nV).setFormula('=SUM(B' + row + ':' + colA1_(1 + nV) + row + ')').setFontWeight('bold');
    row++;
  });
  sh.getRange(row, 1).setValue('Total');
  for (var k = 0; k <= nV; k++) sh.getRange(row, 2 + k).setFormula('=SUM(' + colA1_(2 + k) + first + ':' + colA1_(2 + k) + (row - 1) + ')');
  sh.getRange(row, 1, 1, nV + 2).setFontWeight('bold').setBorder(true, null, null, null, null, null, INK, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(first, 2, row - first + 1, nV + 1).setNumberFormat(MONEY);
  // hide zero rows visually: grey out zeros
  sh.getRange(first, 2, row - first, nV + 1).setFontColor(INK);
  row += 2;

  // 4. month by month (+ chart)
  section('4. Month by month');
  header(['Month'].concat(V.map(function (v) { return 'Villa ' + v.id; })).concat(['Total']));
  villaHeaderColors(row - 1, 2);
  var hdrRow = row - 1;
  first = row;
  var span = 36;                                         // three years of months
  sh.getRange(first, 1, span, 1).setNumberFormat('@');
  sh.getRange(row, 1).setFormula('=IFERROR(SORT(UNIQUE(FILTER(TO_TEXT(' + MON + '),TO_TEXT(' + MON + ')<>""))),"No expenses yet")');

  var monthRows = [];
  for (var r = first; r < first + span; r++) {
    var cells = V.map(function (v) { return '=IF($A' + r + '="","",' + sumWhere([isVilla(v.id), isMonth('$A' + r)]).slice(1) + ')'; });
    cells.push('=IF($A' + r + '="","",' + sumWhere([isMonth('$A' + r)]).slice(1) + ')');
    monthRows.push(cells);
  }
  sh.getRange(first, 2, span, nV + 1).setFormulas(monthRows);
  sh.getRange(first, 2, span, nV + 1).setNumberFormat(MONEY);
  sh.getRange(first, 2 + nV, span, 1).setFontWeight('bold');

  var chart = sh.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange(hdrRow, 1, span + 1, nV + 1))
    .setNumHeaders(1)
    .setOption('isStacked', true)
    .setOption('title', 'Spending per month, per villa')
    .setOption('legend', { position: 'bottom' })
    .setOption('colors', VILLA_COLORS.slice(0, nV))
    .setOption('vAxis', { format: 'short' })
    .setOption('width', 620).setOption('height', 340)
    .setPosition(hdrRow, nV + 4, 0, 0)
    .build();
  sh.insertChart(chart);
  sh.setFrozenRows(0);
}
function colA1_(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

// ============================== AI INVOICE READING ==============================
function readInvoiceWithAI_(b64, mime) {
  var props = PropertiesService.getScriptProperties();
  var claudeKey = props.getProperty('ANTHROPIC_API_KEY');
  var geminiKey = props.getProperty('GEMINI_API_KEY');
  if (!claudeKey && !geminiKey) throw new Error('NO_AI_KEY');
  var prompt = [
    'You are reading a purchase invoice or receipt for a villa business in Bali, Indonesia.',
    'It may be an Indonesian "nota" or "kwitansi", possibly handwritten. Indonesian numbers use "." as the thousands',
    'separator and "," for decimals: "Rp 1.250.000" means 1250000. Today is ' + today_() + '.',
    'Return ONLY a JSON object with these keys:',
    '{"date": "YYYY-MM-DD or null", "vendor": "shop/company name", "description": "short English summary of what was bought",',
    ' "category": one of ' + JSON.stringify(SETUP.categories) + ',',
    ' "amount": final total to pay as a plain number (no separators), "currency": "ISO code, default IDR",',
    ' "invoiceNo": "invoice/nota number or null", "notes": "anything unclear or unreadable, else empty"}'
  ].join('\n');
  var text;
  if (claudeKey) {
    var part = mime === 'application/pdf'
      ? { type: 'document', source: { type: 'base64', media_type: mime, data: b64 } }
      : { type: 'image', source: { type: 'base64', media_type: mime, data: b64 } };
    var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': claudeKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: props.getProperty('CLAUDE_MODEL') || CLAUDE_MODEL_DEFAULT, max_tokens: 1024,
        messages: [{ role: 'user', content: [part, { type: 'text', text: prompt }] }] })
    });
    if (res.getResponseCode() !== 200) throw new Error('Claude API error ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 300));
    var body = JSON.parse(res.getContentText());
    text = (body.content || []).map(function (c) { return c.text || ''; }).join('');
  } else {
    var model = props.getProperty('GEMINI_MODEL') || GEMINI_MODEL_DEFAULT;
    var gres = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + geminiKey, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ contents: [{ parts: [{ inline_data: { mime_type: mime, data: b64 } }, { text: prompt }] }],
        generationConfig: { responseMimeType: 'application/json' } })
    });
    if (gres.getResponseCode() !== 200) throw new Error('Gemini API error ' + gres.getResponseCode() + ': ' + gres.getContentText().slice(0, 300));
    var gbody = JSON.parse(gres.getContentText());
    text = gbody.candidates[0].content.parts.map(function (p) { return p.text || ''; }).join('');
  }
  var m = text.match(/\{[\s\S]*\}/);
  if (!m) throw new Error('Could not read the invoice');
  var o = JSON.parse(m[0]);
  if (o.amount !== null && o.amount !== undefined) o.amount = parseAmount_(o.amount) || '';
  if (SETUP.categories.indexOf(o.category) < 0) o.category = 'Other';
  return o;
}

// ============================== HELPERS ==============================
/**
 * Monthly pay for one worker (contract: daily wage + daily meal allowance per day worked,
 * a fixed monthly bonus if they worked at least one day, minus a % of everything paid).
 * gross = total cost to the villas; net = what the worker receives. Rounded to the nearest rupiah.
 */
function payslip_(u, days) {
  var rate = Number(u.dailyRate || 0), meal = Number(u.mealAllowance || 0), bonus = Number(u.monthlyBonus || 0),
      pct = Number(u.deductionPct || 0);
  var wage = days * rate, mealTotal = days * meal, bonusPaid = days > 0 ? bonus : 0;
  var gross = wage + mealTotal + bonusPaid;
  var deduction = Math.round(gross * pct / 100);
  return { dailyRate: rate, mealAllowance: meal, monthlyBonus: bonus, deductionPct: pct,
           wage: wage, meal: mealTotal, bonus: bonusPaid, gross: gross, deduction: deduction, net: gross - deduction };
}

/** Adds any new columns to a Sheet created by an older version of this script (runs once). */
function ensureSchema_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('SCHEMA_VERSION') === SCHEMA_VERSION) return;
  withLock_(function () {
    Object.keys(COLS).forEach(function (name) {
      var sh = sheet_(name), cols = COLS[name];
      var have = sh.getRange(1, 1, 1, cols.length).getValues()[0];
      for (var i = 0; i < cols.length; i++) {
        if (have[i] === cols[i] || have[i] === label_(name, cols[i])) continue;
        if (have[i] !== '' && have[i] !== undefined) throw new Error('Sheet "' + name + '" column ' + (i + 1) + ' is "' + have[i] + '", expected "' + cols[i] + '". Columns must not be moved.');
        sh.getRange(1, i + 1).setValue(label_(name, cols[i])).setFontWeight('bold').setBackground(OCEAN).setFontColor('#ffffff');
        sh.getRange(2, i + 1, sh.getMaxRows() - 1, 1).setNumberFormat(NUMERIC[cols[i]] || '@');
      }
    });
    props.setProperty('SCHEMA_VERSION', SCHEMA_VERSION);
  });
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function ss_() { return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID')); }
function sheet_(name) { return ss_().getSheetByName(name); }
function rootFolder_() { return DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID')); }
function childFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function attendanceFolder_(month) { return childFolder_(childFolder_(rootFolder_(), 'Attendance photos'), month); }
function villaFolderName_(v) { return 'Villa ' + v.id + ' – ' + v.owner; }
function villaById_(id) { return SETUP.villas.filter(function (v) { return v.id === String(id); })[0]; }
function id_() { return Utilities.getUuid().slice(0, 8); }
function today_() { return Utilities.formatDate(new Date(), SETUP.timezone, 'yyyy-MM-dd'); }
function nowStamp_() { return Utilities.formatDate(new Date(), SETUP.timezone, 'yyyy-MM-dd HH:mm:ss'); }
function nowTime_() { return Utilities.formatDate(new Date(), SETUP.timezone, 'HH:mm'); }
function safe_(s) { return String(s).replace(/[^\w\-]+/g, '').slice(0, 40) || 'x'; }
function pad_(s, n) { s = String(s); while (s.length < n) s += ' '; return s; }
function isTrue_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function randomPin_() { return String(Math.floor(100000 + Math.random() * 900000)); }
function gps_(g) { return g && g.lat ? Number(g.lat).toFixed(6) + ',' + Number(g.lng).toFixed(6) + (g.acc ? ' ±' + Math.round(g.acc) + 'm' : '') : ''; }
function hoursBetween_(a, b) {
  var t = function (s) { var m = String(s).match(/(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})(?::(\d{2}))?/); return m ? Date.UTC(+m[1], m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0)) : NaN; };
  var h = (t(b) - t(a)) / 36e5;
  return isNaN(h) ? '' : Math.round(h * 100) / 100;
}
/** "1.250.000" → 1250000, "1,250,000" → 1250000, "12.5" → 12.5, "Rp 75.000,50" → 75000.5. Always positive. */
function parseAmount_(v) {
  if (typeof v === 'number') return Math.abs(v);
  var s = String(v || '').replace(/[^\d.,]/g, '');
  if (!s) return NaN;
  if (/^\d{1,3}([.,]\d{3})+$/.test(s)) return Number(s.replace(/[.,]/g, ''));
  var m = s.match(/^(.*)[.,](\d{1,2})$/);
  if (m) return Number(m[1].replace(/[.,]/g, '') + '.' + m[2]);
  return Number(s.replace(/[.,]/g, ''));
}
function savePhoto_(b64, folder, name) {
  return folder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', name));
}
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function hashPin_(pin, salt) {
  salt = salt || Utilities.getUuid();
  var h = salt + ':' + pin;
  for (var i = 0; i < 300; i++) {
    h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt)
      .map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
  }
  return { hash: h, salt: salt };
}
function createUser_(p) {
  var h = hashPin_(String(p.pin));
  appendObj_('Users', { id: id_(), name: String(p.name).trim(), role: p.role,
    villas: p.role === 'admin' ? 'all' : (p.villas || ''), dailyRate: p.dailyRate || '',
    mealAllowance: p.mealAllowance || '', monthlyBonus: p.monthlyBonus || '', deductionPct: p.deductionPct || '',
    active: 'TRUE', pinHash: h.hash, salt: h.salt, createdAt: nowStamp_() });
}
function findUserByName_(name) {
  var n = String(name || '').trim().toLowerCase();
  return readTable_('Users').filter(function (u) { return String(u.name).trim().toLowerCase() === n; })[0];
}
function findUserById_(id) { return readTable_('Users').filter(function (u) { return u.id === id; })[0]; }
function publicUser_(u) {
  var num = function (v) { return v === '' || v === undefined ? '' : Number(v); };
  return { id: u.id, name: u.name, role: u.role, villas: String(u.villas || ''), dailyRate: num(u.dailyRate),
           mealAllowance: num(u.mealAllowance), monthlyBonus: num(u.monthlyBonus), deductionPct: num(u.deductionPct),
           active: isTrue_(u.active) };
}
function userVillas_(u) {
  if (u.role === 'admin' || String(u.villas).toLowerCase() === 'all') return SETUP.villas.map(function (v) { return v.id; });
  return String(u.villas || '').split(/[,\s]+/).filter(String);
}
// Owners and admins see and add expenses for every villa (shared books). Workers never reach expense actions.
function canSeeVilla_(u, villaId) { return (u.role === 'admin' || u.role === 'owner') && !!villaById_(villaId); }
function clientConfig_() { return { villas: SETUP.villas, categories: SETUP.categories, currency: SETUP.currency, appName: SETUP.appName, timezone: SETUP.timezone }; }
function auth_(r, roles) {
  var raw = r.token && PropertiesService.getScriptProperties().getProperty('S_' + r.token);
  var s = raw && JSON.parse(raw);
  if (!s || s.exp < Date.now()) throw new Error('SESSION_EXPIRED');
  var u = findUserById_(s.uid);
  if (!u || !isTrue_(u.active)) throw new Error('SESSION_EXPIRED');
  u.villaList = userVillas_(u);
  if (roles && roles.indexOf(u.role) < 0) throw new Error('Not allowed');
  return u;
}
function purgeSessions_() {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties(), now = Date.now();
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('S_') === 0) { try { if (JSON.parse(all[k]).exp < now) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); } }
  });
}
function cleanAtt_(a) {
  var o = strip_(a);
  return o;
}
function strip_(o) {
  var c = {};
  for (var k in o) if (k !== '_row' && k !== 'pinHash' && k !== 'salt') c[k] = o[k];
  return c;
}
function readTable_(name) {
  var sh = sheet_(name);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var cols = COLS[name];
  var vals = sh.getRange(2, 1, last - 1, cols.length).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    if (vals[i].join('') === '') continue;
    var o = { _row: i + 2 };
    for (var j = 0; j < cols.length; j++) {
      var v = vals[i][j];
      var c = cols[j];
      if (v instanceof Date) v = Utilities.formatDate(v, SETUP.timezone,
        c === 'date' || c === 'due' ? 'yyyy-MM-dd' : c === 'month' ? 'yyyy-MM' : 'yyyy-MM-dd HH:mm:ss');
      o[c] = (NUMERIC[c] && typeof v === 'number') ? v : String(v);
    }
    out.push(o);
  }
  return out;
}
function appendObj_(name, obj) {
  var cols = COLS[name];
  sheet_(name).appendRow(cols.map(function (c) { return obj[c] === undefined ? '' : obj[c]; }));
}
function updateRow_(name, row, patch) {
  var cols = COLS[name];
  var sh = sheet_(name);
  var range = sh.getRange(row, 1, 1, cols.length);
  var vals = range.getValues()[0];
  cols.forEach(function (c, i) { if (patch[c] !== undefined) vals[i] = patch[c]; });
  range.setValues([vals]);
}
