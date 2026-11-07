/**
 * _Interlink_Masters.gs
 * -----------------------------------------------------------------------------
 * Thin wrapper around the Masters library (added as `Masters` dependency in
 * this script's Libraries). Caches reads in CacheService for 5 min so repeated
 * dropdown reads don't thrash the library.
 *
 * Paste this file verbatim into Finance / IMS / PM. It's self-contained —
 * no touches to any existing production code.
 *
 * One-time setup per system (do once after pasting this file):
 *   1. Script editor → Libraries → Add library
 *      Script ID: 1zWL2xvoZIZ1HAk-RrngKIeA4Ps4fdByYW9_hsMEm8lMPA2P8rOgTzSxC
 *      Version:   1   (or latest)
 *      Identifier: Masters
 *   2. Run IM_installInterlinkMenu() once from the editor — adds the
 *      🔗 Interlink menu to the sheet (via installable onOpen trigger, so
 *      the system's existing onOpen stays untouched).
 */

var IM_VERSION    = '0.4.0';
var IM_CACHE_TTL  = 300;   // 5 min
var IM_CACHE_KEY  = 'IM_v' + IM_VERSION + '_';

// -----------------------------------------------------------------------------
// Reads (cached)
// -----------------------------------------------------------------------------
function IM_getProjects()  { return IM_cache_('projects',  function () { return Masters.getProjects(); }); }
function IM_getCustomers() { return IM_cache_('customers', function () { return Masters.getCustomers(); }); }
function IM_getEmployees() { return IM_cache_('employees', function () { return Masters.getEmployees(); }); }
function IM_getSites()     { return IM_cache_('sites',     function () { return Masters.getSites(); }); }

/**
 * Parties with role flags. Call with a flag object to narrow:
 *   IM_getParties({ isSupplier:  true })  → suppliers only
 *   IM_getParties({ isJobWorker: true })  → job workers only
 *   IM_getParties({})                     → all parties
 */
function IM_getParties(filter) {
  var key = 'parties_' + JSON.stringify(filter || {});
  return IM_cache_(key, function () { return Masters.getParties(filter || {}); });
}

/** Freeform name → Masters record. For migration-time resolution. */
function IM_resolveName(entity, text) {
  return Masters.resolveFreeformName(entity, text);
}

// -----------------------------------------------------------------------------
// Dropdown-friendly label lists — "{ID} — {Name}"
// -----------------------------------------------------------------------------
function IM_projectLabels()  { return IM_asLabels_(IM_getProjects(),  'Project_ID',  'Name'); }
function IM_customerLabels() { return IM_asLabels_(IM_getCustomers(), 'Customer_ID', 'Name'); }
function IM_partyLabels(f)   { return IM_asLabels_(IM_getParties(f),  'Party_ID',    'Name'); }
function IM_employeeLabels() { return IM_asLabels_(IM_getEmployees(), 'Emp_ID',      'Name'); }
function IM_siteLabels()     { return IM_asLabels_(IM_getSites(),     'Site_ID',     'Name'); }

function IM_asLabels_(rows, idCol, nameCol) {
  return (rows || [])
    .filter(function (r) { return String(r['Active?'] || 'Yes').toLowerCase() !== 'no'; })
    .map(function (r) { return r[idCol] + ' — ' + r[nameCol]; });
}

// Just the IDs (for lookups where you don't want the label)
function IM_projectIds()  { return IM_getProjects().map(function (r)  { return r.Project_ID;  }); }
function IM_customerIds() { return IM_getCustomers().map(function (r) { return r.Customer_ID; }); }

// -----------------------------------------------------------------------------
// Cache plumbing
// -----------------------------------------------------------------------------
function IM_cache_(key, factory) {
  var cache = CacheService.getScriptCache();
  var hit = cache.get(IM_CACHE_KEY + key);
  if (hit) {
    try { return JSON.parse(hit); } catch (_) {}
  }
  var fresh = factory();
  try {
    var str = JSON.stringify(fresh);
    if (str.length < 95000) cache.put(IM_CACHE_KEY + key, str, IM_CACHE_TTL);
  } catch (_) {}
  return fresh;
}

function IM_clearCache() {
  var cache = CacheService.getScriptCache();
  cache.removeAll([
    IM_CACHE_KEY + 'projects',
    IM_CACHE_KEY + 'customers',
    IM_CACHE_KEY + 'employees',
    IM_CACHE_KEY + 'sites',
    IM_CACHE_KEY + 'parties_{}',
    IM_CACHE_KEY + 'parties_{"isSupplier":true}',
    IM_CACHE_KEY + 'parties_{"isJobWorker":true}',
    IM_CACHE_KEY + 'parties_{"isTransporter":true}',
    IM_CACHE_KEY + 'parties_{"isServiceProv":true}'
  ]);
  SpreadsheetApp.getActive().toast('Masters cache cleared', '🔗 Interlink', 3);
}

// -----------------------------------------------------------------------------
// Menu — attached via installable onOpen trigger so the system's existing
// simple onOpen stays untouched. Run IM_installInterlinkMenu() ONCE.
// -----------------------------------------------------------------------------
function IM_installInterlinkMenu() {
  var ss = SpreadsheetApp.getActive();
  // De-dup: remove any existing IM_onOpen_ trigger first
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'IM_onOpen_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('IM_onOpen_').forSpreadsheet(ss).onOpen().create();
  // Also run once now so the menu appears without reopening
  IM_onOpen_();
  SpreadsheetApp.getUi().alert(
    'Interlink menu installed',
    'A 🔗 Interlink menu will appear every time this sheet is opened. ' +
    'The system\'s existing menu is unchanged.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function IM_onOpen_() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🔗 Interlink')
    .addItem('Refresh dropdowns from Masters', 'IM_refreshDropdowns')
    .addItem('Clear Masters cache',            'IM_clearCache')
    .addSeparator()
    .addItem('IMS material costs: import now', 'IM_importImsCostsNow')
    .addSeparator()
    .addItem('About',                          'IM_about_')
    .addToUi();
  try { Masters.registerSystemSheet('finance', SpreadsheetApp.getActive().getId()); } catch (e) {}  // lets PM's project delete find this sheet
  try { IM_ensureImportTrigger_(); } catch (e) {}  // self-heal the 10-minute IMS cost import
}

function IM_about_() {
  var pv = '?';
  try { pv = Masters.getVersion(); } catch (e) {}
  SpreadsheetApp.getUi().alert(
    '🔗 Interlink v' + IM_VERSION,
    'This system reads master data (Projects, Customers, Parties, Employees, ' +
    'Sites) from the Masters library.\n\n' +
    'Masters version: ' + pv + '\n' +
    'Cache TTL: ' + IM_CACHE_TTL + ' sec',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// -----------------------------------------------------------------------------
// Dropdown refresher — stub for v0.4 Phase 1. Each system will fill in its
// own IM_refreshDropdowns() in a sibling _Interlink_Dropdowns.gs file that
// knows which sheets+columns to re-validate.
// -----------------------------------------------------------------------------
function IM_refreshDropdowns() {
  SpreadsheetApp.getUi().alert(
    'Dropdowns not wired yet',
    'This system doesn\'t have per-sheet dropdown rules configured yet. ' +
    'That lives in _Interlink_Dropdowns.gs (next step).',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// =============================================================================
// Phase 7 (v2) — IMS material costs → Project_Costs: automatic, grouped
// Finance PULLS (a time trigger every 10 min, on open, and 🔗 Interlink → import now); nothing
// else writes into this workbook. Material a project drew from IMS stock (Material Requisition lines)
// becomes a 'Material' cost.
//
//  * PM / Masters is the only truth for project names: a requisition counts when its Project BEGINS
//    with a PRJ-### id (the label picked from the IMS dropdown). The Finance project is the Masters
//    label "PRJ-001 — Name"; no pairing with IMS or Finance names is needed. The project is added to
//    M_Project (Code PRJ-001, Name) when Finance has no project with that code yet.
//  * Many partial issues do NOT create many rows: one Project_Costs row per project + item, which
//    grows (qty and amount) as more is issued. Note key "IMS:PRJ-001:ITEMCODE#1".
//  * Every requisition line is priced once and frozen in the hidden sheet IMS_Cost_Log (qty x the IMS
//    'Avg Rate (Rs.)' of that moment, GST 0 since GST was paid on the purchase bill), so a later rate
//    change never rewrites old cost and nothing is ever counted twice.
//  * A group row is only updated while its Status is Pending. Once approved / rejected (or soft
//    deleted) it is never changed; later issues of that item go to a new row "#2".
//  * Lines with no Avg Rate yet, or for a PRJ id Masters does not know, wait and are retried.
//  * Lines already booked by the old per-line import (key "IMS:REQ-0012:SR5") are skipped.
// Vendor bills tagged "Count in project cost" must not cover stock that is also issued via
// requisitions, or the same material is counted twice.
// One-time: run IM_installImportTrigger() from the editor (creates the 10-minute trigger).
// =============================================================================
var IM_IMS_SHEET_ID = '1CglvPxUWZnIHNeToGEDz8-Xj3eNfvaF351j_ogpRTgk';
var IM_LOG_SHEET    = 'IMS_Cost_Log';
var IM_LOG_HEAD     = ['Line', 'Date', 'Project', 'Item Code', 'Item', 'Qty', 'Unit', 'Rate', 'Amount', 'Group', 'Status'];
var IML = { LINE: 0, DATE: 1, PRJ: 2, CODE: 3, ITEM: 4, QTY: 5, UNIT: 6, RATE: 7, AMOUNT: 8, GROUP: 9, STATUS: 10 };
var IM_IMPORT_FN    = 'IM_importImsCostsTrigger';
var IM_IMPORT_MINUTES = 10;

/** Keys ("IMS:REQ-0012:SR5") booked line-by-line by the old import. Soft-deleted rows don't count. */
function IM_legacyKeys_() {
  var out = {};
  var sh = _getProjectCostsSheet();
  if (!sh || sh.getLastRow() < 2) return out;
  var width = Math.min(sh.getLastColumn(), PCC_TOTAL);
  sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues().forEach(function (r) {
    var note = _s(r[PCC.NOTES]);
    if (note.indexOf('IMS:REQ-') === 0 && _s(r[PCC.DELETED]) !== 'Yes') out[note] = true;
  });
  return out;
}

/** IMS Item Master → { code: { name, rate } }, rate = 'Avg Rate (Rs.)'. */
function IM_avgRates_(items) {
  var out = {};
  if (!items || items.getLastRow() < 2) return out;
  var head = items.getRange(1, 1, 1, items.getLastColumn()).getValues()[0];
  var cRate = head.indexOf('Avg Rate (Rs.)'), cDesc = head.indexOf('Description'), cName = head.indexOf('Item Name');
  if (cRate < 0) return out;
  items.getRange(2, 1, items.getLastRow() - 1, items.getLastColumn()).getValues().forEach(function (r) {
    if (r[0]) out[r[0]] = { name: r[cDesc] || r[cName] || '', rate: Number(r[cRate]) || 0 };
  });
  return out;
}

function IM_dateStr_(d) {
  return d instanceof Date ? Utilities.formatDate(d, 'GMT+5:30', 'yyyy-MM-dd') : String(d || '');
}

function IM_round_(n, p) { var k = Math.pow(10, p); return Math.round(n * k) / k; }

function IM_logSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(IM_LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(IM_LOG_SHEET);
    sh.getRange(1, 1, 1, IM_LOG_HEAD.length).setValues([IM_LOG_HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.hideSheet();
  }
  return sh;
}

/** Existing group rows in Project_Costs: { groupKey → { row, status, deleted } } and { "PRJ|CODE" → highest #n }. */
function IM_groupRows_() {
  var byKey = {}, top = {};
  var sh = _getProjectCostsSheet();
  if (!sh || sh.getLastRow() < 2) return { byKey: byKey, top: top };
  var width = Math.min(sh.getLastColumn(), PCC_TOTAL);
  sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues().forEach(function (r, i) {
    var m = /^IMS:(PRJ-\d+):(.+)#(\d+)$/.exec(_s(r[PCC.NOTES]));
    if (!m) return;
    byKey[m[0]] = { row: i + 2, status: _s(r[PCC.STATUS]), deleted: _s(r[PCC.DELETED]) === 'Yes' };
    var k = m[1] + '|' + m[2], n = parseInt(m[3], 10);
    if (!top[k] || n > top[k].n) top[k] = { n: n, key: m[0] };
  });
  return { byKey: byKey, top: top };
}

/** Add PRJ-### to Finance's M_Project (Code | Name | Active) unless a project with that code exists. Returns a warning or ''. */
function IM_ensureFinanceProject_(prj, name) {
  var sh = _getSS().getSheetByName(P.PROJECT);
  if (!sh) return '';
  var rows = sh.getLastRow() >= 2 ? sh.getRange(2, 1, sh.getLastRow() - 1, PRC_TOTAL).getValues() : [];
  for (var i = 0; i < rows.length; i++) {
    if (_s(rows[i][PRC.CODE]).toUpperCase() === prj) {
      return _s(rows[i][PRC.NAME]) === name ? '' : 'Finance already has code ' + prj + ' for "' + _s(rows[i][PRC.NAME]) + '" (Masters: "' + name + '")';
    }
  }
  var row = new Array(PRC_TOTAL).fill('');
  row[PRC.CODE] = prj; row[PRC.NAME] = name; row[PRC.ACTIVE] = 'Yes';
  sh.getRange(sh.getLastRow() + 1, 1, 1, PRC_TOTAL).setValues([row]);
  return '';
}

/** The import. Safe to run any time, as often as you like: it only ever adds what is new. */
function IM_importImsCosts() {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) return { busy: true };
  try { return IM_importImsCostsLocked_(); } finally { lock.releaseLock(); }
}

function IM_importImsCostsLocked_() {
  var ims = SpreadsheetApp.openById(IM_IMS_SHEET_ID);
  var req = ims.getSheetByName('Material Requisition Register'), items = ims.getSheetByName('Item Master');
  if (!req || !items) throw new Error('IMS sheets not found (Material Requisition Register / Item Master).');
  var res = { newLines: 0, created: 0, updated: 0, closed: 0, waitingRate: 0, waitingProject: 0, warnings: [] };

  var names = {};  // PRJ-### → Masters name
  (Masters.getProjects() || []).forEach(function (p) { names[String(p.Project_ID).toUpperCase()] = String(p.Name); });
  var rates = IM_avgRates_(items), legacy = IM_legacyKeys_();

  // 1. the PRJ requisition lines IMS has right now
  var lines = [], present = {};
  if (req.getLastRow() >= 2) req.getRange(2, 1, req.getLastRow() - 1, 8).getValues().forEach(function (r) {
    var reqNo = r[1], code = r[4], qty = Number(r[6]), m = /^PRJ-\d+/i.exec(String(r[3] || '').trim());
    if (!reqNo || !code || !(qty > 0) || !m) return;
    var key = reqNo + ':SR' + r[0];
    present[key] = true;
    lines.push({ key: key, date: IM_dateStr_(r[2]), prj: m[0].toUpperCase(), code: code, qty: qty, unit: r[7],
                 item: (rates[code] && rates[code].name) || r[5] || code });
  });

  // 2. price + group the lines not seen before
  var log = IM_logSheet_();
  var logRows = log.getLastRow() >= 2 ? log.getRange(2, 1, log.getLastRow() - 1, IM_LOG_HEAD.length).getValues() : [];
  var seen = {}; logRows.forEach(function (r) { seen[r[IML.LINE]] = true; });
  var groups = IM_groupRows_();
  var open = {};  // "PRJ|CODE" → { key, n } the group new lines go to
  Object.keys(groups.top).forEach(function (k) {
    var g = groups.byKey[groups.top[k].key];
    if (g && !g.deleted && g.status === 'Pending') open[k] = groups.top[k];
  });
  var newRows = [];
  lines.forEach(function (l) {
    if (seen[l.key]) return;
    if (legacy['IMS:' + l.key]) { newRows.push([l.key, l.date, l.prj, l.code, l.item, l.qty, l.unit, '', '', '', 'Legacy']); return; }
    if (!names[l.prj]) { res.waitingProject++; return; }
    var rate = rates[l.code] ? rates[l.code].rate : 0;
    if (!(rate > 0)) { res.waitingRate++; return; }
    var k = l.prj + '|' + l.code;
    if (!open[k]) {
      var n = groups.top[k] ? groups.top[k].n + 1 : 1;
      open[k] = { n: n, key: 'IMS:' + l.prj + ':' + l.code + '#' + n };
      groups.top[k] = open[k];
    }
    newRows.push([l.key, l.date, l.prj, l.code, l.item, l.qty, l.unit, rate, IM_round_(l.qty * rate, 2), open[k].key, 'Imported']);
  });
  if (newRows.length) {
    log.getRange(log.getLastRow() + 1, 1, newRows.length, IM_LOG_HEAD.length).setValues(newRows);
    res.newLines = newRows.filter(function (r) { return r[IML.STATUS] === 'Imported'; }).length;
  }

  // 3. every group row = the sum of its still-present logged lines
  var all = logRows.concat(newRows), sums = {};
  all.forEach(function (r) {
    if (r[IML.STATUS] !== 'Imported') return;
    var g = sums[r[IML.GROUP]] = sums[r[IML.GROUP]] || { key: r[IML.GROUP], prj: r[IML.PRJ], code: r[IML.CODE], item: r[IML.ITEM], unit: r[IML.UNIT], qty: 0, amount: 0, date: '', lines: 0 };
    if (!present[r[IML.LINE]]) return;                       // removed from IMS since: no longer counts
    g.qty += Number(r[IML.QTY]); g.amount += Number(r[IML.AMOUNT]); g.lines++;
    if (String(r[IML.DATE]) > g.date) g.date = String(r[IML.DATE]);
  });
  var pc = _getProjectCostsSheet(), entries = [], touched = {};
  var sub = PropertiesService.getScriptProperties().getProperty('IM_LAST_SUBCAT') || '';
  Object.keys(sums).forEach(function (key) {
    var g = sums[key], ex = groups.byKey[key];
    g.qty = IM_round_(g.qty, 3); g.amount = IM_round_(g.amount, 2);
    if (ex && (ex.deleted || ex.status !== 'Pending')) return;        // decided by a person: never touched
    var desc = 'Issued from IMS stock · ' + g.lines + ' requisition line(s)';
    if (ex) {
      if (g.lines === 0) {                                              // every line was removed from IMS
        pc.getRange(ex.row, PCC.DELETED + 1).setValue('Yes'); res.closed++; return;
      }
      var cur = pc.getRange(ex.row, 1, 1, PCC_TOTAL).getValues()[0];
      if (Math.abs(Number(cur[PCC.QUANTITY]) - g.qty) < 0.0005 && Math.abs(Number(cur[PCC.AMOUNT]) - g.amount) < 0.005) return;  // unchanged
      pc.getRange(ex.row, PCC.ENTRY_DATE + 1).setValue(g.date);
      pc.getRange(ex.row, PCC.DESCRIPTION + 1).setValue(desc);
      pc.getRange(ex.row, PCC.QUANTITY + 1).setValue(g.qty);
      pc.getRange(ex.row, PCC.RATE_PER_UNIT + 1).setValue(g.qty ? g.amount / g.qty : 0);
      pc.getRange(ex.row, PCC.AMOUNT_EX_GST + 1).setValue(g.amount);
      pc.getRange(ex.row, PCC.AMOUNT + 1).setValue(g.amount);
      res.updated++; touched[g.prj] = true;
    } else if (g.lines > 0) {
      entries.push({ entry_type: 'Material', project: g.prj + ' — ' + names[g.prj], entry_date: g.date, submitted_by: 'IMS import',
        data: { materialName: String(g.item), quantity: g.qty, unit: String(g.unit), rate: g.amount / g.qty, gst_rate: 0,
                subcategory: sub, description: desc, notes: g.key } });
      touched[g.prj] = true;
    }
  });
  if (entries.length) {
    var out = submitProjectCosts(entries);
    if (!out || !out.success) throw new Error('Finance did not accept the cost rows: ' + ((out && out.message) || 'error') + ' (nothing is lost: they are retried on the next run)');
    res.created = out.count;
  }
  Object.keys(touched).forEach(function (prj) {
    var w = IM_ensureFinanceProject_(prj, names[prj]);
    if (w) res.warnings.push(w);
  });
  return res;
}

/** Trigger / menu entry points + heartbeat (log sheet M1: if it is stale, the trigger is not running). */
function IM_importImsCostsTrigger() { IM_runImport_(true); }
function IM_importImsCostsNow() { IM_runImport_(false); }

function IM_runImport_(quiet) {
  var res;
  try { res = IM_importImsCosts(); }
  catch (e) {
    if (!quiet) SpreadsheetApp.getUi().alert('IMS material costs', 'Import failed: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    else Logger.log('IMS cost import failed: ' + e.message);
    return null;
  }
  if (res.busy) { if (!quiet) SpreadsheetApp.getActive().toast('Another import is running; try again in a minute.', 'IMS material costs', 6); return res; }
  var n = ScriptApp.getProjectTriggers().filter(function (t) { return t.getHandlerFunction() === IM_IMPORT_FN; }).length;
  try {
    IM_logSheet_().getRange('M1').setValue('Last import ' + Utilities.formatDate(new Date(), 'GMT+5:30', 'dd/MM HH:mm') +
      ' · +' + res.newLines + ' line(s), ' + res.created + ' row(s) created, ' + res.updated + ' updated · triggers: ' + n);
  } catch (e) {}
  if (!quiet) {
    var msg = res.newLines + ' new requisition line(s) imported.\n' + res.created + ' cost row(s) created, ' + res.updated + ' updated' +
      (res.closed ? ', ' + res.closed + ' closed (lines removed in IMS)' : '') + '.' +
      (res.waitingRate ? '\n' + res.waitingRate + ' line(s) wait for an Avg Rate in IMS Item Master.' : '') +
      (res.waitingProject ? '\n' + res.waitingProject + ' line(s) wait: PRJ id not found in Masters.' : '') +
      (res.warnings.length ? '\n\nCheck: ' + res.warnings.join('\n') : '');
    SpreadsheetApp.getUi().alert('IMS material costs', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  }
  return res;
}

/** Create the 10-minute trigger (once, from the editor; it also authorises the script). */
function IM_installImportTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === IM_IMPORT_FN) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger(IM_IMPORT_FN).timeBased().everyMinutes(IM_IMPORT_MINUTES).create();
  var res = IM_runImport_(true);
  var msg = 'IMS cost import trigger created (every ' + IM_IMPORT_MINUTES + ' minutes). First run: ' +
            (res && !res.busy ? res.newLines + ' line(s), ' + res.created + ' row(s) created.' : 'will run at the next tick.');
  try { SpreadsheetApp.getUi().alert('IMS material costs', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { Logger.log(msg); }
}

/** Recreate the trigger if it has gone missing. Called when Finance opens. */
function IM_ensureImportTrigger_() {
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === IM_IMPORT_FN; });
  if (!has) ScriptApp.newTrigger(IM_IMPORT_FN).timeBased().everyMinutes(IM_IMPORT_MINUTES).create();
}
