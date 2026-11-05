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
    .addItem('IMS material costs: 1. Prepare', 'IM_prepareCostImport')
    .addItem('IMS material costs: 2. Post ticked lines', 'IM_postCostImport')
    .addSeparator()
    .addItem('About',                          'IM_about_')
    .addToUi();
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
// Phase 7 — IMS material costs → Project_Costs
// Finance PULLS; nothing else ever writes into this workbook. Material that a project
// drew from IMS stock (Material Requisition lines) is booked as a 'Material' cost.
//   1. IM_prepareCostImport  reads IMS requisitions that are not booked yet into the review
//                            sheet IMS_Cost_Import (one row per requisition line).
//   2. You review: tick lines, pick the Finance project and Subcategory, fix a rate if needed.
//   3. IM_postCostImport     books the ticked lines through submitProjectCosts() (Finance's own
//                            ID / category / GST rules; Status Pending like any cost entry).
// Cost = IMS Item Master 'Avg Rate (Rs.)' x requisition qty, GST blank (GST was paid on the
// purchase bill). Each booked row carries the key "IMS:REQ-0012:SR5" in Notes, so a line can
// never be booked twice (a soft-deleted cost row frees its line to be booked again).
// The PRJ-### (or old requisition text) -> Finance project pairing is remembered in Masters
// (sheet Project_Map, via Masters.getProjectMap / saveProjectMap), so every system sees one answer;
// it is pre-filled when exactly one Finance project label starts with that same PRJ-### id.
// A local IMS_Project_Map sheet from before the move is copied into Masters on the next Prepare,
// then renamed and hidden.
// Vendor bills tagged "Count in project cost" must not cover stock that is also issued via
// requisitions, or the same material is counted twice.
// =============================================================================
var IM_IMS_SHEET_ID = '1CglvPxUWZnIHNeToGEDz8-Xj3eNfvaF351j_ogpRTgk';
var IM_COST_SHEET   = 'IMS_Cost_Import';
var IM_MAP_SHEET    = 'IMS_Project_Map';  // legacy local copy, moved into Masters by IM_migrateProjectMap_
var IM_COST_HEADERS = ['Post?', 'Requisition', 'Date', 'IMS Project', 'PRJ', 'Item Code', 'Item', 'Qty', 'Unit',
                       'Avg Rate', 'Amount', 'Finance Project', 'Subcategory', 'Key', 'Status'];
// 0-based indexes into a review-sheet row
var IMC = { POST: 0, REQ: 1, DATE: 2, IMSPROJ: 3, PRJ: 4, CODE: 5, ITEM: 6, QTY: 7, UNIT: 8,
            RATE: 9, AMOUNT: 10, FINPROJ: 11, SUBCAT: 12, KEY: 13, STATUS: 14 };

/** Keys ("IMS:REQ-0012:SR5") of IMS lines already booked in Project_Costs. Soft-deleted rows don't count. */
function IM_postedKeys_() {
  var out = {};
  var sh = _getProjectCostsSheet();
  if (!sh || sh.getLastRow() < 2) return out;
  var width = Math.min(sh.getLastColumn(), PCC_TOTAL);
  sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues().forEach(function (r) {
    var note = _s(r[PCC.NOTES]);
    if (note.indexOf('IMS:') === 0 && _s(r[PCC.DELETED]) !== 'Yes') out[note] = _s(r[PCC.ID]);
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

/** { KEY: 'Finance project label' } from Masters (keys upper-cased). */
function IM_readProjectMap_() {
  return Masters.getProjectMap() || {};
}

/** pairs: { 'PRJ-001': 'Finance project label' } → upsert into the Masters Project_Map. */
function IM_saveProjectMap_(pairs) {
  Masters.saveProjectMap(pairs);
}

/** Copy a leftover local IMS_Project_Map sheet into Masters once, then rename + hide it. Returns pairs moved. */
function IM_migrateProjectMap_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(IM_MAP_SHEET);
  if (!sh) return 0;
  var pairs = {}, n = 0;
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
    if (r[0] && r[1]) { pairs[String(r[0]).toUpperCase()] = String(r[1]); n++; }
  });
  if (n) Masters.saveProjectMap(pairs);  // throws if Masters is unreachable: the sheet stays as it is
  sh.setName('_old ' + IM_MAP_SHEET); sh.hideSheet();
  return n;
}

/** Pre-fill: the one Finance project label that starts with this exact PRJ-### id, else ''. */
function IM_guessFinanceProject_(prj, names) {
  var hit = names.filter(function (n) { var u = String(n).toUpperCase(); return u === prj || u.indexOf(prj + ' ') === 0; });
  return hit.length === 1 ? hit[0] : '';
}

/** Step 1: IMS requisition lines not yet booked → review sheet. */
function IM_prepareCostImport() {
  var ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi();
  var ims = SpreadsheetApp.openById(IM_IMS_SHEET_ID);
  var req = ims.getSheetByName('Material Requisition Register'), items = ims.getSheetByName('Item Master');
  if (!req || !items) { ui.alert('IMS sheets not found (Material Requisition Register / Item Master).'); return; }

  var done = IM_postedKeys_();
  var rates = IM_avgRates_(items);
  IM_migrateProjectMap_();
  var map = IM_readProjectMap_();
  var names = (typeof getProjectNames === 'function') ? (getProjectNames() || []) : [];
  var props = PropertiesService.getScriptProperties();
  var lastSub = props.getProperty('IM_LAST_SUBCAT') || '';

  // keep what the user already chose on an earlier Prepare (by line key)
  var sh = ss.getSheetByName(IM_COST_SHEET) || ss.insertSheet(IM_COST_SHEET);
  var kept = {};
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, IM_COST_HEADERS.length).getValues().forEach(function (r) {
      if (r[IMC.KEY] && !/^Posted/.test(String(r[IMC.STATUS]))) kept[r[IMC.KEY]] = r;
    });
  }

  var rows = [], noPrj = 0, ignored = {};  // ignored: project text without a PRJ-### id → line count
  if (req.getLastRow() >= 2) {
    req.getRange(2, 1, req.getLastRow() - 1, 8).getValues().forEach(function (r) {
      var sr = r[0], reqNo = r[1], proj = String(r[3] || ''), code = r[4], qty = Number(r[6]);
      if (!reqNo || !code || !(qty > 0)) return;
      var m = /PRJ-\d+/i.exec(proj);
      // Old requisitions carry a plain site/project name: the exact text (upper-cased) is the map key,
      // so it is paired with a Finance project once by hand and remembered. Never fuzzy-matched.
      if (!m && !proj.trim()) { noPrj++; ignored['(blank)'] = (ignored['(blank)'] || 0) + 1; return; }
      var key = 'IMS:' + reqNo + ':SR' + sr;
      if (done[key]) return;
      var prj = m ? m[0].toUpperCase() : proj.trim().toUpperCase(), it = rates[code] || {}, old = kept[key];
      var rate = it.rate > 0 ? it.rate : (old && Number(old[IMC.RATE]) > 0 ? Number(old[IMC.RATE]) : '');
      var fin = (old && old[IMC.FINPROJ]) || map[prj] || (m ? IM_guessFinanceProject_(prj, names) : '');
      var sub = (old && old[IMC.SUBCAT]) || lastSub;
      var tick = old ? old[IMC.POST] === true : !!(fin && rate > 0);
      rows.push([tick, reqNo, IM_dateStr_(r[2]), proj, prj, code, it.name || r[5], qty, r[7], rate, '', fin, sub, key,
                 rate > 0 ? '' : 'No Avg Rate in IMS: type a rate']);
    });
  }

  rows.sort(function (a, b) { return a[IMC.PRJ] < b[IMC.PRJ] ? -1 : a[IMC.PRJ] > b[IMC.PRJ] ? 1 : 0; });  // grouped by project: pick once, fill down
  sh.clear();
  sh.getRange(1, 1, 1, IM_COST_HEADERS.length).setValues([IM_COST_HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  if (rows.length) {
    var n = rows.length;
    sh.getRange(2, 1, n, IM_COST_HEADERS.length).setValues(rows);
    var f = [];
    for (var i = 0; i < n; i++) f.push(['=IF(AND(ISNUMBER(H' + (i + 2) + '),ISNUMBER(J' + (i + 2) + ')),H' + (i + 2) + '*J' + (i + 2) + ',"")']);
    sh.getRange(2, IMC.AMOUNT + 1, n, 1).setFormulas(f).setNumberFormat('#,##0.00');
    sh.getRange(2, IMC.POST + 1, n, 1).insertCheckboxes();
    if (names.length) sh.getRange(2, IMC.FINPROJ + 1, n, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(names, true).setAllowInvalid(false).build());
    var cats = [];
    try { cats = (getCategories() || []).map(function (c) { return c.category; }).filter(String); } catch (e) {}
    if (cats.length) sh.getRange(2, IMC.SUBCAT + 1, n, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(cats, true).setAllowInvalid(true).build());
    sh.getRange(2, IMC.RATE + 1, n, 1).setNumberFormat('#,##0.00');
    sh.setColumnWidths(1, 1, 60); sh.setColumnWidth(IMC.ITEM + 1, 240); sh.setColumnWidth(IMC.FINPROJ + 1, 240);
    sh.setColumnWidth(IMC.SUBCAT + 1, 180); sh.setColumnWidth(IMC.STATUS + 1, 260);
  }
  ss.setActiveSheet(sh);
  var ready = rows.filter(function (r) { return r[IMC.POST] === true; }).length;
  ui.alert('IMS material costs',
    rows.length + ' unbooked requisition line(s) loaded; ' + ready + ' ticked and ready.\n' +
    (rows.length - ready ? (rows.length - ready) + ' need a Finance project, a rate or your tick.\n' : '') +
    (noPrj ? noPrj + ' line(s) ignored: the requisition has no Project at all.\n' : '') +
    '\nReview the "' + IM_COST_SHEET + '" sheet, then run "2. Post ticked lines".', ui.ButtonSet.OK);
}

/** Step 2: book the ticked, valid lines through submitProjectCosts(). */
function IM_postCostImport() {
  var ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi();
  var sh = ss.getSheetByName(IM_COST_SHEET);
  if (!sh || sh.getLastRow() < 2) { ui.alert('Nothing to post. Run "1. Prepare" first.'); return; }

  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, IM_COST_HEADERS.length).getValues();
  var done = IM_postedKeys_();
  var entries = [], at = [], total = 0, projects = {}, skipped = [];
  vals.forEach(function (r, i) {
    if (r[IMC.POST] !== true || /^Posted/.test(String(r[IMC.STATUS]))) return;
    var qty = Number(r[IMC.QTY]), rate = Number(r[IMC.RATE]), fin = String(r[IMC.FINPROJ] || '').trim(), key = r[IMC.KEY];
    var why = !fin ? 'pick a Finance project' : !(qty > 0) ? 'bad qty' : !(rate > 0) ? 'no rate' :
              done[key] ? 'already booked as ' + done[key] : '';
    if (why) { sh.getRange(i + 2, IMC.STATUS + 1).setValue('Not posted: ' + why); skipped.push(r[IMC.REQ] + ' ' + r[IMC.CODE] + ': ' + why); return; }
    entries.push({
      entry_type: 'Material', project: fin, entry_date: IM_dateStr_(r[IMC.DATE]), submitted_by: 'IMS import',
      data: { materialName: String(r[IMC.ITEM]), quantity: qty, unit: String(r[IMC.UNIT]), rate: rate, gst_rate: 0,
              subcategory: String(r[IMC.SUBCAT] || ''),
              description: 'Issued from IMS stock · ' + r[IMC.REQ] + ' · ' + r[IMC.CODE], notes: key }
    });
    at.push(i + 2);
    total += qty * rate;
    projects[fin] = (projects[fin] || 0) + qty * rate;
  });
  if (!entries.length) { ui.alert('Nothing to post.' + (skipped.length ? '\n\n' + skipped.join('\n') : '')); return; }

  var lines = Object.keys(projects).map(function (p) { return '  ' + p + ': ₹' + projects[p].toFixed(2); }).join('\n');
  var ok = ui.alert('Book ' + entries.length + ' material cost line(s)?',
    'Total ₹' + total.toFixed(2) + ' (no GST) as Project_Costs "Material" entries, Status Pending:\n' + lines +
    (skipped.length ? '\n\nSkipped:\n' + skipped.join('\n') : ''), ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  var res = submitProjectCosts(entries);
  if (!res || !res.success) {
    // submitProjectCosts appends one by one: mark whatever did land, so a retry cannot double-book it
    var landed = IM_postedKeys_();
    at.forEach(function (row, k) {
      var key = entries[k].data.notes;
      sh.getRange(row, IMC.STATUS + 1).setValue(landed[key] ? 'Posted ' + landed[key] : 'Not posted: ' + ((res && res.message) || 'error'));
      if (landed[key]) sh.getRange(row, IMC.POST + 1).setValue(false);
    });
    ui.alert('Finance did not accept all lines: ' + ((res && res.message) || 'error') +
             '\nLines that did land are marked Posted; nothing is booked twice if you run Post again.');
    return;
  }
  var pairs = {};
  res.ids.forEach(function (id, k) {
    sh.getRange(at[k], IMC.STATUS + 1).setValue('Posted ' + id);
    sh.getRange(at[k], IMC.POST + 1).setValue(false);
    pairs[String(vals[at[k] - 2][IMC.PRJ]).toUpperCase()] = entries[k].project;
  });
  try { IM_saveProjectMap_(pairs); }
  catch (e) { ss.toast('Booked, but the project pairing could not be saved to Masters: ' + e.message, 'IMS material costs', 10); }
  var lastSub = String(vals[at[at.length - 1] - 2][IMC.SUBCAT] || '');
  if (lastSub) PropertiesService.getScriptProperties().setProperty('IM_LAST_SUBCAT', lastSub);
  ui.alert('Booked ' + res.count + ' line(s), ₹' + total.toFixed(2) + '. They appear in Project_Costs as Pending (' + res.ids[0] +
           (res.ids.length > 1 ? ' … ' + res.ids[res.ids.length - 1] : '') + ').');
}
