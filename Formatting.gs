// ============================================================
// 3 Star Projects — Unified Finance System
// SHEET FORMATTING  (brand styling for every tab)
// File: Formatting.gs   (depends on Masters.gs BRAND + _getSS)
//
// Run beautifyAllSheets() once to style every existing sheet.
// _buildSheet() also calls _beautifySheet() so new sheets (and future
// FY tabs) are styled automatically — no need to re-run after a new year.
//
// Styling applied per sheet:
//   • Brand header row (navy-blue fill, white bold, taller, frozen)
//   • First column frozen on wide sheets
//   • Zebra striping (auto-extends to new rows)
//   • ₹ Indian number format on money columns; dd-mmm-yyyy on date columns
//   • Right-aligned numeric columns
//   • Status colour-coding (Paid/Settled green · Pending/Unpaid amber · Cancelled red)
//   • Negative amounts shown in red
//   • Tab colours grouped by module (masters / ledger / AR / payroll)
// ============================================================

var FMT = {
  headerBg:  '#1F4E79',   // brand primary
  headerFg:  '#FFFFFF',
  zebra:     '#EEF3F8',   // very light blue-grey
  gridline:  '#D7DEE6',
  money:     '\u20B9#,##,##0.00',          // ₹ with Indian grouping
  date:      'dd-mmm-yyyy',
  green:     '#15803D', greenBg:'#E7F6EC',
  amber:     '#92400E', amberBg:'#FEF3C7',
  red:       '#B91C1C', redBg:'#FDE7E7'
};

// Money / date / number column detection by header name
var MONEY_COLS = ['amount','order_amount','invoiced_amount','pending_amount','taxable_value',
  'cgst','sgst','igst','invoice_amount','outstanding','paid','return_amount','unapplied_advance',
  'gross','daily_rate','ot_earnings','advance_deducted','penalty','penalty_ot','penalty_salary',
  'absent_deduction','half_day_deduction','net_salary','gross_monthly_salary','opening_balance',
  'total_ordered','total_invoiced','total_paid','returns','advance_held','total_spent'];
var DATE_HINTS  = ['date','delivery','activity'];   // any header containing these → date format
var NUM_COLS    = ['month_days','days_present','days_absent','ot_hours','half_days','gst_rate','version'];

function _isMoneyHeader(h){ h=String(h).toLowerCase(); return MONEY_COLS.indexOf(h) !== -1; }
function _isDateHeader(h){ h=String(h).toLowerCase();
  if (h.indexOf('created')===0 || h.indexOf('calculated')===0) return false;  // these are text timestamps
  for (var i=0;i<DATE_HINTS.length;i++) if (h.indexOf(DATE_HINTS[i])!==-1) return true; return false; }
function _isNumHeader(h){ h=String(h).toLowerCase(); return NUM_COLS.indexOf(h) !== -1; }

function _colLetter(n){ var s=''; while(n>0){ var m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26); } return s; }

function _tabColorFor(name){
  if (name === 'Customer_Directory' || name === 'Vendor_Directory') return '#15803D'; // rollups — green
  if (name.indexOf('M_') === 0) return '#2C7DA0';                 // masters — teal
  if (name.indexOf('Transactions_') === 0 || name === 'Personal_Transactions') return '#1F4E79'; // ledger — blue
  if (['Orders','Invoices','Payments','Returns'].indexOf(name) !== -1) return '#F47C2C';          // AR — orange
  if (name === 'Salary') return '#E94E1B';                        // payroll — red-orange
  return '#5C6675';                                               // other — grey
}

// ── Core: style one sheet based on its header row ───────────
function _beautifySheet(sheet){
  try {
    if (!sheet) return;
    var nCols = sheet.getLastColumn(); if (nCols < 1) return;
    var headers = sheet.getRange(1,1,1,nCols).getValues()[0];

    // Header row
    sheet.setRowHeight(1, 34);
    var hdr = sheet.getRange(1,1,1,nCols);
    hdr.setBackground(FMT.headerBg).setFontColor(FMT.headerFg)
       .setFontWeight('bold').setFontSize(11).setVerticalAlignment('middle')
       .setBorder(null,null,true,null,null,null,'#0E2D49',SpreadsheetApp.BorderStyle.SOLID_THICK);
    sheet.setFrozenRows(1);
    if (nCols >= 11) sheet.setFrozenColumns(1);

    // Gridlines / font for the whole used grid
    var lastCol = _colLetter(nCols);
    sheet.getRange('A1:'+lastCol+'1').setHorizontalAlignment('left');

    // Per-column number formats + alignment
    for (var c=0;c<nCols;c++){
      var h = headers[c], col = c+1, L = _colLetter(col), rng = sheet.getRange(L+'2:'+L);
      if (_isMoneyHeader(h)){ rng.setNumberFormat(FMT.money).setHorizontalAlignment('right');
        hdr.getCell(1,col).setHorizontalAlignment('right'); }
      else if (_isDateHeader(h)){ rng.setNumberFormat(FMT.date).setHorizontalAlignment('center');
        hdr.getCell(1,col).setHorizontalAlignment('center'); }
      else if (_isNumHeader(h)){ rng.setHorizontalAlignment('center');
        hdr.getCell(1,col).setHorizontalAlignment('center'); }
    }

    // Remove old bandings to avoid stacking
    sheet.getBandings().forEach(function(b){ b.remove(); });

    // Conditional formats: zebra + negatives + status colours
    var rules = [];
    var dataRange = sheet.getRange('A2:'+lastCol);

    // zebra (auto-extends to new rows)
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=ISEVEN(ROW())').setBackground(FMT.zebra).setRanges([dataRange]).build());

    for (var c2=0;c2<nCols;c2++){
      var hh = String(headers[c2]).toLowerCase(), L2 = _colLetter(c2+1), colRange = sheet.getRange(L2+'2:'+L2);
      // negative money in red
      if (_isMoneyHeader(headers[c2])){
        rules.push(SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThan(0).setFontColor(FMT.red).setRanges([colRange]).build());
      }
      // status-style columns
      if (hh==='status' || hh==='settle_status' || hh==='recovery' || hh==='active'){
        ['Paid','Settled','Active','Recovered','Closed','Yes'].forEach(function(v){
          rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(v)
            .setBackground(FMT.greenBg).setFontColor(FMT.green).setBold(true).setRanges([colRange]).build()); });
        ['Pending','Partial','Open','Unpaid'].forEach(function(v){
          rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(v)
            .setBackground(FMT.amberBg).setFontColor(FMT.amber).setBold(true).setRanges([colRange]).build()); });
        ['Cancelled','No'].forEach(function(v){
          rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(v)
            .setBackground(FMT.redBg).setFontColor(FMT.red).setBold(true).setRanges([colRange]).build()); });
      }
    }
    sheet.setConditionalFormatRules(rules);

    // Tidy view
    sheet.setHiddenGridlines(false);
    sheet.setTabColor(_tabColorFor(sheet.getName()));
  } catch(e){ Logger.log('_beautifySheet '+(sheet?sheet.getName():'?')+': '+e.message); }
}

// ── Entry point: style every sheet in the spreadsheet ──────
function beautifyAllSheets(){
  var ss = _getSS(), sheets = ss.getSheets(), n = 0;
  sheets.forEach(function(sh){
    var name = sh.getName();
    if (sh.getLastColumn() < 1 || sh.getLastRow() < 1) return;
    // skip stray/utility sheets without a header row of our making
    _beautifySheet(sh); n++;
  });
  var msg = 'Styled '+n+' sheet'+(n===1?'':'s')+'.';
  try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log(msg); }
  return {success:true, count:n};
}