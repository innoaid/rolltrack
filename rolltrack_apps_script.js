// ════════════════════════════════════════════════════════════════
// RollTrack Pro — Google Apps Script
// Innovation AID Sdn Bhd
// ════════════════════════════════════════════════════════════════

var SPREADSHEET_ID = '1O3Hvc0D-wMcBKLcC5IQKAboR1maFI2QuSi6XlIFZq9U';

var SUBCONS = { 'SC01': 'Team Atik' };

// ── Helpers ──────────────────────────────────────────────────────
function getSpreadsheet() { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function getSheet(name)   { return getSpreadsheet().getSheetByName(name); }

function sheetToObjects(sheet) {
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  var result  = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (!row[0] && !row[1]) continue; // skip completely empty rows
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      var v = row[c];
      obj[headers[c]] = v instanceof Date ? v.toISOString() : v;
    }
    result.push(obj);
  }
  return result;
}

function nowStr() {
  return Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', "yyyy-MM-dd'T'HH:mm:ss");
}

// ── doGet ────────────────────────────────────────────────────────
function doGet(e) {
  var p  = e ? (e.parameter || {}) : {};
  var cb = p.callback || '';
  var result;
  try {
    switch (p.action) {
      case 'getAll':
        result = getAll();
        break;
      case 'getSubcon':
        result = getSubcon(p.code);
        break;
      case 'submitSubconForm':
        result = submitSubconForm(p);
        break;
      case 'approveSubmission':
        result = approveSubmission(p.submissionId);
        break;
      case 'rejectSubmission':
        result = rejectSubmission(p.submissionId, p.reason);
        break;
      case 'stockIn':
        result = stockIn(p);
        break;
      case 'addQuotation':
        result = addQuotation(p);
        break;
      case 'getPayments':
        result = getPayments(p.subconCode);
        break;
      case 'getSubconRates':
        result = getSubconRates();
        break;
      case 'calculatePayment':
        result = calculatePayment(p.subconCode, p.quotationNo, Number(p.rollsInstalled) || 0);
        break;
      case 'markPayment':
        result = markPayment(p);
        break;
      case 'getQuotations':
        result = { success: true, quotations: getQuotations() };
        break;
      default:
        result = { success: false, error: 'Unknown action: ' + (p.action || '(none)') };
    }
  } catch (err) {
    result = { success: false, error: err.message || String(err) };
  }
  var json = JSON.stringify(result);
  var out  = cb ? cb + '(' + json + ')' : json;
  return ContentService.createTextOutput(out)
    .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════
// CONFIG
// ════════════════════════════════════════════════════════════════

function getConfig() {
  var sheet = getSheet('Stock');
  if (!sheet) return { warehouse: 0, avgCost: 0, totalQty: 0, totalCost: 0 };
  var data = sheet.getDataRange().getValues();
  var cfg  = {};
  for (var r = 0; r < data.length; r++) {
    if (data[r][0]) cfg[String(data[r][0])] = Number(data[r][1]) || 0;
  }
  return {
    warehouse:  cfg.warehouse  || 0,
    avgCost:    cfg.avg_cost   || 0,   // sheet key is avg_cost; returned as avgCost for UI compat
    totalQty:   cfg.total_qty  || 0,
    totalCost:  cfg.total_cost || 0
  };
}

// setConfigKey writes to the Stock sheet.
// Use the exact key strings stored in column A: 'warehouse', 'avg_cost', 'total_qty', 'total_cost'.
// Creates the row if it does not exist.
function setConfigKey(key, value) {
  var sheet = getSheet('Stock');
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  for (var r = 0; r < data.length; r++) {
    if (String(data[r][0]) === String(key)) {
      sheet.getRange(r + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

// ════════════════════════════════════════════════════════════════
// getAll — main data load for admin dashboard
// ════════════════════════════════════════════════════════════════

function getAll() {
  return {
    success:            true,
    stock:              getConfig(),
    subconBalances:     getSubconBalances(),
    pendingSubmissions: getPendingSubmissions(),
    quotations:         getQuotations(),
    recentLog:          getRecentLog(100)
  };
}

// ════════════════════════════════════════════════════════════════
// SUBCONS
// SubconBalances sheet columns (by position — header row not relied on):
//   A(0) SubconCode | B(1) SubconName | C(2) TotalPickup
//   D(3) TotalInstalled | E(4) Balance | F(5) LastUpdated
// ════════════════════════════════════════════════════════════════

function getSubconBalances() {
  var sheet = getSheet('SubconBalances');
  if (!sheet) return [];
  var rows = sheetToObjects(sheet);
  return rows.filter(function(r) { return r.SubconCode; }).map(function(r) {
    return {
      code:           String(r.SubconCode),
      name:           String(r.SubconName || SUBCONS[String(r.SubconCode)] || r.SubconCode),
      totalPickup:    Number(r.TotalPickup) || 0,
      totalInstalled: Number(r.TotalInstalled) || 0,
      balance:        Number(r.Balance) || 0,
      lastUpdated:    String(r.LastUpdated || '')
    };
  });
}

// getSubcon — used by the subcon mobile form to load own data
function getSubcon(code) {
  var sheet = getSheet('SubconBalances');
  if (!sheet) return { success: false, error: 'SubconBalances sheet not found' };
  var rows = sheetToObjects(sheet);

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (String(r.SubconCode).trim() !== String(code).trim()) continue;

    // Build active quotations list for the install dropdown
    var quotes  = getQuotations();
    var activeQ = [];
    for (var q = 0; q < quotes.length; q++) {
      var qt = quotes[q];
      if (qt.status === 'active' || qt.status === 'upcoming') {
        Logger.log('getSubcon qt keys: ' + JSON.stringify(qt));
        activeQ.push({ no: qt.quotationNo || qt.QuotationNo || '', project: qt.projectName || qt.ProjectName || '', client: qt.clientName || qt.ClientName || '' });
      }
    }

    return {
      success:        true,
      code:           String(r.SubconCode),
      name:           String(SUBCONS[String(r.SubconCode)] || r.SubconName || code),
      totalPickup:    Number(r.TotalPickup) || 0,
      totalInstalled: Number(r.TotalInstalled) || 0,
      balance:        Number(r.Balance) || 0,
      quotations:     activeQ
    };
  }
  return { success: false, error: 'Subcon not found: ' + code };
}

// updateSubconBalance — called by approveSubmission to keep SubconBalances in sync.
// Creates the row for the subcon if it does not yet exist.
function updateSubconBalance(subconCode, formType, qty) {
  var sheet = getSheet('SubconBalances');
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]) !== String(subconCode)) continue;
    var pickup    = Number(data[r][2]) || 0;
    var installed = Number(data[r][3]) || 0;
    if (formType === 'pickup') {
      pickup += qty;
    } else if (formType === 'install') {
      installed += qty;
    } else if (formType === 'return' || formType === 'returned') {
      pickup = Math.max(0, pickup - qty);
    }
    sheet.getRange(r + 1, 3).setValue(pickup);
    sheet.getRange(r + 1, 4).setValue(installed);
    sheet.getRange(r + 1, 5).setValue(pickup - installed);
    sheet.getRange(r + 1, 6).setValue(new Date());
    return;
  }
  // Row missing — create it
  var name       = SUBCONS[subconCode] || subconCode;
  var newPickup  = (formType === 'pickup')  ? qty : 0;
  var newInstall = (formType === 'install') ? qty : 0;
  sheet.appendRow([subconCode, name, newPickup, newInstall, newPickup - newInstall, new Date()]);
}

// ════════════════════════════════════════════════════════════════
// SUBMISSIONS
// ════════════════════════════════════════════════════════════════

function getPendingSubmissions() {
  var sheet = getSheet('Submissions');
  if (!sheet) return [];
  var rows = sheetToObjects(sheet);
  var result = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (String(r.Status || '').trim().toLowerCase() !== 'pending') continue;
    result.push({
      id:          r.SubmissionID   || '',
      timestamp:   r.Timestamp      || '',
      subconCode:  r.SubconCode     || '',
      subconName:  r.SubconName     || '',
      formType:    r.FormType       || '',
      quotationNo: r.QuotationNo    || '',
      qty:         Number(r.Qty)    || 0,
      date:        r.ActivityDate   || '',
      notes:       r.Notes          || '',
      status:      r.Status         || ''
    });
  }
  return result;
}

function submitSubconForm(p) {
  var sheet = getSheet('Submissions');
  if (!sheet) return { success: false, error: 'Submissions sheet not found' };

  var subId = 'SUB-' + Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'yyyyMMddHHmmss') +
              '-' + Math.floor(Math.random() * 1000).toString().padStart(3, '0');

  // Columns: A=Timestamp B=SubconCode C=SubconName D=FormType E=QuotationNo
  //          F=Qty G=ActivityDate H=Notes I=PhotoURL J=Status
  //          K=ApprovedBy L=ApprovedAt M=RejectionReason N=SubmissionID
  sheet.appendRow([
    new Date(),           // A  Timestamp
    p.subconCode  || '',  // B  SubconCode
    p.subconName  || '',  // C  SubconName
    p.formType    || '',  // D  FormType
    p.quotationNo || '',  // E  QuotationNo
    Number(p.qty) || 0,   // F  Qty
    p.date        || '',  // G  ActivityDate
    p.notes       || '',  // H  Notes
    p.photoURL    || '',  // I  PhotoURL
    'pending',            // J  Status
    '',                   // K  ApprovedBy
    '',                   // L  ApprovedAt
    '',                   // M  RejectionReason
    subId                 // N  SubmissionID
  ]);

  return { success: true, submissionId: subId };
}

function approveSubmission(submissionId) {
  var sheet = getSheet('Submissions');
  if (!sheet) return { success: false, error: 'Submissions sheet not found' };
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var idx     = {};
  headers.forEach(function(h, i) { idx[h] = i; });

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (String(row[idx['SubmissionID']]) !== String(submissionId)) continue;
    if (String(row[idx['Status']]) !== 'pending') {
      return { success: false, error: 'Submission already processed' };
    }

    var formType   = String(row[idx['FormType']]);
    var qty        = Number(row[idx['Qty']]) || 0;
    var subconCode = row[idx['SubconCode']];
    var subconName = row[idx['SubconName']];
    var quotNo     = row[idx['QuotationNo']];
    var notes      = row[idx['Notes']] || '';

    // Update warehouse stock (pickup draws from warehouse; return puts back)
    var cfg = getConfig();
    if (formType === 'pickup') {
      setConfigKey('warehouse', Math.max(0, (cfg.warehouse || 0) - qty));
    } else if (formType === 'return' || formType === 'returned') {
      setConfigKey('warehouse', (cfg.warehouse || 0) + qty);
    }

    // Update SubconBalances sheet totals
    updateSubconBalance(subconCode, formType, qty);

    // Mark submission as approved FIRST so payment calc includes this row
    sheet.getRange(r + 1, idx['Status'] + 1).setValue('approved');
    SpreadsheetApp.flush(); // ensure write is visible to subsequent reads

    // Update quotation installed count; recalculate payment from all approved installs
    if (formType === 'install' && quotNo) {
      updateQuotationInstalled(quotNo, qty);
      var payCalc = calculatePaymentForQuotation(quotNo, String(subconCode));
      if (payCalc.success) {
        upsertPaymentRecord(String(quotNo), String(subconCode), payCalc);
      }
    }

    // Log movement
    addLog({ type: formType, subconCode: subconCode, subconName: subconName,
             quotationNo: quotNo, qty: qty, notes: notes });

    return { success: true };
  }
  return { success: false, error: 'Submission not found: ' + submissionId };
}

function rejectSubmission(submissionId, reason) {
  var sheet = getSheet('Submissions');
  if (!sheet) return { success: false, error: 'Submissions sheet not found' };
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var idx     = {};
  headers.forEach(function(h, i) { idx[h] = i; });

  for (var r = 1; r < data.length; r++) {
    if (String(data[r][idx['SubmissionID']]) !== String(submissionId)) continue;
    var savedReason = reason || 'Rejected by admin';
    sheet.getRange(r + 1, idx['Status'] + 1).setValue('rejected');
    if (idx['RejectionReason'] !== undefined) {
      sheet.getRange(r + 1, idx['RejectionReason'] + 1).setValue(savedReason);
    }
    return {
      success:     true,
      reason:      savedReason,
      quotationNo: String(data[r][idx['QuotationNo']] || ''),
      subconCode:  String(data[r][idx['SubconCode']]  || ''),
      formType:    String(data[r][idx['FormType']]     || '')
    };
  }
  return { success: false, error: 'Submission not found: ' + submissionId };
}

// ════════════════════════════════════════════════════════════════
// QUOTATIONS
// ════════════════════════════════════════════════════════════════

function getQuotations() {
  var sheet = getSheet('Quotations');
  if (!sheet) return [];
  var data    = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  var result  = [];
  for (var r = 1; r < data.length; r++) {
    if (!data[r][0]) continue;
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      var v = data[r][c];
      obj[headers[c].charAt(0).toLowerCase() + headers[c].slice(1)] =
        v instanceof Date ? v.toISOString() : v;
    }
    result.push(obj);
  }
  return result;
}

function addQuotation(p) {
  var sheet  = getSheet('Quotations');
  if (!sheet) return { success: false, error: 'Quotations sheet not found' };
  var quotNo = p.quotationNo || '';
  if (!quotNo) return { success: false, error: 'Quotation number required' };

  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var qIdx    = headers.indexOf('QuotationNo');
  var iIdx    = headers.indexOf('RollsInstalled');
  var sIdx    = headers.indexOf('Status');
  var bIdx    = headers.indexOf('Blocks');

  // Columns: QuotationNo | Date | ClientName | ProjectName | SiteAddress |
  //          MembraneType | RatePerSqft | TotalSqft | EstRolls |
  //          MembraneValue | TotalValue | Blocks | RollsInstalled | Status
  var row = [
    quotNo,                                                                  // 0  QuotationNo
    p.date         || Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'yyyy-MM-dd'), // 1  Date
    p.clientName   || '',                                                    // 2  ClientName
    p.projectName  || '',                                                    // 3  ProjectName
    p.siteAddress  || '',                                                    // 4  SiteAddress
    p.membraneType || '',                                                    // 5  MembraneType
    parseFloat(p.ratePerSqft)   || 0,                                        // 6  RatePerSqft
    parseFloat(p.totalSqft)     || 0,                                        // 7  TotalSqft
    Math.ceil((parseFloat(p.totalSqft) || 0) / 80),                          // 8  EstRolls
    parseFloat(p.membraneValue) || 0,                                        // 9  MembraneValue
    parseFloat(p.totalValue)    || 0,                                        // 10 TotalValue
    p.blocks       || '',                                                    // 11 Blocks
    0,                                                                       // 12 RollsInstalled (preserved on update)
    p.status || 'active'                                                     // 13 Status
  ];

  for (var r = 1; r < data.length; r++) {
    if (String(data[r][qIdx]) === String(quotNo)) {
      // Preserve RollsInstalled, SiteAddress and Blocks from existing row
      if (iIdx >= 0) row[12] = Number(data[r][iIdx]) || 0;
      if (!row[4]  && data[r][4])  row[4]  = data[r][4];   // keep existing SiteAddress
      if (!row[11] && bIdx >= 0)   row[11] = data[r][bIdx]; // keep existing Blocks
      sheet.getRange(r + 1, 1, 1, row.length).setValues([row]);
      return { success: true, updated: true };
    }
  }

  sheet.appendRow(row);
  return { success: true, created: true };
}

function updateQuotationInstalled(quotationNo, addQty) {
  var sheet = getSheet('Quotations');
  if (!sheet) return 0;
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var qIdx    = headers.indexOf('QuotationNo');
  var iIdx    = headers.indexOf('RollsInstalled');
  if (iIdx < 0) return 0;

  for (var r = 1; r < data.length; r++) {
    if (String(data[r][qIdx]) === String(quotationNo)) {
      var newTotal = (Number(data[r][iIdx]) || 0) + addQty;
      sheet.getRange(r + 1, iIdx + 1).setValue(newTotal);
      return newTotal;
    }
  }
  return 0;
}

// ════════════════════════════════════════════════════════════════
// STOCK IN
// ════════════════════════════════════════════════════════════════

function stockIn(p) {
  var qty  = parseInt(p.qty)           || 0;
  var cost = parseFloat(p.costPerRoll) || 0;
  if (qty < 1) return { success: false, error: 'Invalid quantity' };

  // ── 1. Calculate new Stock values ──────────────────────────────
  var cfg          = getConfig();
  var newWh        = (cfg.warehouse || 0) + qty;
  var newTotalQty  = (cfg.totalQty  || 0) + qty;
  var newTotalCost = (cfg.totalCost || 0) + (qty * cost);
  var newAvgCost   = newTotalQty > 0 ? newTotalCost / newTotalQty : 0;

  // ── 2. Write back to Stock sheet (upserts; creates rows if missing) ──
  setConfigKey('warehouse',  newWh);
  setConfigKey('avg_cost',   newAvgCost);
  setConfigKey('total_qty',  newTotalQty);
  setConfigKey('total_cost', newTotalCost);

  // ── 3. Append delivery record to StockIn sheet ────────────────
  var siSheet = getSheet('StockIn');
  if (siSheet) {
    siSheet.appendRow([
      new Date(),
      p.supplier     || '',
      p.membraneType || '',
      p.doNumber     || '',
      qty,
      cost,
      qty * cost,
      newWh
    ]);
  }

  // ── 4. Also append to movement Log ───────────────────────────
  var logSheet = getSheet('Log');
  if (logSheet) {
    logSheet.appendRow([
      new Date(), 'in', '', p.supplier || '',
      p.doNumber || '', qty, p.membraneType || '', cost
    ]);
  }

  return { success: true, warehouse: newWh, avgCost: newAvgCost };
}

// ════════════════════════════════════════════════════════════════
// LOG
// ════════════════════════════════════════════════════════════════

function getRecentLog(limit) {
  var sheet = getSheet('Log');
  if (!sheet) return [];
  var data    = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  var result  = [];
  for (var r = data.length - 1; r >= 1; r--) {
    if (!data[r][0]) continue;
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      var v = data[r][c];
      obj[headers[c].charAt(0).toLowerCase() + headers[c].slice(1)] =
        v instanceof Date ? v.toISOString() : v;
    }
    result.push(obj);
    if (limit && result.length >= limit) break;
  }
  return result;
}

function addLog(entry) {
  var sheet = getSheet('Log');
  if (!sheet) return;
  sheet.appendRow([
    new Date(),
    entry.type        || '',
    entry.subconCode  || '',
    entry.subconName  || '',
    entry.quotationNo || '',
    entry.qty         || 0,
    entry.notes       || '',
    entry.costPerRoll || 0
  ]);
}

// ════════════════════════════════════════════════════════════════
// PAYMENT FUNCTIONS
// ════════════════════════════════════════════════════════════════

// ── calculateTieredRate ──────────────────────────────────────────
// Returns rate per roll, total, payment1 (50%), payment2 (50%)
// Applies flat tier based on total rolls (not accumulative)
function calculateTieredRate(subconCode, rollsInstalled) {
  var sheet = getSheet('SubconRates');
  if (!sheet) return { success: false, error: 'SubconRates sheet not found. Run setupPaymentSheets() first.' };

  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var idx     = {};
  headers.forEach(function(h, i) { idx[h] = i; });

  var rolls = Number(rollsInstalled) || 0;

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (String(row[idx['SubconCode']]) !== String(subconCode)) continue;

    // Sheet columns: Tier1MaxRolls | Tier1Rate | Tier2MinRolls | Tier2MaxRolls | Tier2Rate | Tier3Rate
    var t1Max  = Number(row[idx['Tier1MaxRolls']]) || 4;
    var t1Rate = Number(row[idx['Tier1Rate']])     || 190;
    var t2Min  = Number(row[idx['Tier2MinRolls']]) || 5;
    var t2Max  = Number(row[idx['Tier2MaxRolls']]) || 9;
    var t2Rate = Number(row[idx['Tier2Rate']])     || 170;
    var t3Rate = Number(row[idx['Tier3Rate']])     || 150;

    var rate;
    if (rolls <= t1Max)                        { rate = t1Rate; }
    else if (rolls >= t2Min && rolls <= t2Max) { rate = t2Rate; }
    else if (rolls > t2Max)                    { rate = t3Rate; }
    else                                       { rate = t1Rate; }

    var total = rolls * rate;
    return { success: true, rate: rate, total: total, payment1: total * 0.5, payment2: total * 0.5 };
  }

  return { success: false, error: 'Subcon not found in rates: ' + subconCode };
}

// ── getSubconRates ────────────────────────────────────────────────
function getSubconRates() {
  var sheet = getSheet('SubconRates');
  if (!sheet) return { success: false, error: 'SubconRates sheet not found' };
  return { success: true, rates: sheetToObjects(sheet) };
}

// ── getPayments ───────────────────────────────────────────────────
function getPayments(subconCode) {
  var sheet = getSheet('Payments');
  if (!sheet) return { success: false, error: 'Payments sheet not found' };

  var all = sheetToObjects(sheet);
  var payments = subconCode
    ? all.filter(function(p) { return String(p['SubconCode']) === String(subconCode); })
    : all;

  // Enrich with project/client name from Quotations
  var qtList = getQuotations();
  var qtMap  = {};
  qtList.forEach(function(q) { qtMap[q.quotationNo] = q; });

  payments.forEach(function(pay) {
    var q = qtMap[pay['QuotationNo']];
    if (q) {
      pay.projectName = q.projectName || '';
      pay.clientName  = q.clientName  || '';
    }
  });

  return { success: true, payments: payments };
}

// ── calculatePayment ──────────────────────────────────────────────
function calculatePayment(subconCode, quotationNo, rollsInstalled) {
  var calc = calculateTieredRate(subconCode, rollsInstalled);
  if (!calc.success) return calc;
  return {
    success:        true,
    subconCode:     subconCode,
    quotationNo:    quotationNo,
    rollsInstalled: rollsInstalled,
    rate:           calc.rate,
    total:          calc.total,
    payment1:       calc.payment1,
    payment2:       calc.payment2
  };
}

// ── calculatePaymentForQuotation ──────────────────────────────────
// Sums qty from ALL approved install submissions for a quotation,
// then applies tiered rate to the total.
function calculatePaymentForQuotation(quotationNo, subconCode) {
  var sheet = getSheet('Submissions');
  if (!sheet) return { success: false, error: 'Submissions sheet not found' };
  var rows = sheetToObjects(sheet);
  var totalRolls = 0;
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (String(r.QuotationNo)  !== String(quotationNo)) continue;
    if (String(r.SubconCode)   !== String(subconCode))  continue;
    if (String(r.FormType)     !== 'install')            continue;
    if (String(r.Status).trim().toLowerCase() !== 'approved') continue;
    totalRolls += Number(r.Qty) || 0;
  }
  if (totalRolls === 0) return { success: false, error: 'No approved installs for ' + quotationNo };
  var calc = calculateTieredRate(subconCode, totalRolls);
  if (!calc.success) return calc;
  calc.totalRolls = totalRolls;
  return calc;
}

// ── upsertPaymentRecord ──────────────────────────────────────────
// Creates or updates the payment record for a quotation.
// On update: recalculates amounts but preserves existing payment statuses.
function upsertPaymentRecord(quotationNo, subconCode, calc) {
  var sheet = getSheet('Payments');
  if (!sheet) return { success: false, error: 'Payments sheet not found' };
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var idx     = {};
  headers.forEach(function(h, i) { idx[h] = i; });
  var qIdx = idx['QuotationNo'];

  // Check if record exists — update in place
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][qIdx]) !== String(quotationNo)) continue;
    // Update rolls, rate, amounts — preserve payment statuses/dates/refs
    sheet.getRange(r + 1, idx['RollsInstalled'] + 1).setValue(calc.totalRolls);
    sheet.getRange(r + 1, idx['RateApplied']    + 1).setValue(calc.rate);
    sheet.getRange(r + 1, idx['TotalAmount']    + 1).setValue(calc.total);
    sheet.getRange(r + 1, idx['Payment1Amount'] + 1).setValue(calc.payment1);
    sheet.getRange(r + 1, idx['Payment2Amount'] + 1).setValue(calc.payment2);
    return { success: true, updated: true, paymentID: data[r][idx['PaymentID']] };
  }

  // Create new record
  var subconName = SUBCONS[subconCode] || subconCode;
  var subSheet   = getSheet('SubconBalances');
  if (subSheet) {
    var subRows = sheetToObjects(subSheet);
    for (var s = 0; s < subRows.length; s++) {
      if (String(subRows[s].SubconCode) === String(subconCode)) {
        subconName = String(subRows[s].SubconName || subconName);
        break;
      }
    }
  }

  var paymentID = 'PAY-' + Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'yyyyMMddHHmmss');
  sheet.appendRow([
    paymentID,         // PaymentID
    quotationNo,       // QuotationNo
    subconCode,        // SubconCode
    subconName,        // SubconName
    calc.totalRolls,   // RollsInstalled
    calc.rate,         // RateApplied
    calc.total,        // TotalAmount
    calc.payment1,     // Payment1Amount
    'unpaid',          // Payment1Status
    '',                // Payment1Date
    '',                // Payment1Reference
    calc.payment2,     // Payment2Amount
    'unpaid',          // Payment2Status
    '',                // Payment2Date
    '',                // Payment2Reference
    nowStr()           // CreatedAt
  ]);
  return { success: true, created: true, paymentID: paymentID };
}

// ── createPaymentRecord ───────────────────────────────────────────
// Legacy wrapper — kept for backward compatibility.
function createPaymentRecord(quotationNo, subconCode, rollsInstalled) {
  var sheet = getSheet('Payments');
  if (!sheet) return { success: false, error: 'Payments sheet not found' };

  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var qIdx    = headers.indexOf('QuotationNo');

  // Skip if record already exists for this quotation
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][qIdx]) === String(quotationNo)) {
      return { success: true, message: 'Record already exists', paymentID: data[r][0] };
    }
  }

  var calc = calculateTieredRate(subconCode, rollsInstalled);
  if (!calc.success) return calc;

  // Resolve subcon name from SubconBalances (col A = code, col B = name)
  var subconName = SUBCONS[subconCode] || subconCode;
  var subSheet   = getSheet('SubconBalances');
  if (subSheet) {
    var subData = subSheet.getDataRange().getValues();
    for (var s = 1; s < subData.length; s++) {
      if (String(subData[s][0]) === String(subconCode)) {
        subconName = String(subData[s][1] || subconName);
        break;
      }
    }
  }

  var paymentID = 'PAY-' + Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'yyyyMMddHHmmss');

  sheet.appendRow([
    paymentID,        // PaymentID
    quotationNo,      // QuotationNo
    subconCode,       // SubconCode
    subconName,       // SubconName
    rollsInstalled,   // RollsInstalled
    calc.rate,        // RateApplied
    calc.total,       // TotalAmount
    calc.payment1,    // Payment1Amount
    'unpaid',         // Payment1Status
    '',               // Payment1Date
    '',               // Payment1Reference
    calc.payment2,    // Payment2Amount
    'unpaid',         // Payment2Status
    '',               // Payment2Date
    '',               // Payment2Reference
    nowStr()          // CreatedAt
  ]);

  return {
    success:   true,
    paymentID: paymentID,
    rate:      calc.rate,
    total:     calc.total,
    payment1:  calc.payment1,
    payment2:  calc.payment2
  };
}

// ── markPayment ───────────────────────────────────────────────────
// Updates Payment 1 or Payment 2 fields. Status values: unpaid, paid, partial
function markPayment(p) {
  var paymentID     = p.paymentID;
  var paymentNumber = String(p.paymentNumber);
  var status        = p.status    || 'paid';
  var date          = p.date      || '';
  var reference     = p.reference || '';

  if (!paymentID || !paymentNumber) {
    return { success: false, error: 'Missing paymentID or paymentNumber' };
  }

  var sheet = getSheet('Payments');
  if (!sheet) return { success: false, error: 'Payments sheet not found' };

  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var pidIdx  = headers.indexOf('PaymentID');

  for (var r = 1; r < data.length; r++) {
    if (String(data[r][pidIdx]) !== String(paymentID)) continue;

    if (paymentNumber === '1') {
      var s1 = headers.indexOf('Payment1Status');
      var d1 = headers.indexOf('Payment1Date');
      var r1 = headers.indexOf('Payment1Reference');
      if (s1 >= 0) sheet.getRange(r + 1, s1 + 1).setValue(status);
      if (d1 >= 0) sheet.getRange(r + 1, d1 + 1).setValue(date);
      if (r1 >= 0) sheet.getRange(r + 1, r1 + 1).setValue(reference);
    } else if (paymentNumber === '2') {
      var s2 = headers.indexOf('Payment2Status');
      var d2 = headers.indexOf('Payment2Date');
      var r2 = headers.indexOf('Payment2Reference');
      if (s2 >= 0) sheet.getRange(r + 1, s2 + 1).setValue(status);
      if (d2 >= 0) sheet.getRange(r + 1, d2 + 1).setValue(date);
      if (r2 >= 0) sheet.getRange(r + 1, r2 + 1).setValue(reference);
    }

    return { success: true };
  }

  return { success: false, error: 'Payment record not found: ' + paymentID };
}

// ════════════════════════════════════════════════════════════════
// SETUP — Run these once from the Script Editor after deploying.
// ════════════════════════════════════════════════════════════════

// setupStock() — creates/initialises the Stock and StockIn sheets.
// Run this ONCE if the Stock sheet is empty or missing.
function setupStock() {
  var ss = getSpreadsheet();

  // ── Stock sheet (key-value config) ──
  var stockSh = ss.getSheetByName('Stock');
  if (!stockSh) {
    stockSh = ss.insertSheet('Stock');
    Logger.log('Stock sheet created.');
  }
  var stockData = stockSh.getDataRange().getValues();
  var existing  = {};
  stockData.forEach(function(row) { if (row[0]) existing[String(row[0])] = true; });

  var requiredKeys = ['warehouse', 'avg_cost', 'total_qty', 'total_cost'];
  requiredKeys.forEach(function(k) {
    if (!existing[k]) {
      stockSh.appendRow([k, 0]);
      Logger.log('Added Stock row: ' + k);
    } else {
      Logger.log('Stock row already exists: ' + k);
    }
  });

  // ── StockIn sheet (delivery log) ──
  var siSh = ss.getSheetByName('StockIn');
  if (!siSh) {
    siSh = ss.insertSheet('StockIn');
    siSh.appendRow([
      'Timestamp', 'Supplier', 'MembraneType', 'DONumber',
      'Qty', 'CostPerRoll', 'TotalCost', 'WarehouseAfter'
    ]);
    Logger.log('StockIn sheet created.');
  } else {
    Logger.log('StockIn sheet already exists — skipping creation.');
  }

  Logger.log('setupStock() complete.');
}

function setupPaymentSheets() {
  var ss = getSpreadsheet();

  // ── SubconRates ──
  var rateSh = ss.getSheetByName('SubconRates');
  if (!rateSh) {
    rateSh = ss.insertSheet('SubconRates');
    rateSh.appendRow([
      'SubconCode','SubconName',
      'Tier1MaxRolls','Tier1Rate',
      'Tier2MinRolls','Tier2MaxRolls','Tier2Rate',
      'Tier3MinRolls','Tier3Rate'
    ]);
    rateSh.appendRow(['SC01','Team Atik', 4, 190, 5, 9, 170, 10, 150]);
    Logger.log('SubconRates sheet created with SC01 data.');
  } else {
    Logger.log('SubconRates sheet already exists — skipping creation.');
  }

  // ── Payments ──
  var paySh = ss.getSheetByName('Payments');
  if (!paySh) {
    paySh = ss.insertSheet('Payments');
    paySh.appendRow([
      'PaymentID','QuotationNo','SubconCode','SubconName',
      'RollsInstalled','RateApplied','TotalAmount',
      'Payment1Amount','Payment1Status','Payment1Date','Payment1Reference',
      'Payment2Amount','Payment2Status','Payment2Date','Payment2Reference',
      'CreatedAt'
    ]);
    Logger.log('Payments sheet created.');
  } else {
    Logger.log('Payments sheet already exists — skipping creation.');
  }

  Logger.log('Setup complete. Redeploy the web app to activate the new endpoints.');
}
