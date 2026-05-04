// ════════════════════════════════════════════════════════════════
// RollTrack Pro — Google Apps Script
// Innovation AID Sdn Bhd
// ════════════════════════════════════════════════════════════════

var SPREADSHEET_ID = '1O3Hvc0D-wMcBKLcC5IQKAboR1maFI2QuSi6XlIFZq9U';

var SUBCONS = { 'SC01': 'Md Atik', 'SC02': 'Md Shahazan', 'SC03': 'Md Mohiuddin', 'SC04': 'Md Foysel' };

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
      case 'approveSubmissionWithEdit':
        result = approveSubmissionWithEdit(p.submissionId, p.additionalCosts, p.editReason);
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
      case 'recalcPayment':
        result = recalcPayment(p.quotationNo, p.subconCode);
        break;
      case 'undoApproval':
        result = undoApproval(p.quotationNo, p.subconCode);
        break;
      case 'getAllSubmissions':
        result = getAllSubmissions();
        break;
      case 'getQuotations':
        result = { success: true, quotations: getQuotations() };
        break;
      case 'login':
        result = handleLogin(p);
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

    // Build active quotations list for the install dropdown — filtered by assigned subcon
    var quotes  = getQuotations();
    var activeQ = [];
    for (var q = 0; q < quotes.length; q++) {
      var qt = quotes[q];
      if (qt.status !== 'active' && qt.status !== 'upcoming') continue;
      var assigned = qt.assignedSubcon || '';
      if (assigned && assigned !== String(code).trim()) continue; // skip if assigned to different subcon
      activeQ.push({ no: qt.quotationNo || '', project: qt.projectName || '', client: qt.clientName || '', assignedSubcon: assigned });
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
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  var puIdx = headers.indexOf('PUSealant');
  if (puIdx < 0) puIdx = 14; // fallback to column O
  var acIdx = headers.indexOf('AdditionalCosts');
  if (acIdx < 0) acIdx = 15; // fallback to column P
  var rows = sheetToObjects(sheet);
  var result = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (String(r.Status || '').trim().toLowerCase() !== 'pending') continue;
    var ac = [];
    try { var raw = r.AdditionalCosts || data[i + 1][acIdx] || ''; if (raw) ac = JSON.parse(raw); } catch(e) {}
    var origAC = [];
    try { if (r.OriginalAdditionalCosts) origAC = JSON.parse(r.OriginalAdditionalCosts) || []; } catch(e) {}
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
      status:      r.Status         || '',
      rejectionReason: r.RejectionReason || '',
      originalAdditionalCosts: origAC,
      editedBy:    r.EditedBy       || '',
      editedAt:    r.EditedAt       || '',
      editReason:  r.EditReason     || '',
      puSealant:   Number(r.PUSealant || data[i + 1][puIdx]) || 0,
      additionalCosts: ac
    });
  }
  return result;
}

function getAllSubmissions() {
  var sheet = getSheet('Submissions');
  if (!sheet) return { success: true, submissions: [] };
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, submissions: [] };
  var headers = data[0];
  var puIdx = headers.indexOf('PUSealant');
  if (puIdx < 0) puIdx = 14;
  var acIdx = headers.indexOf('AdditionalCosts');
  if (acIdx < 0) acIdx = 15;
  var rows = sheetToObjects(sheet);
  var result = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var ac = [];
    try { var raw = r.AdditionalCosts || data[i + 1][acIdx] || ''; if (raw) ac = JSON.parse(raw); } catch(e) {}
    var origAC2 = [];
    try { if (r.OriginalAdditionalCosts) origAC2 = JSON.parse(r.OriginalAdditionalCosts) || []; } catch(e) {}
    result.push({
      id:           r.SubmissionID   || '',
      timestamp:    r.Timestamp      || '',
      subconCode:   r.SubconCode     || '',
      subconName:   r.SubconName     || '',
      formType:     r.FormType       || '',
      quotationNo:  r.QuotationNo    || '',
      qty:          Number(r.Qty)    || 0,
      activityDate: r.ActivityDate   || '',
      notes:        r.Notes          || '',
      status:       r.Status         || '',
      rejectionReason: r.RejectionReason || '',
      originalAdditionalCosts: origAC2,
      editedBy:     r.EditedBy       || '',
      editedAt:     r.EditedAt       || '',
      editReason:   r.EditReason     || '',
      puSealant:    Number(r.PUSealant || data[i + 1][puIdx]) || 0,
      additionalCosts: ac
    });
  }
  return { success: true, submissions: result };
}

function submitSubconForm(p) {
  var sheet = getSheet('Submissions');
  if (!sheet) return { success: false, error: 'Submissions sheet not found' };

  var subId = 'SUB-' + Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'yyyyMMddHHmmss') +
              '-' + Math.floor(Math.random() * 1000).toString().padStart(3, '0');

  // Columns: A=Timestamp B=SubconCode C=SubconName D=FormType E=QuotationNo
  //          F=Qty G=ActivityDate H=Notes I=PhotoURL J=Status
  //          K=ApprovedBy L=ApprovedAt M=RejectionReason N=SubmissionID
  //          O=PUSealant P=AdditionalCosts
  var additionalCosts = '[]';
  if (p.additionalCosts) {
    try { JSON.parse(p.additionalCosts); additionalCosts = p.additionalCosts; } catch(e) { additionalCosts = '[]'; }
  }
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
    subId,                // N  SubmissionID
    Number(p.puSealant) || 0,  // O  PUSealant
    additionalCosts       // P  AdditionalCosts (JSON)
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

    // Update quotation installed count; create/update payment record for installs
    if ((formType === 'install' || formType === 'Install') && quotNo) {
      updateQuotationInstalled(quotNo, qty);
      try {
        var payCalc = calculatePaymentForQuotation(quotNo, String(subconCode));
        if (payCalc.success) {
          upsertPaymentRecord(String(quotNo), String(subconCode), payCalc);
        } else {
          // Fallback: create payment from this submission's qty alone
          createPaymentRecord(String(quotNo), String(subconCode), String(subconName), qty);
        }
      } catch(e) {
        Logger.log('Payment calc error, falling back to createPaymentRecord: ' + e);
        try {
          createPaymentRecord(String(quotNo), String(subconCode), String(subconName), qty);
        } catch(e2) {
          Logger.log('createPaymentRecord error: ' + e2);
        }
      }
    }

    // Auto-complete quotation status on install approval
    if ((formType === 'install' || formType === 'Install') && quotNo) {
      var qtSheet = getSheet('Quotations');
      if (qtSheet) {
        var qtData = qtSheet.getDataRange().getValues();
        for (var qi = 1; qi < qtData.length; qi++) {
          if (String(qtData[qi][0]).trim() === String(quotNo).trim()) {
            qtSheet.getRange(qi + 1, 14).setValue('completed');
            break;
          }
        }
      }
    }

    // Log movement
    addLog({ type: formType, subconCode: subconCode, subconName: subconName,
             quotationNo: quotNo, qty: qty, notes: notes });

    return { success: true };
  }
  return { success: false, error: 'Submission not found: ' + submissionId };
}

// approveSubmissionWithEdit — admin edits the submission's additional costs
// during the approval click. The submission's original AdditionalCosts are
// snapshotted into OriginalAdditionalCosts; the new value is written back to
// AdditionalCosts; EditedBy / EditedAt / EditReason record the action; then
// the standard approveSubmission pipeline runs so the payment record + balances
// reflect the edited values.
function approveSubmissionWithEdit(submissionId, additionalCostsJson, editReason) {
  var sheet = getSheet('Submissions');
  if (!sheet) return { success: false, error: 'Submissions sheet not found' };
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var idx     = {};
  headers.forEach(function(h, i) { idx[h] = i; });

  // Auto-add edit-tracking columns if missing.
  ['OriginalAdditionalCosts', 'EditedBy', 'EditedAt', 'EditReason'].forEach(function(name) {
    if (idx[name] === undefined) {
      var nextCol = headers.length + 1;
      sheet.getRange(1, nextCol).setValue(name);
      idx[name] = headers.length;
      headers.push(name);
    }
  });

  for (var r = 1; r < data.length; r++) {
    if (String(data[r][idx['SubmissionID']]) !== String(submissionId)) continue;

    // Validate the new payload as JSON; reject silently if malformed.
    var newAC = '[]';
    if (additionalCostsJson) {
      try { JSON.parse(additionalCostsJson); newAC = additionalCostsJson; }
      catch(e) { return { success: false, error: 'Invalid additionalCosts JSON' }; }
    }

    // Snapshot original (only the first time; don't clobber on re-edit).
    var existingOriginal = String(data[r][idx['OriginalAdditionalCosts']] || '').trim();
    if (!existingOriginal) {
      var originalAC = data[r][idx['AdditionalCosts']] || '[]';
      sheet.getRange(r + 1, idx['OriginalAdditionalCosts'] + 1).setValue(originalAC);
    }

    // Write the edited value into AdditionalCosts so the calc pipeline picks it up.
    sheet.getRange(r + 1, idx['AdditionalCosts'] + 1).setValue(newAC);
    sheet.getRange(r + 1, idx['EditedBy']        + 1).setValue('admin');
    sheet.getRange(r + 1, idx['EditedAt']        + 1).setValue(new Date());
    if (editReason) {
      sheet.getRange(r + 1, idx['EditReason'] + 1).setValue(editReason);
    }
    SpreadsheetApp.flush();

    // Hand off to the normal approval pipeline; payment calc reads AdditionalCosts fresh.
    return approveSubmission(submissionId);
  }
  return { success: false, error: 'Submission not found: ' + submissionId };
}

// REJECTION FLOW:
// 1. Submission status → 'rejected'
// 2. Rejection reason saved
// 3. NO payment record created
// 4. Quotation stays 'active' — subcon can resubmit
// 5. SubconBalances unchanged
// 6. Subcon sees rejection in MY HISTORY with reason
// 7. Quotation reappears in subcon dropdown for resubmission
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
    // J = Status → 'rejected'
    sheet.getRange(r + 1, idx['Status'] + 1).setValue('rejected');
    // K = ApprovedBy (who actioned the rejection)
    if (idx['ApprovedBy'] !== undefined) {
      sheet.getRange(r + 1, idx['ApprovedBy'] + 1).setValue('admin');
    }
    // L = ApprovedAt (when rejected)
    if (idx['ApprovedAt'] !== undefined) {
      sheet.getRange(r + 1, idx['ApprovedAt'] + 1).setValue(new Date());
    }
    // M = RejectionReason
    if (idx['RejectionReason'] !== undefined) {
      sheet.getRange(r + 1, idx['RejectionReason'] + 1).setValue(savedReason);
    }
    // No payment record, no Quotations status change, no SubconBalances mutation.
    return { success: true, message: 'Rejected' };
  }
  return { success: false, error: 'Submission not found: ' + submissionId };
}

// ── undoApproval ──────────────────────────────────────────────────
// Reverts an approved install for (quotationNo, subconCode):
//   - Refuse if the Payments row's P1 or P2 is already 'paid'.
//   - Submissions row(s) for that pair: Status approved → pending,
//     clear ApprovedBy / ApprovedAt. Sum the qty reverted.
//   - SubconBalances: decrement TotalInstalled (via updateSubconBalance,
//     which is pure-additive over qty so a negative qty just subtracts).
//   - Quotations: decrement RollsInstalled. Demote 'completed' → 'active'
//     if RollsInstalled drops back below estRolls.
//   - Payments: delete the now-empty (qno, scCode) row.
//   - Append an 'undo_approval' entry to the Log.
//
// Defensive: handles N>=1 approved install rows in case legacy data
// violates the user's "1 quotation = 1 submission" rule.
function undoApproval(quotationNo, subconCode) {
  if (!quotationNo || !subconCode) {
    return { success: false, error: 'Missing quotationNo or subconCode' };
  }
  var paySheet = getSheet('Payments');
  if (!paySheet) return { success: false, error: 'Payments sheet not found' };
  var payData    = paySheet.getDataRange().getValues();
  var payHeaders = payData[0];
  var payQIdx    = payHeaders.indexOf('QuotationNo');
  var paySCIdx   = payHeaders.indexOf('SubconCode');
  var payP1Idx   = payHeaders.indexOf('Payment1Status');
  var payP2Idx   = payHeaders.indexOf('Payment2Status');
  if (payQIdx < 0 || paySCIdx < 0 || payP1Idx < 0 || payP2Idx < 0) {
    return { success: false, error: 'Payments sheet missing expected headers' };
  }

  var payRow = -1;
  for (var pr = 1; pr < payData.length; pr++) {
    if (String(payData[pr][payQIdx])  === String(quotationNo) &&
        String(payData[pr][paySCIdx]) === String(subconCode)) {
      payRow = pr;
      break;
    }
  }
  if (payRow < 0) {
    return { success: false, error: 'No payment record for ' + quotationNo + ' / ' + subconCode };
  }
  var p1st = String(payData[payRow][payP1Idx] || '').trim().toLowerCase();
  var p2st = String(payData[payRow][payP2Idx] || '').trim().toLowerCase();
  if (p1st === 'paid' || p2st === 'paid') {
    return { success: false, error: 'Payment already disbursed — cannot undo. Reverse the bank transfer first.' };
  }

  // Flip every approved install Submission for this (qno, scCode).
  var subSheet = getSheet('Submissions');
  if (!subSheet) return { success: false, error: 'Submissions sheet not found' };
  var subData  = subSheet.getDataRange().getValues();
  var subHdrs  = subData[0];
  var sQIdx    = subHdrs.indexOf('QuotationNo');
  var sSCIdx   = subHdrs.indexOf('SubconCode');
  var sFTIdx   = subHdrs.indexOf('FormType');
  var sStIdx   = subHdrs.indexOf('Status');
  var sQtyIdx  = subHdrs.indexOf('Qty');
  var sIdIdx   = subHdrs.indexOf('SubmissionID');
  var sAByIdx  = subHdrs.indexOf('ApprovedBy');
  var sAAtIdx  = subHdrs.indexOf('ApprovedAt');
  if (sQIdx < 0 || sSCIdx < 0 || sFTIdx < 0 || sStIdx < 0 || sQtyIdx < 0) {
    return { success: false, error: 'Submissions sheet missing expected headers' };
  }

  var revertedQty   = 0;
  var revertedIds   = [];
  var subconNameSeen = '';
  for (var r = 1; r < subData.length; r++) {
    if (String(subData[r][sQIdx])  !== String(quotationNo)) continue;
    if (String(subData[r][sSCIdx]) !== String(subconCode))  continue;
    if (String(subData[r][sFTIdx]).toLowerCase() !== 'install') continue;
    if (String(subData[r][sStIdx]).trim().toLowerCase() !== 'approved') continue;
    subSheet.getRange(r + 1, sStIdx + 1).setValue('pending');
    if (sAByIdx >= 0) subSheet.getRange(r + 1, sAByIdx + 1).setValue('');
    if (sAAtIdx >= 0) subSheet.getRange(r + 1, sAAtIdx + 1).setValue('');
    revertedQty += Number(subData[r][sQtyIdx]) || 0;
    if (sIdIdx >= 0) revertedIds.push(String(subData[r][sIdIdx]));
    var nameIdx = subHdrs.indexOf('SubconName');
    if (nameIdx >= 0 && !subconNameSeen) subconNameSeen = String(subData[r][nameIdx] || '');
  }
  if (revertedQty === 0 && !revertedIds.length) {
    // No approved installs found, but Payments row exists — orphan. Drop it.
    paySheet.deleteRow(payRow + 1);
    return { success: false, error: 'No approved install submissions found; cleared orphan payment row.' };
  }
  SpreadsheetApp.flush();

  // Decrement balances. updateSubconBalance + updateQuotationInstalled are
  // pure-additive helpers — passing a negative qty just subtracts.
  updateSubconBalance(String(subconCode), 'install', -revertedQty);
  updateQuotationInstalled(String(quotationNo), -revertedQty);

  // Demote Quotation status if it was auto-set to 'completed' on approval.
  var qtSheet = getSheet('Quotations');
  if (qtSheet) {
    var qtData = qtSheet.getDataRange().getValues();
    for (var qi = 1; qi < qtData.length; qi++) {
      if (String(qtData[qi][0]).trim() !== String(quotationNo).trim()) continue;
      var estRolls       = Number(qtData[qi][8])  || 0;  // col 9 (I) — EstRolls
      var rollsInstalled = Number(qtData[qi][12]) || 0;  // col 13 (M) — RollsInstalled
      var status         = String(qtData[qi][13] || '').trim().toLowerCase();  // col 14 (N) — Status
      if (status === 'completed' && rollsInstalled < estRolls) {
        qtSheet.getRange(qi + 1, 14).setValue('active');
      }
      break;
    }
  }

  // Drop the now-empty Payments row.
  paySheet.deleteRow(payRow + 1);

  addLog({
    type:        'undo_approval',
    subconCode:  subconCode,
    subconName:  subconNameSeen || (SUBCONS[subconCode] || subconCode),
    quotationNo: quotationNo,
    qty:         -revertedQty,
    notes:       'reverted approval (' + revertedIds.length + ' submission' + (revertedIds.length === 1 ? '' : 's') + ')'
  });

  return {
    success:        true,
    revertedRolls:  revertedQty,
    revertedCount:  revertedIds.length,
    quotationNo:    quotationNo,
    subconCode:     subconCode
  };
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
  // Find AssignedSubcon column (may be beyond getDataRange if header missing)
  var asIdx = -1;
  for (var h = 0; h < headers.length; h++) {
    if (String(headers[h]).trim() === 'AssignedSubcon') { asIdx = h; break; }
  }
  // If header not found, check column O (index 14) directly
  if (asIdx < 0) asIdx = 14;
  var result  = [];
  for (var r = 1; r < data.length; r++) {
    if (!data[r][0]) continue;
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      var key = String(headers[c]).trim();
      if (!key) continue;
      var v = data[r][c];
      obj[key.charAt(0).toLowerCase() + key.slice(1)] =
        v instanceof Date ? v.toISOString() : v;
    }
    // Ensure assignedSubcon is always set
    if (!obj.assignedSubcon && data[r].length > asIdx) {
      obj.assignedSubcon = String(data[r][asIdx] || '').trim();
    }
    result.push(obj);
  }
  return result;
}

function fixQuotationsHeader() {
  var sheet = getSheet('Quotations');
  if (!sheet) return;
  var header = sheet.getRange(1, 15).getValue();
  if (!header || String(header).trim() === '') {
    sheet.getRange(1, 15).setValue('AssignedSubcon');
    Logger.log('Set Quotations O1 = AssignedSubcon');
  } else {
    Logger.log('Quotations O1 already set: ' + header);
  }
}

function fixSubmissionsHeader() {
  var sheet = getSheet('Submissions');
  if (!sheet) return;
  var header = sheet.getRange(1, 15).getValue();
  if (!header || String(header).trim() === '') {
    sheet.getRange(1, 15).setValue('PUSealant');
    Logger.log('Set Submissions O1 = PUSealant');
  } else {
    Logger.log('Submissions O1 already set: ' + header);
  }
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
  //          MembraneValue | TotalValue | Blocks | RollsInstalled | Status | AssignedSubcon
  var sqft     = parseFloat(p.totalSqft)    || 0;
  var rate     = parseFloat(p.ratePerSqft)  || 0;
  var estRolls = Math.ceil(sqft / 80) || 0;
  var cfg      = getConfig();
  var memValue = estRolls * (cfg.avgCost || 0);       // material cost = EstRolls × avg_cost
  var totValue = parseFloat(p.totalValue) || (sqft * rate); // contract value = TotalSqft × RatePerSqft

  var row = [
    quotNo,                                                                  // 0  QuotationNo
    p.date         || Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'yyyy-MM-dd'), // 1  Date
    p.clientName   || '',                                                    // 2  ClientName
    p.projectName  || '',                                                    // 3  ProjectName
    p.siteAddress  || '',                                                    // 4  SiteAddress
    p.membraneType || '',                                                    // 5  MembraneType
    rate,                                                                    // 6  RatePerSqft
    sqft,                                                                    // 7  TotalSqft
    estRolls,                                                                // 8  EstRolls
    memValue,                                                                // 9  MembraneValue = EstRolls × avgCost
    totValue,                                                                // 10 TotalValue = TotalSqft × RatePerSqft
    p.blocks       || '',                                                    // 11 Blocks
    0,                                                                       // 12 RollsInstalled (preserved on update)
    p.status || 'active',                                                    // 13 Status
    p.assignedSubcon || ''                                                   // 14 AssignedSubcon
  ];

  for (var r = 1; r < data.length; r++) {
    if (String(data[r][qIdx]) === String(quotNo)) {
      // Preserve RollsInstalled, SiteAddress and Blocks from existing row
      if (iIdx >= 0) row[12] = Number(data[r][iIdx]) || 0;
      if (!row[4]  && data[r][4])  row[4]  = data[r][4];   // keep existing SiteAddress
      if (!row[11] && bIdx >= 0)   row[11] = data[r][bIdx]; // keep existing Blocks
      // Keep existing AssignedSubcon if not provided
      var asIdx = headers.indexOf('AssignedSubcon');
      if (!row[14] && asIdx >= 0 && data[r][asIdx]) row[14] = data[r][asIdx];
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
  if (!sheet) return { success: false, error: 'SubconRates sheet not found. Run setupAllSubconRates() first.' };

  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var idx     = {};
  headers.forEach(function(h, i) { idx[h] = i; });

  var rolls = Number(rollsInstalled) || 0;

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (String(row[idx['SubconCode']]) !== String(subconCode)) continue;

    // Check for SpecialRates JSON (SC04 style)
    var specialCol = idx['SpecialRates'];
    var specialRaw = (specialCol !== undefined) ? String(row[specialCol] || '').trim() : '';
    if (specialRaw && specialRaw !== '') {
      try {
        var tiers = JSON.parse(specialRaw);
        for (var t = 0; t < tiers.length; t++) {
          var tier = tiers[t];
          var minR = tier.minRolls || 0;
          var maxR = tier.maxRolls || 999999;
          if (rolls >= minR && rolls <= maxR) {
            if (tier.flatTotal) {
              var ft = Number(tier.flatTotal);
              return { success: true, rate: Math.round(ft / rolls), total: ft, payment1: ft * 0.5, payment2: ft * 0.5 };
            }
            var sRate = Number(tier.rate);
            var sTotal = rolls * sRate;
            return { success: true, rate: sRate, total: sTotal, payment1: sTotal * 0.5, payment2: sTotal * 0.5 };
          }
        }
        return { success: false, error: 'No matching tier for ' + rolls + ' rolls (SC04)' };
      } catch(e) {
        Logger.log('Error parsing SpecialRates for ' + subconCode + ': ' + e);
      }
    }

    // Standard 3-tier structure (SC01-SC03)
    var t1Max  = Number(row[idx['Tier1MaxRolls']]) || 4;
    var t1Rate = Number(row[idx['Tier1Rate']])     || 200;
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
    // Parse additional costs
    try { pay.additionalCosts = pay['AdditionalCosts'] ? JSON.parse(pay['AdditionalCosts']) : []; } catch(e) { pay.additionalCosts = []; }
    pay.additionalTotal = Number(pay['AdditionalTotal']) || 0;
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

// ── getFirstSubmissionAt ──────────────────────────────────────────
// Earliest approved install submission timestamp for (quotation, subcon).
// Used as the base for payment due-date calc so the due date stays
// anchored to when the work was reported, not when admin clicked Approve.
function getFirstSubmissionAt(quotationNo, subconCode) {
  var sheet = getSheet('Submissions');
  if (!sheet) return '';
  var rows = sheetToObjects(sheet);
  var earliest = null;
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (String(r.QuotationNo) !== String(quotationNo)) continue;
    if (String(r.SubconCode)  !== String(subconCode))  continue;
    if (String(r.FormType).toLowerCase() !== 'install') continue;
    if (String(r.Status).trim().toLowerCase() !== 'approved') continue;
    var ts = r.Timestamp;
    if (!ts) continue;
    var d = (ts instanceof Date) ? ts : new Date(ts);
    if (isNaN(d.getTime())) continue;
    if (!earliest || d < earliest) earliest = d;
  }
  return earliest ? Utilities.formatDate(earliest, 'Asia/Kuala_Lumpur', "yyyy-MM-dd'T'HH:mm:ss") : '';
}

// ── calculatePaymentForQuotation ──────────────────────────────────
// Sums qty from ALL approved install submissions for a quotation,
// then applies tiered rate to the total.
function calculatePaymentForQuotation(quotationNo, subconCode) {
  var sheet = getSheet('Submissions');
  if (!sheet) return { success: false, error: 'Submissions sheet not found' };
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var acIdx = headers.indexOf('AdditionalCosts');
  if (acIdx < 0) acIdx = 15;
  var rows = sheetToObjects(sheet);
  var totalRolls = 0;
  var allAdditionalCosts = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (String(r.QuotationNo)  !== String(quotationNo)) continue;
    if (String(r.SubconCode)   !== String(subconCode))  continue;
    if (String(r.FormType).toLowerCase() !== 'install')   continue;
    if (String(r.Status).trim().toLowerCase() !== 'approved') continue;
    totalRolls += Number(r.Qty) || 0;
    try {
      var raw = r.AdditionalCosts || data[i + 1][acIdx] || '';
      if (raw) { var ac = JSON.parse(raw); if (Array.isArray(ac)) allAdditionalCosts = allAdditionalCosts.concat(ac); }
    } catch(e) {}
  }
  if (totalRolls === 0) return { success: false, error: 'No approved installs for ' + quotationNo };
  var calc = calculateTieredRate(subconCode, totalRolls);
  if (!calc.success) return calc;
  calc.totalRolls = totalRolls;
  // Expose base values; the writer (upsertPaymentRecord) applies the split.
  calc.basePayment1    = calc.payment1;
  calc.basePayment2    = calc.payment2;
  calc.baseTotal       = calc.total;
  calc.additionalCosts = allAdditionalCosts;
  calc.additionalTotal = allAdditionalCosts.reduce(function(s, c) { return s + (Number(c.amount) || 0); }, 0);
  return calc;
}

// recalcPayment — admin action to re-run calc + upsert for a single quotation.
// Use it after enabling split-50-50 to convert a legacy untagged record to the
// new mode (the upsert's "upgrade unpaid legacy" branch handles the actual flip
// and writes the SplitMode tag). Honours the Payment*Status guards so paid
// amounts are never overwritten.
function recalcPayment(quotationNo, subconCode) {
  if (!quotationNo) return { success: false, error: 'quotationNo required' };
  if (!subconCode)  return { success: false, error: 'subconCode required' };
  var calc = calculatePaymentForQuotation(String(quotationNo), String(subconCode));
  if (!calc.success) return calc;
  return upsertPaymentRecord(String(quotationNo), String(subconCode), calc);
}

// applySplit — picks Payment1Amount / Payment2Amount / TotalAmount based on splitMode.
//   'split-50-50' → additional is split evenly across P1 and P2  (new default)
//   'p1-only'     → additional all goes to P1                    (legacy)
function applySplit(calc, splitMode) {
  var add = Number(calc.additionalTotal) || 0;
  if (splitMode === 'split-50-50') {
    calc.payment1 = (Number(calc.basePayment1) || 0) + add / 2;
    calc.payment2 = (Number(calc.basePayment2) || 0) + add / 2;
  } else { // 'p1-only' (legacy)
    calc.payment1 = (Number(calc.basePayment1) || 0) + add;
    calc.payment2 = (Number(calc.basePayment2) || 0);
  }
  calc.total = (Number(calc.baseTotal) || 0) + add;
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

  // Auto-add SplitMode header (column S) on legacy sheets so writes work.
  if (idx['SplitMode'] === undefined) {
    var nextCol = headers.length + 1;
    sheet.getRange(1, nextCol).setValue('SplitMode');
    idx['SplitMode'] = headers.length;
    headers.push('SplitMode');
  }
  // Auto-add FirstSubmissionAt header (column T) — anchors due-date calc to
  // submission time instead of approval time.
  if (idx['FirstSubmissionAt'] === undefined) {
    var nextCol2 = headers.length + 1;
    sheet.getRange(1, nextCol2).setValue('FirstSubmissionAt');
    idx['FirstSubmissionAt'] = headers.length;
    headers.push('FirstSubmissionAt');
  }
  var smIdx  = idx['SplitMode'];
  var fsaIdx = idx['FirstSubmissionAt'];
  var qIdx   = idx['QuotationNo'];

  // Check if record exists — update in place
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][qIdx]) !== String(quotationNo)) continue;

    // Resolve the split mode for this row.
    //   explicit 'split-50-50' / 'p1-only' → use as-is
    //   empty (untagged legacy)            → upgrade to split-50-50 unless P1
    //                                        has already been paid (then freeze
    //                                        as p1-only so paid amount is safe)
    var existingMode = String((data[r][smIdx] !== undefined ? data[r][smIdx] : '') || '').trim();
    var splitMode;
    var writeBackMode = false;
    if (existingMode === 'split-50-50' || existingMode === 'p1-only') {
      splitMode = existingMode;
    } else {
      var p1Paid = String(data[r][idx['Payment1Status']] || '').trim().toLowerCase() === 'paid';
      splitMode = p1Paid ? 'p1-only' : 'split-50-50';
      writeBackMode = true; // tag the row so future upserts are unambiguous
    }
    applySplit(calc, splitMode);

    sheet.getRange(r + 1, idx['RollsInstalled'] + 1).setValue(calc.totalRolls);
    sheet.getRange(r + 1, idx['RateApplied']    + 1).setValue(calc.rate);
    sheet.getRange(r + 1, idx['TotalAmount']    + 1).setValue(calc.total);
    // Don't clobber Payment1Amount if it has already been paid out — a paid
    // amount is a real disbursement and editing it would silently change the
    // payment record after the fact.
    var p1AlreadyPaid = String(data[r][idx['Payment1Status']] || '').trim().toLowerCase() === 'paid';
    var p2AlreadyPaid = String(data[r][idx['Payment2Status']] || '').trim().toLowerCase() === 'paid';
    if (!p1AlreadyPaid) sheet.getRange(r + 1, idx['Payment1Amount'] + 1).setValue(calc.payment1);
    if (!p2AlreadyPaid) sheet.getRange(r + 1, idx['Payment2Amount'] + 1).setValue(calc.payment2);
    // Update additional costs columns
    var acColIdx = idx['AdditionalCosts'];
    var atColIdx = idx['AdditionalTotal'];
    if (acColIdx !== undefined) sheet.getRange(r + 1, acColIdx + 1).setValue(calc.additionalCosts ? JSON.stringify(calc.additionalCosts) : '[]');
    if (atColIdx !== undefined) sheet.getRange(r + 1, atColIdx + 1).setValue(calc.additionalTotal || 0);
    if (writeBackMode) sheet.getRange(r + 1, smIdx + 1).setValue(splitMode);
    // Backfill FirstSubmissionAt if missing (preserve original value if set).
    var existingFsa = data[r][fsaIdx];
    if (!existingFsa) {
      var fsa = getFirstSubmissionAt(quotationNo, subconCode);
      if (fsa) sheet.getRange(r + 1, fsaIdx + 1).setValue(fsa);
    }
    return { success: true, updated: true, paymentID: data[r][idx['PaymentID']] };
  }

  // Create new record — uses new split logic
  applySplit(calc, 'split-50-50');

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
    nowStr(),          // CreatedAt
    calc.additionalCosts ? JSON.stringify(calc.additionalCosts) : '[]',  // Q  AdditionalCosts
    calc.additionalTotal || 0,  // R  AdditionalTotal
    'split-50-50',     // S  SplitMode
    getFirstSubmissionAt(quotationNo, subconCode)  // T  FirstSubmissionAt
  ]);
  return { success: true, created: true, paymentID: paymentID };
}

// ── createPaymentRecord ───────────────────────────────────────────
// Creates a payment record for a quotation. Skips if one already exists.
// Accepts (quotationNo, subconCode, subconName, rollsInstalled) or
// legacy (quotationNo, subconCode, rollsInstalled) — detects by arg type.
function createPaymentRecord(quotationNo, subconCode, arg3, arg4) {
  var subconName, rollsInstalled;
  if (arg4 !== undefined) {
    // 4-arg call: (quotationNo, subconCode, subconName, rollsInstalled)
    subconName     = String(arg3);
    rollsInstalled = Number(arg4) || 0;
  } else {
    // 3-arg legacy call: (quotationNo, subconCode, rollsInstalled)
    rollsInstalled = Number(arg3) || 0;
    subconName     = SUBCONS[subconCode] || subconCode;
  }

  var sheet = getSheet('Payments');
  if (!sheet) return { success: false, error: 'Payments sheet not found' };

  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var qIdx    = headers.indexOf('QuotationNo');
  var scIdx   = headers.indexOf('SubconCode');

  // Skip if record already exists for this quotation+subcon
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][qIdx]) === String(quotationNo) &&
        (scIdx < 0 || String(data[r][scIdx]) === String(subconCode))) {
      return { success: true, message: 'Record already exists', paymentID: data[r][0] };
    }
  }

  var calc = calculateTieredRate(subconCode, rollsInstalled);
  if (!calc.success) return calc;

  // Resolve subcon name if not provided
  if (!subconName || subconName === subconCode) {
    subconName = SUBCONS[subconCode] || subconCode;
    var subSheet = getSheet('SubconBalances');
    if (subSheet) {
      var subRows = sheetToObjects(subSheet);
      for (var s = 0; s < subRows.length; s++) {
        if (String(subRows[s].SubconCode) === String(subconCode)) {
          subconName = String(subRows[s].SubconName || subconName);
          break;
        }
      }
    }
  }

  var paymentID = 'PAY-' + Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'yyyyMMddHHmmss');

  // Auto-add SplitMode (S) and FirstSubmissionAt (T) headers on legacy sheets so the appended row aligns.
  var crHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (crHeaders.indexOf('SplitMode') < 0) {
    sheet.getRange(1, crHeaders.length + 1).setValue('SplitMode');
    crHeaders.push('SplitMode');
  }
  if (crHeaders.indexOf('FirstSubmissionAt') < 0) {
    sheet.getRange(1, crHeaders.length + 1).setValue('FirstSubmissionAt');
    crHeaders.push('FirstSubmissionAt');
  }

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
    nowStr(),         // CreatedAt
    '[]',             // Q  AdditionalCosts
    0,                // R  AdditionalTotal
    'split-50-50',    // S  SplitMode
    getFirstSubmissionAt(quotationNo, subconCode)  // T  FirstSubmissionAt
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

// ── backfillFirstSubmissionAt ─────────────────────────────────────
// One-shot migration: populates FirstSubmissionAt (col T) on every
// existing Payments row using the earliest approved install submission
// for that (QuotationNo, SubconCode). Safe to re-run — only fills empty
// cells. Run from the Apps Script editor, no doGet exposure.
function backfillFirstSubmissionAt() {
  var sheet = getSheet('Payments');
  if (!sheet) return 'Payments sheet not found';
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('FirstSubmissionAt') < 0) {
    sheet.getRange(1, headers.length + 1).setValue('FirstSubmissionAt');
    headers.push('FirstSubmissionAt');
  }
  var fsaIdx = headers.indexOf('FirstSubmissionAt');
  var qIdx   = headers.indexOf('QuotationNo');
  var scIdx  = headers.indexOf('SubconCode');
  var data   = sheet.getDataRange().getValues();
  var filled = 0;
  for (var r = 1; r < data.length; r++) {
    if (data[r][fsaIdx]) continue;
    var ts = getFirstSubmissionAt(data[r][qIdx], data[r][scIdx]);
    if (ts) {
      sheet.getRange(r + 1, fsaIdx + 1).setValue(ts);
      filled++;
    }
  }
  Logger.log('backfillFirstSubmissionAt: filled ' + filled + ' rows.');
  return 'Filled ' + filled + ' rows.';
}

// ════════════════════════════════════════════════════════════════
// LOGIN / CREDENTIALS
// ════════════════════════════════════════════════════════════════

function handleLogin(p) {
  var sheet = getSheet('Credentials');
  if (!sheet) return { success: false, error: 'Credentials sheet not found. Run setupCredentials().' };
  var data = sheet.getDataRange().getValues();
  var userCode = String(p.userCode || '').trim().toUpperCase();
  var pin      = String(p.pin || '').trim();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() === userCode &&
        String(data[i][4]).trim() === pin &&
        (data[i][5] === true || String(data[i][5]).toUpperCase() === 'TRUE')) {
      return {
        success:  true,
        userCode: String(data[i][0]),
        userName: String(data[i][1]),
        userType: String(data[i][2])
      };
    }
  }
  return { success: false, error: 'Invalid code or PIN' };
}

function setupCredentials() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Credentials');
  if (!sheet) {
    sheet = ss.insertSheet('Credentials');
  }
  var existing = sheet.getDataRange().getValues();
  if (existing.length <= 1) {
    sheet.clear();
    sheet.appendRow(['UserCode', 'UserName', 'UserType', 'Phone', 'PIN', 'Active']);
    sheet.appendRow(['ADMIN', 'Admin', 'admin', '', 'admin123', true]);
    sheet.appendRow(['SC01', 'Md Atik', 'subcon', '', '1234', true]);
    sheet.appendRow(['SC02', 'Md Shahazan', 'subcon', '', '1234', true]);
    sheet.appendRow(['SC03', 'Md Mohiuddin', 'subcon', '', '1234', true]);
    sheet.appendRow(['SC04', 'Md Foysel', 'subcon', '', '1234', true]);
    Logger.log('Credentials sheet created with default users');
  } else {
    Logger.log('Credentials sheet already has data (' + existing.length + ' rows)');
  }
}

function setupAllSubcons() {
  var ss = getSpreadsheet();
  var sh = ss.getSheetByName('SubconBalances');
  if (!sh) {
    sh = ss.insertSheet('SubconBalances');
    sh.appendRow(['SubconCode','SubconName','TotalPickup','TotalInstalled','Balance','LastUpdated']);
  }
  var existing = sh.getDataRange().getValues().map(function(r){ return String(r[0]); });
  var subcons = [
    ['SC01','Md Atik',0,0,0,''],
    ['SC02','Md Shahazan',0,0,0,''],
    ['SC03','Md Mohiuddin',0,0,0,''],
    ['SC04','Md Foysel',0,0,0,'']
  ];
  subcons.forEach(function(s) {
    if (existing.indexOf(s[0]) === -1) sh.appendRow(s);
  });
  Logger.log('setupAllSubcons complete');
}

function setupAllSubconRates() {
  var ss = getSpreadsheet();
  var sh = ss.getSheetByName('SubconRates');
  if (!sh) {
    sh = ss.insertSheet('SubconRates');
  }
  sh.clearContents();
  sh.appendRow(['SubconCode','SubconName','Tier1MaxRolls','Tier1Rate','Tier2MinRolls','Tier2MaxRolls','Tier2Rate','Tier3Rate','SpecialRates']);

  var sc04Rates = JSON.stringify([
    {"maxRolls":3,"flatTotal":500},
    {"minRolls":4,"maxRolls":6,"rate":170},
    {"minRolls":7,"maxRolls":9,"rate":165},
    {"minRolls":10,"maxRolls":15,"rate":160},
    {"minRolls":16,"maxRolls":20,"rate":155},
    {"minRolls":21,"maxRolls":25,"rate":150},
    {"minRolls":26,"maxRolls":30,"rate":145},
    {"minRolls":31,"maxRolls":45,"rate":130},
    {"minRolls":46,"maxRolls":60,"rate":125},
    {"minRolls":61,"maxRolls":100,"rate":110},
    {"minRolls":101,"rate":100}
  ]);

  sh.appendRow(['SC01','Md Atik',4,200,5,9,170,150,'']);
  sh.appendRow(['SC02','Md Shahazan',4,200,5,9,170,150,'']);
  sh.appendRow(['SC03','Md Mohiuddin',4,200,5,9,170,150,'']);
  sh.appendRow(['SC04','Md Foysel',0,0,0,0,0,0,sc04Rates]);

  Logger.log('setupAllSubconRates complete');
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
