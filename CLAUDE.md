# RollTrack — Innovation AID Sdn Bhd

Roof-membrane install + payment tracker. Admin dashboard + per-subcon mobile forms, backed by a Google Apps Script Web App and a Google Sheet.

## Stack

- **Frontend:** static HTML + vanilla JS, no build step. The admin page is one big HTML file.
- **Backend:** Google Apps Script (`rolltrack_apps_script.js`). One `doGet(e)` dispatches `e.parameter.action` to handlers and JSONP-wraps the response via `ContentService`.
- **Storage:** a single Google Sheet workbook with multiple tabs.
- **Wire:** every API call is JSONP. Frontend's `jsonp(params)` (`rolltrack_admin.html` ~line 522) appends a unique `callback=cb_xxx` and resolves when the script tag fires the callback. 20s timeout.

## Files

| File | Role |
|---|---|
| `rolltrack_apps_script.js` | All backend logic. ~2000 lines. |
| `rolltrack_admin.html` | Admin dashboard: Approvals, Payments, Quotations, Subcons, Stock, Reports. ~2200 lines. |
| `subcon_form_SC01.html` ... `SC04.html` | Per-subcon mobile forms (install / pickup / return / MY HISTORY). **Near-identical** — most bugfixes need the same edit applied to all four. |
| `login.html` | PIN login. |
| `logo.png` | Innovation AID logo. |

## Subcons

| Code | Name |
|---|---|
| SC01 | Md Atik |
| SC02 | Md Shahazan |
| SC03 | Md Mohiuddin |
| SC04 | Md Foysel |

The mapping appears in three places (none yet DRY'd):

- `rolltrack_apps_script.js` — `SUBCONS` (~line 8)
- `rolltrack_admin.html` — `SUBCONS_MAP` (~line 517) — values include the suffix `(SCxx)`
- `generateReport()` / `generateBankTallyReport()` — inline `NAME_BY_CODE`

If you're tempted to add another mapping, hoist them all instead.

## Sheets and columns

### Submissions

A=Timestamp · B=SubconCode · C=SubconName · D=FormType · E=QuotationNo · F=Qty · G=ActivityDate · H=Notes · I=PhotoURL · J=Status · K=ApprovedBy · L=ApprovedAt · M=RejectionReason · N=SubmissionID · O=PUSealant · P=AdditionalCosts (JSON) · Q=OriginalAdditionalCosts · R=EditedBy · S=EditedAt · T=EditReason

`Status` flows: `pending` → `approved` | `rejected`. Rejection only writes J/K/L/M; no payment, no balance change. Q–T are auto-added by `approveSubmissionWithEdit` on first use.

### Payments

A=PaymentID · B=QuotationNo · C=SubconCode · D=SubconName · E=RollsInstalled · F=RateApplied · G=TotalAmount · H=Payment1Amount · I=Payment1Status · J=Payment1Date · K=Payment1Reference · L=Payment2Amount · M=Payment2Status · N=Payment2Date · O=Payment2Reference · P=CreatedAt · Q=AdditionalCosts · R=AdditionalTotal · S=SplitMode

`SplitMode`:
- `'p1-only'` — legacy: 100% of additional costs on P1, P2 is rolls-only.
- `'split-50-50'` — current default: additional split evenly across P1 and P2.
- empty (legacy untagged) — `upsertPaymentRecord` resolves on touch: if P1 is already paid, freeze as `'p1-only'` so the historical paid amount is preserved; else upgrade to `'split-50-50'` and tag the row.

S column auto-creates on the first upsert if missing.

### Quotations

QuotationNo · Date · ClientName · ProjectName · SiteAddress · MembraneType · RatePerSqft · TotalSqft · EstRolls · MembraneValue · TotalValue · Blocks · RollsInstalled · Status · AssignedSubcon (col O — auto-fixed by `fixQuotationsHeader` if blank)

`Status`: `active` (open for installs) · `completed` (auto-set on first install approval) · `upcoming`.

### SubconBalances

SubconCode · SubconName · TotalPickup · TotalInstalled · Balance · UpdatedAt

Mutated only by `updateSubconBalance` from `approveSubmission`. Never touched by reject.

### SubconRates

Per-subcon tier rates. Most subcons follow the default `calculateTieredRate` (≤4 → 190, 5–9 → 170, 10+ → 150). SC04 has different tiers per `setupAllSubconRates`.

### Stock / StockIn

Warehouse roll inventory. Updated by `approveSubmission` on pickup / return.

## Key flows

### Install submission → approval

1. Subcon submits via `submitSubconForm`. Row in Submissions, `Status='pending'`.
2. Admin sees it in Approvals (`getPendingSubmissions` filters `pending`).
3. Admin clicks **✅ Approve** OR **✅ Approve with adjustment** (only differs when admin edited the additional-cost amounts inline):
   - `approveSubmission(submissionId)`:
     - Marks `Status=approved`, flushes.
     - `updateSubconBalance` (TotalPickup / TotalInstalled / Balance).
     - For installs: `updateQuotationInstalled` and auto-set quotation Status to `completed`.
     - For installs: `calculatePaymentForQuotation` → `upsertPaymentRecord` (creates payment row or updates existing).
     - `addLog` movement entry.
   - `approveSubmissionWithEdit(submissionId, additionalCostsJson, editReason)`:
     - Snapshots original `AdditionalCosts` to `OriginalAdditionalCosts` (only the first time — re-edits don't clobber the original).
     - Writes new `AdditionalCosts` to col P.
     - Records `EditedBy='admin'`, `EditedAt`, optional `EditReason`.
     - Delegates to `approveSubmission` so the payment pipeline picks up the edited values.

### Rejection

- `rejectSubmission(submissionId, reason)` writes Status='rejected' + reason + ApprovedBy + ApprovedAt onto the Submissions row. **Nothing else.**
- Quotation stays `active`; Balances unchanged; no payment record.
- Subcon-side install dropdown filters only on `status === 'pending'`, so the quotation reappears for resubmission.
- Subcon's MY HISTORY shows the rejected card with reason, red left border, and "Please resubmit / Sila hantar semula" prompt.
- Resubmission is a fresh row in Submissions; nothing structurally links it back to the rejected one.

### Payment math

`calculateTieredRate(subconCode, rolls)` — base 50/50 split of `rolls × rate`. No additional baked in.

`calculatePaymentForQuotation(qno, scCode)` — sums every `approved` install row in Submissions for that (qno, scCode). Returns `{rate, totalRolls, basePayment1, basePayment2, baseTotal, additionalCosts, additionalTotal}`. **Does not** apply the split — that happens at write time.

`applySplit(calc, splitMode)`:
- `'split-50-50'`: `payment1 = base1 + add/2`, `payment2 = base2 + add/2`
- `'p1-only'` (legacy): `payment1 = base1 + add`, `payment2 = base2`
- `total = baseTotal + add` either way

`upsertPaymentRecord(qno, scCode, calc)` — single chokepoint that writes Payments rows:
- Auto-adds the `SplitMode` header on column S if missing.
- Existing row: resolves splitMode (explicit → as-is; empty → upgrade if P1 unpaid, freeze as `p1-only` if P1 already paid; tags the row both ways).
- **Guard**: never overwrites `Payment1Amount` if `Payment1Status === 'paid'` (same for P2). Protects historical paid amounts on every recompute.
- New row: tagged `'split-50-50'`, all calc applied via `applySplit`.

`recalcPayment(qno, scCode)` — admin action that re-runs calc + upsert for one quotation. Use it via JSONP URL to fix individual records after policy changes (e.g. flipping a legacy unpaid row to split-50-50 without waiting for a fresh install).

### Mark Paid

- **Per-card** Mark Paid → `submitMarkPaid` → `markPayment` action with `paymentID`, `paymentNumber` (1 or 2), `status='paid'`, `date`, optional `reference`.
- **Mark Week as Paid** batch → `confirmBatchPay` loops `markPayment` for each item from `buildWeeklyItems` (overdue + dueThisWeek).
- `buildWeeklyItems` skips P2 if P1 is unpaid — P2 is gated on P1 in this system. The on-screen card already labels those P2 rows "not due yet".
- Modal-close-clears-state pattern (see Gotchas) — both flows snapshot `payingRecord` / `batchPayItems` into locals before calling `closePayModal()` / `closeBatchPay()`.

### Reports (Payments tab)

- **Generate Report** → `generateReport()` — status overview per subcon. Table columns: Quotation · Customer · Address · Rolls · Rate · Amount · P1 (with `[STATUS]` tag + Due) · P2 (same). Single subcon = one table; "All Subcons" = one table per subcon under an `<h3>` heading. Each table has a TOTAL row summing rolls/amount/P1/P2.
- **Bank Tally** → `generateBankTallyReport()` — one row per actual P1/P2 disbursement whose `Payment*Date` falls in the `bankFrom`–`bankTo` window. Sorted chronologically (matches a bank statement). Per-report date inputs default to first-of-current-month → today. "All Subcons" mode adds a "Subtotal by subcon" mini-table at the bottom for accounting cross-check.

Both reports open in a new tab via `window.open('', '_blank')` + `document.write(html)`. Print stylesheet hides the print button during print.

## Cross-cutting gotchas

### Host caches `rolltrack_admin.html` aggressively

We hit this multiple times. After every push, the user often sees stale code even with hard-refresh. Workarounds, in order of strength:

1. Append a query string to the URL (`?v=YYYYMMDDx`) — different URL, cache-bypass at every layer.
2. DevTools → Application → Service Workers → Unregister anything.
3. DevTools → Application → Storage → "Clear site data".
4. DevTools → Network → "Disable cache" (only while DevTools open).
5. Incognito window — bypasses browser caches but NOT host caches.

**Always verify** the loaded code is current: DevTools → Sources → search for an identifier introduced in the latest commit.

### Apps Script doesn't auto-update on git push

The Web App keeps serving whatever was last deployed. After ANY change to `rolltrack_apps_script.js`:

1. Paste the new contents into the Apps Script editor (no clasp config in repo — manual paste).
2. Save (Ctrl+S).
3. **Deploy → Manage deployments** (top-right).
4. Pencil icon next to the existing deployment — **NOT** "New deployment". A "New deployment" creates a fresh URL that `rolltrack_admin.html` line ~498 doesn't know about.
5. Version dropdown → **New version** → **Deploy**.

The URL stays the same. If you ever need to change the URL, update `rolltrack_admin.html` line ~498 (`const API = '...'`).

`Unknown action: <name>` from a doGet response always means "deployed code doesn't have that case yet" — i.e. forgot to redeploy.

### Each subcon HTML must be edited four times

`subcon_form_SC01.html`–`SC04.html` share the same `applySubconData`, `loadSubconData`, `renderHistory`, etc. They are copies, not symlinks. Bug fixes / feature additions need the same Edit applied to all four. Verify with `diff <(sed -n 'N,Mp' subcon_form_SC01.html) <(sed -n 'N,Mp' subcon_form_SC02.html)` for the changed range.

### JSONP error semantics

`jsonp(params)` rejects with one of:
- `'timeout'` — 20s timeout, Apps Script too slow or hung.
- `'network error'` — script tag `onerror` fired (HTTP error, response wasn't valid JS, etc).

Both reject paths surface as `toast('Connection error')` in the various catch handlers.

**Important**: any TypeError thrown synchronously inside the `try { await jsonp(...) }` block ALSO falls into the catch and shows "Connection error". So before assuming it's a network issue, check whether the catch is firing on a synchronous bug. We hit this with the modal-close-clears-state bug below.

### Modal close clears form state — snapshot first

These functions all clear shared state:
- `closePayModal()` → `payingRecord = null`
- `closeBatchPay()` → `batchPayItems = []`
- `closeRejModal()` → clears `rejectingId / rejectingFormType / rejectingQuotNo`

If you read those vars AFTER calling the close, you get null/empty. Always:

```js
const savedFoo = payingRecord.foo;
const savedBar = payingRecord.bar;
closePayModal();
// ... use savedFoo / savedBar ...
```

The rejection-flow code (line ~997) is the canonical pattern.

### Payment record P1+P2 sum invariant (mostly)

Normally `Payment1Amount + Payment2Amount === TotalAmount`. The "preserve paid amounts" guard breaks this in one pathological case: if P1 was paid under one set of rolls/additional, and a later install approval changes the rolls/additional, P1 stays at the old paid amount but P2 + TotalAmount get recomputed. This is intentional — paid amounts must not change retroactively — but be aware when summing or reconciling.

### `markPayment` column lookup is header-based

`markPayment` uses `headers.indexOf('Payment1Status')` etc., not hardcoded column numbers. Adding columns Q/R/S to the right of the existing payment columns is safe — it doesn't shift anything earlier.

## Conventions

- **API URL** is hardcoded in `rolltrack_admin.html` line ~498 (`const API = '...'`). Update there if redeploying creates a new URL (it shouldn't if you use Manage deployments → Edit existing).
- **Spreadsheet ID** is hardcoded in `rolltrack_apps_script.js` line ~6 (`var SPREADSHEET_ID = '...'`).
- **Date formatter** `fmtDate(ts)` — `en-MY` locale, day-short-month-year.
- **Money formatter** `fmtN(n)` — comma-thousand separators, two decimals. Always prefix `RM ` manually in the markup.
- **Status strings** are lowercase (`'pending'`, `'approved'`, `'rejected'`, `'paid'`, `'unpaid'`). Compare with `String(x).toLowerCase()` for safety.
- **Existing code uses emojis** in toasts, badges, button labels — keep consistent. Don't add them to backend logs or sheet values.

## Quick verification before commit

Brace/paren/bracket/backtick balance check, run per file:

```bash
python -c "s=open('rolltrack_admin.html','r',encoding='utf-8').read(); print('braces',s.count('{')-s.count('}'),'parens',s.count('(')-s.count(')'),'brkts',s.count('[')-s.count(']'),'backticks_odd',s.count('\`')%2)"
```

Each number should be 0. Catches the most common template-literal breakage and unbalanced edits. Repeat for `rolltrack_apps_script.js` and any subcon HTML you touched.

`git diff --stat` should show only the files you intended to change.

## Commit + deploy checklist

For a frontend-only change:
1. Verify balance check above.
2. Commit + push.
3. Tell the user: "HTML auto-live, hard-refresh / cache-bust".

For an Apps Script change:
1. Same as above.
2. Tell the user: "Apps Script needs redeploy — Deploy → Manage deployments → pencil → New version → Deploy".
3. Optionally include a JSONP URL the user can paste into a browser to verify the new action exists in the deployed code.

For a subcon-HTML change:
1. Apply to all four SC files.
2. Diff the four files in the touched range to confirm they stayed identical.
