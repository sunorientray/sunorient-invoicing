# Sun Orient invoicing: run procedure

This is what each scheduled run follows. A run starts in a fresh workspace with no memory of earlier
runs, so all state lives in **Gmail labels** and in the **approval email threads**.

Mailbox: the connected Gmail (receives raymond@ and klineops@ mail via POP).
Sign-off on every email: **Bob**. Never "Claude" or "invoice assistant".

## 0. Setup (every run)

```bash
pip install --break-system-packages openpyxl reportlab pypdf   # if missing; LibreOffice (soffice) must be present
python3 tests/test_regression.py                                # must print ALL PASS before doing anything else
```

If the regression test fails, stop and email raymond@sunorient.com.sg with the error. Do not invoice.

Gmail labels (create if missing):

| Label | Meaning |
|---|---|
| `SOInv/Processed` | Source email has been turned into an invoice draft. Never process again |
| `SOInv/AwaitingApproval` | Approval request sent; waiting on the approver's reply |
| `SOInv/AwaitingNT` | K Line invoice built; Ng Terminal invoice not received yet |
| `SOInv/Approved` | Approver replied APPROVED; final files sent |

## 1. Find new work

Search (exclude anything already labelled `SOInv/Processed`):

- **K Line confirmation report / tally docs**: subject contains `Confirmation Report` or `Tally Documents`,
  from `@sunorient.com.sg` (the supervisor forwards what Ng Terminal staff send at the end of cargo ops),
  with an `.xls` attachment and usually `Tally Documents.zip`.
- **Ng Terminal invoice**: from `@ngterminal.com.sg`, or forwarded by a Sun Orient address, with a
  `TAX INVOICE` PDF whose Job No is a K Line vessel.
- **ECL job**: the supervisor's handling sheet (e.g. `Scan<date>.pdf`) and/or the Andy Services bill (`Bill_new*.pdf`).
- **Approval replies** on threads labelled `SOInv/AwaitingApproval`.

## 2. K Line

1. Parse the `.xls` with `parse_kline_report()`. From `Tally Documents.zip`, take the signed
   `Confirmation Report.pdf` and check its text layer against the `.xls` numbers. If any value differs, stop and flag it.
2. **PIC / Attn**: sender of the latest `<VESSEL> V.<voy> - TENTATIVE BERTH APPLICATION` email.
   Write it as `MR <FIRSTNAME>`. If none is found, use `MR MALCOLM` and ask in the approval email.
3. **Sunday/PH shifts**: day 1 = berthing date, day 2 = next day. Mark `ph_shifts` for any shift that falls on a
   Sunday or Singapore public holiday. Read the remarks: work done on another date (e.g. MAFI stuffing)
   uses that date's status.
4. **Extra work from remarks** (MAFI stuffing, signalman, key collection, etc.): move the quantities into an
   `extra_blocks` entry with a title line (`MAFI STUFFING OPERATION ON dd/mm/yyyy`, B/L numbers), using
   `deduct` so nothing is billed twice. List every such line under "Please confirm" in the approval email.
5. **Invoice number**: `YY-MM##`. Take the highest number in sent approval emails (subject
   `For approval: ... invoice YY-MM##`) and add 1. K Line and ECL share one sequence. Always ask the approver to confirm it.
6. Generate: `render_invoice()` → PDF, `write_edi()` → CSV, `build_stack()` = invoice + signed confirmation report
   + fuel top-up sheet (if fuel was billed) + any receipts. **Never put the Ng Terminal invoice in the stack.**
7. **Check against Ng Terminal**: once its invoice is in, compare quantities line by line (Ng Terminal's unit prices
   differ from ours, so compare quantities only). If it hasn't arrived, label the thread `SOInv/AwaitingNT` and check again next run.
8. Email **raymond@sunorient.com.sg**: stack + EDI CSV + Ng Terminal invoice, with the checks summarised.

## 3. ECL

1. Read the handling sheet (handwritten scan, so zoom into it) and the Andy Services bill. Check them against each other.
2. Email **ECLOps@sunorient.com.sg** for approval of the job quantities, listing any mismatches.
3. After ECL Ops approves: build the invoice with `ecl_blocks()` and a stack of invoice + handling sheet
   (**not** the Andy Services bill), then email **raymond@sunorient.com.sg** for approval.

## 4. After approval

When the approver replies APPROVED (or with corrections): apply the corrections, regenerate without the DRY RUN mark,
reply on the same thread with the final files, and label it `SOInv/Approved`.
Uploading to the client portal is a separate step (not automated yet).

## Rules

- Nothing goes to a client. Every email goes to raymond@ or ECLOps@ only.
- If something is unclear, flag it in the approval email rather than guess silently.
- Rates live only in `engine/rates.json`. When a rate changes, update the file and re-run the tests.
