# sunorient-invoicing

Invoice generator and run procedure for Sun Orient (S) Pte Ltd's K Line and ECL billing.

- `engine/invoicegen.py`: parses K Line confirmation reports, calculates the line items, and renders the invoice PDF on
  letterhead, the K Line IMS EDI CSV and the invoice stack.
- `engine/rates.json`: the rate card and tariff codes. **Update rates here only.**
- `WORKFLOW.md`: the step-by-step procedure each scheduled run follows (Gmail in → approval email out).
- `tests/test_regression.py`: must reproduce real invoices 26-0901 ($17,584.43) and 26-0906 ($17,904.76) to the cent.

Private repository. It contains client rates and sample operational documents.
