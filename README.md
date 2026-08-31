# Luika Drilling Daily Report Agent

## Start here: `luika_daily_report_agent.ipynb`

The notebook is the walkthrough version — it reads Hamza's **PDF** directly, one step
per cell, with the parsed output visible at every stage. Run it top to bottom once to
see how each figure is produced, then use Section 12 (batch) daily.

```bash
pip install pdfplumber anthropic pandas matplotlib jupyter
export ANTHROPIC_API_KEY=sk-...
jupyter lab luika_daily_report_agent.ipynb      # or just open it in VS Code
```

Drop the day's PDFs into `inbox/` and run the last cell. Everything you would ever
need to edit — rate card, monthly plan, tolerances — is in Section 2.

The two `.py` scripts below are the same logic packaged for unattended running.

---

Turns Hamza's daily rig sheet into a CEO-facing update, with the arithmetic checked
before anyone reads a sentence.

## The design rule

**Python does the arithmetic. Claude does the language.**

The model never adds, divides or judges a number. It receives a JSON block of
already-computed fields plus an already-computed exception list, and writes prose.
This is deliberate. A language model asked to "check the numbers" will sometimes
produce a confident, wrong figure, and you will be the one defending it. Arithmetic
that must be defensible belongs in code you can point at.

## Setup

```bash
python -m venv .venv && source .venv/bin/activate
pip install anthropic openpyxl pdfplumber
export ANTHROPIC_API_KEY=sk-...
```

In VS Code: open the folder, select `.venv` as the interpreter (Cmd/Ctrl+Shift+P →
"Python: Select Interpreter"), then F5 to run with the debugger. Set breakpoints in
`compute()` — that is where every validation rule lives and where you will spend
your time as the rules get sharper.

## Run

```bash
python daily_report_agent.py inbox/Daily_Report26.xlsx        # normal
python daily_report_agent.py inbox/Daily_Report26.pdf --dry-run   # no API call
```

Outputs:
- `out/<date>_ceo_update.md` — the update to send
- `history.csv` — one row per day, appended. This is the point of the whole thing.
  After a month you can chart metres per rig hour, fuel per metre and flag counts,
  and answer "is the rig getting better or worse" instead of "how was yesterday".

## Get the Excel, not the PDF

`extract_xlsx()` reads by searching for label text and taking the value to its
right, so it survives Hamza moving a row. `extract_pdf()` is positional guesswork
and loses fields — on the 25 and 26 August files it recovers the shift figures but
drops the client supervisor names entirely. Every PDF-sourced run is tagged with a
`SOURCE:` exception for that reason. One WhatsApp message to Hamza asking for the
`.xlsx` instead of a printout removes an entire class of failure.

## Validation rules currently implemented

| Rule | Checks |
|---|---|
| R1 | Metre counter (end − start) matches the "Meter Drilled" box |
| R2 | Rig hours present, positive, and not longer than the shift window |
| R3 | Fuel recorded; computes litres per rig hour and per metre |
| R4 | Hole ID, core type, driller and location are not blank |
| R5 | Hole ID changed when the metre counter resets to zero |
| R7 | No rig-hour gap or reversal against the previous day's report |
| R8 | Week-to-date table total matches the shift log, and core type matches |
| R9 | A plan figure exists for the day |

Rules to add once you have the answers: expected rate of penetration by depth band
and core size; fuel litres per hour outside a rolling band; consumables on hand
below reorder level; days since last service against the service interval.

## The report template needs new fields

The agent can only check what is written down. These fields do not exist on the
current sheet and are the reason the CEO's request is only one-quarter answered:

**Machine status**
- Downtime hours, split into: mechanical breakdown, waiting on spares, waiting on
  client, weather, rig move, no water, shift change. One line per event with hours.
- Fault log — what failed, what was done, is it closed or still open.
- Hour-meter reading at last service, and hours until next service due.
- Rig condition at handover: green / amber / red, with the supervisor's initials.

**Consumables, tools and spares**
- Closing stock of drill bits by size, core barrels, rods, casing, drilling mud
  and polymer, fuel in the bowser.
- Reorder level for each, and a flag when stock is below it.
- Items ordered but not yet delivered, with the date ordered and expected date.
- Anything that stopped or slowed work because it was not on site.

**Accountability**
- Supervisor name and sign-off (Hamza's name appears nowhere on the current sheet).
- Client supervisor sign-off on the metres — this is what protects the invoice.

**Commercial**
- Rate band applied and the invoice value for the day. The current sheet has an
  "Invoice Total" cell that is too narrow to display its own number.

---

## `master_audit.py` — run this before the invoice goes out

```bash
python master_audit.py "Master August (2026).xlsx"
```

Reconciles the three layers that do not currently talk to each other:

```
DAILY REPORT  (shift log, overwritten daily)
   ->  ACTUAL DRIL   (hand-keyed daily grid by depth band and core size)
        ->  TAX INVOICE  (formula-linked to the band totals)
```

Exits non-zero if anything CRITICAL is found, so it can be used as a gate before
an invoice is issued. On the August 2026 workbook it returns 13 findings, one of
which is worth TZS 45,240,000.

Checks: INV-1 invoice line pulls its own band's monthly total; INV-2/3/4 the
invoice foots and VAT is 18%; INV-5 the invoice number is a real reference;
ACT-1 band totals equal the daily-total row; GOAL-1 the headline percentage
compares to plan-to-date not full-month plan; HIST-1/2 daily history is kept and
the shift log ties to the grid; COST-1/2 cost fields are populated and
consumables flow through; REF-1 no sheet is full of #REF! errors.
