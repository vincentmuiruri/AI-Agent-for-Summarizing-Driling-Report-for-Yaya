# Setup guide — Luika Daily Drilling Report Agent (VS Code)

Written for a first-time run on a clean machine. Steps 1–4 are once only. Steps 5–8
are the daily routine and take about two minutes.

---

## 1. Install the prerequisites

| | |
|---|---|
| Python 3.10 or newer | python.org/downloads — **tick "Add Python to PATH"** on Windows |
| VS Code | code.visualstudio.com |
| VS Code extensions | Open the Extensions panel (`Ctrl+Shift+X`) and install **Python** and **Jupyter**, both by Microsoft |

Check it worked. Open a terminal (`Ctrl+` `` ` ``) and run:

```bash
python --version      # Windows
python3 --version     # macOS / Linux
```

Anything from 3.10 up is fine.

---

## 2. Create the project folder

```
yaya-drilling/
├─ luika_daily_report_agent.ipynb
├─ master_audit.py
├─ daily_report_agent.py
├─ inbox/      <- Hamza's PDFs go here
└─ out/        <- the daily updates land here
```

Create `yaya-drilling` anywhere sensible (Documents is fine), drop the three files
in, and add the two empty folders. Then in VS Code: **File → Open Folder →
yaya-drilling**.

Open the folder, not the notebook file on its own. If you open only the file, the
notebook writes `inbox/` and `out/` wherever VS Code last happened to be, and you
will spend ten minutes looking for your output.

---

## 3. Create the environment and install the libraries

In VS Code press `Ctrl+Shift+P` and run **Python: Create Environment** → **Venv** →
pick your Python version. VS Code builds a `.venv` folder and switches to it.

Then, in the VS Code terminal:

```bash
pip install pdfplumber anthropic pandas matplotlib openpyxl ipykernel python-dotenv
```

Two to three minutes on a normal connection.

---

## 4. Set the Anthropic API key

Get a key from console.anthropic.com → **API Keys**.

Create a file called `.env` in the `yaya-drilling` folder containing one line:

```
ANTHROPIC_API_KEY=sk-ant-...
```

And a file called `.gitignore` containing:

```
.env
.venv/
history.csv
out/
inbox/
```

**Do not type the key into the notebook itself.** The notebook has a commented-out
line offering to do that. Leave it commented. You will end up emailing this notebook
to somebody, and a key pasted into cell 23 goes with it.

Then add this as the **first line** of the Section 11 code cell, above
`# os.environ[...]`:

```python
from dotenv import load_dotenv; load_dotenv()
```

Cost is negligible — one short call per day, well under a dollar a month.

---

## 5. Get today's report into `inbox/`

Hamza sends the PDF by WhatsApp or email. Save it into `yaya-drilling/inbox/`.

- **WhatsApp Desktop** — right-click the attachment → Save as → navigate to `inbox`.
- **Outlook / Gmail** — Download, then drag the file from your Downloads folder into
  the `inbox` folder shown in the VS Code Explorer panel on the left.

**Leave the previous days' PDFs in the folder.** Rule R7 — the one that catches an
hour meter running backwards or five hours going unaccounted for — compares today
against yesterday. With a single file in the folder it has nothing to compare to and
stays silent. Keep at least the last week in there.

The filename does not matter. The notebook reads the date off the page and sorts by
it.

---

## 6. Check the configuration (start of each month only)

Open the notebook and look at **Section 2**.

- `RATE_CARD` — changes only when the Shanta contract changes.
- `MONTH_PLAN_METRES` and `MONTH_PLAN_VALUE` — **change every month**. Take them
  from the `PLANNED DRIL` sheet (`AH25`) and the `TAX INVOICE` sheet (`Y38`) of that
  month's master workbook.

If you forget, the money check in Section 9 compares against last month's plan and
will either miss a real gap or invent a false one. It is the one setting worth a
diary reminder on the 1st.

---

## 7. Run it

**First time:** work down the notebook cell by cell with `Shift+Enter`. You will be
asked to pick a kernel — choose the `.venv` one. Read the output of each section so
you know what the thing is doing before you send anything to the CEO off the back of
it.

**Every day after that:** just run the last cell in **Section 12**.

```python
results = run_folder(dry_run=False)     # dry_run=False calls Claude
```

Leave it at `dry_run=True` if you only want the flags and no narrative — that runs
free and offline.

It prints a summary table and, underneath, every `CRITICAL` finding in full. It also
writes `out/2026-08-31_ceo_update.md` and appends a row to `history.csv`.

---

## 8. Read it, then send it

Open the file in `out/`. It is about 220 words in five blocks: production, machine
status, consumables, issues, data flags.

**Before you forward it:**

1. Read every `CRITICAL` flag. `R8` and `MONEY-1` mean the metres that generate
   revenue disagree with the metres the client supervisor signed for. Those stop the
   invoice; they do not go in a routine daily update without a covering line from
   you.
2. `R2` and `R7` — hour-meter problems — go to Hamza the same day, while he can
   still remember what the meter actually read. A day later it is a guess.
3. Sanity-check the day total against what Hamza told you on the phone. You are
   speaking to him daily anyway; this is that conversation's paperwork, not a
   replacement for it.

**To send:** open the `.md` file, `Ctrl+A`, `Ctrl+C`, paste into the email or
WhatsApp. It is short by design — do not attach a file for 220 words. Put your own
one-line read at the top:

> *Rig on plan for the day. One item needs a decision — see the flag on core type.*

If you ever do need it as a PDF, install the **Markdown PDF** extension in VS Code,
right-click the file, and choose *Markdown PDF: Export (pdf)*.

---

## Monthly, before the invoice goes to Shanta

```bash
python master_audit.py "Master August (2026).xlsx"
```

Reconciles the shift log, the `ACTUAL DRIL` grid and the `TAX INVOICE` sheet. Exits
non-zero if anything critical is found, so it can gate the invoice run rather than
sit beside it.

---

## When something breaks

| Symptom | Cause | Fix |
|---|---|---|
| `ModuleNotFoundError` | wrong kernel selected | `Ctrl+Shift+P` → *Python: Select Interpreter* → the `.venv` one, then restart the kernel |
| `KeyError: 'ANTHROPIC_API_KEY'` | `.env` not loaded | check the `load_dotenv()` line from step 4 is at the top of the Section 11 cell |
| Money check says *"No month-total column"* | normal | that column only prints in the last week of the month |
| A field comes out blank or wrong | PDF text extraction | expected — see below |
| `R7` never fires | only one PDF in `inbox/` | keep the last week of files in there |

## The standing limitation

The report is an Excel sheet printed to PDF, so the cell grid is gone and the parser
is reading text in reading order. It is good enough — it found every issue in the
August reports — but it is inherently brittle, and every PDF-sourced run is tagged
with a `SOURCE` flag for that reason.

**One WhatsApp message to Hamza asking for the `.xlsx` instead of a printout removes
this entire class of risk.** Ask once; it costs him nothing.
