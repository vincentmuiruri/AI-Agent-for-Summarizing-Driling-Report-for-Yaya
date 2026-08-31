#!/usr/bin/env python3
"""
Yaya Resources — Master Workbook Audit (Luika / Shanta)
--------------------------------------------------------
Run this on Master_<Month>_<Year>.xlsx BEFORE the invoice goes to Shanta.

It reconciles the three layers that currently do not talk to each other:

    DAILY REPORT (shift log, typed and overwritten each day)
        -> ACTUAL DRIL   (hand-keyed daily grid, per depth band and core size)
            -> TAX INVOICE (formula-linked to ACTUAL DRIL band totals)

The single most important check is INV-1: every invoice line must pull from its
own band's monthly total column. In the August 2026 workbook the HQ 000-100m
line pulled from the wrong cell and dropped 120 m off the invoice.

Usage:
    python master_audit.py "Master August (2026).xlsx"
    python master_audit.py "Master August (2026).xlsx" --json
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import openpyxl

# ---------------------------------------------------------------------------
# Layout of the workbook as it stands. Change here, not in the logic below.
# ---------------------------------------------------------------------------

SH_ACTUAL, SH_PLAN, SH_INV = "ACTUAL DRIL", "PLANNED DRIL", "TAX INVOICE"
SH_DAILY, SH_MONTH, SH_MASTER = "DAILY REPORT", "MONTH REPORT", "MASTAR"

# invoice row -> (band label, ACTUAL DRIL row holding that band)
INVOICE_BANDS = {
    23: ("NQ 000-100m", 10), 24: ("NQ 101-250m", 11), 25: ("NQ 251-400m", 12),
    26: ("NQ 401-500m", 13), 27: ("NQ 501-750m", 14), 28: ("PQ 000-100m", 15),
    29: ("PQ 101-250m", 16), 30: ("HQ 000-100m", 17), 31: ("HQ 101-250m", 18),
    32: ("HQ 251-400m", 19),
}
BAND_TOTAL_COL = "AH"      # monthly total column on ACTUAL DRIL / PLANNED DRIL
FIRST_BD_COL, LAST_BD_COL = 3, 33   # C..AG = BD1..BD31
BAND_ROWS = list(range(10, 20))
DAILY_TOTAL_ROW = 24       # ACTUAL DRIL row of per-day totals
PLAN_TOTAL_ROW = 25        # PLANNED DRIL row of per-day totals
VAT_RATE = 0.18


def _f(v):
    return float(v) if isinstance(v, (int, float)) else None


class Audit:
    def __init__(self, path: Path):
        self.path = path
        self.v = openpyxl.load_workbook(path, data_only=True)   # cached values
        self.f = openpyxl.load_workbook(path, data_only=False)  # formulas
        self.findings: list[dict] = []
        self.facts: dict = {}

    def flag(self, code, severity, message, money=None):
        self.findings.append(
            {"code": code, "severity": severity, "message": message, "tzs": money}
        )

    # -- INV-1: every invoice line must pull its own band's monthly total ----
    def check_invoice_links(self):
        inv_f, inv_v = self.f[SH_INV], self.v[SH_INV]
        act = self.v[SH_ACTUAL]
        lost = 0.0
        for row, (label, act_row) in INVOICE_BANDS.items():
            expected = f"'{SH_ACTUAL}'!{BAND_TOTAL_COL}{act_row}"
            got = inv_f[f"F{row}"].value
            if not isinstance(got, str) or expected not in got.replace("$", ""):
                drilled = _f(act[f"{BAND_TOTAL_COL}{act_row}"].value) or 0.0
                billed = _f(inv_v[f"F{row}"].value) or 0.0
                rate = _f(inv_v[f"G{row}"].value) or 0.0
                gap = (drilled - billed) * rate
                lost += gap
                self.flag(
                    "INV-1", "CRITICAL",
                    f"Invoice line {label} (F{row}) reads {got!r}; it must read "
                    f"={expected}. Drilled {drilled:g} m, billed {billed:g} m at "
                    f"TZS {rate:,.0f}/m.",
                    money=gap if gap else None,
                )
        if lost:
            self.flag(
                "INV-1T", "CRITICAL",
                f"Total under-billing from broken invoice links: "
                f"TZS {lost:,.0f} ex-VAT / TZS {lost * (1 + VAT_RATE):,.0f} incl VAT.",
                money=lost,
            )
        self.facts["under_billed_ex_vat"] = round(lost, 2)

    # -- INV-2: invoice totals must foot -------------------------------------
    def check_invoice_arithmetic(self):
        inv = self.v[SH_INV]
        lines = sum(
            (_f(inv[f"F{r}"].value) or 0) * (_f(inv[f"G{r}"].value) or 0)
            for r in INVOICE_BANDS
        )
        sub = _f(inv["H44"].value) or 0
        vat, total = _f(inv["H45"].value) or 0, _f(inv["H46"].value) or 0
        self.facts.update(subtotal=sub, vat=vat, total=total)
        if abs(lines - sub) > 1:
            self.flag("INV-2", "HIGH",
                      f"Line items sum to TZS {lines:,.0f} but the sub-total cell says "
                      f"TZS {sub:,.0f}.")
        if abs(sub * VAT_RATE - vat) > 1:
            self.flag("INV-3", "HIGH", f"VAT is not {VAT_RATE:.0%} of the sub-total.")
        if abs(sub + vat - total) > 1:
            self.flag("INV-4", "HIGH", "Payable total does not equal sub-total plus VAT.")
        inv_no = self.v[SH_INV]["D13"].value
        if isinstance(inv_no, (int, float)):
            self.flag("INV-5", "MEDIUM",
                      f"Invoice number is a raw number ({inv_no}), not a document "
                      f"reference. TRA-facing documents need a stable sequential number.")

    # -- ACT-1: band rows must foot to the daily-total row -------------------
    def check_actual_grid(self):
        act = self.v[SH_ACTUAL]
        band_sum = sum(_f(act[f"{BAND_TOTAL_COL}{r}"].value) or 0 for r in BAND_ROWS)
        day_sum = _f(act[f"{BAND_TOTAL_COL}{DAILY_TOTAL_ROW}"].value) or 0
        self.facts["metres_drilled_mtd"] = day_sum
        if abs(band_sum - day_sum) > 0.01:
            self.flag("ACT-1", "HIGH",
                      f"Depth-band totals sum to {band_sum:g} m but the daily-total row "
                      f"says {day_sum:g} m.")

        # last populated business day, and plan to that same day
        plan = self.v[SH_PLAN]
        last_bd = 0
        for i, col in enumerate(range(FIRST_BD_COL, LAST_BD_COL + 1), start=1):
            if (_f(act.cell(row=DAILY_TOTAL_ROW, column=col).value) or 0) > 0:
                last_bd = i
        act_td = sum(_f(act.cell(row=DAILY_TOTAL_ROW, column=c).value) or 0
                     for c in range(FIRST_BD_COL, FIRST_BD_COL + last_bd))
        plan_td = sum(_f(plan.cell(row=PLAN_TOTAL_ROW, column=c).value) or 0
                      for c in range(FIRST_BD_COL, FIRST_BD_COL + last_bd))
        self.facts.update(last_business_day=last_bd, actual_to_date=act_td,
                          plan_to_date=plan_td)
        if plan_td:
            self.facts["attainment_to_date"] = round(act_td / plan_td, 4)

    # -- GOAL-1: the headline % must compare like with like ------------------
    def check_goal_chasing(self):
        daily_f, daily_v = self.f[SH_DAILY], self.v[SH_DAILY]
        formula = daily_f["J4"].value
        shown = _f(daily_v["J4"].value)
        if isinstance(formula, str) and "Y38" in formula:
            self.flag("GOAL-1", "HIGH",
                      f"'Goal Chasing' ({shown:.1%} as shown) divides month-to-date "
                      f"invoice value by the FULL-month plan value. It cannot reach 100% "
                      f"until the last day of the month and tells the reader nothing "
                      f"about whether the rig is on plan today. Compare to plan-to-date.")

    # -- HIST-1: the shift log keeps no history ------------------------------
    def check_daily_history(self):
        d = self.v[SH_DAILY]
        self.facts["daily_sheet_date"] = str(d["A6"].value)
        self.flag("HIST-1", "HIGH",
                  f"The DAILY REPORT sheet holds one day only (currently "
                  f"{self.facts['daily_sheet_date']}) and is overwritten each morning. "
                  f"Rig hours, fuel, hole ID, driller and the operations note are "
                  f"destroyed daily. No availability or cost-per-hour analysis is "
                  f"possible from this workbook.")
        # nothing links the shift metres to the ACTUAL DRIL grid
        self.flag("HIST-2", "HIGH",
                  "Nothing links the shift log (B17 + H17) to the ACTUAL DRIL grid. "
                  "The metres that generate the invoice are re-keyed by hand from the "
                  "metres the client supervisor signs. Two numbers, no reconciliation.")

    # -- COST-1: is any cost actually captured? ------------------------------
    def check_costs(self):
        m = self.v[SH_MONTH]
        fuel_l = _f(m["G19"].value)
        empty = [lbl for lbl, cell in
                 [("Fuel cost", "G20"), ("Total cost", "G21"),
                  ("Maintenance $/m", "J17"), ("Fuel $/m", "J18"),
                  ("Labour $/m", "J19"), ("Consumable $/m", "J20"),
                  ("Total ops $/m", "J21")]
                 if not (_f(m[cell].value) or 0)]
        if empty:
            self.flag("COST-1", "CRITICAL",
                      "Every cost field on MONTH REPORT is zero or blank: "
                      + ", ".join(empty) +
                      ". Revenue per metre is known; cost per metre is not. Margin on "
                      "this contract cannot be stated.")
        cons = _f(self.v[SH_MASTER]["D44"].value)
        mtd = self.facts.get("metres_drilled_mtd")
        if cons and mtd:
            self.facts["consumables_tzs"] = cons
            self.facts["consumables_per_metre"] = round(cons / mtd, 0)
            self.flag("COST-2", "MEDIUM",
                      f"Consumables of TZS {cons:,.0f} are captured on MASTAR "
                      f"(TZS {cons / mtd:,.0f} per metre) but do not flow into the "
                      f"'Consumable ($/m)' field on MONTH REPORT.")
        if fuel_l and mtd:
            self.facts["fuel_litres_mtd"] = fuel_l
            self.facts["fuel_litres_per_metre"] = round(fuel_l / mtd, 2)

    # -- REF-1: broken sheets ------------------------------------------------
    def check_broken_refs(self):
        for ws in self.v.worksheets:
            bad = sum(1 for row in ws.iter_rows() for c in row
                      if isinstance(c.value, str) and "#REF!" in c.value)
            if bad:
                self.flag("REF-1", "HIGH" if bad > 50 else "MEDIUM",
                          f"Sheet '{ws.title}' contains {bad} #REF! errors "
                          f"(state: {ws.sheet_state}). It is not producing usable output.")

    def run(self):
        for fn in (self.check_invoice_links, self.check_invoice_arithmetic,
                   self.check_actual_grid, self.check_goal_chasing,
                   self.check_daily_history, self.check_costs,
                   self.check_broken_refs):
            try:
                fn()
            except Exception as e:                     # a broken sheet must not
                self.flag("AUDIT-ERR", "MEDIUM",       # abort the whole audit
                          f"{fn.__name__} could not complete: {e}")
        return self


def report(a: Audit) -> str:
    order = {"CRITICAL": 0, "HIGH": 1, "MEDIUM": 2, "LOW": 3}
    out = [f"WORKBOOK AUDIT — {a.path.name}", "=" * 62, ""]
    f = a.facts
    if "actual_to_date" in f:
        out.append(f"Metres to BD{f['last_business_day']}: {f['actual_to_date']:g} actual "
                   f"vs {f['plan_to_date']:g} plan "
                   f"({f.get('attainment_to_date', 0):.1%} of plan-to-date)")
    if f.get("subtotal"):
        out.append(f"Invoice as built: TZS {f['subtotal']:,.0f} ex-VAT")
    if f.get("under_billed_ex_vat"):
        u = f["under_billed_ex_vat"]
        out.append(f"Invoice if corrected: TZS {f['subtotal'] + u:,.0f} ex-VAT "
                   f"(understated by TZS {u:,.0f})")
    if f.get("consumables_per_metre"):
        out.append(f"Consumables: TZS {f['consumables_per_metre']:,.0f}/m  |  "
                   f"Fuel: {f.get('fuel_litres_per_metre', 'n/r')} L/m")
    out += ["", f"{len(a.findings)} finding(s)", ""]
    for x in sorted(a.findings, key=lambda d: order.get(d["severity"], 9)):
        out.append(f"[{x['severity']:8}] {x['code']:8} {x['message']}")
        out.append("")
    return "\n".join(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("path")
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()
    p = Path(args.path)
    if not p.exists():
        sys.exit(f"No such file: {p}")
    a = Audit(p).run()
    if args.json:
        print(json.dumps({"facts": a.facts, "findings": a.findings},
                         indent=2, default=str))
    else:
        print(report(a))
    # non-zero exit if anything critical, so this can gate an invoice run
    sys.exit(1 if any(x["severity"] == "CRITICAL" for x in a.findings) else 0)


if __name__ == "__main__":
    main()
