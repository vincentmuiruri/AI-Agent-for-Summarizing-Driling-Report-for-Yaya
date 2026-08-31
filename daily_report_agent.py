#!/usr/bin/env python3
"""
Yaya Resources — Luika Drilling Daily Report Agent
--------------------------------------------------
Turns Hamza's daily rig report into a CEO-facing update.

Design rule that matters: PYTHON DOES THE ARITHMETIC, CLAUDE DOES THE LANGUAGE.
The model is never asked to add, subtract, divide or judge a number. It receives
already-computed fields and an already-computed exception list, and writes prose.
This is what stops the agent inventing a figure you then have to defend.

Usage:
    python daily_report_agent.py inbox/Daily_Report26.xlsx
    python daily_report_agent.py inbox/Daily_Report26.pdf --dry-run

Install:
    pip install anthropic openpyxl pdfplumber
    export ANTHROPIC_API_KEY=sk-...
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
from dataclasses import dataclass, asdict, field
from datetime import datetime, date
from pathlib import Path
from typing import Optional

MODEL = "claude-sonnet-5"
HISTORY = Path("history.csv")
OUTDIR = Path("out")

# Labels as they appear in Hamza's sheet. Left-most occurrence = DAY shift,
# next occurrence to the right = NIGHT shift.
LABELS = {
    "rig_no": "Rig No.",
    "start_time": "Start Time",
    "end_shift": "End Shift",
    "start_rig_hours": "Start Rig Hours",
    "end_rig_hours": "End Rig Hours",
    "start_meter": "Start Meter",
    "end_meter": "End Meter",
    "meter_drilled": "Meter Drilled",
    "core_type": "Core Type",
    "fuel_litres": "Fuel Liters",
    "driller": "Driller Name",
    "client_supervisor": "Client Supervisor",
    "location": "Location",
    "hole_id": "Hole ID",
}


# ----------------------------------------------------------------------------
# 1. EXTRACT
# ----------------------------------------------------------------------------

@dataclass
class Shift:
    name: str
    rig_no: Optional[str] = None
    start_time: Optional[str] = None
    end_shift: Optional[str] = None
    start_rig_hours: Optional[float] = None
    end_rig_hours: Optional[float] = None
    start_meter: Optional[float] = None
    end_meter: Optional[float] = None
    meter_drilled: Optional[float] = None
    core_type: Optional[str] = None
    fuel_litres: Optional[float] = None
    driller: Optional[str] = None
    client_supervisor: Optional[str] = None
    location: Optional[str] = None
    hole_id: Optional[str] = None

    # derived
    rig_hours: Optional[float] = None
    metres_from_counter: Optional[float] = None
    rop_m_per_hr: Optional[float] = None
    fuel_per_rig_hour: Optional[float] = None
    fuel_per_metre: Optional[float] = None
    scheduled_hours: Optional[float] = None


@dataclass
class DailyReport:
    report_date: Optional[str] = None
    source_file: str = ""
    day: Shift = field(default_factory=lambda: Shift("day"))
    night: Shift = field(default_factory=lambda: Shift("night"))
    wtd_today_metres: Optional[float] = None
    wtd_today_plan: Optional[float] = None
    wtd_today_core_band: Optional[str] = None
    operation_update: str = ""
    total_metres: Optional[float] = None
    total_rig_hours: Optional[float] = None
    total_fuel: Optional[float] = None
    day_rop: Optional[float] = None
    exceptions: list = field(default_factory=list)


def _num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(",", "").replace("(", "-").replace(")", "").strip()
    if s in {"", "-", "—", "#REF!", "#DIV/0!"}:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def extract_xlsx(path: Path) -> DailyReport:
    """Preferred path. Reads by label, not by hardcoded cell address, so the
    parser survives Hamza nudging a row up or down."""
    from openpyxl import load_workbook

    wb = load_workbook(path, data_only=True)  # data_only -> cached values, not formulas
    ws = wb["DAILY REPORT"] if "DAILY REPORT" in wb.sheetnames else wb.active

    grid = {}
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None:
                grid[(c.row, c.column)] = c.value

    def hits(label: str):
        """All (row, col) where a cell's text starts with the label."""
        out = []
        for (r, col), v in grid.items():
            if isinstance(v, str) and v.strip().lower().startswith(label.lower()):
                out.append((r, col))
        return sorted(out, key=lambda rc: rc[1])  # left to right = day, then night

    def value_right(r, col, span=6):
        for k in range(1, span + 1):
            v = grid.get((r, col + k))
            if v is not None and str(v).strip() != "":
                return v
        return None

    rep = DailyReport(source_file=str(path))

    for field_name, label in LABELS.items():
        found = hits(label)
        for shift_obj, rc in zip((rep.day, rep.night), found):
            raw = value_right(*rc)
            if field_name in {
                "start_rig_hours", "end_rig_hours", "start_meter",
                "end_meter", "meter_drilled", "fuel_litres",
            }:
                setattr(shift_obj, field_name, _num(raw))
            else:
                setattr(shift_obj, field_name, None if raw is None else str(raw).strip())

    # report date: first date-looking cell in the top 12 rows
    for (r, c), v in sorted(grid.items()):
        if r <= 12:
            if isinstance(v, (datetime, date)):
                rep.report_date = v.strftime("%Y-%m-%d")
                break
            if isinstance(v, str) and re.match(r"^\d{1,2}-[A-Za-z]{3}-\d{2,4}$", v.strip()):
                rep.report_date = v.strip()
                break

    # WTD table: today's column is the one matching the sheet's own BD pointer
    try:
        act = load_workbook(path, data_only=True)["ACTUAL DRIL"]
        bd = int(act["B6"].value)
        col = 2 + bd                      # C = BD1
        rep.wtd_today_metres = _num(act.cell(row=24, column=col).value)
        for r in range(10, 20):           # find the band that carries today's metres
            if (_num(act.cell(row=r, column=col).value) or 0) > 0:
                rep.wtd_today_core_band = f"{act[f'B{r}'].value} {act[f'A{r}'].value}"
        pln = load_workbook(path, data_only=True)["PLANNED DRIL"]
        rep.wtd_today_plan = _num(pln.cell(row=25, column=col).value)
    except (KeyError, TypeError, ValueError):
        pass

    # operation update: the text under the OPERATION UPDATE label
    lbl = [rc for rc, v in grid.items()
           if isinstance(v, str) and "UPDATE" in v.upper() and len(v) < 40]
    if lbl:
        r0, c0 = sorted(lbl)[0]
        below = [v for (r, c), v in sorted(grid.items())
                 if r0 < r <= r0 + 6 and isinstance(v, str) and len(v.strip()) > 3]
        rep.operation_update = " ".join(x.strip() for x in below)
    if not rep.operation_update:
        texts = [v for v in grid.values() if isinstance(v, str) and len(v) > 60]
        rep.operation_update = max(texts, key=len).strip() if texts else ""
    if not rep.operation_update:
        rep.exceptions.append(
            "R10: no operations note written for the day. If the rig ran clean, that "
            "must be stated; a blank box is not the same as 'no issues'.")

    return rep


def extract_pdf(path: Path) -> DailyReport:
    """Fallback only. A PDF of a spreadsheet loses the cell grid, so this is
    line-order guesswork and WILL break when the layout changes. Get the .xlsx."""
    import pdfplumber

    with pdfplumber.open(path) as pdf:
        text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    rep = DailyReport(source_file=str(path))
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    def pair(label):
        """The template prints each label twice on one line: day value then night
        value ("Start Rig Hours 1748.0 Start Rig Hours 1738.0"). Split on the
        label itself and keep the non-empty segments in order."""
        for l in lines:
            if l.lower().startswith(label.lower()):
                segs = [s.strip() for s in re.split(re.escape(label), l, flags=re.I)]
                return [s for s in segs if s]
        return []

    for field_name, label in LABELS.items():
        toks = pair(label)
        for shift_obj, tok in zip((rep.day, rep.night), toks):
            if field_name in {"start_rig_hours", "end_rig_hours", "start_meter",
                              "end_meter", "meter_drilled", "fuel_litres"}:
                setattr(shift_obj, field_name, _num(tok))
            else:
                setattr(shift_obj, field_name, tok)

    m = re.search(r"(\d{1,2}-[A-Za-z]{3}-\d{2,4})", text)
    if m:
        rep.report_date = m.group(1)

    m = re.search(r"OPERAT\w*\s*UPDATE:?\s*(.+?)(?:\n\s*\n|W E E K|$)", text, re.S | re.I)
    if m:
        rep.operation_update = " ".join(m.group(1).split())

    rep.exceptions.append(
        "SOURCE: parsed from PDF, not Excel. Field mapping is positional and "
        "unverified. Treat every figure as unconfirmed until the .xlsx is supplied."
    )
    return rep


# ----------------------------------------------------------------------------
# 2. COMPUTE + VALIDATE  (all arithmetic lives here, none in the model)
# ----------------------------------------------------------------------------

def clock_hours(start: str, end: str) -> Optional[float]:
    try:
        sh, sm = [int(x) for x in re.split(r"[:.]", start)[:2]]
        eh, em = [int(x) for x in re.split(r"[:.]", end)[:2]]
    except Exception:
        return None
    s, e = sh * 60 + sm, eh * 60 + em
    if e <= s:
        e += 24 * 60  # crosses midnight
    return round((e - s) / 60, 2)


def compute(rep: DailyReport) -> DailyReport:
    ex = rep.exceptions

    for s in (rep.day, rep.night):
        if s.start_rig_hours is not None and s.end_rig_hours is not None:
            s.rig_hours = round(s.end_rig_hours - s.start_rig_hours, 2)
        if s.start_meter is not None and s.end_meter is not None:
            s.metres_from_counter = round(s.end_meter - s.start_meter, 2)
        if s.start_time and s.end_shift:
            s.scheduled_hours = clock_hours(s.start_time, s.end_shift)

        # R1 counter vs reported metres
        if s.metres_from_counter is not None and s.meter_drilled is not None:
            if abs(s.metres_from_counter - s.meter_drilled) > 0.01:
                ex.append(
                    f"R1 {s.name.upper()}: metre counter says "
                    f"{s.metres_from_counter:.1f} m ({s.start_meter}->{s.end_meter}) but "
                    f"'Meter Drilled' says {s.meter_drilled:.1f} m."
                )

        # R2 rig hours sane
        if s.rig_hours is None:
            ex.append(
                f"R2 {s.name.upper()}: rig hour meter readings missing, so this shift "
                f"contributes nothing to the day's rig hours. Day totals are understated."
            )
        else:
            if s.rig_hours <= 0:
                ex.append(f"R2 {s.name.upper()}: rig hours {s.rig_hours} is zero or negative.")
            elif s.scheduled_hours and s.rig_hours > s.scheduled_hours + 0.25:
                ex.append(
                    f"R2 {s.name.upper()}: rig ran {s.rig_hours:.1f} h inside a "
                    f"{s.scheduled_hours:.2f} h shift window "
                    f"({s.start_time}-{s.end_shift}). Impossible as stated — either the "
                    f"shift times or the hour-meter readings are wrong."
                )

        # R3 rates of penetration and fuel
        if s.rig_hours and s.metres_from_counter is not None and s.rig_hours > 0:
            s.rop_m_per_hr = round(s.metres_from_counter / s.rig_hours, 2)
        if s.fuel_litres is None:
            ex.append(f"R3 {s.name.upper()}: fuel litres not recorded. Cost per metre cannot be built.")
        else:
            if s.rig_hours:
                s.fuel_per_rig_hour = round(s.fuel_litres / s.rig_hours, 2)
            if s.metres_from_counter:
                s.fuel_per_metre = round(s.fuel_litres / s.metres_from_counter, 2)

        # R4 mandatory fields
        for f_ in ("hole_id", "core_type", "driller", "location"):
            if not getattr(s, f_):
                ex.append(f"R4 {s.name.upper()}: '{f_}' is blank.")

    # R5 new hole started but hole ID unchanged
    if rep.day.start_meter == 0 and rep.night.hole_id and rep.day.hole_id == rep.night.hole_id:
        ex.append(
            f"R5: metre counter reset to 0 (new hole started) but Hole ID is still "
            f"{rep.day.hole_id}. Previous hole's final depth and the new hole ID are not recorded."
        )

    # R6 shift-to-shift continuity within the day
    hrs = [s.rig_hours for s in (rep.day, rep.night) if s.rig_hours]
    rep.total_rig_hours = round(sum(hrs), 2) if hrs else None
    mets = [s.metres_from_counter for s in (rep.day, rep.night) if s.metres_from_counter is not None]
    rep.total_metres = round(sum(mets), 2) if mets else None
    fuels = [s.fuel_litres for s in (rep.day, rep.night) if s.fuel_litres is not None]
    rep.total_fuel = round(sum(fuels), 2) if fuels else None
    if rep.total_metres is not None and rep.total_rig_hours:
        rep.day_rop = round(rep.total_metres / rep.total_rig_hours, 2)

    # R7 continuity against yesterday
    prev = _last_history_row()
    if prev:
        try:
            prev_end = float(prev["max_rig_hours"])
            todays = [x for s in (rep.day, rep.night)
                      for x in (s.start_rig_hours, s.end_rig_hours) if x is not None]
            if todays:
                gap = round(min(todays) - prev_end, 2)
                if gap > 0.1:
                    ex.append(
                        f"R7: {gap:.1f} rig hours unaccounted between the previous report "
                        f"(ended {prev_end:.1f}) and this one (starts {min(todays):.1f}). "
                        f"Rig time ran with no shift recorded against it."
                    )
                elif gap < -0.1:
                    ex.append(
                        f"R7: hour meter goes backwards by {abs(gap):.1f} h versus the "
                        f"previous report. Readings are being transcribed wrongly."
                    )
        except (KeyError, TypeError, ValueError):
            pass

    # R8 WTD table vs shift log
    if rep.wtd_today_metres is not None and rep.total_metres is not None:
        if abs(rep.wtd_today_metres - rep.total_metres) > 0.01:
            ex.append(
                f"R8: week-to-date table records {rep.wtd_today_metres:.1f} m for today but "
                f"the shift log totals {rep.total_metres:.1f} m. The WTD table is what feeds "
                f"the invoice — this is a billing exposure, not a typo."
            )
    if rep.wtd_today_core_band and rep.day.core_type:
        if rep.day.core_type.upper() not in rep.wtd_today_core_band.upper():
            ex.append(
                f"R8b: shift log says core type {rep.day.core_type} but the WTD table books "
                f"today under '{rep.wtd_today_core_band}'. Core sizes bill at different rates."
            )
    if rep.wtd_today_plan in (None, 0):
        ex.append("R9: no plan metres set for today, so variance against plan is meaningless.")

    return rep


def _last_history_row():
    if not HISTORY.exists():
        return None
    with HISTORY.open() as f:
        rows = list(csv.DictReader(f))
    return rows[-1] if rows else None


def append_history(rep: DailyReport):
    HISTORY.parent.mkdir(parents=True, exist_ok=True)
    all_hours = [x for s in (rep.day, rep.night)
                 for x in (s.start_rig_hours, s.end_rig_hours) if x is not None]
    row = {
        "report_date": rep.report_date,
        "rig_no": rep.day.rig_no or rep.night.rig_no,
        "hole_id": rep.day.hole_id,
        "metres": rep.total_metres,
        "rig_hours": rep.total_rig_hours,
        "rop_m_per_hr": rep.day_rop,
        "fuel_litres": rep.total_fuel,
        "max_rig_hours": max(all_hours) if all_hours else "",
        "n_exceptions": len(rep.exceptions),
        "source_file": rep.source_file,
    }
    new = not HISTORY.exists()
    with HISTORY.open("a", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(row))
        if new:
            w.writeheader()
        w.writerow(row)


# ----------------------------------------------------------------------------
# 3. NARRATE  (Claude writes; it does not calculate)
# ----------------------------------------------------------------------------

SYSTEM = """You write the daily drilling update for the CEO of Yaya Resources, a
Tanzanian mining contractor. The rig is on Shanta's New Luika site.

Absolute rules:
- Use ONLY the numbers in the JSON you are given. Do not add, subtract, average,
  convert or estimate any figure. If a number is not in the JSON, write "not recorded".
- Do not smooth over the exceptions list. Every exception must appear.
- No filler, no praise, no restating the question. Short sentences.
- If the exceptions list is empty, say the checks passed. Do not invent concerns.

Output exactly these five blocks and nothing else:

PRODUCTION — metres and rig hours by shift, and the day total.
MACHINE STATUS — what the report says about the rig running or not running. If the
  report contains no fault log, service record or downtime reason, say so plainly.
CONSUMABLES, TOOLS AND SPARES — report what is stated. If nothing is stated, write
  "Not covered in the report submitted."
ISSUES AND ACTIONS — the site's own operations note, then what it does not answer.
DATA INTEGRITY FLAGS — the exceptions list, one line each, plain English.

Under 220 words."""


def narrate(rep: DailyReport, dry_run: bool = False) -> str:
    payload = json.dumps(asdict(rep), indent=2, default=str)
    if dry_run:
        return "[dry-run] payload that would go to Claude:\n\n" + payload

    from anthropic import Anthropic

    client = Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
    msg = client.messages.create(
        model=MODEL,
        max_tokens=1200,
        system=SYSTEM,
        messages=[{"role": "user", "content":
                   f"Rig report data:\n\n{payload}\n\nWrite the CEO update."}],
    )
    return "".join(b.text for b in msg.content if b.type == "text")


# ----------------------------------------------------------------------------
# 4. RUN
# ----------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("path", help="Hamza's report (.xlsx preferred, .pdf fallback)")
    ap.add_argument("--dry-run", action="store_true", help="skip the API call")
    args = ap.parse_args()

    p = Path(args.path)
    if not p.exists():
        sys.exit(f"No such file: {p}")

    rep = extract_xlsx(p) if p.suffix.lower() in {".xlsx", ".xlsm"} else extract_pdf(p)
    rep = compute(rep)

    body = narrate(rep, dry_run=args.dry_run)

    OUTDIR.mkdir(exist_ok=True)
    stamp = rep.report_date or datetime.now().strftime("%Y-%m-%d")
    out = OUTDIR / f"{stamp}_ceo_update.md"
    header = (
        f"# Luika drilling — daily update, {stamp}\n"
        f"Rig {rep.day.rig_no or 'n/r'} | source: {p.name} | "
        f"{len(rep.exceptions)} data flag(s)\n\n"
    )
    out.write_text(header + body + "\n")
    append_history(rep)

    print(header + body)
    print(f"\nWritten: {out}\nHistory: {HISTORY}")


if __name__ == "__main__":
    main()
