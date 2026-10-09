"""Domus budget app — Streamlit. Layout and levy math follow Domus Excel packs."""

from __future__ import annotations

import json
import re
from io import BytesIO
from pathlib import Path

from meeting_deck import render_meeting_html

import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt
from lxml import etree

st.set_page_config(page_title="Domus Budget", layout="wide", initial_sidebar_state="expanded")

NOW_YEAR = 2026
YELLOW, NAVY, BLUE, RED, TOTAL, SECTION = "FFFF99", "1F4E79", "0000FF", "FFC7CE", "D9E2F3", "2E75B6"
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
MONEY = '#,##0.00;(#,##0.00);0.00'


def uid() -> str:
    import random, string
    return "".join(random.choices(string.ascii_lowercase + string.digits, k=8))


def money(n: float) -> str:
    return f"R {n:,.2f}"


def pct_from_amounts(actual: float, yearly: float) -> float:
    """% increase from last year’s actual to this year’s budgeted yearly."""
    a = float(actual or 0)
    y = float(yearly or 0)
    if a < 0.5:
        return 0.0
    return (y / a) * 100.0 - 100.0


def clean_note(v) -> str:
    if v is None:
        return ""
    try:
        if isinstance(v, float) and pd.isna(v):
            return ""
    except Exception:
        pass
    s = str(v).strip()
    if s.lower() in ("nan", "none", "nat", "<na>"):
        return ""
    return s


def norm(s: str) -> str:
    s = str(s or "").lower()
    s = re.sub(r"[–—-]", " ", s)
    s = re.sub(r"[^a-z0-9/& ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def row(desc: str, note: str = "") -> dict:
    return {
        "id": uid(), "desc": desc, "actual": 0.0, "pct": 0.0, "yearly": 0.0,
        "insurance": 0.0, "owner_recovery": 0.0, "note": note, "is_recovery": False,
        "reduces_levy": False, "claim_against": "", "edit_mode": "pct",
    }


def default_sections() -> dict:
    return {
        "levy": [
            row("Ordinary Levies", "Admin levy. Does not include estate/HOA pass-throughs or insurance billed on its own PQ column."),
            row("Reserve Fund Contribution", "Type the yearly amount, or use 15% of ordinary."),
            row("CSOS Levy (Income)", "This complex’s own CSOS. Not the master-estate CSOS."),
            row("Insurance billed to owners", "Own column on the owner invoice. Leave 0 if insurance stays inside the ordinary levy."),
            row("Levy - Boathouse"),
            row("Levy - Boatport"),
            row("Special Levy"),
        ],
        "other": [
            row("Interest on Arrear Levies", "Usually not budgeted (do not rely on arrears)."),
            row("Investment Income"),
            row("Penalty Income"),
            row("Eskom / Electricity meters fixed charge", "Only if this complex bills a fixed charge."),
            row("Communal electricity recovered"),
            row("Rental Income"),
            row("Garage Rental Income"),
            row("Gate Registration / Services"),
            row("Clubhouse Rental"),
        ],
        "hoa_income": [
            row("Estate / HOA Levies recovered", "Master-estate levy billed to owners on the PQ. Not ordinary."),
            row("Estate / HOA CSOS recovered", "Master-estate CSOS. Own PQ column."),
        ],
        "hoa_expense": [
            row("Estate / HOA Levies paid", "Paid to the master estate. Not in ordinary."),
            row("Estate / HOA CSOS paid", "Master-estate CSOS paid. Not in ordinary."),
        ],
        "recoveries_other": [
            row("Insurance claims recovered", "Leave Comes off blank if the payout is already on the repair line. Otherwise pick that repair. Do not do both."),
            row("Legal Fees Recovered", "Pick Legal Expense. The legal bill stays on Expenditure, and this amount comes off it."),
            row("Charged to an owner", "We paid for the repair, then invoiced the owner. Example: a broken window. Pick the repair line this pays."),
        ],
        "municipal": [
            row("Electricity"),
            {**row("Less: Electricity recovered from owners"), "is_recovery": True},
            row("Water"),
            {**row("Less: Water recovered from owners"), "is_recovery": True},
            row("Sewerage"),
            {**row("Less: Sewerage recovered from owners"), "is_recovery": True},
            row("Refuse Removal"),
            {**row("Less: Refuse recovered from owners"), "is_recovery": True},
            {**row("Less: Sewer plant electricity recovered"), "is_recovery": True},
            row("Rates / Property Tax"),
        ],
        "expenditure": [
            row("Accounting Fees"), row("Audit Fees"), row("Bank Charges"),
            row("CSOS Levies (Expense)"), row("Insurance"), row("Management Fees"),
            row("Legal Expense"), row("Security / Guarding"), row("Cleaning Materials"),
            row("Computer Expenses"), row("Printing and Stationery"),
            row("Telephone and Internet"), row("Meeting Expenses"),
            row("Health & Safety"), row("Protective Clothing"),
            row("Office / General Expenses"), row("Property Valuation"),
            row("Garden service (contract)"), row("Motor Vehicle Expense"),
        ],
        "rm": [
            row("Electrical"), row("Fire Equipment"), row("General Building"),
            row("Plumbing / Sewerage"), row("Gate & Intercom"), row("Garden Expenses"),
            row("Roofs & Gutters"), row("Painting / Waterproofing"), row("Pool"),
            row("Electric Fence"), row("CCTV"), row("Paving / Roadways"),
            row("Lifts"), row("Equipment Repairs"), row("Other R&M"),
        ],
        "personnel": [
            row("Salaries & Wages"), row("Casual / Relief Wages"), row("PAYE / UIF"),
            row("Travel"), row("Bonuses & Overtime"), row("WCA / COIDA"),
            row("Pension / Provident Fund"), row("Staff Welfare"), row("Caretaker Fees"),
        ],
        "tax": [row("Taxation Payable", "Based on taxable investment income. Do not leave this blank if the FS has tax.")],
        "special": [
            row("Special Project 1"), row("Special Project 2"), row("Special Project 3"),
        ],
        "fixed": [
            row("Prepaid meters estimate (monthly)"),
            row("Eskom fixed charge (monthly)"),
            row("Communal charge (monthly)"),
        ],
    }


def budget_row(desc: str, actual: float, note: str = "", yearly: float | None = None, pct: float | None = None, **extra) -> dict:
    """Actual is the full-year base. Pass pct (10 means +10%) or yearly. Otherwise the budget starts equal to Actual."""
    rec = row(desc, note)
    rec["actual"] = float(actual)
    if pct is not None:
        rec["pct"] = float(pct)
        rec["yearly"] = round(float(actual) * (1 + float(pct) / 100.0), 2)
        rec["edit_mode"] = "pct"
    else:
        rec["yearly"] = float(actual if yearly is None else yearly)
        rec["pct"] = pct_from_amounts(rec["actual"], rec["yearly"]) if rec["actual"] >= 0.5 else 0.0
        rec["edit_mode"] = "pct"
    rec.update(extra)
    return rec


def key_west_sections() -> dict:
    """Key West Body Corporate draft from the ledger 1 Jan 2026 to 9 Oct 2026.
    Monthly costs are a full year. Once-off jobs are not stretched."""
    return {
        "levy": [
            budget_row(
                "Ordinary Levies",
                5119247.24,
                "Actual is the full-year ordinary levy. Budgeted yearly is worked out from the costs.",
            ),
            row("Reserve Fund Contribution", "The bank balance sits in the reserve box on the left, not in this Actual. Type the contribution, or pick 15%, 25% or 100% of repairs."),
            budget_row("CSOS Levy (Income)", 65568.24, "Own column. Not inside ordinary levies. With the PQ file loaded, the formula can replace the budget."),
        ],
        "other": [
            budget_row("Interest on Arrear Levies", 101584.94, "Not ticked. Do not rely on owners paying late."),
            budget_row("Interest Received - Investec", 59445.66, "Not ticked. Bank interest stays in the reserve.", pct=6),
            budget_row("Interest Received - FNB", 12713.10, "Not ticked. Bank interest stays in the reserve.", pct=4),
            budget_row("Penalty Income", 35324.25, "Not in this budget.", pct=-100),
            budget_row("Cost Recovered", 1040, "Broken window, unit 15, invoiced to the owner. Not in this budget.", pct=-100),
            budget_row("Levy Boathouse", 130590.60, "On Other Income, not on the levy block. Tick Lowers the levy only if this money must bring the ordinary levy down. It is off for now.", pct=-100, reduces_levy=True),
            budget_row("Levy Boatport", 75270.06, "On Other Income, not on the levy block. Tick Lowers the levy only if this money must bring the ordinary levy down. It is off for now.", pct=-100, reduces_levy=True),
            budget_row("Clubhouse Rental", 7500, "Not ticked, so it stays on the sheet and does not change the levy."),
        ],
        "hoa_income": [],
        "hoa_expense": [],
        "recoveries_other": [
            budget_row(
                "Additional Insurance", 92593.53,
                "Claim money already received. Not in this budget. Do not tick it as income.",
                pct=-100,
            ),
            budget_row("CSOS Collections", 3000, "A collection fee. Not in this budget.", pct=-100),
            budget_row(
                "Legal Fees Recovered", 224307.11,
                "Owners were billed this. Not in this budget unless it happens again. If you put a budget here, it comes off Legal Expense.",
                pct=-100,
                claim_against="Legal Expense",
            ),
        ],
        "municipal": [
            budget_row("Electricity", 1957000, "10 city bills, averaged to a year."),
            {**budget_row("Less: Electricity recovered from owners", 1534900, "Comes off the city bill. About R422,100 stays in the levy."), "is_recovery": True},
            budget_row("Sewerage and water plant", 355300, "No recovery line. The full amount stays in the levy."),
            budget_row("Refuse Removal", 311200, "Latest bill × 12. August looked like two months."),
            budget_row("Rates, unit 130", 7800),
            budget_row("Water, common property", 0, "The bill was credited back. Leave at 0 unless a new bill is expected."),
        ],
        "expenditure": [
            budget_row("Accounting Fees", 5775, "Once a year. Not times 12."),
            budget_row("Audit Fees", 13500, "Once a year. Not times 12."),
            budget_row("Bank Charges", 7700),
            budget_row("Cleaning Materials", 46800),
            budget_row("Computer Expenses", 2500),
            budget_row("CSOS Levies (Expense)", 65200, "Same as CSOS income. Not inside ordinary levies."),
            budget_row("Donation", 5000, "Was coded under legal. Leave it out if the trustees do not want it."),
            budget_row("Fuel (generator, petrol, diesel)", 92500),
            budget_row("Garden service (contract)", 595143, "Servest. This is the contract, so it sits here, not under Repairs."),
            budget_row("Insurance Premium", 376316, "Inside the ordinary levy. Latest premium × 12."),
            budget_row("Internet, clubhouse", 19300),
            budget_row("Legal Expense", 40000, "Not the R267,000. Almost all of that was billed back to owners."),
            budget_row("Management Fees", 221393, "New September fee × 12."),
            budget_row("Meeting Expenses", 46000, "Mostly the AGM. Not times 12."),
            budget_row("Meeting Refreshments", 4000),
            budget_row("Office / General Expenses", 1000),
            budget_row("Printing and Stationery", 4000),
            budget_row("Security / Guarding", 1101240, "Bidvest. New monthly fee × 12."),
            budget_row("Subscriptions", 1500),
            budget_row("Telephone", 2900),
            budget_row("Telkom Fibre", 2500, "One bill in the ledger. Not times 12."),
            budget_row("TimeTec / intercom", 52400, "Normal monthly fee × 12. Ask before adding January’s extra of about R53,000."),
        ],
        "rm": [
            budget_row("Borehole repairs", 19400, "Repairs only. The small borehole project is under Special."),
            budget_row("Building repairs", 5400),
            budget_row("Electrical", 34000),
            budget_row("Fire Equipment", 18000, "Not times 12."),
            budget_row("Garden equipment", 4800),
            budget_row("Garden Expenses", 28500, "General garden costs. The tree job is under Special."),
            budget_row("Garden refuse", 27600),
            budget_row("Gate & Intercom", 28000, "Repairs. The TimeTec fee is under Expenditure."),
            budget_row("General Building", 102000),
            budget_row("Jungle gym", 1800, "Once-off. Set to 0 if it will not happen again."),
            budget_row("Other R&M", 2000),
            budget_row("Plumbing / Sewerage", 19500),
            budget_row("Pool", 24800, "Was coded as admin. It belongs here."),
            budget_row("Roofs & Gutters", 38000, "Jobs already done. Not times 12."),
            budget_row("Tools", 1000),
            budget_row("Tractor repairs", 24000),
            budget_row("Waterproofing", 50000, "Set to 0 if no more waterproofing is planned."),
        ],
        "personnel": [
            budget_row("Bonuses & Overtime", 35000, "Not times 12."),
            budget_row("Casual / Relief Wages", 117500),
            budget_row("Funeral insurance", 5400),
            budget_row("PAYE / UIF", 157900, "Use R146,000 if the high January amount was a once-off."),
            budget_row("Payroll costs", 19300),
            budget_row("Pension / Provident Fund", 29300),
            budget_row("Protective Clothing", 20400),
            budget_row("Salaries & Wages", 980700),
            budget_row("Staff expenses", 2000),
            budget_row("Staff Welfare", 10000, "Staff refreshments."),
            budget_row("Travel", 17000),
            budget_row("Union fees", 6100),
            budget_row("WCA / COIDA", 10300, "The net for the year. Not times 12."),
        ],
        "tax": [
            budget_row(
                "Taxation Payable", 13000,
                "Estimate only, about R11,000 to R15,000, if arrears interest and penalties stay unticked. Do not use the R286,107 SARS payment. That is old tax. Confirm with the auditor.",
            ),
        ],
        "special": [
            budget_row("Aluminium windows", 99730, "This year only. Not times 12. The reserve pays unless you tick Special into the levy."),
            budget_row("Borehole project", 2899, "Separate from borehole repairs."),
            budget_row("Chemical dosing pumps", 16880),
            budget_row("Jetty", 39917, "Already moved off Repairs. Do not enter it again there."),
            budget_row("Roadways", 15950),
            budget_row("Staircase lights", 5232),
            budget_row("Tree felling", 29900),
            budget_row("Water-plant sand beds", 50370),
        ],
        "fixed": [],
    }


def is_muni_recovery(desc: str, flag: bool = False) -> bool:
    """Any municipal line with 'recovered' or 'Less:' is subtracted. The name need not say water or electricity."""
    d = (desc or "").lower()
    if "recover" in d or d.startswith("less:"):
        return True
    return False


def net_of(r: dict) -> float:
    y = float(r.get("yearly") or 0)
    ins = float(r.get("insurance") or 0)
    own = float(r.get("owner_recovery") or 0)
    if r.get("is_recovery"):
        return -abs(y)
    # Payout or owner recovery reduces that line. It is not income, and it must not drive the levy negative.
    return max(0.0, y - ins - own)


def sum_net(items: list) -> float:
    return sum(net_of(r) for r in items)


def claims_against(state: dict) -> dict:
    """Yearly rands in Claims, keyed by the cost line they pay."""
    out: dict[str, float] = {}
    for r in (state.get("sections") or {}).get("recoveries_other") or []:
        target = norm(r.get("claim_against") or "")
        amt = abs(float(r.get("yearly") or 0))
        if not target or amt < 0.5:
            continue
        out[target] = out.get(target, 0.0) + amt
    return out


def cost_net(state: dict, r: dict) -> float:
    """What this cost adds to the levy, after payouts and claims."""
    taken = claims_against(state).get(norm(r.get("desc") or ""), 0.0)
    return max(0.0, net_of(r) - taken)


def income_used(state: dict) -> float:
    """Other income the user ticked. Nothing is ticked unless they tick it."""
    total = 0.0
    for r in (state.get("sections") or {}).get("other") or []:
        if r.get("reduces_levy"):
            total += max(0.0, float(r.get("yearly") or 0))
    return total


def insurance_on_pq(state: dict) -> bool:
    return state.get("insurance_mode") == "pq"


def insurance_expense_amount(state: dict) -> float:
    total = 0.0
    for r in state["sections"].get("expenditure") or []:
        if family(r.get("desc") or "") == "insurance":
            total += net_of(r)
    return total


def is_ins_bill_line(desc: str) -> bool:
    f = family(desc or "")
    if f == "ins_bill":
        return True
    return bool(re.search(r"insurance recovered|levy\s*[-–]\s*insurance", desc or "", re.I))


def insurance_bill_amount(state: dict) -> float:
    """What we put on the owner invoice when insurance is extra."""
    typed = float(state.get("insurance_bill_yearly") or 0)
    if typed > 0.5:
        return typed
    for r in state["sections"].get("levy") or []:
        if is_ins_bill_line(r.get("desc") or "") and float(r.get("yearly") or 0) > 0.5:
            return float(r["yearly"])
    return insurance_expense_amount(state)


def equal_charge_keys(desc: str) -> set:
    d = (desc or "").lower()
    keys = set()
    if "eskom" in d or ("fixed" in d and "electr" in d):
        keys.add("eskom")
    if "communal" in d:
        keys.add("communal")
    if "prepaid" in d and "meter" in d:
        keys.add("prepaid")
    return keys


def billed_as_equal_charge(desc: str, state: dict) -> bool:
    """True if this cost is billed as a same-rand extra on the invoice (not in ordinary levies)."""
    keys = equal_charge_keys(desc)
    if not keys:
        return False
    for r in state.get("sections", {}).get("fixed") or []:
        if float(r.get("yearly") or 0) < 0.5:
            continue
        if keys & equal_charge_keys(r.get("desc") or ""):
            return True
    return False


def skip_from_ordinary(r: dict, state: dict) -> bool:
    f = family(r.get("desc") or "")
    d = (r.get("desc") or "").lower()
    if f in ("hoa_levy_exp", "hoa_csos_exp", "hoa_levy_inc", "hoa_csos_inc", "csos_exp", "csos_inc"):
        return True
    if "xanadu" in d or "eco park" in d or "csos" in d:
        return True
    if f == "insurance" and insurance_on_pq(state):
        return True
    if billed_as_equal_charge(r.get("desc") or "", state):
        return True
    return False


def csos_monthly_for_levy(levy_monthly: float) -> float:
    """CSOS Practice Directive: 2% of (monthly admin levy − R500), capped at R40 per unit."""
    return min(40.0, max(0.0, (float(levy_monthly) - 500.0) * 0.02))


def scheme_csos_yearly(state: dict, levy_yearly: float) -> float:
    units = state.get("pq") or []
    if units:
        total_m = 0.0
        for u in units:
            pq = float(u.get("PQ") or 0)
            total_m += csos_monthly_for_levy(levy_yearly * pq / 12.0)
        return total_m * 12.0
    n = max(len(units), 1)
    return csos_monthly_for_levy(levy_yearly / 12.0 / n) * n * 12.0


def municipal_split(state: dict, field: str) -> tuple[float, float]:
    """City-bill total and recovered total for Actual or Budgeted yearly. Recoveries stay positive."""
    gross = 0.0
    rec = 0.0
    for r in state["sections"].get("municipal") or []:
        y = abs(float(r.get(field) or 0))
        if is_muni_recovery(r.get("desc") or "", r.get("is_recovery")):
            rec += y
        else:
            gross += y
    return gross, rec


def municipal_gross_and_rec(state: dict) -> tuple[float, float]:
    """This year’s budget: city bill vs owner recoveries."""
    return municipal_split(state, "yearly")


def municipal_bucket(desc: str) -> str:
    d = (desc or "").lower()
    if "plant" in d and ("electric" in d or "sewer" in d):
        return "Sewer plant electricity"
    if "electric" in d:
        return "Electricity"
    if "water" in d:
        return "Water"
    if "sewer" in d or "effluent" in d:
        return "Sewerage"
    if "refuse" in d or "waste" in d:
        return "Refuse"
    if "rates" in d or "property tax" in d:
        return "Rates"
    return "Other municipal"


def municipal_gaps(state: dict) -> list:
    """Per-service gross vs recovered. Positive gap = under-recovery (levy pays the shortfall)."""
    buckets: dict[str, dict] = {}
    for r in state["sections"].get("municipal") or []:
        desc = r.get("desc") or ""
        y = abs(float(r.get("yearly") or 0))
        if y < 0.5 and abs(float(r.get("actual") or 0)) < 0.5:
            continue
        name = municipal_bucket(desc)
        slot = buckets.setdefault(name, {"name": name, "gross": 0.0, "rec": 0.0, "gross_descs": [], "rec_descs": []})
        if is_muni_recovery(desc, r.get("is_recovery")):
            slot["rec"] += y
            slot["rec_descs"].append(desc)
        else:
            slot["gross"] += y
            slot["gross_descs"].append(desc)
    out = []
    for name, slot in buckets.items():
        gap = slot["gross"] - slot["rec"]
        slot["gap"] = gap
        out.append(slot)
    order = ["Electricity", "Water", "Sewerage", "Refuse", "Sewer plant electricity", "Rates", "Other municipal"]
    out.sort(key=lambda x: order.index(x["name"]) if x["name"] in order else 99)
    return out


def municipal_net(state: dict) -> float:
    g, rec = municipal_gross_and_rec(state)
    return g - rec


def ordinary_total(state: dict) -> float:
    s = state["sections"]
    total = municipal_net(state)
    total += sum(cost_net(state, r) for r in s["rm"] if not skip_from_ordinary(r, state))
    total += sum_net(s["personnel"]) + sum_net(s["tax"])
    total += sum(cost_net(state, r) for r in s["expenditure"] if not skip_from_ordinary(r, state))
    if state.get("special_in_ordinary"):
        total += sum_net(s["special"])
    total -= income_used(state)
    return max(0.0, total)


def ordinary_actual(state: dict) -> float:
    for r in state["sections"].get("levy") or []:
        if family(r.get("desc") or "") == "ordinary":
            return float(r.get("actual") or 0)
    return 0.0


def approved_ordinary(state: dict) -> float:
    """What owners are charged. Costs stay the reference until the meeting types a % or a rand."""
    costs = ordinary_total(state)
    mode = state.get("levy_approve_mode") or "costs"
    if mode == "amount":
        amt = float(state.get("levy_approved_amount") or 0)
        return amt if amt > 0.5 else costs
    if mode == "pct":
        actual = ordinary_actual(state)
        if actual > 0.5:
            return actual * (1 + float(state.get("levy_approved_pct") or 0) / 100.0)
    return costs


def levy_pieces(state: dict) -> list:
    s = state["sections"]
    return [
        ("Net municipal (gross minus recoveries)", municipal_net(state)),
        ("Expenditure", sum(cost_net(state, r) for r in s["expenditure"] if not skip_from_ordinary(r, state))),
        ("R&M after insurance and recovery income", sum(cost_net(state, r) for r in s["rm"] if not skip_from_ordinary(r, state))),
        ("Personnel", sum_net(s["personnel"])),
        ("Tax", sum_net(s["tax"])),
        ("Special (only if ticked)", sum_net(s["special"]) if state.get("special_in_ordinary") else 0.0),
        ("Income used to lower the levy", -income_used(state)),
        ("Insurance premium (billed on PQ, not in levy)", insurance_expense_amount(state) if insurance_on_pq(state) else 0.0),
    ]


def estimate_income_tax(state: dict) -> tuple[float, float]:
    """s 10(1)(e) estimate: levies are exempt. Other income above R50,000 × 27%. Not a SARS assessment."""
    other = 0.0
    for r in state["sections"].get("other") or []:
        d = (r.get("desc") or "").lower()
        if any(k in d for k in ("interest", "invest", "rental", "rent", "penalty", "garage", "clubhouse")):
            other += max(0.0, float(r.get("yearly") or 0))
    taxable = max(0.0, other - 50000.0)
    return other, round(taxable * 0.27, 2)


def odd_budget_lines(state: dict) -> list:
    """Lines that usually explain a huge levy %."""
    flags = []
    for key, items in state["sections"].items():
        for it in items:
            desc = it.get("desc") or ""
            act = float(it.get("actual") or 0)
            y = float(it.get("yearly") or 0)
            pct = float(it.get("pct") or 0)
            if pct > 80:
                flags.append(f"{desc}: % is {pct:.0f} (actual {money(act)} → yearly {money(y)})")
            elif act > 1 and y > act * 8:
                flags.append(f"{desc}: yearly {money(y)} is far above actual {money(act)}")
            if re.search(r"airtime|gate", desc, re.I) and y > 20000:
                flags.append(f"{desc}: looks like another line’s amount landed here")
    rec = sum(net_of(r) for r in state["sections"].get("municipal") or [] if r.get("is_recovery") or "recover" in (r.get("desc") or "").lower())
    gross = sum(net_of(r) for r in state["sections"].get("municipal") or [] if not (r.get("is_recovery") or "recover" in (r.get("desc") or "").lower()))
    if gross > 500000 and rec < gross * 0.4:
        flags.append(
            f"Municipal recoveries {money(abs(rec))} look low against the city bill {money(gross)}. "
            "Check that each recovery was typed as a positive rand."
        )
    return flags


def is_own_scheme_csos(desc: str) -> bool:
    """This complex’s CSOS (not estate/Xanadu, not a collection fee)."""
    d = (desc or "").lower()
    if "csos" not in d:
        return False
    if re.search(r"xanadu|eco park|master|\bhoa\b|estate", d):
        return False
    if "collect" in d:
        return False
    return True


def prev_admin_contributions(state: dict) -> float:
    """Last year’s administrative (ordinary) levies. Scaled to 12 months if Actual is part-year."""
    for r in state["sections"].get("levy") or []:
        if family(r.get("desc") or "") == "ordinary":
            a = float(r.get("actual") or 0)
            months = max(1, int(state.get("actual_months") or 12))
            if 0 < months < 12 and a > 0:
                return a / months * 12.0
            return a
    return 0.0


def reserve_project_spend(state: dict) -> float:
    """This year’s special projects, if the reserve pays them. Zero when they are inside the levy."""
    if state.get("special_in_ordinary"):
        return 0.0
    return sum(max(0.0, float(r.get("yearly") or 0)) for r in (state["sections"].get("special") or []))


def projected_reserve(state: dict) -> float:
    """Opening balance already includes interest earned to date. Do not add it again."""
    opening = float(state.get("reserve_balance") or 0)
    return opening + reserve_contribution(state) - reserve_project_spend(state)


def rm_budget(state: dict) -> float:
    return sum(
        cost_net(state, r)
        for r in (state["sections"].get("rm") or [])
        if not skip_from_ordinary(r, state)
    )


def bc_reserve_rule(state: dict) -> dict:
    """STSMA Regulation 2. Only for a sectional-title body corporate, not an HOA."""
    bal = float(state.get("reserve_balance") or 0)
    prev = prev_admin_contributions(state)
    rm = rm_budget(state)
    admin = ordinary_total(state)
    if prev < 1:
        return {
            "band": "unknown",
            "minimum": 0.0,
            "ratio": None,
            "prev": prev,
            "rm": rm,
            "admin": admin,
        }
    ratio = bal / prev
    if ratio < 0.25:
        return {"band": "15", "minimum": admin * 0.15, "ratio": ratio, "prev": prev, "rm": rm, "admin": admin}
    if ratio < 1.0:
        return {"band": "rm", "minimum": max(0.0, rm), "ratio": ratio, "prev": prev, "rm": rm, "admin": admin}
    return {"band": "none", "minimum": 0.0, "ratio": ratio, "prev": prev, "rm": rm, "admin": admin}


def reserve_contribution(state: dict) -> float:
    """Same four choices for a body corporate and an HOA. The Act does not force the choice."""
    mode = state.get("reserve_mode") or "amount"
    if mode in ("pct15", "15pct", "legal"):
        return prev_admin_contributions(state) * 0.15
    if mode == "pct25":
        return prev_admin_contributions(state) * 0.25
    if mode == "rm100":
        return rm_budget(state)
    return float(state.get("reserve_amount") or 0)


def apply_levy_lines(state: dict) -> None:
    ord_amt = ordinary_total(state)
    for r in state["sections"]["levy"]:
        f = family(r["desc"])
        if f == "ordinary":
            r["yearly"] = ord_amt
            a = float(r.get("actual") or 0)
            r["pct"] = 0.0 if a == 0 else (ord_amt / a) * 100 - 100
        if f == "reserve":
            r["yearly"] = reserve_contribution(state)
            a = float(r.get("actual") or 0)
            r["pct"] = 0.0 if a == 0 else (float(r["yearly"]) / a) * 100 - 100
    charged = approved_ordinary(state)
    if state.get("auto_csos", True):
        csos_y = scheme_csos_yearly(state, charged)
    else:
        csos_y = 0.0
        for r in state["sections"]["levy"]:
            if is_own_scheme_csos(r.get("desc") or "") and family(r["desc"]) != "ordinary":
                csos_y = float(r.get("yearly") or 0)
                break
    if csos_y or state.get("auto_csos", True):
        found_exp = False
        for r in state["sections"]["levy"]:
            if is_own_scheme_csos(r.get("desc") or ""):
                r["yearly"] = csos_y
                a = float(r.get("actual") or 0)
                r["pct"] = 0.0 if a == 0 else (csos_y / a) * 100 - 100
        for r in state["sections"].get("expenditure") or []:
            if is_own_scheme_csos(r.get("desc") or ""):
                r["yearly"] = csos_y
                a = float(r.get("actual") or 0)
                r["pct"] = 0.0 if a == 0 else (csos_y / a) * 100 - 100
                found_exp = True
        if csos_y > 0.5 and not found_exp:
            rec = row("CSOS Levies (Expense)", "Same amount as CSOS income — we collect from owners and pay CSOS.")
            rec["yearly"] = csos_y
            state["sections"]["expenditure"].append(rec)
    if state.get("has_master_hoa"):
        hoa_levy_y = sum(
            float(r.get("yearly") or 0)
            for r in (state["sections"].get("hoa_income") or [])
            if family(r.get("desc") or "") != "hoa_csos_inc"
        )
        hoa_csos_y = scheme_csos_yearly(state, hoa_levy_y)
        for r in state["sections"].get("hoa_income") or []:
            if family(r.get("desc") or "") == "hoa_csos_inc":
                r["yearly"] = hoa_csos_y
                a = float(r.get("actual") or 0)
                r["pct"] = 0.0 if a == 0 else (hoa_csos_y / a) * 100 - 100
        for r in state["sections"].get("hoa_expense") or []:
            if family(r.get("desc") or "") == "hoa_csos_exp":
                r["yearly"] = hoa_csos_y
                a = float(r.get("actual") or 0)
                r["pct"] = 0.0 if a == 0 else (hoa_csos_y / a) * 100 - 100
    if insurance_on_pq(state):
        typed = float(state.get("insurance_bill_yearly") or 0)
        found = False
        for r in state["sections"]["levy"]:
            if not is_ins_bill_line(r.get("desc") or ""):
                continue
            found = True
            current = float(r.get("yearly") or 0)
            if typed > 0.5:
                r["yearly"] = typed
            elif current < 0.5:
                r["yearly"] = insurance_expense_amount(state)
            # else keep the amount the user typed on the line
            a = float(r.get("actual") or 0)
            y = float(r.get("yearly") or 0)
            r["pct"] = 0.0 if a == 0 else (y / a) * 100 - 100
        if not found:
            bill = typed if typed > 0.5 else insurance_expense_amount(state)
            if bill:
                rec = row("Insurance recovered", "Billed to owners on the PQ. Change this amount if the quote differs from last year’s premium.")
                rec["yearly"] = bill
                rec["actual"] = 0.0
                state["sections"]["levy"].append(rec)


def pq_bill_lines(state: dict) -> list:
    """Owner-invoice columns. split 'pq' = × participation quota; 'equal' = same rand each unit."""
    s = state["sections"]
    out = []

    def add(name, yearly, split="pq"):
        out.append((str(name), float(yearly or 0), split))

    for r in s.get("levy", []):
        if family(r["desc"]) == "ordinary":
            add("Levies", approved_ordinary(state))
            break
    if not out:
        add("Levies", approved_ordinary(state))
    if state.get("has_master_hoa"):
        for r in s.get("hoa_income", []):
            if float(r.get("yearly") or 0) or float(r.get("actual") or 0):
                add(r["desc"], r.get("yearly"))
    if insurance_on_pq(state):
        add("Insurance", insurance_bill_amount(state))
    for r in s.get("levy", []):
        if family(r["desc"]) == "reserve" and (float(r.get("yearly") or 0) or float(r.get("actual") or 0)):
            add("Reserve Fund", r.get("yearly"))
            break
    for r in s.get("levy", []):
        if family(r["desc"]) == "csos_inc" and (float(r.get("yearly") or 0) or float(r.get("actual") or 0)):
            add("CSOS", r.get("yearly"))
            break
    for r in s.get("levy", []):
        d = r["desc"].lower()
        if any(x in d for x in ("boathouse", "boatport", "special levy")) and (
            float(r.get("yearly") or 0) or float(r.get("actual") or 0)
        ):
            add(r["desc"], r.get("yearly"))
    for r in s.get("fixed") or []:
        y = float(r.get("yearly") or 0)
        if y < 0.5:
            continue
        d = (r.get("desc") or "").lower()
        if "insurance" in d and insurance_on_pq(state):
            continue
        add(r.get("desc") or "Extra charge", y, "equal")
    return out


def family(desc: str) -> str | None:
    d = norm(desc)
    master = bool(re.search(r"xanadu|eco park|master scheme|master hoa|\bhoa\b|estate levy|estate levies", d))
    if master:
        if "csos" in d:
            if re.search(r"hoa csos|csos recovered|csos income|levy\s*[-–].*csos", d):
                return "hoa_csos_inc"
            if re.search(r"csos paid|csos expense", d):
                return "hoa_csos_exp"
            if d.startswith("levy"):
                return "hoa_csos_inc"
            return "hoa_csos_inc"
        if d.startswith("levy") or "recovered" in d or "from owners" in d:
            return "hoa_levy_inc"
        if re.search(r"paid|expense", d) or re.match(r"^xanadu(\s+hoa|\s+eco\s*park)?$", d):
            return "hoa_levy_exp"
        return "hoa_levy_inc"
    if re.search(r"insurance\s*(claim|payout)", d):
        return "ins_claim"
    if re.search(r"levy\s*[-–]\s*insurance|insurance recovered", d):
        return "ins_bill"
    if re.search(r"insurance\s*(recovered|billed|additional)", d):
        return "ins_bill"
    if "eskom" in d or ("fixed" in d and "electr" in d) or "meters recovered" in d:
        return "eskom"
    if "reserve" in d:
        return "reserve"
    if "special levy" in d:
        return "special_levy"
    if "csos" in d:
        if "collect" in d:
            return "csos_col"
        if re.search(r"income|recovered", d):
            return "csos_inc"
        if re.search(r"paid|expense|^csos levy$|^csos levies$", d):
            return "csos_exp"
        return "csos_inc"
    if "boathouse" in d:
        return "boathouse"
    if "boatport" in d:
        return "boatport"
    if d in ("levies", "levy") or "ordinary" in d or re.search(r"^levies?\b", d):
        return "ordinary"
    if "electricity" in d and "recover" in d and "commun" in d:
        return "elec_comm"
    if "electricity" in d and "recover" in d:
        return "elec_rec"
    if "water" in d and "recover" in d:
        return "water_rec"
    if "sewer" in d and "recover" in d:
        return "sewer_rec"
    if "refuse" in d and "recover" in d:
        return "refuse_rec"
    if "interest" in d and "arrear" in d:
        return "int_arr"
    if re.search(r"interest|investment|marketlink", d) and re.search(r"bank|invest|marketlink", d):
        return "invest"
    if "penalty" in d:
        return "penalty"
    if re.search(r"sewer(age)?\s*plant", d) and re.search(r"electric|usage", d):
        return "elec_rec"
    if "rental" in d and "garage" in d:
        return "garage"
    if ("rental" in d or "rent received" in d) and "electric" not in d:
        return "rental"
    if "electricity" in d and "recover" not in d:
        return "elec_g"
    if re.search(r"^water$|water\s*(charge|expense)", d):
        return "water_g"
    if "sewer" in d and "recover" not in d:
        return "sewer_g"
    if "refuse" in d and "recover" not in d:
        return "refuse_g"
    if "management" in d and "fee" in d:
        return "mgmt"
    if re.search(r"^insurance$|insurance\s*premium", d):
        return "insurance"
    if "security" in d or "guarding" in d:
        return "security"
    if re.search(r"salar|staff wages", d):
        return "salaries"
    return None


def section_for(desc: str) -> str:
    f = family(desc)
    d = norm(desc)
    if f in ("ordinary", "reserve", "csos_inc", "boathouse", "boatport", "ins_bill", "special_levy"):
        return "levy"
    if f in ("hoa_levy_inc", "hoa_csos_inc"):
        return "hoa_income"
    if f in ("hoa_levy_exp", "hoa_csos_exp"):
        return "hoa_expense"
    if f == "csos_exp":
        return "expenditure"
    if f == "csos_col":
        return "other"
    if f in ("elec_rec", "water_rec", "sewer_rec", "refuse_rec"):
        return "municipal"
    if f in ("elec_g", "water_g", "sewer_g", "refuse_g"):
        return "municipal"
    if f in ("eskom", "elec_comm", "int_arr", "invest", "penalty", "rental", "garage"):
        return "other"
    if f == "ins_claim":
        return "recoveries_other"
    if f == "salaries":
        return "personnel"
    if re.search(r"garden(ing)?\s*(expense|general|repair)", d):
        return "rm"
    if re.search(r"garden(ing)?\s*(service|contract)", d):
        return "expenditure"
    if re.search(r"site cleaning", d):
        return "rm"
    if re.search(r"\b(repair|maintenance|plumb|paint|roof|gutter|pool|electrical|fire equipment|gate|paving)\b", d):
        return "rm"
    if re.search(r"\b(wages|salary|paye|uif|bonus|overtime|casual|relief|wca|coida|caretaker|staff)\b", d):
        return "personnel"
    if re.search(r"\b(income tax|taxation)\b", d):
        return "tax"
    if re.search(r"\b(special project|improvement|jungle gym|damp|aluminium)\b", d):
        return "special"
    if "legal" in d and "recover" in d:
        return "recoveries_other"
    if "recover" in d:
        return "recoveries_other"
    return "expenditure"


def num(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        n = float(v)
        return 0.0 if abs(n) < 0.01 else n
    s = str(v).strip()
    if not s or s in ("-", "–"):
        return 0.0
    s2 = re.sub(r"[R$\s,]", "", s).replace("(", "-").replace(")", "")
    try:
        n = float(s2)
    except ValueError:
        return None
    return 0.0 if abs(n) < 0.01 else n


def extract_wcu(uploaded) -> list:
    xl = pd.ExcelFile(uploaded)
    out = []
    for sheet in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet, header=None)
        if df.empty or df.shape[1] < 4:
            continue
        header = None
        for i in range(min(15, len(df))):
            vals = [str(v).strip().lower() for v in df.iloc[i].tolist()]
            if "actual" in vals and any("budget" in v for v in vals):
                header = i
                break
        if header is None:
            continue
        prev = [str(v).strip().lower() for v in df.iloc[header - 1].tolist()] if header else [""] * df.shape[1]
        cur = [str(v).strip().lower() for v in df.iloc[header].tolist()]
        merged = [(prev[i] if i < len(prev) else "") + " " + (cur[i] if i < len(cur) else "") for i in range(df.shape[1])]
        ytd = next((i for i, t in enumerate(merged) if "ytd" in t and "actual" in t and "var" not in t), None)
        if ytd is None:
            ytd = next((i for i, t in enumerate(cur) if t == "actual"), None)
        if ytd is None:
            continue
        for r in range(header + 1, len(df)):
            raw = str(df.iloc[r, 0] or "").strip()
            if not raw or re.match(r"^(total|surplus|shortfall)", raw, re.I):
                continue
            m = re.match(r"^(\d{3,5}\s*/\s*\d{2,4})\s*[-–—:]\s*(.+)$", raw)
            if m:
                gl = re.sub(r"\s", "", m.group(1))
                desc = m.group(2).strip()
                if gl.endswith("/000"):
                    continue
            else:
                desc = raw
                if len(desc) < 3:
                    continue
            actual = abs(num(df.iloc[r, ytd]) or 0.0)
            if actual < 0.5:
                continue
            out.append({"desc": desc, "actual": actual})
    return out


AFS_HEADING = re.compile(
    r"^(gross revenue|other income|expenditure|levy income|recoveries|municipal charges|"
    r"figures in r|detailed income|statement of|financial statements for|"
    r"operating costs?|the supplementary|page \d+|index$|note\(s\)|unaudited|"
    r"registration number|sectional scheme|figures in rand)$",
    re.I,
)
AFS_TOTAL = re.compile(
    r"\b(profit|surplus|deficit|total net|total income|total expenditure|for the year|"
    r"balance at|gross surplus)\b",
    re.I,
)
AFS_BS = re.compile(
    r"property, plant|cash and cash|trade and other|levies in arrears|levies in advance|"
    r"retained (income|earnings)|current tax liability|borrowings",
    re.I,
)
NUM_TAIL = re.compile(
    r"(\(?-?\d{1,3}(?:,\d{3})+(?:\.\d+)?\)?|\(?-?\d+\.\d+\)?|(?<![A-Za-z0-9])-|(?<![A-Za-z])\d{1,7})\s*$"
)


def _afs_num(tok: str):
    tok = tok.strip()
    if tok == "-":
        return 0.0
    neg = tok.startswith("(") and tok.endswith(")")
    s = tok.replace("(", "").replace(")", "").replace(",", "")
    try:
        n = float(s)
    except ValueError:
        return None
    return -n if neg else n


def _peel_afs(line: str):
    nums = []
    rest = line.rstrip()
    for _ in range(4):
        m = NUM_TAIL.search(rest)
        if not m:
            break
        n = _afs_num(m.group(1))
        if n is None:
            break
        nums.append(n)
        rest = rest[: m.start()].rstrip()
    nums.reverse()
    rest = re.sub(r"\s+\d{1,2}$", "", rest).strip()  # note number e.g. Insurance 9
    rest = re.sub(r"\b20\d{2}\s*$", "", rest).strip()
    return rest, nums


def _ocr_image(pil) -> str:
    try:
        import pytesseract
        return pytesseract.image_to_string(pil) or ""
    except Exception:
        return ""


def _pdf_page_texts(uploaded) -> list:
    uploaded.seek(0)
    with pdfplumber.open(uploaded) as pdf:
        pages = []
        for page in pdf.pages:
            t = page.extract_text(x_tolerance=2, y_tolerance=3) or page.extract_text() or ""
            pages.append(t)
        if sum(len(t) for t in pages) >= 400:
            return pages
        # Scanned PDF (Thornhill / Depotel Kruger packs): OCR the last pages — Detailed Income Statement is at the back.
        n = len(pdf.pages)
        start = max(0, n - 6)
        for i in range(start, n):
            try:
                im = pdf.pages[i].to_image(resolution=140)
                pil = im.original.convert("RGB")
                text = _ocr_image(pil)
                if len(text) > len(pages[i]):
                    pages[i] = text
            except Exception:
                continue
        # Name is on page 1
        if n and len(pages[0]) < 40:
            try:
                im = pdf.pages[0].to_image(resolution=110)
                pages[0] = _ocr_image(im.original.convert("RGB")) or pages[0]
            except Exception:
                pass
        return pages


def extract_afs_pdf(uploaded) -> tuple[list, str]:
    """Read the Detailed Income Statement. Line names stay as on the AFS."""
    pages = _pdf_page_texts(uploaded)
    name = ""
    for t in pages[:3]:
        for line in t.splitlines():
            line = re.sub(r"\s+", " ", line).strip()
            if re.search(r"homeowners|body corporate|association npc", line, re.I) and len(line) > 8:
                name = line
                break
        if name:
            break
    hits = [
        i
        for i, t in enumerate(pages)
        if re.search(r"detailed\s*income\s*statement", t, re.I) and len(t) > 200
    ]
    if not hits:
        hits = [
            i
            for i, t in enumerate(pages)
            if re.search(r"statement of comprehensive income", t, re.I) and len(t) > 200
        ]
    if not hits:
        return [], name
    run = [hits[-1]]
    for i in reversed(hits[:-1]):
        if i == run[0] - 1:
            run.insert(0, i)
        else:
            break
    chunks = [pages[run[0]]]
    for t in pages[run[0] + 1 :]:
        chunks.append(t)
        if re.search(r"supplementary information presented|notes to the financial statements", t, re.I):
            break
    rows = []
    for raw in "\n".join(chunks).splitlines():
        line = re.sub(r"\s+", " ", raw).strip()
        if not line:
            continue
        desc, nums = _peel_afs(line)
        if len(desc) < 3 or not nums:
            continue
        if AFS_HEADING.search(desc) or AFS_TOTAL.search(desc) or AFS_BS.search(desc):
            continue
        if re.match(r"^\d+$", desc):
            continue
        actual = abs(float(nums[0]))
        rows.append({"desc": desc, "actual": actual})
    return rows, name


def sections_from_afs(rows: list) -> dict:
    secs = {k: [] for k in default_sections()}
    saw_csos_inc = False
    for src in rows:
        desc = src["desc"]
        actual = float(src["actual"] or 0)
        fam = family(desc)
        if fam == "csos_col":
            continue
        if re.match(r"^(repairs and maintenance|salaries(& ?wages)?)$", desc, re.I):
            continue
        sec = section_for(desc)
        if fam == "csos_inc":
            if saw_csos_inc:
                sec = "expenditure"
            else:
                saw_csos_inc = True
                sec = "levy"
        non_cash = bool(re.search(r"depreciation|scrapping|fair value|impairment", desc, re.I))
        item = row(
            desc,
            "Non-cash on the financial statements. Left at R0 so it does not inflate levies."
            if non_cash
            else "Line name from this complex’s financial statements.",
        )
        item["actual"] = actual
        item["yearly"] = 0.0 if non_cash else actual
        item["pct"] = 0.0
        item["is_recovery"] = "recover" in norm(desc) and sec == "municipal"
        if fam == "ins_claim":
            item["note"] = "Put this on Recovery income and pick the repair, or type the payout on the repair line. Do not do both."
            item["yearly"] = 0.0
            sec = "recoveries_other"
        secs[sec].append(item)
    if not any(family(i["desc"]) == "ordinary" for i in secs["levy"]):
        ordinary = row(
            "Ordinary Levies",
            "Added so levies can be calculated. Name will follow the AFS if it had a levy line.",
        )
        ordinary["yearly"] = 0.0
        secs["levy"].insert(0, ordinary)
    if not any("reserve" in i["desc"].lower() for i in secs["levy"]):
        insert_at = 1 if secs["levy"] else 0
        secs["levy"].insert(
            insert_at,
            row("Reserve Fund Contribution", "Not always a separate AFS line. Type the yearly amount trustees want."),
        )
    if not secs["special"]:
        secs["special"] = [row("Special Project 1"), row("Special Project 2")]
    if not secs["tax"]:
        secs["tax"] = [row("Taxation Payable", "Based on taxable investment income.")]
    return secs


def _dedupe_headers(headers: list) -> list:
    seen, cols = {}, []
    for c in headers:
        c = str(c).strip() or "col"
        if c in seen:
            seen[c] += 1
            cols.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            cols.append(c)
    return cols


def _pq_from_frame(raw: pd.DataFrame) -> tuple[list, str]:
    """Find a header row with a unit column and a PQ / ratio column. Sum near 100 is treated as percent."""
    if raw is None or raw.empty:
        raise ValueError("Empty sheet.")
    header_idx = None
    for i in range(min(25, len(raw))):
        cells = [str(v).strip().lower() for v in raw.iloc[i].tolist()]
        has_unit = any(re.search(r"unit|erf|stand|plot|door|owner|customer|section|\bcode\b", c) for c in cells)
        has_pq = any(re.search(r"\bpq\b|quota|ratio|share", c) for c in cells)
        if has_unit and has_pq:
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("No header row with a unit and a PQ or ratio column.")
    cols = _dedupe_headers(raw.iloc[header_idx].tolist())
    body = raw.iloc[header_idx + 1 :].copy()
    body.columns = cols

    unit_candidates = [
        c for c in body.columns
        if re.search(r"customer code|owner code|account code|unit|erf|stand|plot|door|owner|code|section", str(c), re.I)
        and "size" not in str(c).lower()
    ]
    unit_col = None
    for c in unit_candidates:
        sample = body[c].astype(str).head(25)
        if "customer" in str(c).lower() and sample.str.contains(r"[A-Za-z]", regex=True).any():
            unit_col = c
            break
    if unit_col is None:
        for c in unit_candidates:
            sample = body[c].astype(str).head(25)
            if sample.str.contains(r"[A-Za-z]", regex=True).any():
                unit_col = c
                break
    if unit_col is None and unit_candidates:
        unit_col = unit_candidates[0]

    pq_candidates = [c for c in body.columns if re.search(r"\bpq\b|quota|ratio|^share$", str(c), re.I)]
    pq_col, best = None, -1.0
    for c in pq_candidates:
        nums = pd.to_numeric(body[c], errors="coerce").fillna(0.0)
        pos = nums[nums > 0]
        if len(pos) < 2:
            continue
        s = float(pos.sum())
        score = float(len(pos))
        if abs(s - 1) < 0.05:
            score += 200
        elif abs(s - 1) < 0.25:
            score += 80
        if abs(s - 100) < 5:
            score += 120
        elif abs(s - 100) < 25:
            score += 60
        if score > best:
            best, pq_col = score, c
        # WeConnectU leaves the PQ column at 0. The real share is Ratio 1.
        ratio1 = next((c for c in pq_candidates if re.fullmatch(r"ratio\s*1", str(c).strip(), re.I)), None)
        if ratio1 is not None:
            nums = pd.to_numeric(body[ratio1], errors="coerce").fillna(0.0)
            if float(nums[nums > 0].sum()) > 0:
                pq_col = ratio1
    if not unit_col or not pq_col:
        raise ValueError("Need a unit column and a PQ or ratio column. Found: " + ", ".join(cols))

    clean = pd.DataFrame({
        "Unit": body[unit_col].astype(str).str.strip(),
        "PQ": pd.to_numeric(body[pq_col], errors="coerce").fillna(0.0),
    })
    clean = clean[~clean["Unit"].str.lower().isin({"", "nan", "none", "total", "totals"})]
    clean = clean[~clean["Unit"].str.contains(r"^total\b", case=False, na=False)].reset_index(drop=True)
    if clean.empty:
        raise ValueError(f"No unit rows under {unit_col}.")
    total = float(clean["PQ"].sum())
    note = ""
    if 50 < total < 150:
        clean["PQ"] = clean["PQ"] / 100.0
        total = float(clean["PQ"].sum())
        note = " The file used percentages, so they were divided by 100."
    if abs(total - 1) > 0.02:
        note += f" Warning: ratios add to {total:.4f}, not 1.000. Levies will not add up until they do."
    msg = f"Loaded {len(clean)} units from {pq_col}. Ratio total {total:.6f}.{note}"
    return clean.to_dict("records"), msg


def parse_pq_upload(uploaded):
    """WeConnectU unit file, a two-column sheet, or this app’s workbook (the PQ sheet, not the first tab)."""
    name = str(getattr(uploaded, "name", "") or "").lower()
    if hasattr(uploaded, "getvalue"):
        raw_bytes = uploaded.getvalue()
    elif hasattr(uploaded, "read"):
        raw_bytes = uploaded.read()
    else:
        raw_bytes = uploaded
    bio = BytesIO(raw_bytes)
    frames = []
    if name.endswith(".csv") or (isinstance(raw_bytes, (bytes, bytearray)) and not name.endswith((".xlsx", ".xls", ".xlsm"))):
        frames.append(pd.read_csv(bio, header=None))
    else:
        xl = pd.ExcelFile(bio)
        names = list(xl.sheet_names)
        ordered = sorted(names, key=lambda n: (0 if re.search(r"pq|ratio|unit|levy", n, re.I) else 1))
        for n in ordered:
            if str(n).startswith("_"):
                continue
            frames.append(pd.read_excel(xl, sheet_name=n, header=None))
    errors = []
    for frame in frames:
        try:
            return _pq_from_frame(frame)
        except Exception as e:
            errors.append(str(e))
    raise ValueError(errors[-1] if errors else "Could not find unit ratios in this file.")


def match_into(extracted: list, sections: dict) -> tuple[dict, int]:
    nxt = {k: [dict(x) for x in v] for k, v in sections.items()}
    used = set()
    added = 0
    flat = [(k, it) for k, items in nxt.items() for it in items]
    leftover = []
    for src in extracted:
        fam = family(src["desc"])
        if fam == "csos_col":
            continue
        best, score = None, 0.0
        for key, it in flat:
            mark = f"{key}:{it['id']}"
            if mark in used:
                continue
            if fam and family(it["desc"]) and fam != family(it["desc"]):
                continue
            if ("reserve" in norm(src["desc"]) and "eskom" in norm(it["desc"])) or (
                "eskom" in norm(src["desc"]) and "reserve" in norm(it["desc"])
            ):
                continue
            nd, ni = norm(src["desc"]), norm(it["desc"])
            sc = 0.0
            if fam and family(it["desc"]) == fam:
                sc = 0.96
            if nd == ni:
                sc = 1.0
            elif ni in nd or nd in ni:
                sc = max(sc, 0.86)
            else:
                stop = {"levy", "levies", "income", "other", "general", "expense", "expenses", "fee", "fees", "and", "the"}
                a = {w for w in nd.split() if len(w) > 3 and w not in stop}
                b = {w for w in ni.split() if len(w) > 3 and w not in stop}
                if a and b:
                    sc = max(sc, len(a & b) / max(len(a), len(b)))
            if sc > 0.72 and sc > score:
                if src.get("actual") and it.get("actual"):
                    a1, a2 = abs(float(src["actual"])), abs(float(it.get("actual") or 0))
                    if a2 > 1 and max(a1, a2) / max(min(a1, a2), 1) > 15 and sc < 0.98:
                        continue
                score, best = sc, (key, it)
        if best:
            key, it = best
            used.add(f"{key}:{it['id']}")
            it["actual"] = src["actual"]
            if fam in ("hoa_levy_inc", "hoa_csos_inc", "hoa_levy_exp", "hoa_csos_exp", "ins_bill"):
                it["desc"] = src["desc"]
            if it.get("is_recovery"):
                it["yearly"] = src["actual"]
            else:
                it["yearly"] = src["actual"] * (1 + float(it.get("pct") or 0) / 100)
        else:
            leftover.append(src)
    for src in leftover:
        fam = family(src["desc"])
        if fam == "csos_col":
            continue
        sec = section_for(src["desc"])
        if fam:
            existing = next((i for i in nxt[sec] if family(i["desc"]) == fam), None)
            if existing:
                existing["actual"] = float(existing["actual"] or 0) + src["actual"]
                continue
        extra = row(src["desc"], "Added from WeConnectU for this complex")
        extra["actual"] = src["actual"]
        extra["yearly"] = src["actual"]
        extra["is_recovery"] = "recover" in norm(src["desc"]) and sec == "municipal"
        nxt[sec].append(extra)
        added += 1
    return nxt, added


def items_to_df(items: list, rm: bool, recover: bool = False, income: bool = False, claims: bool = False, claim_map: dict | None = None) -> pd.DataFrame:
    recs = []
    show_extra = rm or recover
    for it in items:
        actual = float(it.get("actual") or 0)
        yearly = float(it.get("yearly") or 0)
        stored = float(it.get("pct") or 0)
        pct = pct_from_amounts(actual, yearly) if actual >= 0.5 else stored
        ins = float(it.get("insurance") or 0)
        own = float(it.get("owner_recovery") or 0)
        taken = float((claim_map or {}).get(norm(it.get("desc") or "")) or 0)
        net = max(0.0, yearly - ins - own - taken) if (show_extra or taken) else abs(yearly)
        rec = {
            "Description": it["desc"],
            "Actual": actual,
            "% Increase": round(pct, 2),
            "Budgeted yearly": yearly,
            "Monthly": net / 12.0,
            "Notes": it.get("note") or "",
        }
        if rm:
            rec["Insurance payout"] = ins
        if recover:
            rec["Recovered from some owners"] = own
        if income:
            rec["Lower the levies"] = bool(it.get("reduces_levy"))
        if claims:
            rec["Comes off this cost"] = (it.get("claim_against") or "").strip() or "— none —"
        recs.append(rec)
    cols = ["Description", "Actual", "% Increase", "Budgeted yearly", "Monthly", "Notes"]
    if income:
        cols = ["Description", "Actual", "% Increase", "Budgeted yearly", "Monthly", "Lower the levies", "Notes"]
    if claims:
        cols = ["Description", "Actual", "% Increase", "Budgeted yearly", "Monthly", "Comes off this cost", "Notes"]
    if rm:
        cols = ["Description", "Actual", "% Increase", "Budgeted yearly", "Monthly", "Insurance payout", "Notes"]
    if recover:
        cols = ["Description", "Actual", "% Increase", "Budgeted yearly", "Recovered from some owners", "Monthly", "Notes"]
        if rm:
            cols = ["Description", "Actual", "% Increase", "Budgeted yearly", "Insurance payout", "Recovered from some owners", "Monthly", "Notes"]
    if not recs:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame(recs)[cols]


def cell_num(v) -> float:
    """A number from the grid. Blank, R, spaces and commas are fine."""
    if v is None or isinstance(v, bool):
        return 0.0
    try:
        if isinstance(v, float) and pd.isna(v):
            return 0.0
    except Exception:
        pass
    if isinstance(v, str):
        s = v.strip().replace("R", "").replace("r", "").replace(" ", "").replace(",", "").replace("%", "")
        if not s or s.lower() in ("nan", "none"):
            return 0.0
        try:
            return float(s)
        except ValueError:
            return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def save_editor(edited: pd.DataFrame, previous: list, rm: bool, municipal: bool = False) -> list:
    out = []
    records = edited.to_dict("records")
    for i, rec in enumerate(records):
        desc = str(rec.get("Description") or "").strip()
        if not desc or desc.lower() in ("none", "nan"):
            continue
        prev = {}
        if i < len(previous) and norm(previous[i].get("desc") or "") == norm(desc):
            prev = previous[i]
        else:
            hits = [p for p in previous if norm(p.get("desc") or "") == norm(desc)]
            if len(hits) == 1:
                prev = hits[0]
            elif i < len(previous):
                prev = previous[i]
        actual = abs(cell_num(rec.get("Actual")))
        pct = cell_num(rec.get("% Increase"))
        yearly = abs(cell_num(rec.get("Budgeted yearly")))
        ins = cell_num(rec.get("Insurance payout")) if rm else float(prev.get("insurance") or 0)
        if "Recovered from some owners" in rec:
            own = abs(cell_num(rec.get("Recovered from some owners")))
        else:
            own = float(prev.get("owner_recovery") or 0)
        if "Lower the levies" in rec:
            flag = rec.get("Lower the levies")
            reduces = False if flag is None or (isinstance(flag, float) and pd.isna(flag)) else bool(flag)
        else:
            reduces = bool(prev.get("reduces_levy"))
        if "Comes off this cost" in rec:
            against = clean_note(rec.get("Comes off this cost"))
            if against.lower() in ("", "— none —", "- none -", "none"):
                against = ""
        else:
            against = prev.get("claim_against") or ""
        old_actual = abs(float(prev.get("actual") or 0))
        old_y = abs(float(prev.get("yearly") or 0))
        shown_pct = round(pct_from_amounts(old_actual, old_y), 2) if old_actual >= 0.5 else round(float(prev.get("pct") or 0), 2)
        pct_changed = abs(round(pct, 2) - shown_pct) > 0.05
        y_changed = abs(yearly - old_y) > 0.5
        if pct_changed and actual >= 0.5:
            yearly = round(actual * (1 + pct / 100.0), 2)
            edit_mode = "pct"
        elif y_changed:
            pct = pct_from_amounts(actual, yearly) if actual >= 0.5 else pct
            edit_mode = "amount"
        else:
            edit_mode = prev.get("edit_mode") or "pct"
            if actual >= 0.5 and edit_mode != "amount":
                pct = pct_from_amounts(actual, yearly)
        if "Notes" not in rec or rec.get("Notes") is None or (
            isinstance(rec.get("Notes"), float) and pd.isna(rec.get("Notes"))
        ):
            note = prev.get("note") or ""
        else:
            note = clean_note(rec.get("Notes"))
        recovery = municipal and is_muni_recovery(desc, prev.get("is_recovery"))
        out.append({
            "id": prev.get("id") or uid(),
            "desc": desc,
            "actual": actual,
            "pct": pct,
            "yearly": yearly,
            "insurance": ins,
            "owner_recovery": own,
            "note": note,
            "is_recovery": recovery,
            "reduces_levy": reduces,
            "claim_against": against,
            "edit_mode": edit_mode,
        })
    return out


def ymp_years_from_nums(nums: list, start_year: int | None = None) -> list:
    """Old pack: First cycle | Frequency | Current estimate | 2025 | 2026 | … then roll so Year 1 = this budget year."""
    start_year = int(start_year or NOW_YEAR)
    n = [float(x) for x in nums if x is not None]
    if n and 2020 <= n[0] <= 2040:
        n = n[1:]
        if n and 1 <= n[0] <= 20:
            n = n[1:]
        if n and (len(n) >= 11 or (n[0] >= 50 and len(n) >= 2 and n[0] > max([x for x in n[1:] if x] or [0]) * 1.2)):
            n = n[1:]  # drop Current Estimate
        offset = max(0, start_year - 2025)
        if offset:
            n = n[offset:] if len(n) > offset else [0.0] * 10
    if len(n) > 10:
        n = n[:10]
    return (n + [0.0] * 10)[:10]


def parse_ymp_paste(text: str, start_year: int | None = None) -> list:
    rows = []
    for line in (text or "").splitlines():
        line = line.replace("\xa0", " ").rstrip()
        if not line.strip():
            continue
        if "\t" in line:
            rows.append([c.strip() for c in line.split("\t")])
        else:
            rows.append([c.strip() for c in re.split(r"\s{2,}", line)])
    if not rows:
        return []
    width = max(len(r) for r in rows)
    padded = [r + [""] * (width - len(r)) for r in rows]
    return parse_ymp_sheet(pd.DataFrame(padded), start_year)


def parse_ymp_sheet(df, start_year: int | None = None) -> list:
    """Read a 10 YMP sheet. Calendar headers (2025, 2026, …) are aligned so Year 1 = this budget year (2026)."""
    start_year = int(start_year or NOW_YEAR)
    header_row = None
    year_cols = []  # column indexes in Year1..Year10 order (already aligned)
    cal_pairs = []  # (col, calendar_year)
    for i in range(min(12, len(df))):
        raw = [df.iat[i, j] for j in range(df.shape[1])]
        labels = [str(v).strip().lower() for v in raw]
        ylabels = [j for j, v in enumerate(labels) if re.match(r"^(year|tear)\s*\d+", v) or re.match(r"^y\d+$", v)]
        cals = []
        for j, v in enumerate(raw):
            s = str(v).strip().replace(".0", "")
            if re.fullmatch(r"20[2-4]\d", s):
                cals.append((j, int(s)))
        if len(cals) >= 8:
            header_row = i
            cal_pairs = cals
            by_year = {yr: col for col, yr in cals}
            year_cols = [by_year.get(start_year + k) for k in range(10)]
            break
        if len(ylabels) >= 8:
            header_row, year_cols = i, ylabels[:10]
            break
    projects = []
    start = (header_row + 1) if header_row is not None else 0
    for r in range(start, len(df)):
        cells = [df.iat[r, j] if j < df.shape[1] else None for j in range(df.shape[1])]
        desc = ""
        for v in cells[:6]:
            s = str(v or "").strip()
            if not s or s.lower() in ("nan", "none"):
                continue
            if num(s) is not None and not re.search(r"[A-Za-z]", s):
                continue
            desc = s
            break
        if not desc or re.match(
            r"^(total|planned maintenance|project|projects|first|cycle|frequency|current|estimate|"
            r"year\s*\d+|tear\s*\d+|10 year|body corporate)$",
            desc,
            re.I,
        ):
            continue
        if re.search(r"maintenance plan|body corporate", desc, re.I) and not re.search(r"paint|roof|door|window|valuation|lift|pool|fence", desc, re.I):
            continue
        if year_cols:
            years = []
            for c in year_cols:
                if c is None or c >= len(cells):
                    years.append(0.0)
                else:
                    years.append(float(num(cells[c]) or 0))
            years = (years + [0.0] * 10)[:10]
            skip = {c for c in year_cols if c is not None}
            meta = []
            for j, v in enumerate(cells):
                if j in skip:
                    continue
                n = num(v)
                if n is not None:
                    meta.append(n)
            first = int(meta[0]) if meta and 2020 <= meta[0] <= 2040 else ""
            freq = int(meta[1]) if len(meta) > 1 and 1 <= meta[1] <= 15 else ""
            est = float(meta[2]) if len(meta) > 2 else (float(meta[0]) if meta and not first else 0.0)
        else:
            nums = [num(v) for v in cells]
            nums = [x for x in nums if x is not None]
            years = ymp_years_from_nums(nums, start_year)
            first, freq, est = "", "", 0.0
            if nums and 2020 <= nums[0] <= 2040:
                first = int(nums[0])
                if len(nums) > 1 and 1 <= nums[1] <= 20:
                    freq = int(nums[1])
                    if len(nums) > 2 and nums[2] >= 50:
                        est = float(nums[2])
            elif nums:
                est = float(nums[0]) if nums[0] >= 50 else 0.0
        if not desc:
            continue
        if not any(abs(y) > 0.5 for y in years) and not est and not first:
            continue
        projects.append({
            "desc": desc,
            "years": years,
            "first_cycle": first,
            "freq": freq,
            "estimate": est or (years[0] if years[0] else 0.0),
        })
    return projects


def plan_start_year(state: dict) -> int:
    m = re.search(r"20\d{2}", str(state.get("fin_year") or ""))
    y = int(m.group()) if m else NOW_YEAR
    return max(y, NOW_YEAR)


def _rgb(h: str) -> RGBColor:
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _r0(n: float) -> str:
    return f"R {abs(float(n or 0)):,.0f}"


def _typical_share(state: dict):
    """Most common positive PQ, and a plain name for that owner."""
    pos = [float(u.get("PQ") or 0) for u in (state.get("pq") or []) if float(u.get("PQ") or 0) > 1e-8]
    if not pos:
        return None, "the complex"
    from collections import Counter
    share = Counter(round(p, 6) for p in pos).most_common(1)[0][0]
    if max(pos) - share < 1e-5 and min(pos) > share * 0.9:
        return share, "each owner"
    return share, "a full erf"


def _story_costs(state: dict) -> list:
    """Plain groups that add up to the ordinary levy."""
    groups: dict[str, list] = {}
    order: list[str] = []

    def add(label, amount, line):
        amount = float(amount or 0)
        if amount < 1:
            return
        if label not in groups:
            groups[label] = [0.0, line]
            order.append(label)
        groups[label][0] += amount

    net = municipal_net(state)
    add("The shared city bill", net, "Electricity and water for the common property, after owners pay back what they use.")
    blurbs = {
        "Security and the gate": "Someone watching, and a gate that opens.",
        "Gardens": "The grounds you come home to.",
        "Cleaning": "The complex kept clean.",
        "Repairs": "What breaks, fixed before it gets worse.",
        "Day-to-day management": "Bills, owners, contractors and the work in between.",
        "Insurance": "Cover when something serious happens.",
        "Tax": "Tax on interest and rental. Your levy itself is not taxed.",
        "Other running costs": "Audit, bank charges, legal, meetings and the smaller bills.",
    }
    sources = list(state["sections"].get("expenditure") or []) + list(state["sections"].get("rm") or []) + list(state["sections"].get("personnel") or [])
    for r in sources:
        if skip_from_ordinary(r, state):
            continue
        d = (r.get("desc") or "").lower()
        y = cost_net(state, r)
        if "secur" in d or "gate" in d or "fence" in d or "camera" in d:
            label = "Security and the gate"
        elif "garden" in d:
            label = "Gardens"
        elif "clean" in d or "refuse" in d:
            label = "Cleaning"
        elif "manag" in d:
            label = "Day-to-day management"
        elif "insur" in d:
            label = "Insurance"
        elif any(k in d for k in ("plumb", "sewer", "electric", "maint", "repair", "fire")):
            label = "Repairs"
        else:
            label = "Other running costs"
        add(label, y, blurbs[label])
    for r in state["sections"].get("tax") or []:
        add("Tax", net_of(r), blurbs["Tax"])
    if state.get("special_in_ordinary"):
        add("Special projects", sum_net(state["sections"].get("special") or []), "Big jobs the owners asked to pay this year, not from the reserve.")
    rows = [(label, groups[label][0], groups[label][1]) for label in order]
    rows.sort(key=lambda x: -x[1])
    if len(rows) > 6:
        head, tail = rows[:5], rows[5:]
        extra = sum(a for _, a, _ in tail)
        if extra > 1:
            head.append(("Other running costs", extra, blurbs["Other running costs"]))
        rows = head
    return rows


def generate_pptx(state: dict) -> BytesIO:
    """PowerPoint of the same owner story as the HTML deck."""
    pack = meeting_pack(state)
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = f"{pack['name']} — for the owners"
    prs.core_properties.author = "Domus Property Management"
    blank = prs.slide_layouts[6]
    logo = Path(__file__).parent / "domus_logo.jpeg"
    INK, MINT, CREAM, WHITE, MUTED, SOFT = "111111", "70F8C8", "F6F6F4", "FFFFFF", "6B6B6B", "F3F3F1"

    def morph(sld, ms=1600):
        el = sld._element
        for child in list(el):
            if child.tag.endswith("transition"):
                el.remove(child)
        el.append(etree.fromstring(
            '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
            'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" spd="slow" p14:dur="%d">'
            '<p159:morph xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" option="byObject"/>'
            '</p:transition>' % ms
        ))

    def slide(dark=False, hero=False, full=False, dur=1600):
        s = prs.slides.add_slide(blank)
        fill = s.background.fill
        fill.solid()
        fill.fore_color.rgb = _rgb(INK if dark else CREAM)
        morph(s, dur)
        if logo.exists():
            if full:
                pic = s.shapes.add_picture(str(logo), Inches(-2.1), Inches(-0.08), width=Inches(17.55))
            elif hero:
                pic = s.shapes.add_picture(str(logo), Inches(0.5), Inches(2.05), width=Inches(4.7))
            else:
                pic = s.shapes.add_picture(str(logo), Inches(0.38), Inches(0.16), width=Inches(3.15))
            pic.name = "!!logo"
        if full:
            rule = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(3.7), Inches(10.9), Inches(0.08))
        elif hero:
            rule = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(4.25), Inches(4.7), Inches(0.07))
        else:
            rule = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.38), Inches(1.68), Inches(12.55), Inches(0.045))
        rule.fill.solid()
        rule.fill.fore_color.rgb = _rgb(MINT)
        rule.line.fill.background()
        rule.name = "!!rule"
        if pack.get("year") and not hero and not full:
            text(s, 6.4, 0.55, 6.4, 0.35, [(pack["year"], 13, False, WHITE if dark else INK)], align="right")
        if not hero and not full:
            text(s, 0.4, 7.08, 8, 0.28, [(f"Domus  ·  {pack['name']}", 11, False, MINT if dark else MUTED)])
        return s

    def rect(s, l, t, w, h, color, radius=0.08):
        sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
        sh.fill.solid()
        sh.fill.fore_color.rgb = _rgb(color)
        sh.line.fill.background()
        try:
            sh.adjustments[0] = radius
        except Exception:
            pass
        return sh

    def text(s, l, t, w, h, lines, align="left"):
        box = s.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
        tf = box.text_frame
        tf.word_wrap = True
        tf.auto_size = None
        tf.margin_left = tf.margin_right = tf.margin_top = Emu(0)
        align_e = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}[align]
        for i, item in enumerate(lines):
            msg, size, bold, color = item[0], item[1], item[2], item[3]
            space = item[4] if len(item) > 4 else 4
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align_e
            p.space_after = Pt(space)
            run = p.add_run()
            run.text = str(msg)
            run.font.name = "Georgia" if bold and size >= 20 else "Calibri"
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.color.rgb = _rgb(color)
        return box

    def kicker(s, msg, dark=False):
        text(s, 0.45, 1.82, 12, 0.32, [(msg.upper(), 12, True, MINT if dark else INK)])

    # Opening: the logo fills the screen, then zooms out onto the first page
    slide(True, full=True)
    s = slide(True, hero=True, dur=2600)
    text(s, 5.6, 2.15, 7.2, 1.15, [(pack["name"], 36, True, WHITE, 0)])
    text(s, 5.6, 3.35, 7.2, 0.9, [(pack.get("year") or "", 16, False, MINT, 6), ("Prepared by Domus, so every owner can see what the year will cost.", 16, False, "E4E4E4", 0)])
    tiles = [
        (str(int(pack["units"])), "owners sharing the cost"),
        (_r0(pack["owner_month"]), f"a month for {pack['who']}"),
        (_r0(pack["collect_year"]), "for the whole year"),
    ]
    for i, (num, cap) in enumerate(tiles):
        x = 0.45 + i * 4.2
        rect(s, x, 4.55, 3.95, 1.85, "1C1C1C")
        text(s, x + 0.25, 4.7, 3.5, 1.5, [(num, 28, True, WHITE, 4), (cap, 14, False, "CFCFCF", 0)])

    # 2 Three questions
    s = slide()
    kicker(s, "Start here")
    text(s, 0.45, 2.15, 12, 0.7, [("Three quiet questions", 32, True, INK, 2), ("The rest of this presentation answers them, one at a time.", 16, False, MUTED, 0)])
    qs = [
        ("Money in", "WHAT OWNERS PAY", "Each month, every owner pays a share. That payment is the levy."),
        ("Money out", "WHAT IT KEEPS GOING", "Security, gardens, insurance, repairs, and the people who look after the bills."),
        ("Money saved", "SET ASIDE", "Part of the payment is kept for the larger jobs, shown on their own slides."),
    ]
    for i, (title, tag, body) in enumerate(qs):
        x = 0.45 + i * 4.2
        rect(s, x, 3.35, 3.95, 3.15, WHITE)
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(3.35), Inches(3.95), Inches(0.08))
        bar.fill.solid(); bar.fill.fore_color.rgb = _rgb(MINT); bar.line.fill.background()
        text(s, x + 0.25, 3.6, 3.45, 2.7, [(title, 22, True, INK, 8), (tag, 12, True, INK, 8), (body, 15, False, INK, 0)])

    # 3 Four figures
    s = slide()
    kicker(s, "At a glance")
    text(s, 0.45, 2.15, 12, 0.55, [("The year in four figures", 32, True, INK)])
    pct = pack.get("pct")
    fourth = "—" if pct is None else f"{'+' if pct >= 0 else '−'}{abs(pct):.1f}%"
    fourth_cap = "compared with last year" if pct is None else ("higher than last year" if pct >= 0.5 else ("lower than last year" if pct <= -0.5 else "about the same as last year"))
    figs = [
        (_r0(pack["ordinary"]), "to run the complex", "Ordinary levy for the year"),
        (_r0(pack["ordinary"] / 12), "each month, together", "From all the owners"),
        (_r0(pack["owner_month"]), f"for {pack['who']}", "Everything on the monthly bill"),
        (fourth, fourth_cap, "Only where the cost itself changed"),
    ]
    for i, (num, title, cap) in enumerate(figs):
        x = 0.4 + i * 3.2
        rect(s, x, 3.05, 3.02, 3.2, WHITE)
        text(s, x + 0.18, 3.3, 2.66, 2.7, [(num, 26, True, INK, 10), (title, 15, True, INK, 6), (cap, 13, False, MUTED, 0)])

    # 4 Where it goes
    s = slide()
    kicker(s, "The split")
    text(s, 0.45, 2.15, 12, 0.7, [("Where every rand of the levy goes", 30, True, INK, 2), ("Each bar is one part of the ordinary levy.", 15, False, MUTED, 0)])
    costs = pack.get("costs") or []
    top = max((c["amount"] for c in costs), default=1) or 1
    for i, c in enumerate(costs[:6]):
        y = 3.05 + i * 0.62
        text(s, 0.45, y, 3.6, 0.5, [(c["label"], 14, True, INK)])
        track = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.2), Inches(y + 0.08), Inches(6.3), Inches(0.28))
        track.fill.solid(); track.fill.fore_color.rgb = _rgb("E6E6E6"); track.line.fill.background()
        bar = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.2), Inches(y + 0.08), Inches(max(0.15, 6.3 * c["amount"] / top)), Inches(0.28))
        bar.fill.solid(); bar.fill.fore_color.rgb = _rgb(INK if i % 2 == 0 else "3A3A3A"); bar.line.fill.background()
        if i == 1:
            bar.fill.fore_color.rgb = _rgb(MINT)
        text(s, 10.6, y, 2.2, 0.45, [(_r0(c["amount"]), 14, True, INK)], align="right")

    # 5 Monthly bill
    s = slide()
    kicker(s, "What you actually pay")
    text(s, 0.45, 2.15, 12, 0.6, [("Your monthly bill, line by line", 30, True, INK, 2), (f"For {pack['who']}.", 15, False, MUTED, 0)])
    head = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.45), Inches(3.0), Inches(12.4), Inches(0.42))
    head.fill.solid(); head.fill.fore_color.rgb = _rgb(INK); head.line.fill.background()
    text(s, 0.6, 3.05, 4, 0.32, [("On the statement", 13, True, WHITE)])
    text(s, 5.2, 3.05, 2.2, 0.32, [("A month", 13, True, WHITE)])
    text(s, 7.6, 3.05, 4.8, 0.32, [("What it is for", 13, True, WHITE)])
    for i, line in enumerate((pack.get("invoice") or [])[:5]):
        y = 3.42 + i * 0.52
        if i % 2 == 0:
            bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.45), Inches(y), Inches(12.4), Inches(0.52))
            bg.fill.solid(); bg.fill.fore_color.rgb = _rgb(SOFT); bg.line.fill.background()
        text(s, 0.6, y + 0.08, 4.4, 0.36, [(line["name"], 14, True, INK)])
        text(s, 5.2, y + 0.08, 2.2, 0.36, [(_r0(line["monthly"]), 14, True, INK)])
        text(s, 7.6, y + 0.08, 5, 0.36, [(line["plain"], 13, False, MUTED)])
    text(s, 0.5, 6.35, 6, 0.4, [("Total", 16, True, INK)])
    text(s, 7.5, 6.2, 5.2, 0.5, [(_r0(pack["owner_month"]), 26, True, INK)], align="right")

    # 6 Why
    if pack.get("risers"):
        s = slide()
        kicker(s, "The honest answer")
        title = "Why the levy is higher" if (pack.get("pct") or 0) >= 0.5 else "What changed from last year"
        text(s, 0.45, 2.15, 12, 0.85, [(title, 30, True, INK, 4), (pack.get("why_lead") or "", 15, False, MUTED, 0)])
        most = max(r["more"] for r in pack["risers"]) or 1
        for i, r in enumerate(pack["risers"][:5]):
            y = 3.2 + i * 0.7
            text(s, 0.45, y, 3.5, 0.45, [(r["label"][:28], 14, True, INK)])
            bar = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.1), Inches(y + 0.06), Inches(max(0.2, 6.2 * r["more"] / most)), Inches(0.3))
            bar.fill.solid(); bar.fill.fore_color.rgb = _rgb(INK if i else MINT); bar.line.fill.background()
            text(s, 10.5, y, 2.3, 0.4, [("+" + _r0(r["more"]), 14, True, INK)], align="right")

    # 7 Municipal
    if pack.get("muni"):
        m = pack["muni"]
        s = slide()
        kicker(s, "Municipal")
        text(s, 0.45, 2.15, 12, 0.85, [("The city account, on its own", 30, True, INK, 4), ("Home use is paid back by the owner who used it. Only the shared part remains in the levy.", 15, False, MUTED, 0)])
        tiles = [("We pay the city", m["gross"]), ("Owners pay back", m["rec"]), ("Left in the levy", max(m["net"], 0))]
        for i, (label, amount) in enumerate(tiles):
            x = 0.45 + i * 4.2
            rect(s, x, 3.3, 3.95, 2.15, INK if i == 2 else WHITE)
            text(s, x + 0.25, 3.5, 3.5, 1.8, [
                (label, 14, False, "CFCFCF" if i == 2 else MUTED, 8),
                (_r0(amount), 28, True, MINT if i == 2 else INK, 0),
            ])

    # 8 Repairs
    if pack.get("repairs"):
        s = slide()
        kicker(s, "When something breaks")
        text(s, 0.45, 2.15, 12, 0.7, [("The smaller repairs", 30, True, INK, 2), ("Everyday breakages. The larger planned work is on the next slides.", 15, False, MUTED, 0)])
        items = pack["repairs"][:4]
        w = 12.4 / len(items) - 0.18
        for i, r in enumerate(items):
            x = 0.45 + i * (w + 0.18)
            rect(s, x, 3.2, w, 3.0, WHITE)
            text(s, x + 0.2, 3.45, w - 0.4, 2.5, [(_r0(r["yearly"]), 26, True, INK, 10), (r["desc"], 16, True, INK, 8), (r["plain"], 13, False, MUTED, 0)])

    # 9 Reserve
    s = slide(True)
    kicker(s, "The reserve", True)
    text(s, 0.45, 2.15, 12, 0.7, [("The reserve fund", 32, True, WHITE, 2), ("What is already saved, what is added, and what this year’s work will use.", 16, False, "CFCFCF", 0)])
    boxes = [("Already saved", pack["opening"]), ("Added this year", pack["reserve"]), ("Used this year", pack["projects"])]
    for i, (label, amount) in enumerate(boxes):
        x = 0.45 + i * 4.2
        rect(s, x, 3.25, 3.95, 1.9, "1C1C1C")
        text(s, x + 0.22, 3.4, 3.5, 1.55, [(label, 14, False, "BDBDBD", 6), (_r0(amount), 26, True, WHITE, 0)])
    rect(s, 0.45, 5.4, 12.4, 1.15, "1C1C1C")
    text(s, 0.7, 5.65, 12, 0.7, [(f"Still in the fund at year-end:  {_r0(pack['projected'])}", 24, True, WHITE)])

    # 10 Jobs
    if pack.get("jobs"):
        s = slide()
        kicker(s, "This year’s big jobs")
        text(s, 0.45, 2.15, 12, 0.55, [("Work planned for this year", 30, True, INK)])
        items = pack["jobs"][:4]
        w = 12.4 / len(items) - 0.18
        for i, j in enumerate(items):
            x = 0.45 + i * (w + 0.18)
            rect(s, x, 3.0, w, 2.3, WHITE)
            text(s, x + 0.18, 3.2, w - 0.35, 1.9, [(_r0(j["yearly"]), 24, True, INK, 8), (j["desc"], 15, True, INK, 0)])
        rect(s, 0.45, 5.55, 12.4, 1.05, INK)
        text(s, 0.7, 5.75, 12, 0.65, [(f"Together: {_r0(pack['projects'])}    ·    Taken from the reserve fund.", 20, True, WHITE)])

    # 11 Ten year
    if pack.get("years"):
        s = slide()
        kicker(s, "Looking ahead")
        text(s, 0.45, 2.15, 12, 0.6, [("The ten-year plan", 30, True, INK, 2), ("The taller column is the busiest year.", 15, False, MUTED, 0)])
        years = pack["years"]
        mx = max(y["amount"] for y in years) or 1
        gap = 12.4 / len(years)
        for i, y in enumerate(years):
            h = 0.08 if y["amount"] < 1 else max(0.15, 2.6 * y["amount"] / mx)
            x = 0.55 + i * gap
            col = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(5.55 - h), Inches(gap * 0.62), Inches(h))
            hot = y["amount"] >= mx * 0.98 and y["amount"] > 1
            col.fill.solid(); col.fill.fore_color.rgb = _rgb(MINT if hot else INK); col.line.fill.background()
            text(s, x - 0.05, 5.6, gap * 0.8, 0.55, [(y["label"], 11, False, MUTED)], align="center")

    # 12 Approve
    s = slide(True)
    kicker(s, "Please approve", True)
    text(s, 0.45, 2.15, 12, 0.55, [("For your approval", 32, True, WHITE)])
    asks = [
        ("1", "Approve the budget", (
            f"Trustees approve {_r0(pack.get('approved', pack['ordinary']))}. "
            f"The costs need {_r0(pack['ordinary'])}. Gap {_r0(pack.get('levy_gap') or 0)}."
            if (pack.get("levy_gap") or 0) > 1
            else f"Ordinary levies of {_r0(pack['ordinary'])} for the year."
        )),
        ("2", f"Approve what {pack['who']} pays", f"{_r0(pack['owner_month'])} a month."),
        ("3", "Approve the reserve", f"{_r0(pack['reserve'])} added to the reserve this year."),
    ]
    if pack.get("projects", 0) > 1:
        asks.append(("4", "Note the planned work", f"{_r0(pack['projects'])} for the jobs on the plan."))
    for i, (n, title, body) in enumerate(asks):
        y = 2.9 + i * 0.95
        rect(s, 0.45, y, 12.4, 0.85, "1C1C1C")
        text(s, 0.7, y + 0.12, 1, 0.6, [(n, 22, True, MINT)])
        text(s, 1.6, y + 0.1, 10.8, 0.65, [(title, 18, True, WHITE, 0), (body, 13, False, "CFCFCF", 0)])

    # 13 Close — Morph grows the logo back out
    s = slide(True, hero=True)
    text(s, 5.6, 2.3, 7, 1.3, [("Thank you.", 44, True, WHITE, 8)])
    text(s, 5.6, 3.7, 7, 1.1, [("Questions are welcome. Domus will walk through any line with you.", 18, False, "E4E4E4")])

    buf = BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


def _child_plain(desc: str) -> str:
    d = (desc or "").lower()
    if "secur" in d or "gate" in d or "camera" in d:
        return "People who watch the complex, and a gate that works."
    if "garden" in d:
        return "The grass, the plants and the garden work."
    if "clean" in d or "refuse" in d:
        return "Keeping the grounds clean and taking rubbish away."
    if "manag" in d:
        return "The office that pays the bills and helps the owners."
    if "insur" in d:
        return "Cover if the buildings are damaged."
    if "audit" in d:
        return "A check that the books are right."
    if "account" in d:
        return "Preparing the books and the statements."
    if "bank" in d:
        return "The bank’s fee on the complex account."
    if "legal" in d:
        return "Help when the complex needs advice."
    if "plumb" in d or "sewer" in d:
        return "Pipes and drains on the shared property."
    if "electric" in d or "light" in d:
        return "Lights and power for the shared areas."
    if "tax" in d:
        return "Tax on interest and rental. The levy itself is not taxed."
    if "fire" in d:
        return "Checking the fire equipment."
    if "paint" in d or "palisade" in d:
        return "Paint and care so the buildings last."
    if "loan" in d:
        return "A loan repayment, taken from the reserve."
    return "A cost of looking after the complex."


def _invoice_plain(name: str, split: str) -> str:
    n = (name or "").lower()
    if "reserve" in n:
        return "Saved for the big jobs. Not spent on the monthly bills."
    if "csos" in n:
        return "A small legal amount. We collect it and pay it over. We keep none of it."
    if "insur" in n:
        return "Cover for the buildings, shown on its own line so you can see it."
    if split == "equal":
        return "The same rand for every owner. Not based on the size of the erf."
    if n.startswith("lev"):
        return "Day-to-day running of the complex."
    return "Your share of this charge."


def meeting_pack(state: dict) -> dict:
    costs = ordinary_total(state)
    approved = approved_ordinary(state)
    ordinary = costs
    reserve = reserve_contribution(state)
    csos = scheme_csos_yearly(state, approved) if state.get("auto_csos", True) else 0.0
    share, who = _typical_share(state)
    units = [u for u in (state.get("pq") or []) if float(u.get("PQ") or 0) > 1e-8]
    n_units = len(units)
    if not share:
        who = "the whole complex"
    gross, recovered = municipal_gross_and_rec(state)

    invoice = []
    for name, yearly, split in pq_bill_lines(state):
        if abs(yearly) < 0.5:
            continue
        if split == "equal":
            monthly = yearly / 12 / max(n_units, 1)
        elif share:
            monthly = yearly / 12 * share
        else:
            monthly = yearly / 12
        invoice.append({"name": name, "yearly": yearly, "monthly": monthly, "plain": _invoice_plain(name, split)})
    owner_month = sum(x["monthly"] for x in invoice) or ((ordinary + reserve + csos) / 12 * (share or 1))
    collect_year = sum(x["yearly"] for x in invoice) or (ordinary + reserve + csos)

    last_actual = 0.0
    for r in state["sections"].get("levy") or []:
        if family(r.get("desc") or "") == "ordinary":
            last_actual = float(r.get("actual") or 0)
            break
    pct = None
    if last_actual > 1 and ordinary > 1:
        pct = (ordinary / last_actual) * 100 - 100

    costs = [{"label": a, "amount": b, "blurb": c} for a, b, c in _story_costs(state)]

    risers = []
    pool = (
        list(state["sections"].get("expenditure") or [])
        + list(state["sections"].get("rm") or [])
        + list(state["sections"].get("personnel") or [])
        + list(state["sections"].get("tax") or [])
    )
    for r in pool:
        if skip_from_ordinary(r, state):
            continue
        a = float(r.get("actual") or 0)
        y = cost_net(state, r)
        more = y - a
        if more > 500 and a > 1:
            risers.append({
                "label": r.get("desc") or "Cost",
                "more": more,
                "plain": f"Last year this cost {_r0(a)}. This year we allowed {_r0(y)}.",
            })
    muni_last = municipal_split(state, "actual")
    muni_more = (gross - recovered) - (muni_last[0] - muni_last[1])
    if muni_more > 500:
        risers.append({
            "label": "Shared city bill",
            "more": muni_more,
            "plain": "The part of the city bill that owners do not pay back on their own meters.",
        })
    risers.sort(key=lambda x: -x["more"])

    big = []
    for r in pool:
        if skip_from_ordinary(r, state):
            continue
        y = cost_net(state, r)
        if y > 1:
            big.append({"label": r.get("desc") or "Cost", "yearly": y})
    big.sort(key=lambda x: -x["yearly"])
    big_names = {b["label"] for b in big[:6]}

    repairs = []
    for r in state["sections"].get("rm") or []:
        if skip_from_ordinary(r, state):
            continue
        if (r.get("desc") or "") in big_names:
            continue
        y = cost_net(state, r)
        if y > 1:
            repairs.append({"desc": r.get("desc"), "yearly": y, "plain": _child_plain(r.get("desc") or "")})
    repairs.sort(key=lambda x: -x["yearly"])

    jobs = []
    if not state.get("special_in_ordinary"):
        for r in state["sections"].get("special") or []:
            y = float(r.get("yearly") or 0)
            if y > 1:
                jobs.append({"desc": r.get("desc") or "Project", "yearly": y})

    years = []
    start = plan_start_year(state)
    for i in range(10):
        amt = 0.0
        for proj in state.get("ymp") or []:
            ys = proj.get("years") or []
            if i < len(ys):
                amt += float(ys[i] or 0)
        if amt > 1 or any(float(x or 0) > 1 for proj in (state.get("ymp") or []) for x in (proj.get("years") or [])):
            years.append({"label": str(start + i), "amount": amt})
    if years and not any(y["amount"] > 1 for y in years):
        years = []

    calm = "Every number started from last year’s real cost. We did not guess."
    if recovered > 1000:
        calm = "The electricity and water an owner uses is paid back, and taken off the levy. " + calm
    if insurance_on_pq(state):
        calm = "Insurance is on its own line, so you can see it. " + calm

    if pct is None:
        why_lead = "These lines cost more than last year."
    elif pct >= 0.5:
        why_lead = f"The ordinary levy is {pct:.0f}% higher than last year. Most of that comes from the lines below."
    elif pct <= -0.5:
        why_lead = f"The ordinary levy is {abs(pct):.0f}% lower than last year. These lines still moved."
    else:
        why_lead = "The levy is about the same as last year. These are the lines that moved."

    bill_note = "A smaller erf pays a smaller share of the levy, the reserve and CSOS."
    if not share:
        bill_note = "Load the PQ sheet and this slide will show one owner’s share."
    elif gross > 1:
        bill_note = "Your own home’s electricity and water are not inside this total. Those follow your meter."

    return {
        "name": state.get("complex_name") or "This complex",
        "year": state.get("fin_year") or "",
        "units": n_units,
        "who": who,
        "ordinary": ordinary,
        "approved": approved,
        "levy_gap": ordinary - approved,
        "reserve": reserve,
        "owner_month": owner_month,
        "collect_year": collect_year,
        "pct": pct,
        "costs": costs[:6],
        "invoice": invoice,
        "risers": risers[:5],
        "why_lead": why_lead,
        "big": big[:6],
        "muni": {"gross": gross, "rec": recovered, "net": gross - recovered, "gaps": municipal_gaps(state)} if gross > 1 or recovered > 1 else None,
        "repairs": repairs[:4],
        "opening": float(state.get("reserve_balance") or 0),
        "projects": reserve_project_spend(state),
        "projected": projected_reserve(state),
        "jobs": jobs[:4],
        "years": years,
        "calm": calm,
        "bill_note": bill_note,
    }


def generate_meeting_html(state: dict) -> bytes:
    logo = Path(__file__).parent / "domus_logo.jpeg"
    return render_meeting_html(meeting_pack(state), logo)


def generate_excel(state: dict) -> BytesIO:
    apply_levy_lines(state)
    s = state["sections"]
    wb = Workbook()
    ws = wb.active
    ws.title = "BUDGET"
    # Old pack: Description | GL Code | Actual | % | Budgeted Yearly | Monthly | Comments
    for i, w in enumerate([3, 42, 12, 14, 10, 16, 14, 36], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    def fill(c, color):
        c.fill = PatternFill("solid", fgColor=color)

    def inp(c, val, fmt=None):
        c.value = val
        fill(c, YELLOW)
        c.font = Font(name="Calibri", color=BLUE, size=10)
        c.border = THIN
        if fmt:
            c.number_format = fmt

    def fml(c, f, bg=None):
        c.value = f"={f}"
        c.font = Font(name="Calibri", size=10)
        c.border = THIN
        c.number_format = MONEY
        if bg:
            fill(c, bg)

    ws.merge_cells("B2:G2")
    ws["B2"] = state.get("complex_name") or "BODY CORPORATE / HOA BUDGET"
    ws["B2"].font = Font(name="Calibri", bold=True, size=16, color=NAVY)
    ws["B3"] = state.get("fin_year") or ""
    ws["B3"].font = Font(bold=True, size=12)
    ws["B5"] = "Current reserve fund (already in the bank)"
    inp(ws["D5"], float(state.get("reserve_balance") or 0), MONEY)
    ws["B6"] = "This year’s reserve contribution"
    ws["B7"] = "This year’s projects paid from the reserve"
    ws["B8"] = "Projected reserve at year-end"
    fml(ws["D8"], "D5+D6-D7")
    ws["B9"] = "Yellow cell = the one you type. If the % is yellow, Budgeted Yearly = Actual × (1 + %). If Budgeted Yearly is yellow, the % = Yearly ÷ Actual − 1. Monthly = Yearly ÷ 12."
    ws["B9"].font = Font(italic=True, size=9, color="666666")

    r = 11

    def bar(title):
        nonlocal r
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        for col in range(2, 9):
            fill(ws.cell(r, col), SECTION)
            ws.cell(r, col).border = THIN
        ws.cell(r, 2).value = title
        ws.cell(r, 2).font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
        r += 1

    def hdr():
        nonlocal r
        labs = ["Description", "GL Code", "Actual", "%", "Budgeted Yearly", "Monthly", "Comments / Notes"]
        for i, lab in enumerate(labs, 2):
            c = ws.cell(r, i, lab)
            fill(c, NAVY)
            c.font = Font(bold=True, color="FFFFFF", size=10)
            c.border = THIN
        r += 1

    levy_rows = {}
    named_rows = {}
    levy_comp_rows = []

    def write(items, recovery_as_income=False, in_levy=False, owner_box=False, levy_tick=False, claim_col=False):
        """Old pack formulas: F = D*(1+E), G = F/12. Recoveries shown as positive income when asked."""
        nonlocal r
        if not items:
            start = r
            r += 1
            return start, start
        start = r
        if items and levy_tick:
            hc = ws.cell(start - 1, 9, "Lowers the levy")
            fill(hc, NAVY)
            hc.font = Font(bold=True, color="FFFFFF", size=10)
            hc.border = THIN
            ws.column_dimensions["I"].width = 18
        if items and claim_col:
            hc = ws.cell(start - 1, 9, "Comes off this cost")
            fill(hc, NAVY)
            hc.font = Font(bold=True, color="FFFFFF", size=10)
            hc.border = THIN
            ws.column_dimensions["I"].width = 28
        for it in items:
            desc = it.get("desc") or ""
            inp(ws.cell(r, 2), desc)
            inp(ws.cell(r, 3), it.get("gl") or "")
            act = abs(float(it.get("actual") or 0))
            pct = float(it.get("pct") or 0)
            y = float(it.get("yearly") or 0)
            fam = family(desc)
            if fam == "ordinary":
                levy_rows["ordinary"] = r
            if fam == "csos_inc" or (is_own_scheme_csos(desc) and fam != "csos_exp" and "expense" not in desc.lower()):
                levy_rows["csos"] = r
            if fam == "reserve":
                levy_rows["reserve"] = r
            if fam == "ins_bill" or re.search(r"insurance recovered|levy\s*[-–]\s*insurance", desc, re.I):
                levy_rows["insurance"] = r
            named_rows[desc] = r
            if in_levy and not skip_from_ordinary(it, state):
                levy_comp_rows.append(r)
            ins = float(it.get("insurance") or 0)
            rec = bool(it.get("is_recovery")) and not recovery_as_income
            inp(ws.cell(r, 4), act, MONEY)
            ins = float(it.get("insurance") or 0)
            has_claim = False
            for claim in (state.get("sections") or {}).get("recoveries_other") or []:
                if norm(claim.get("claim_against") or "") == norm(desc) and abs(float(claim.get("yearly") or 0)) > 0.5 and named_rows.get(claim.get("desc") or ""):
                    has_claim = True
                    break
            use_amount = (
                it.get("edit_mode") == "amount"
                and fam not in ("ordinary", "reserve", "csos_inc", "csos_exp")
                and not rec
                and ins < 0.5
                and not owner_box
                and not has_claim
                and act >= 0.5
            )
            if fam == "ordinary":
                fml(ws.cell(r, 5), f"IF(D{r}=0,0,F{r}/D{r}-1)")
                ws.cell(r, 5).number_format = "0.00%"
            elif not use_amount:
                show_pct = pct_from_amounts(act, y) if act >= 0.5 else pct
                inp(ws.cell(r, 5), show_pct / 100.0, "0.00%")
            if fam == "ordinary":
                pass  # F filled after totals
            elif fam == "reserve" and state.get("reserve_mode") in ("pct15", "15pct", "legal") and levy_rows.get("ordinary"):
                months = max(1, int(state.get("actual_months") or 12))
                scale = f"*12/{months}" if months < 12 else ""
                fml(ws.cell(r, 6), f"0.15*D{levy_rows['ordinary']}{scale}")
            elif fam == "reserve" and state.get("reserve_mode") == "pct25" and levy_rows.get("ordinary"):
                months = max(1, int(state.get("actual_months") or 12))
                scale = f"*12/{months}" if months < 12 else ""
                fml(ws.cell(r, 6), f"0.25*D{levy_rows['ordinary']}{scale}")
            elif fam == "reserve" and state.get("reserve_mode") == "rm100":
                inp(ws.cell(r, 6), float(reserve_contribution(state) or y or 0), MONEY)
            elif fam == "reserve":
                inp(ws.cell(r, 6), float(reserve_contribution(state) or y or 0), MONEY)
            elif is_own_scheme_csos(desc) and fam == "csos_exp" and levy_rows.get("csos"):
                fml(ws.cell(r, 6), f"F{levy_rows['csos']}")
            elif use_amount:
                inp(ws.cell(r, 6), abs(y), MONEY)
                fml(ws.cell(r, 5), f"IF(D{r}=0,0,F{r}/D{r}-1)")
                ws.cell(r, 5).number_format = "0.00%"
            elif act < 0.5 and abs(y) > 0.5 and not rec:
                inp(ws.cell(r, 6), abs(y), MONEY)
            elif rec and not recovery_as_income:
                fml(ws.cell(r, 6), f"-ABS(D{r}*(1+E{r}))")
            elif ins > 0.5:
                fml(ws.cell(r, 6), f"MAX(0,D{r}*(1+E{r})-{ins})")
            else:
                fml(ws.cell(r, 6), f"D{r}*(1+E{r})")
            fml(ws.cell(r, 7), f"F{r}/12")
            note = clean_note(it.get("note"))
            if ins:
                extra = "Insurance payout " + f"{ins:,.2f}"
                note = f"{note} | {extra}".strip(" |") if note else extra
            own = float(it.get("owner_recovery") or 0)
            if owner_box:
                if items and r == start:
                    hc = ws.cell(start - 1, 9, "Recovered from owners")
                    fill(hc, NAVY)
                    hc.font = Font(bold=True, color="FFFFFF", size=10)
                    hc.border = THIN
                    ws.column_dimensions["I"].width = 22
                inp(ws.cell(r, 9), own, MONEY)
                cur = ws.cell(r, 6).value
                if isinstance(cur, str) and cur.startswith("="):
                    ws.cell(r, 6).value = f"=MAX(0,({cur[1:]})-N(I{r}))"
                elif cur is not None:
                    ws.cell(r, 6).value = f"=MAX(0,{float(cur)}-N(I{r}))"
                    ws.cell(r, 6).number_format = MONEY
                if own:
                    extra = "Recovered from some owners " + f"{own:,.2f}" + ". Not income. Taken off this line."
                    note = f"{note} | {extra}".strip(" |") if note else extra
            claim_bits = []
            for claim in (state.get("sections") or {}).get("recoveries_other") or []:
                if norm(claim.get("claim_against") or "") != norm(desc):
                    continue
                if abs(float(claim.get("yearly") or 0)) < 0.5:
                    continue
                crow = named_rows.get(claim.get("desc") or "")
                if crow:
                    claim_bits.append(f"N(F{crow})")
            if claim_bits:
                cur = ws.cell(r, 6).value
                minus = "-".join(claim_bits)
                if isinstance(cur, str) and str(cur).startswith("="):
                    ws.cell(r, 6).value = f"=MAX(0,({cur[1:]})-{minus})"
                elif cur is not None:
                    ws.cell(r, 6).value = f"=MAX(0,{float(cur)}-{minus})"
                    ws.cell(r, 6).number_format = MONEY
            if levy_tick:
                inp(ws.cell(r, 9), "Yes" if it.get("reduces_levy") else "No")
            if claim_col:
                inp(ws.cell(r, 9), it.get("claim_against") or "")
            cnote = ws.cell(r, 8, note)
            cnote.font = Font(name="Calibri", size=9, italic=True, color="1F4E79")
            cnote.alignment = Alignment(wrap_text=True, vertical="top")
            cnote.border = THIN
            r += 1
        return start, r - 1

    def tot(label, start, end):
        nonlocal r
        ws.cell(r, 2, label).font = Font(bold=True)
        fml(ws.cell(r, 4), f"SUM(D{start}:D{end})", TOTAL)
        fml(ws.cell(r, 6), f"SUM(F{start}:F{end})", TOTAL)
        fml(ws.cell(r, 7), f"F{r}/12", TOTAL)
        row_n = r
        r += 2
        return row_n

    bar("INCOME")
    hdr()
    a, b = write(s["levy"])
    approved_row = None
    if levy_rows.get("ordinary"):
        ord_row = levy_rows["ordinary"]
        bar("MEETING DECISION — Levies received")
        ws.cell(r, 2, "What the costs need").font = Font(bold=True)
        fml(ws.cell(r, 5), f"E{ord_row}")
        ws.cell(r, 5).number_format = "0.00%"
        fml(ws.cell(r, 6), f"F{ord_row}")
        fml(ws.cell(r, 7), f"G{ord_row}")
        r += 1
        ws.cell(r, 2, "Approved — type a % or a rand")
        ws.cell(r, 8, "Type the % the meeting agrees, for example 10%. Or type a rand over the yearly amount and leave the %.")
        ws.cell(r, 8).font = Font(name="Calibri", size=9, italic=True, color="1F4E79")
        ws.cell(r, 8).alignment = Alignment(wrap_text=True, vertical="center")
        pct_cell = ws.cell(r, 5)
        mode = state.get("levy_approve_mode") or "costs"
        if mode == "pct":
            inp(pct_cell, float(state.get("levy_approved_pct") or 0) / 100.0, "0.00%")
        else:
            pct_cell.value = f"=E{ord_row}"
            pct_cell.number_format = "0.00%"
            fill(pct_cell, YELLOW)
            pct_cell.font = Font(name="Calibri", color=BLUE, size=10)
            pct_cell.border = THIN
        year_cell = ws.cell(r, 6)
        if mode == "amount" and float(state.get("levy_approved_amount") or 0) > 0.5:
            inp(year_cell, float(state.get("levy_approved_amount") or 0), MONEY)
        else:
            year_cell.value = f"=D{ord_row}*(1+E{r})"
            year_cell.number_format = MONEY
            fill(year_cell, YELLOW)
            year_cell.font = Font(name="Calibri", color=BLUE, size=10)
            year_cell.border = THIN
        fml(ws.cell(r, 7), f"F{r}/12", "C6EFCE")
        approved_row = r
        r += 1
        ws.cell(r, 2, "Gap — costs minus approved. Above zero means the costs are still higher.")
        ws.cell(r, 2).font = Font(italic=True, size=9, color="9C0006")
        fml(ws.cell(r, 6), f"F{ord_row}-F{approved_row}", RED)
        fml(ws.cell(r, 7), f"F{r}/12", RED)
        r += 2
    inc_tot = tot("TOTAL INCOME", a, b)
    bar("OTHER INCOME")
    hdr()
    a, b = write(list(s.get("other") or []), levy_tick=True)
    tot("TOTAL OTHER INCOME", a, b)
    bar("RECOVERY INCOME")
    hdr()
    a, b = write(list(s.get("recoveries_other") or []), claim_col=True)
    tot("TOTAL RECOVERY INCOME", a, b)
    hoa_tot = None
    if state.get("has_master_hoa"):
        bar("Recoveries on HOA Costs")
        hdr()
        a, b = write(s.get("hoa_income") or [])
        hoa_tot = tot("TOTAL OTHER RECOVERIES", a, b)
        for it in s.get("hoa_income") or []:
            named_rows[it.get("desc") or ""] = named_rows.get(it.get("desc") or "")
    muni_gross = [x for x in (s.get("municipal") or []) if not is_muni_recovery(x.get("desc") or "", x.get("is_recovery"))]
    muni_rec = [x for x in (s.get("municipal") or []) if is_muni_recovery(x.get("desc") or "", x.get("is_recovery"))]
    bar("Recoveries on Utilities")
    hdr()
    a, b = write(muni_rec, recovery_as_income=True)
    util_tot = tot("TOTAL UTILITY RECOVERIES", a, b)
    bar("Municipal Charges")
    hdr()
    a, b = write(muni_gross)
    muni_g_tot = tot("TOTAL CITY BILL", a, b)
    ws.cell(r, 2, "TOTAL NET MUNICIPAL (city bill minus recovered)").font = Font(bold=True)
    fml(ws.cell(r, 4), f"D{muni_g_tot}-D{util_tot}", RED)
    fml(ws.cell(r, 6), f"F{muni_g_tot}-F{util_tot}", RED)
    fml(ws.cell(r, 7), f"F{r}/12", RED)
    net_muni = r
    r += 1
    ws.cell(r, 2, "UNDER-RECOVERY memo (already inside the net above — do not add it again)")
    ws.cell(r, 2).font = Font(italic=True, size=9, color="9C0006")
    fml(ws.cell(r, 6), f"MAX(0,F{net_muni})", RED)
    fml(ws.cell(r, 7), f"F{r}/12", RED)
    r += 1
    gaps = municipal_gaps(state)
    under_rows = [g for g in gaps if g["gap"] > 1]
    if under_rows:
        ws.cell(r, 2, "Under-recovery by service (city bill minus recovered)").font = Font(bold=True, size=10, color="9C0006")
        r += 1
        for g in under_rows:
            grefs = [f"F{named_rows[d]}" for d in g["gross_descs"] if d in named_rows]
            rrefs = [f"F{named_rows[d]}" for d in g["rec_descs"] if d in named_rows]
            ws.cell(r, 2, f"{g['name']} under-recovered")
            if grefs or rrefs:
                gf = "+".join(grefs) if grefs else "0"
                rf = "+".join(rrefs) if rrefs else "0"
                fml(ws.cell(r, 6), f"MAX(0,({gf})-({rf}))", RED)
            else:
                inp(ws.cell(r, 6), max(0.0, g["gap"]), MONEY)
            fml(ws.cell(r, 7), f"F{r}/12")
            r += 1
    r += 1
    bar("EXPENDITURE")
    hdr()
    exp_items = list(s.get("expenditure") or [])
    if state.get("has_master_hoa"):
        exp_items = exp_items + list(s.get("hoa_expense") or [])
    a, b = write(exp_items, in_levy=True, owner_box=True)
    exp_tot = tot("TOTAL EXPENDITURE", a, b)
    bar("REPAIR AND MAINTENANCE")
    hdr()
    a, b = write(s.get("rm") or [], in_levy=True, owner_box=True)
    rm_tot = tot("Total Repair and Maintenance", a, b)
    if state.get("reserve_mode") == "rm100" and levy_rows.get("reserve"):
        fml(ws.cell(levy_rows["reserve"], 6), f"F{rm_tot}")
        fml(ws.cell(levy_rows["reserve"], 7), f"F{levy_rows['reserve']}/12")
    bar("PERSONNEL")
    hdr()
    a, b = write(s.get("personnel") or [], in_levy=True)
    per_tot = tot("Total Personnel Expenses", a, b)
    bar("INCOME TAX")
    hdr()
    a, b = write(s.get("tax") or [], in_levy=True)
    tax_tot = tot("TOTAL TAX", a, b)
    bar("SPECIAL PROJECTS")
    hdr()
    a, b = write(s.get("special") or [], in_levy=bool(state.get("special_in_ordinary")))
    sp_tot = tot("Total Special Projects Expenses", a, b)
    charged = [x for x in (s.get("fixed") or []) if float(x.get("yearly") or 0) > 0.5]
    if charged:
        bar("EQUAL CHARGES BILLED TO OWNERS (same rand each unit)")
        hdr()
        a, b = write(charged)
        tot("TOTAL EQUAL CHARGES", a, b)

    # Ordinary = net municipal + each cost line that is not billed separately (estate / insurance / garden / CSOS)
    bits = f"F{net_muni}"
    if levy_comp_rows:
        bits += "+" + "+".join(f"F{n}" for n in levy_comp_rows)
    cut = []
    for it in s.get("other") or []:
        if it.get("reduces_levy") and abs(float(it.get("yearly") or 0)) > 0.5:
            n = named_rows.get(it.get("desc") or "")
            if n:
                cut.append(f"F{n}")
    if cut:
        bits += "-" + "-".join(cut)
    bar("ORDINARY LEVY (what we charge)")
    ws.cell(r, 2, "Ordinary levies = net municipal + expenditure + R&M + personnel + tax, minus only the other income marked Yes")
    fml(ws.cell(r, 6), f"MAX(0,{bits})", RED)
    fml(ws.cell(r, 7), f"F{r}/12", RED)
    ord_check = r
    if levy_rows.get("ordinary"):
        ws.cell(levy_rows["ordinary"], 6).value = f"=F{ord_check}"
        ws.cell(levy_rows["ordinary"], 6).font = Font(name="Calibri", size=10)
        ws.cell(levy_rows["ordinary"], 6).number_format = MONEY
        fill(ws.cell(levy_rows["ordinary"], 6), RED)
    r += 2
    if levy_rows.get("reserve"):
        fml(ws["D6"], f"F{levy_rows['reserve']}")
    else:
        ws["D6"] = 0
    if state.get("special_in_ordinary"):
        ws["D7"] = 0
        ws["B7"] = "This year’s projects paid from the reserve (none — they are in the levy)"
    else:
        fml(ws["D7"], f"F{sp_tot}")

    pq = wb.create_sheet("PQ")
    pq["A1"] = "PQ / LEVY SCHEDULE"
    pq["A1"].font = Font(bold=True, size=14, color=NAVY)
    pq["A2"] = state.get("complex_name") or ""
    bills = pq_bill_lines(state)
    pq["B3"] = "Monthly"
    name_to_budget = {
        "Levies": f"BUDGET!G{approved_row}" if approved_row else (f"BUDGET!G{levy_rows['ordinary']}" if levy_rows.get("ordinary") else None),
        "CSOS": f"BUDGET!G{levy_rows['csos']}" if levy_rows.get("csos") else None,
        "Reserve Fund": f"BUDGET!G{levy_rows['reserve']}" if levy_rows.get("reserve") else None,
        "Insurance": f"BUDGET!G{levy_rows['insurance']}" if levy_rows.get("insurance") else None,
    }
    for it in (s.get("hoa_income") or []) + (s.get("fixed") or []):
        nm = it.get("desc") or ""
        if nm in named_rows:
            name_to_budget[nm] = f"BUDGET!G{named_rows[nm]}"
    headers = ["#", "Unit", "PQ"] + [n for n, _, _s in bills] + ["Total"]
    for i, h in enumerate(headers, 1):
        cell = pq.cell(6, i, h)
        fill(cell, NAVY)
        cell.font = Font(bold=True, color="FFFFFF")
    for j, (name, yearly, split) in enumerate(bills):
        col = 4 + j
        pq.cell(3, col, name)
        fill(pq.cell(3, col), NAVY)
        pq.cell(3, col).font = Font(bold=True, color="FFFFFF")
        src = name_to_budget.get(name)
        if src:
            pq.cell(4, col, f"={src}")
        else:
            pq.cell(4, col, float(yearly or 0) / 12)
        pq.cell(4, col).number_format = MONEY
        fill(pq.cell(4, col), YELLOW)
        pq.cell(5, col, "equal / unit" if split == "equal" else "× PQ")
        pq.cell(5, col).font = Font(italic=True, size=8, color="666666")
    units = state.get("pq") or [{"Unit": "UNIT-1", "PQ": 1.0}]
    n_units = max(len(units), 1)
    first_amt, last_amt = 4, 3 + len(bills)
    for i, u in enumerate(units):
        rr = 7 + i
        pq.cell(rr, 1, i + 1)
        pq.cell(rr, 2, str(u.get("Unit", "")))
        pq.cell(rr, 3, float(u.get("PQ") or 0)).number_format = "0.000000"
        for j, (_n, _y, split) in enumerate(bills):
            col = 4 + j
            letter = get_column_letter(col)
            if split == "equal":
                pq.cell(rr, col, f"={letter}$4/{n_units}").number_format = MONEY
            else:
                pq.cell(rr, col, f"=$C{rr}*{letter}$4").number_format = MONEY
        if bills:
            pq.cell(rr, last_amt + 1, f"=SUM({get_column_letter(first_amt)}{rr}:{get_column_letter(last_amt)}{rr})").number_format = MONEY
    last_u = 6 + len(units)
    tot_row = last_u + 1
    pq.cell(tot_row, 2, "TOTAL")
    pq.cell(tot_row, 3, f"=SUM(C7:C{last_u})")
    pq.cell(tot_row, 3).number_format = "0.000000"
    if bills:
        for col in range(first_amt, last_amt + 2):
            letter = get_column_letter(col)
            pq.cell(tot_row, col, f"=SUM({letter}7:{letter}{last_u})").number_format = MONEY
    pq.column_dimensions["A"].width = 6
    pq.column_dimensions["B"].width = 22
    pq.column_dimensions["C"].width = 14

    ymp = wb.create_sheet("10 YMP")
    start_y = plan_start_year(state)
    ymp["A1"] = (state.get("complex_name") or "").upper()
    ymp["A1"].font = Font(bold=True, size=14, color=NAVY)
    ymp["A2"] = "10 YEAR MAINTENANCE PLAN"
    ymp["A2"].font = Font(bold=True, size=12, color=NAVY)
    headers = ["PROJECTS", "First Cycle", "Frequency of Cycles", "Current Estimate"] + [str(start_y + i) for i in range(10)]
    sub = ["", "", "", ""] + [f"Year {i+1}" for i in range(10)]
    for i, h in enumerate(headers, 1):
        c = ymp.cell(4, i, h)
        fill(c, NAVY)
        c.font = Font(bold=True, color="FFFFFF", size=9)
        c.alignment = Alignment(wrap_text=True, horizontal="center")
    for i, h in enumerate(sub, 1):
        c = ymp.cell(5, i, h)
        fill(c, NAVY)
        c.font = Font(bold=True, color="FFFFFF", size=8)
        c.alignment = Alignment(horizontal="center")
    ymp.column_dimensions["A"].width = 38
    ymp.column_dimensions["B"].width = 12
    ymp.column_dimensions["C"].width = 14
    ymp.column_dimensions["D"].width = 16
    projects = [p for p in (state.get("ymp") or []) if (p.get("desc") or "").strip()]
    if not projects:
        projects = [{"desc": "", "years": [0] * 10}]
    first_data = 6
    for i, p in enumerate(projects):
        rr = first_data + i
        years = p.get("years") or [0] * 10
        inp(ymp.cell(rr, 1), p.get("desc") or "")
        fc = p.get("first_cycle") or ""
        ymp.cell(rr, 2, fc if fc != "" else None)
        fr = p.get("freq") or ""
        ymp.cell(rr, 3, fr if fr != "" else None)
        est = float(p.get("estimate") or 0) or float(years[0] or 0)
        inp(ymp.cell(rr, 4), est, MONEY)
        for y in range(10):
            val = float(years[y] if y < len(years) else 0)
            inp(ymp.cell(rr, 5 + y), val, MONEY)
            if val and y == 0:
                ymp.cell(rr, 5 + y).font = Font(bold=True, color="9C0006")
    last = first_data + len(projects) - 1
    tot = last + 1
    ymp.cell(tot, 1, "TOTAL").font = Font(bold=True)
    for col in range(5, 15):
        letter = get_column_letter(col)
        fml(ymp.cell(tot, col), f"SUM({letter}{first_data}:{letter}{last})", TOTAL)

    # Hidden round-trip sheet so Restore uploads the same numbers we downloaded.
    pack = wb.create_sheet("_DOMUS")
    pack.sheet_state = "hidden"
    pack["A1"] = "DOMUS_STATE_V1"
    payload = {
        "complex_name": state.get("complex_name") or "",
        "fin_year": state.get("fin_year") or "",
        "sections": state.get("sections") or {},
        "pq": state.get("pq"),
        "ymp": state.get("ymp") or [],
        "has_master_hoa": bool(state.get("has_master_hoa")),
        "insurance_mode": state.get("insurance_mode") or "levy",
        "insurance_bill_yearly": float(state.get("insurance_bill_yearly") or 0),
        "auto_csos": bool(state.get("auto_csos", True)),
        "reserve_mode": state.get("reserve_mode") or "amount",
        "reserve_amount": float(state.get("reserve_amount") or 0),
        "reserve_balance": float(state.get("reserve_balance") or 0),
        "scheme_type": state.get("scheme_type") or "bc",
        "special_in_ordinary": bool(state.get("special_in_ordinary")),
        "current_monthly_levy": float(state.get("current_monthly_levy") or 0),
        "levy_approved_pct": float(state.get("levy_approved_pct") or 0),
        "levy_approved_amount": float(state.get("levy_approved_amount") or 0),
        "levy_approve_mode": state.get("levy_approve_mode") or "costs",
    }
    blob = json.dumps(payload, ensure_ascii=False)
    # Excel cell cap is 32767; split across rows if needed.
    chunk = 30000
    pack["A2"] = "CHUNKS"
    pack["B2"] = (len(blob) + chunk - 1) // chunk
    for i in range(0, len(blob), chunk):
        pack.cell(3 + i // chunk, 1, blob[i : i + chunk])

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


RESTORE_BARS = [
    ("recovery income", "recoveries_other"),
    ("claims", "recoveries_other"),
    ("other income", "other"),
    ("recoveries on hoa", "hoa_income"),
    ("hoa / estate recovered", "hoa_income"),
    ("hoa / estate paid", "hoa_expense"),
    ("equal charges", "fixed"),
    ("recoveries on utilities", "recoveries_other"),
    ("other recoveries", "recoveries_other"),
    ("municipal", "municipal"),
    ("expenditure", "expenditure"),
    ("repair", "rm"),
    ("personnel", "personnel"),
    ("income tax", "tax"),
    ("special", "special"),
    ("fixed monthly", "fixed"),
    ("levy income", "levy"),
    ("income", "levy"),
]


def _apply_budget_sheet(wb, data: dict) -> None:
    """Trustees edit the sheet in the meeting. Those cells win when the file comes back."""
    if "BUDGET" not in getattr(wb, "sheetnames", []):
        return
    ws = wb["BUDGET"]
    sections = data.setdefault("sections", {})
    current = None
    for r in range(1, int(ws.max_row or 1) + 1):
        desc = str(ws.cell(r, 2).value or "").strip()
        if not desc:
            continue
        low = desc.lower()
        if low in ("description", "gl code"):
            continue
        if low.startswith((
            "meeting decision", "what the costs", "approved —", "approved %", "approved budget",
            "approved monthly", "or type the rand", "gap", "total", "under-recovery",
            "under recovery", "ordinary levies =",
        )):
            continue
        d_val = ws.cell(r, 4).value
        f_val = ws.cell(r, 6).value
        mapped = next((k for title, k in RESTORE_BARS if title in low and "total" not in low and "check" not in low), None)
        if mapped and d_val is None and f_val is None:
            current = mapped
            continue
        if current is None:
            continue
        items = sections.setdefault(current, [])
        found = next((it for it in items if norm(it.get("desc") or "") == norm(desc)), None)
        if found is None:
            found = row(desc, "")
            items.append(found)
        if isinstance(d_val, (int, float)):
            found["actual"] = abs(float(d_val))
        pct_v = ws.cell(r, 5).value
        year_v = f_val
        fam = family(desc)
        note = ws.cell(r, 8).value
        if isinstance(note, str) and note.strip():
            found["note"] = note.strip()
        got = ws.cell(r, 9).value
        if current == "other":
            if isinstance(got, str):
                found["reduces_levy"] = got.strip().lower() in ("yes", "y", "true", "1")
            elif isinstance(got, (int, float)) and not isinstance(got, bool):
                found["reduces_levy"] = float(got) >= 0.5
        elif current == "recoveries_other" and isinstance(got, str) and got.strip():
            found["claim_against"] = got.strip()
        elif isinstance(got, (int, float)) and not isinstance(got, bool):
            found["owner_recovery"] = abs(float(got))
        if fam == "ordinary":
            continue
        if fam == "reserve" and isinstance(year_v, (int, float)):
            found["yearly"] = abs(float(year_v))
            data["reserve_mode"] = "amount"
            data["reserve_amount"] = found["yearly"]
            a = float(found.get("actual") or 0)
            found["pct"] = 0.0 if a < 0.5 else (found["yearly"] / a) * 100 - 100
            continue
        if fam in ("csos_inc", "csos_exp") and isinstance(year_v, (int, float)):
            found["yearly"] = abs(float(year_v))
            data["auto_csos"] = False
            a = float(found.get("actual") or 0)
            found["pct"] = 0.0 if a < 0.5 else (found["yearly"] / a) * 100 - 100
            continue
        if isinstance(year_v, (int, float)):
            found["yearly"] = abs(float(year_v))
            a = float(found.get("actual") or 0)
            if a > 0.5:
                found["pct"] = (found["yearly"] / a) * 100 - 100
        elif isinstance(pct_v, (int, float)):
            p = float(pct_v)
            found["pct"] = p * 100 if abs(p) <= 2 else p
            a = float(found.get("actual") or 0)
            found["yearly"] = a * (1 + float(found["pct"]) / 100.0)


def _read_meeting_from_sheet(wb, data: dict) -> None:
    """A % or rand typed in the meeting wins over the saved choice."""
    if "BUDGET" not in getattr(wb, "sheetnames", []):
        return
    ws = wb["BUDGET"]
    pct_v = None
    rand_v = None
    for r in range(1, int(ws.max_row or 1) + 1):
        label = str(ws.cell(r, 2).value or "")
        if label.startswith("Approved"):
            pct_v = ws.cell(r, 5).value
            year_v = ws.cell(r, 6).value
            if isinstance(year_v, (int, float)) and float(year_v) > 0.5:
                data["levy_approve_mode"] = "amount"
                data["levy_approved_amount"] = float(year_v)
                return
            if isinstance(pct_v, str) and str(pct_v).startswith("="):
                data["levy_approve_mode"] = "costs"
                return
            if isinstance(pct_v, (int, float)):
                pct = float(pct_v)
                if abs(pct) <= 2:
                    pct *= 100.0
                data["levy_approve_mode"] = "pct"
                data["levy_approved_pct"] = pct
            return


def _restore_from_domus_sheet(wb) -> dict | None:
    if "_DOMUS" not in wb.sheetnames:
        return None
    ws = wb["_DOMUS"]
    if str(ws["A1"].value or "") != "DOMUS_STATE_V1":
        return None
    parts = []
    for r in range(3, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v:
            parts.append(str(v))
    if not parts:
        return None
    data = json.loads("".join(parts))
    secs = data.get("sections") or {}
    # Keep default keys so missing sections don't crash the editors.
    base = default_sections()
    for k in base:
        if k in secs and isinstance(secs[k], list) and secs[k]:
            base[k] = secs[k]
        elif k in secs and isinstance(secs[k], list):
            base[k] = secs[k]
    data["sections"] = base
    return data


def restore_from_app_excel(uploaded) -> dict:
    """Reload a file this app previously downloaded so work is not lost."""
    raw = uploaded.read() if hasattr(uploaded, "read") else uploaded
    bio = BytesIO(raw)
    try:
        wb = load_workbook(bio, data_only=False, read_only=False)
        packed = _restore_from_domus_sheet(wb)
        if packed:
            _apply_budget_sheet(wb, packed)
            _read_meeting_from_sheet(wb, packed)
            try:
                recs, _msg = parse_pq_upload(BytesIO(raw))
                if recs:
                    packed["pq"] = recs
            except Exception:
                pass
            return packed
    except Exception:
        packed = None
    bio = BytesIO(raw)
    xl = pd.ExcelFile(bio)
    names = {n.lower(): n for n in xl.sheet_names}
    out = {
        "sections": {k: [] for k in default_sections()},
        "complex_name": "",
        "fin_year": "",
        "pq": None,
        "ymp": [{"desc": "", "years": [0.0] * 10}],
    }
    if "budget" in names:
        df = pd.read_excel(xl, sheet_name=names["budget"], header=None)
        for r in range(min(8, len(df))):
            for c in range(min(8, df.shape[1])):
                v = str(df.iat[r, c] or "").strip().lower()
                if v.startswith("complex") and c + 1 < df.shape[1]:
                    out["complex_name"] = str(df.iat[r, c + 1] or "").strip()
                if v.startswith("year") and c + 1 < df.shape[1]:
                    out["fin_year"] = str(df.iat[r, c + 1] or "").strip()
            b = str(df.iat[r, 1] if df.shape[1] > 1 else "").strip()
            if r == 1 and b and not out["complex_name"]:
                out["complex_name"] = b
            if r == 2 and re.search(r"20\d{2}", b) and not out["fin_year"]:
                out["fin_year"] = b
        current = None
        i_act, i_pct, i_year, i_note = 2, 3, 4, 6
        for r in range(len(df)):
            desc = str(df.iat[r, 1] if df.shape[1] > 1 else "").strip()
            if not desc or desc.lower() in ("nan", "none"):
                continue
            low = desc.lower()
            if low == "description":
                headers = [str(df.iat[r, c] or "").strip().lower() for c in range(df.shape[1])]
                if any("gl" in h for h in headers):
                    i_act, i_pct, i_year, i_note = 3, 4, 5, 7
                else:
                    i_act, i_pct, i_year, i_note = 2, 3, 4, 6
                continue
            mapped = next((k for title, k in RESTORE_BARS if title in low and "total" not in low and "check" not in low), None)
            if mapped:
                current = mapped
                continue
            if current is None:
                continue
            if low.startswith((
                "meeting decision", "what the costs", "approved %", "approved budget",
                "approved monthly", "or type the rand", "gap",
            )):
                continue
            if re.match(r"^(description|total |net |ordinary levy|current reserve|this year’s|projected)", desc, re.I):
                continue
            act_v = df.iat[r, i_act] if df.shape[1] > i_act else None
            pct_v = df.iat[r, i_pct] if df.shape[1] > i_pct else None
            year_v = df.iat[r, i_year] if df.shape[1] > i_year else None
            actual = abs(num(act_v) or 0.0)
            pct_raw = num(pct_v)
            if pct_raw is None:
                pct_raw = 0.0
            pct = pct_raw * 100 if abs(pct_raw) <= 2 else pct_raw
            yearly = num(year_v)
            if yearly is None:
                yearly = actual * (1 + pct / 100)
            else:
                yearly = abs(yearly)
            note = str(df.iat[r, i_note] if df.shape[1] > i_note else "") or ""
            dest = current
            if current == "recoveries_other" and re.search(r"water|electric|sewer|refuse", norm(desc)):
                dest = "municipal"
            item = row(desc, "" if note in ("nan", "None") else note)
            item["actual"] = actual
            item["pct"] = pct
            item["yearly"] = yearly if yearly else actual * (1 + pct / 100)
            item["is_recovery"] = dest == "municipal" and (
                "recover" in norm(desc) or desc.lower().startswith("less:")
            )
            out["sections"][dest].append(item)
    pq_name = names.get("pq")
    if pq_name:
        pq = pd.read_excel(xl, sheet_name=pq_name, header=None)
        header = None
        for i in range(min(12, len(pq))):
            vals = [str(v).strip().lower() for v in pq.iloc[i].tolist()]
            if "unit" in vals and any(v == "pq" for v in vals):
                header = i
                break
        if header is not None:
            cols = [str(c).strip() for c in pq.iloc[header].tolist()]
            unit_i = next((i for i, c in enumerate(cols) if c.lower() == "unit"), 1)
            pq_i = next((i for i, c in enumerate(cols) if c.lower() == "pq"), 2)
            recs = []
            for r in range(header + 1, len(pq)):
                unit = str(pq.iat[r, unit_i] or "").strip()
                if not unit or unit.lower() in ("nan", "total"):
                    continue
                recs.append({"Unit": unit, "PQ": float(num(pq.iat[r, pq_i]) or 0)})
            if recs:
                out["pq"] = recs
    ymp_name = next((n for k, n in names.items() if "ymp" in k or "10" in k or "maintenance" in k), None)
    if ymp_name:
        ymp = pd.read_excel(xl, sheet_name=ymp_name, header=None)
        projects = parse_ymp_sheet(ymp, NOW_YEAR)
        if projects:
            out["ymp"] = projects
    return out


def apply_falcon_view_notes(state: dict) -> None:
    """Correct Falcon View comments that no longer match the figures. Other complexes are left alone."""
    if "falcon" not in (state.get("complex_name") or "").lower():
        return
    notes = {
        "levies received": "Ordinary levy for the year, worked out from the costs. The 27.42% is the change from last year’s R1,259,000. It was not typed in as an increase.",
        "reserve fund contribution": (
            f"{money(reserve_contribution(state))} this year. "
            + (
                "Equal to 100% of this year’s repairs and maintenance. "
                if state.get("reserve_mode") == "rm100"
                else "Collected for the reserve. "
            )
            + "Interest is already in the opening balance and is not added again."
        ),
        "csos levy recovered": "R18,763.91. Owners pay this on its own column. It is not part of the ordinary levy. The expense is the same amount.",
        "csos levy": "R18,763.91, the same as CSOS income. Left out of the ordinary levy so owners are not charged twice.",
        "investment income": "7% on last year’s R85,615.96 = R91,609.08. Already inside the reserve balance, so it is not added to the reserve again and it does not reduce the levy. Kept here only for the tax estimate.",
        "rental income": "R216,695.40, same as last year. Not used to reduce the levy until the trustees confirm what this rental is for.",
        "audit fees": "R12,663.45, being last year’s R11,835 plus 7%. The previous fee was for the 2023/2024 year-end. Confirm against the latest quotation.",
        "security": "R445,484.18. That is 8.24% above last year’s R411,570.72, not 5%.",
        "meeting expenses": "5% on last year’s R5,459.77. Budget R5,732.76.",
        "gate airtime & data": "Budget R3,000. The % looks very high only because last year was R583 after the once-off MTN deposits were left out.",
        "fire equipment & services": "Budget R1,500 for the annual service. The last invoice was R1,380.",
        "electrical maintenance": "5% on last year’s R10,744.50. Budget R11,281.73.",
        "taxation": "Estimate only. Interest R91,609 plus rental R216,695 = R308,304. First R50,000 is exempt. The rest × 27% = R69,742.21. This tax is in the levy. Confirm with the auditor.",
        "loans payable": "R64,545 from the 10-year plan and taken off the reserve. Confirm that this is a loan repayment, not a maintenance job.",
    }
    for items in (state.get("sections") or {}).values():
        for it in items:
            key = norm(it.get("desc") or "")
            new = notes.get(key)
            if new and it.get("note") != new:
                it["note"] = new


def init():
    ss = st.session_state
    ss.setdefault("sections", default_sections())
    ss.setdefault("complex_name", "")
    ss.setdefault("fin_year", "01-03-2026 / 28-02-2027")
    ss.setdefault("reserve_mode", "amount")
    ss.setdefault("reserve_amount", 0.0)
    ss.setdefault("reserve_balance", 0.0)
    ss.setdefault("scheme_type", "bc")
    ss.setdefault("special_in_ordinary", False)
    ss.setdefault("afs_sections", None)
    ss.setdefault("wcu_rows", None)
    ss.setdefault("auto_csos", True)
    ss.setdefault("has_master_hoa", False)
    ss.setdefault("insurance_mode", "levy")
    ss.setdefault("insurance_bill_yearly", 0.0)
    ss.setdefault("pq", None)
    ss.setdefault("ymp", [{"desc": "", "years": [0.0] * 10}])
    ss.setdefault("msg", "")
    ss.setdefault("current_monthly_levy", 0.0)
    ss.setdefault("actual_months", 12)
    ss.setdefault("levy_approve_mode", "costs")
    ss.setdefault("levy_approved_pct", 0.0)
    ss.setdefault("levy_approved_amount", 0.0)
    ss.setdefault("estate_levy_yearly", 0.0)
    ss.setdefault("estate_levy_name", "Estate / master HOA levy")
    ss.setdefault("estate_levy_mode", "separate")
    ss.setdefault("estate_split", "equal")


def section_form(key: str, title: str, help_text: str, rm: bool = False, recover: bool = False, income: bool = False, claims: bool = False):
    st.subheader(title)
    if help_text:
        st.caption(help_text)
    items = st.session_state.sections.get(key) or []
    claim_map = claims_against(st.session_state) if key in ("expenditure", "rm") else None
    cost_names = ["— none —"]
    if claims:
        for sec in ("expenditure", "rm"):
            for it in st.session_state.sections.get(sec) or []:
                name = (it.get("desc") or "").strip()
                if name and name not in cost_names:
                    cost_names.append(name)
        for it in items:
            name = (it.get("claim_against") or "").strip()
            if name and name not in cost_names:
                cost_names.append(name)
    df = items_to_df(items, rm, recover, income=income, claims=claims, claim_map=claim_map)
    nonce = int(st.session_state.get(f"_nonce_{key}") or 0)
    edited = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key=f"grid_{key}_{nonce}",
        column_config={
                "Description": st.column_config.TextColumn("Description", width="medium"),
                "Actual": st.column_config.NumberColumn("Actual", format="%.2f"),
                "% Increase": st.column_config.NumberColumn(
                    "% Increase",
                    format="%.2f",
                    help="Type 10 for +10%, then Save. We set Budgeted yearly = Actual × 1.10",
                ),
                "Budgeted yearly": st.column_config.NumberColumn(
                    "Budgeted yearly",
                    format="%.2f",
                    help="The full bill. Or type the rand amount then Save. We fill in the %.",
                ),
                "Monthly": st.column_config.NumberColumn("Monthly", format="%.2f", disabled=True, help="What goes into the levy, per month. Owner recoveries, insurance payouts and recovery income are already taken off."),
                "Insurance payout": st.column_config.NumberColumn("Insurance payout", format="%.2f"),
                "Recovered from some owners": st.column_config.NumberColumn(
                    "Recovered from some owners",
                    format="%.2f",
                    help="What some owners pay towards this same bill. Not income. Not a levy column. The levy carries the bill minus this.",
                ),
                "Lower the levies": st.column_config.CheckboxColumn(
                    "Lower the levies",
                    help="Tick only the income that should bring the ordinary levy down. Leave it off if this complex does not use that income.",
                ),
                "Comes off this cost": st.column_config.SelectboxColumn(
                    "Comes off this cost",
                    options=cost_names,
                    help="Pick the expense or repair this money pays. That cost stays on the books, and this amount comes off it.",
                ),
                "Notes": st.column_config.TextColumn("Notes", width="large", help="Shows on the Excel Comments / Notes column. Click Save after typing."),
            },
            disabled=["Monthly"],
        )
    saved = st.button("Save this section", type="primary", key=f"save_{key}")
    if saved:
        st.session_state.sections[key] = save_editor(edited, items, rm, municipal=(key == "municipal"))
        st.session_state[f"_nonce_{key}"] = nonce + 1
        if key == "levy":
            for r in st.session_state.sections["levy"]:
                if is_ins_bill_line(r.get("desc") or "") and float(r.get("yearly") or 0) > 0.5:
                    st.session_state["_pending_insurance_bill"] = float(r["yearly"])
                    break
        apply_levy_lines(st.session_state)
        st.success("Saved. A new % sets Budgeted yearly = Actual × (1 + %). A new yearly amount sets the %.")
        st.rerun()
    if key == "municipal":
        g, rec = municipal_gross_and_rec(st.session_state)
        st.caption(
            f"Add the city-bill lines {money(g)}, then subtract the recovered lines {money(rec)}. "
            f"Net in the levy = {money(g - rec)}. Do not add the recovered lines on top of the city bill."
        )
    elif key == "other":
        used = [r for r in st.session_state.sections.get("other") or [] if r.get("reduces_levy") and float(r.get("yearly") or 0) > 0.5]
        if used:
            names = ", ".join(r.get("desc") or "" for r in used)
            st.caption(f"Lowering the levy by {money(income_used(st.session_state))}: {names}. Unticked income stays on the sheet and does not change the levy.")
        else:
            st.caption("Nothing is ticked, so other income does not change the levy. Tick a line only for this complex.")
    elif key == "recoveries_other":
        linked = [r for r in st.session_state.sections.get("recoveries_other") or [] if (r.get("claim_against") or "").strip() and float(r.get("yearly") or 0) > 0.5]
        if linked:
            bits = [f"{r.get('desc')} {money(float(r.get('yearly') or 0))} comes off {r.get('claim_against')}" for r in linked]
            st.caption(" · ".join(bits))
        else:
            st.caption("Pick the cost each recovery pays, then Save. Until you pick one, it does not change the levy.")
    else:
        st.caption(f"Section net total: {money(sum_net(st.session_state.sections[key]))}  ·  Monthly total: {money(sum_net(st.session_state.sections[key]) / 12)}")


def main():
    init()
    apply_falcon_view_notes(st.session_state)
    pending = st.session_state.pop("_pending_insurance_bill", None)
    if pending is not None:
        st.session_state.insurance_bill_yearly = float(pending)
    logo = Path(__file__).parent / "domus_logo.jpeg"
    cols = st.columns([1, 5])
    with cols[0]:
        if logo.exists():
            st.image(str(logo), width=110)
    with cols[1]:
        st.title("Domus Property Management Budget")
        st.caption("Load the financial statement PDF (line names) then WeConnectU Excel (rands). Either order is fine.")

    apply_levy_lines(st.session_state)
    s = st.session_state.sections
    ord_amt = ordinary_total(st.session_state)
    reserve = next((r for r in s["levy"] if "reserve" in r["desc"].lower()), None)
    ordinary_row = next((r for r in s["levy"] if "ordinary" in r["desc"].lower()), None)
    actual_year = float(ordinary_row["actual"]) if ordinary_row else 0.0
    months = max(1, int(st.session_state.get("actual_months") or 12))
    if months < 12 and actual_year > 0:
        actual_year_full = actual_year / months * 12
    else:
        actual_year_full = actual_year
    current_m = float(st.session_state.get("current_monthly_levy") or 0)
    if current_m <= 0 and actual_year_full > 0:
        current_m = actual_year_full / 12
    approved = approved_ordinary(st.session_state)
    new_m = approved / 12
    levy_pct = 0.0 if current_m == 0 else (new_m / current_m) * 100 - 100

    if actual_year > 0 and ord_amt > 0 and 8 <= (ord_amt / actual_year) <= 15:
        st.error(
            f"Ordinary **Actual** looks like one month ({money(actual_year)}), but "
            f"**Budgeted yearly** is a full year ({money(ord_amt)}). "
            f"Put the full-year levy in Actual (about {money(actual_year * 12)}), "
            f"or type the current monthly total in the sidebar. "
            f"Trustees should compare {money(actual_year)} / month with {money(new_m)} / month."
        )

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("What owners pay now (monthly)", money(current_m))
    m2.metric("What owners will pay (monthly)", money(new_m), delta=f"{levy_pct:+.1f}%")
    m3.metric("Approved ordinary levy", money(approved))
    m4.metric("Reserve for the year", money(float(reserve["yearly"]) if reserve else 0))
    xanadu_left = 0.0
    for r in (s.get("expenditure") or []) + (s.get("hoa_expense") or []):
        d = (r.get("desc") or "").lower()
        if "xanadu" in d or "eco park" in d:
            xanadu_left += float(r.get("yearly") or 0)
    if xanadu_left > 1 and not st.session_state.get("has_master_hoa"):
        st.warning(
            f"An estate levy of {money(xanadu_left)} is still on the books. "
            "Tick **We still bill an estate levy to owners** on Income if owners pay it through us. "
            "If owners pay that estate themselves, set that expense to 0."
        )
    if insurance_on_pq(st.session_state) is False and insurance_expense_amount(st.session_state) > 1:
        st.info(
            "Insurance is inside ordinary levies. If owners pay it as its own line on the invoice, "
            "choose **Extra on the owner invoice**."
        )

    with st.expander("Why is the levy this amount? (plain English)", expanded=True):
        st.write(
            "**Ordinary levies (Budgeted yearly)** is the money owners must put in to cover this year’s costs. "
            "It is **not** last year’s levy plus a %."
        )
        st.write(
            "**Ordinary = net municipal + expenditure + R&M + personnel + tax − income you ticked** "
            "(and special projects only if that box is ticked)."
        )
        st.write(
            "**Left out of ordinary:** estate pass-through, CSOS (own PQ column), "
            "insurance when it is extra on the invoice, and any line you set to R0 (for example garden service if owners pay the gardener themselves)."
        )
        for label, amt in levy_pieces(st.session_state):
            if amt or "Special" not in label:
                st.write(f"- {label}: **{money(amt)}**")
        st.write(f"- **What the costs need: {money(ord_amt)}** for the year ({money(ord_amt / 12)} a month).")
        if abs(ord_amt - approved) > 1:
            st.write(
                f"- **What the meeting approves: {money(approved)}** "
                f"({money(approved / 12)} a month). "
                f"Gap **{money(ord_amt - approved)}**. "
                "A gap above zero means a cost must come down, or the reserve covers it."
            )
        flags = odd_budget_lines(st.session_state)
        if flags:
            st.warning("These lines are making the levy jump. Fix them on the tabs, then Save — you do not need to start over.")
            for f in flags:
                st.write(f"- {f}")

    with st.sidebar:
        st.header("Complex")
        st.session_state.complex_name = st.text_input("Complex name", st.session_state.complex_name)
        st.session_state.fin_year = st.text_input("Financial year", st.session_state.fin_year)
        st.caption("Key West only. This replaces the lines on the screen. Download Excel first if you need to keep what is there.")
        if st.button("Load the Key West draft"):
            st.session_state.sections = key_west_sections()
            st.session_state.complex_name = "Key West Body Corporate"
            st.session_state.scheme_type = "bc"
            st.session_state.insurance_mode = "levy"
            st.session_state.insurance_bill_yearly = 0.0
            st.session_state.special_in_ordinary = False
            st.session_state.has_master_hoa = False
            st.session_state.auto_csos = bool(st.session_state.get("pq"))
            st.session_state.reserve_mode = "amount"
            st.session_state.reserve_balance = 1079058.82
            st.session_state.reserve_amount = 1162666.92
            st.session_state.actual_months = 12
            st.session_state.current_monthly_levy = round(5119247.24 / 12, 2)
            if not (st.session_state.fin_year or "").strip():
                st.session_state.fin_year = "1 January 2027 – 31 December 2027"
            for key in (
                "levy", "other", "hoa_income", "hoa_expense", "recoveries_other",
                "municipal", "expenditure", "rm", "personnel", "tax", "special", "fixed",
            ):
                st.session_state[f"_nonce_{key}"] = int(st.session_state.get(f"_nonce_{key}") or 0) + 1
            apply_levy_lines(st.session_state)
            st.session_state.msg = (
                "Key West draft loaded. Income is only Ordinary, Reserve and CSOS. "
                "Boathouse, boatport, clubhouse, interest, penalties and the window recovery are under Other Income. "
                "Boathouse and boatport are ticked, but their budget is 0 until you type a %. "
                "Investec is +6% and FNB is +4%. The reserve bank balance is on the left, and the contribution is R1,162,666.92. "
                "Special projects are not in the levy. Insurance is inside the ordinary levy. "
                "Change the financial year if it is not a calendar year."
            )
            st.rerun()

        st.header("Load last year")
        st.caption(
            "1. Financial statement PDF — line names. "
            "2. WeConnectU Excel — last year’s rands. "
            "3. PQ Excel — each unit’s share."
        )
        pdf_up = st.file_uploader("Annual financial statement (PDF)", type=["pdf"], key="afs_pdf")
        if pdf_up and st.button("Load financial statement", type="primary"):
            try:
                rows, name = extract_afs_pdf(pdf_up)
                if not rows:
                    st.error(
                        "This PDF has no selectable text (it is a scan / picture). "
                        "Wait for the app to reboot with OCR, or load the WeConnectU Excel — that has the same lines. "
                        "Ask the auditor for a PDF you can highlight text in if OCR still fails."
                    )
                else:
                    secs = sections_from_afs(rows)
                    st.session_state.afs_sections = secs
                    if name and not st.session_state.complex_name:
                        st.session_state.complex_name = name
                    if st.session_state.get("wcu_rows"):
                        secs, added = match_into(st.session_state.wcu_rows, secs)
                        st.session_state.msg = (
                            f"FS lines loaded ({len(rows)}). WeConnectU rands applied. Extra lines: {added}."
                        )
                    else:
                        st.session_state.sections = secs
                        st.session_state.msg = (
                            f"Loaded {len(rows)} lines from the financial statements. "
                            "Now load WeConnectU Excel to fill last year’s rands."
                        )
                    if st.session_state.get("wcu_rows"):
                        st.session_state.sections = secs
                    apply_levy_lines(st.session_state)
                    st.rerun()
            except Exception as e:
                st.error(f"Could not read PDF: {e}")

        up = st.file_uploader("WeConnectU Actual vs Budget (Excel)", type=["xlsx", "xls", "xlsm"])
        if up and st.button("Load Excel"):
            try:
                rows = extract_wcu(up)
                if not rows:
                    st.error("No lines found. Export Options → Budget and Actuals.")
                else:
                    st.session_state.wcu_rows = rows
                    base = st.session_state.get("afs_sections") or default_sections()
                    secs, added = match_into(rows, base)
                    st.session_state.sections = secs
                    source = "financial statement lines" if st.session_state.get("afs_sections") else "standard template"
                    st.session_state.msg = f"Loaded {len(rows)} WeConnectU lines onto the {source}. Extra lines: {added}."
                    apply_levy_lines(st.session_state)
                    st.rerun()
            except Exception as e:
                st.error(f"Could not read file: {e}")

        pq_up = st.file_uploader(
            "PQ / unit ratios (Excel or CSV)",
            type=["csv", "xlsx", "xls", "xlsm"],
            key="pq_sidebar",
        )
        st.caption("WeConnectU’s unit file, or this budget’s Excel after you correct the PQ sheet. Then click Load PQs. The ratios must add to 1.")
        if pq_up and st.button("Load PQs"):
            try:
                records, msg = parse_pq_upload(pq_up)
                st.session_state.pq = records
                st.session_state.msg = msg
                st.rerun()
            except Exception as e:
                st.error(f"Could not read PQ file: {e}")

        st.header("Keep your work")
        st.caption("Do not click Start over. Download Excel, then Restore it here. New downloads restore 100% (the file keeps a hidden copy of your numbers).")
        rest = st.file_uploader("Restore this app’s Excel", type=["xlsx"], key="restore_xlsx")
        if rest and st.button("Restore my budget"):
            try:
                data = restore_from_app_excel(rest)
                st.session_state.sections = data["sections"]
                if data.get("complex_name"):
                    st.session_state.complex_name = data["complex_name"]
                if data.get("fin_year"):
                    st.session_state.fin_year = data["fin_year"]
                if data.get("pq"):
                    st.session_state.pq = data["pq"]
                if data.get("ymp"):
                    st.session_state.ymp = data["ymp"]
                for k in (
                    "has_master_hoa", "insurance_mode", "insurance_bill_yearly", "auto_csos",
                    "reserve_mode", "reserve_amount", "reserve_balance", "scheme_type", "special_in_ordinary",
                    "current_monthly_levy", "actual_months",
                    "levy_approve_mode", "levy_approved_pct", "levy_approved_amount",
                ):
                    if k in data and data[k] is not None:
                        st.session_state[k] = data[k]
                apply_levy_lines(st.session_state)
                n_lines = sum(len(v) for v in st.session_state.sections.values())
                st.session_state.msg = f"Budget restored ({n_lines} lines). Nothing was started over."
                st.rerun()
            except Exception as e:
                st.error(f"Could not restore: {e}")

        st.header("What owners pay now")
        st.session_state.current_monthly_levy = st.number_input(
            "Current ordinary levy — all units, one month",
            value=float(st.session_state.current_monthly_levy),
            min_value=0.0,
            step=100.0,
            help="Example: R26 400 per month for the whole complex. Not the yearly total.",
        )
        st.session_state.actual_months = st.number_input(
            "Months covered by the Actual column",
            min_value=1,
            max_value=12,
            value=int(st.session_state.actual_months),
            help="12 = a full year. If WeConnectU is only 6 months, put 6 and we scale up for the %.",
        )
        st.header("Reserve fund")
        st.session_state.scheme_type = st.radio(
            "What kind of scheme is this?",
            ["bc", "hoa"],
            format_func=lambda x: "Sectional title body corporate" if x == "bc" else "Homeowners association (HOA)",
            index=0 if st.session_state.get("scheme_type") != "hoa" else 1,
        )
        st.session_state.reserve_balance = st.number_input(
            "How much is already in the reserve fund?",
            value=float(st.session_state.get("reserve_balance") or 0),
            min_value=0.0,
            step=1000.0,
            help="Money in the reserve bank account now. Not this year’s contribution.",
        )
        mode_now = st.session_state.get("reserve_mode") or "amount"
        if mode_now in ("15pct", "legal"):
            mode_now = "pct15"
        if mode_now not in ("amount", "pct15", "pct25", "rm100"):
            mode_now = "amount"
        st.session_state.reserve_mode = st.radio(
            "This year’s reserve contribution",
            ["amount", "pct15", "pct25", "rm100"],
            index=["amount", "pct15", "pct25", "rm100"].index(mode_now),
            format_func=lambda x: {
                "amount": "Own amount",
                "pct15": "15% of last year’s ordinary levy (Actual)",
                "pct25": "25% of last year’s ordinary levy (Actual)",
                "rm100": "100% of this year’s Repairs and Maintenance",
            }[x],
        )
        if st.session_state.reserve_mode == "amount":
            st.session_state.reserve_amount = st.number_input(
                "Reserve fund contribution (yearly rands)",
                value=float(st.session_state.reserve_amount),
                step=1000.0,
                min_value=0.0,
            )
        contrib = reserve_contribution(st.session_state)
        st.caption(f"This year’s contribution: {money(contrib)}.")
        if st.session_state.reserve_mode in ("pct15", "pct25") and prev_admin_contributions(st.session_state) < 1:
            st.warning("Put last year’s ordinary levy in the Actual column on Levy Income. 15% and 25% use that figure.")
        if st.session_state.scheme_type == "hoa":
            st.caption(
                "An HOA may use any of these. There is no automatic 15% or 25% rule. "
                "Use the option the MOI, constitution or members approved."
            )
        else:
            rule = bc_reserve_rule(st.session_state)
            if rule["band"] != "unknown":
                st.caption(
                    f"Body corporate note only: the Act’s minimum would be {money(rule['minimum'])}. "
                    "That does not change the option you picked above."
                )
        opening = float(st.session_state.reserve_balance or 0)
        projects = reserve_project_spend(st.session_state)
        st.caption(
            f"Already in reserve {money(opening)} + this year {money(contrib)} "
            f"− this year’s projects {money(projects)} "
            f"= projected {money(projected_reserve(st.session_state))}."
        )
        st.caption(
            "Interest already sits inside the reserve balance, so it is not added again. "
            "Leave the interest line under Other Income only for the tax estimate."
        )
        st.caption(
            "Guidance only, not legal advice. Body corporate minimums: STSMA Regulation 2. "
            "HOA amounts follow the MOI or constitution."
        )
        st.session_state.special_in_ordinary = st.checkbox(
            "Add Special Projects into ordinary levies",
            value=st.session_state.special_in_ordinary,
            help="Tick only if special work is paid from levies, not from the reserve fund.",
        )
        st.header("Second levy (master estate)")
        st.caption("Use this when owners also pay another estate / HOA.")
        st.session_state.estate_levy_name = st.text_input(
            "Name on the owner schedule",
            st.session_state.estate_levy_name,
        )
        st.session_state.estate_levy_mode = st.radio(
            "Who collects it?",
            ["separate", "we_collect"],
            format_func=lambda x: (
                "Owners pay the estate themselves — do not put it in our ordinary levy"
                if x == "separate"
                else "We collect it and pay the estate"
            ),
            index=0 if st.session_state.estate_levy_mode == "separate" else 1,
        )
        st.session_state.estate_levy_yearly = st.number_input(
            "Estate levy for the whole complex (yearly rands)",
            value=float(st.session_state.estate_levy_yearly),
            min_value=0.0,
            step=1000.0,
        )
        st.session_state.estate_split = st.radio(
            "How is it split per unit?",
            ["equal", "pq"],
            format_func=lambda x: "Same amount each unit" if x == "equal" else "By PQ",
            index=0 if st.session_state.estate_split == "equal" else 1,
        )
        st.caption("This wipes the screen. Download Excel first if you still need the numbers.")
        if st.button("Erase everything"):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()

    if st.session_state.msg:
        st.success(st.session_state.msg)

    tabs = st.tabs([
        "How it works", "Income", "Municipal", "Expenditure",
        "Repair & Maintenance", "Personnel", "Tax", "Special",
        "PQ / Levies", "10-year plan", "For the meeting", "Download",
    ])

    with tabs[0]:
        st.subheader("How to use this budget")
        st.markdown(
            """
**Do this in order.** Click **Save this section** after every tab you change. Nothing is kept until Save. You can add a row at the bottom of any table if this complex has a line that is not listed.

### 1. Name the complex
In the sidebar, type the **complex name** and the **financial year** (for example 1 March 2027 – 28 February 2028).

Choose **body corporate** or **HOA**. The four reserve choices work for both. The note under the reserve box changes: a body corporate also sees the Act’s minimum, an HOA follows its MOI or constitution. The note does not change the amount. You do.

**Key West:** click **Load the Key West draft** on the left. Income on the sheet is only **Ordinary Levies**, **Reserve Fund Contribution** and **CSOS**. Boathouse, boatport, clubhouse, the two bank-interest lines, penalties and the window recovery sit under **Other Income**, with **Lowers the levy** Yes or No. Actual is the full-year figure. Type a % and Save, and Budgeted yearly moves. It replaces what is on the screen, so download Excel first if you need to keep other work.

### 2. Load last year
On the left, in this order:

1. **Annual financial statement (PDF)** — brings in the line names, so the budget reads like the financial statements.
2. **WeConnectU Actual vs Budget (Excel)** — brings in last year’s rands (the Actual column). Click **Load Excel**.
3. **PQ / unit ratios** — each unit’s share. Click **Load PQs**. You can also upload the same file on the **PQ / Levies** tab.

If the WeConnectU file only covers part of the year, set **Months covered by the Actual column**. 12 means a full year and nothing is stretched. 7 means seven months, and the levy comparison is scaled up to a year. Leave it on **12** unless you mean to stretch.

**Current ordinary levy — all units, one month** is what owners pay now, for the whole complex, for one month. It is only for the comparison at the top. It is not the new budget.

### 3. Type this year’s amounts
On each cost tab you have Actual, % Increase, Budgeted yearly, Monthly, and Notes.

- Press **Enter** in the cell (so the number is kept), then **Save this section**.
- Type **%** and Save → Budgeted yearly = Actual × (1 + %). Example: 10 means plus 10%.
- Type **Budgeted yearly** and Save → the % fills in. % = (yearly ÷ actual) × 100 − 100.
- If Actual is 0, a % cannot make a rand. Type the **Budgeted yearly** instead.
- **Monthly** is always yearly ÷ 12. You cannot type it.
- **Notes** go to the Comments column on the Excel sheet. Save after you type them.
- A new line: type it on the blank row at the bottom, then Save.

### 4. What ordinary levies are
Ordinary levies are **not** last year’s levy plus a %.

**Ordinary = net municipal + expenditure + repairs and maintenance + personnel + tax − only the other income you tick**

Special projects are added only if you tick **Add Special Projects into ordinary levies**.

The **% on Levies received** is the result of that sum. Do not type over it. It will not stick.

The box **Why is the levy this amount?** at the top lists each piece in rand.

**Left out of ordinary** (own column, or R0):

- Reserve fund
- CSOS
- Insurance, when owners pay it on its own invoice line
- A master-estate levy, when owners pay another estate through us
- Garden service, when owners pay the gardener themselves (set that Expenditure line to R0)
- Income you did **not** tick

### 5. Other income — you tick it
On **Income → Other Income** each line has **Lower the levies**.

- Tick a line only if **this** complex uses that money to pay the running costs. Rental is the usual one.
- Nothing is ticked for you. Complexes are not the same.
- Unticked income stays on the sheet and does **not** change the levy.
- Do **not** tick bank interest if that interest is already inside the reserve balance. Using it again would spend it twice. Leave it for the tax estimate.
- Do **not** tick interest on arrears or penalties unless you are sure that money will come in. You cannot count on owners paying late.
- Save. The caption under the table shows how much is coming off the levy.

On the Excel sheet the same choice is the column **Lowers the levy** (Yes or No). The ordinary-levy formula subtracts only the Yes lines. If the meeting changes a Yes amount, the levy moves with it.

### 6. Recovery income — money that pays a cost
**Income → Recovery income** is money that pays a cost we already have. It is not added on top of the levy.

Type the amount, then under **Comes off this cost** pick the expense or repair it pays. The cost stays on the books. The recovery comes off that cost, never below R0.

- **Legal fees recovered:** pick **Legal Expense**. If the bill and the recovery are the same, legal adds nothing to the levy.
- **A repair invoiced to an owner** (a broken window): use **Charged to an owner**, or add a row. Pick the repair line, for example General Building.
- **Insurance:** either type the payout on the repair line, **or** put the recovery here and pick that repair. Do not do both, or it comes off twice. The app warns you if both are filled in.

Until you pick a cost, the recovery does not change the levy.

On the Excel sheet this block is **RECOVERY INCOME**. Column **Comes off this cost** names the line, and that line’s yearly formula subtracts the recovery.

### 7. Municipal
On **Municipal**, the city bill and the recovery are separate lines.

- Type the city bill as a positive amount (electricity, water, sewerage, refuse, rates).
- Type what owners pay back as a positive amount on the **Less: … recovered** line. Do not type a minus. Any line with the word **recovered** in the name is subtracted, including a sewer-plant recovery.
- **Net municipal** = city bill minus recovered. That net is what goes into the levy.
- If recovered is less than the city bill, the difference is an **under-recovery**. It is already inside the net. Do not add it again. The sheet lists it so the trustees can see it and attend to it.

### 8. Expenditure and repairs
**Expenditure** is the operating costs (accounting, audit, security, legal, garden contract, and so on). Not repairs, not salaries, not tax.

**Repair and Maintenance** is the repair lines.

On both:

- **Budgeted yearly** is the full bill.
- **Recovered from some owners** is what only some owners pay towards that same bill (for example a private refuse company for a few units). It is not income, not municipal, and not a column on the PQ. The levy carries the bill minus that amount.
- On repairs, **Insurance payout** is what the insurer pays towards that job. It is a deduction. Do not type the new budget in the payout column.

### 9. Garden and other equal charges
If owners pay the gardener themselves, set **Garden service** on Expenditure to **R0**. Do not also put it under Fixed monthly charges. **Garden Expenses** on Repair and Maintenance stays in the levy (small garden repairs).

**Fixed monthly charges** on the Income tab are extras that are the **same rand for every unit** (prepaid meters, an Eskom fixed charge, a communal charge). Type the **yearly total for the whole complex**. Each owner pays that total ÷ 12 ÷ number of units, not by PQ. Do not put insurance or garden here.

### 10. Reserve fund
Type **How much is already in the reserve fund?** That is the bank balance now, not this year’s contribution.

Then pick one. It works for a body corporate and an HOA:

- Own amount
- 15% of last year’s ordinary levy (the Actual on Levies received)
- 25% of last year’s ordinary levy (that same Actual)
- 100% of this year’s Repairs and Maintenance (after recovery income and payouts)

The amount shows on **Reserve Fund Contribution** (yearly and monthly) and in the box at the top of the Excel sheet. Those two are the same figure.

If last year had no reserve contribution, Actual on that line stays 0. That is fine.

**Projection** = already in reserve + this year’s contribution − this year’s special projects (only if the reserve pays them).

Interest already sits inside the opening balance. It is not added again. Leave the interest line under Other Income for the tax estimate, and leave **Lower the levies** off.

The body-corporate minimum on the side is guidance only, not legal advice. It does not override the option you picked.

### 11. Insurance
On the Income tab, choose one:

- **Inside the ordinary levy** — no extra column. The premium on Expenditure is part of the levy.
- **Extra on the owner invoice** — owners pay it on its own PQ column. Keep the premium on Expenditure. It is **not** inside ordinary levies. Type the yearly amount to bill, or leave 0 to use the premium.

### 12. Master estate
Tick **We still bill an estate levy to owners** only if owners pay another estate **through us**. Those amounts are extra PQ columns, not part of ordinary levies.

If owners pay that estate themselves, leave the tick off and set that expense to R0. It must not sit in our levy.

In the sidebar you can also name that levy, type the yearly total, and split it **the same rand each unit** or **by PQ**.

### 13. CSOS
We collect CSOS from owners and pay the same amount to CSOS. Income and expense are the same rand, so owners are not charged twice. It is its own PQ column, not inside ordinary levies.

Leave **Calculate CSOS** ticked. The formula is 2% of (monthly admin levy − R500), maximum R40 per unit per month. It follows the **approved** levy, not only the cost total.

Untick it only if you must type a different CSOS figure yourself.

### 14. Income tax
Levies are not taxed. Interest, investment income and rent can be.

On the **Tax** tab, **Put this estimate on the tax line** uses Other Income: the first R50,000 is exempt, the rest × 27%. Or type the auditor’s yearly figure and Save. That tax is a cost inside the ordinary levy.

Owners do not pay SARS every month. The monthly column is only so the levy can fund one yearly amount.

If the estimate is R0, nothing is paid. If tax is payable, SARS usually wants provisional tax: one payment six months into the year, and one at year-end. A top-up can be due a few months after year-end. A February year-end is often 31 August, 28 February, and a top-up by 30 September if needed. The return is filed after year-end. Confirm the dates and the amount with the auditor.

### 15. Personnel and special projects
**Personnel** is salaries, casuals, PAYE / UIF, travel, bonuses and the like.

**Special Projects** is Year 1 of the 10-year plan. Tick **Add Special Projects into ordinary levies** only if the levy must pay for that work this year. If the reserve pays it, leave the tick off. It then comes off the reserve projection, not the levy.

### 16. PQ / what each owner pays
Upload the unit file in the sidebar or on **PQ / Levies**. If the PQ column is 0, the app uses **Ratio 1**. The ratios must add up to about 1.000.

Each owner’s share of a levy column = their PQ × that column’s monthly total.

Equal charges (and an estate levy set to “same amount”) are split by the number of units, not by PQ.

The schedule uses the **approved** ordinary levy once the meeting has set one. Until then it uses what the costs need.

### 17. 10-year plan
Year 1 is **this budget year** (2026 while we are in 2026), not last year’s first column.

Upload the plan, or paste the whole block from Excel, including the year headings. An older sheet that starts in 2025 is rolled forward: last year’s next year becomes this Year 1.

**Copy Year 1 into Special Projects** only when the levy must pay those jobs this year. Then go back and tick the special-projects box if that is the decision. Save the plan after you edit it.

### 18. The meeting
You prepare the budget. The trustees decide in the meeting.

**Download Excel** on the Download tab. Yellow cells are the ones they may type in. The sheet has formulas, so a change to a % or a rand updates the yearly amount, the monthly amount, the levy and the PQ.

Under Levies received there is one yellow line: **Approved — type a % or a rand**.

- They type the % the meeting agrees (10 means plus 10% on last year’s levies).
- Or they type a rand amount over the yearly figure and leave the %.
- **What the costs need** stays as the reference.
- **Gap** is the costs minus the approved amount. Above zero means the costs are still higher. They cut a cost, or they take the gap from the reserve.
- The PQ uses the **approved** amount.

On a normal cost line, the **yellow** cell is the one you type. The other one calculates, and Monthly follows.

- If you set that line with a **%**, the % is yellow. Budgeted yearly = Actual × (1 + %). Change the % and the yearly amount moves.
- If you set that line with a **rand**, Budgeted yearly is yellow. The % = yearly ÷ Actual − 1. Change the rand and the % moves.
- A line with an insurance payout, recovery income, or “recovered from some owners” always keeps the % yellow, so the yearly amount stays a formula.

They can also change notes, **Lowers the levy** (Yes or No), **Comes off this cost**, and the reserve when it is an own amount.

Do **not** type on the % next to Levies received itself. That % is worked out from the costs. It jumps back.

If you already know the decision before you download, you can set the same choice on the Income tab under **Meeting decision**. The sheet still shows it, so the meeting can change it.

### 19. When the sheet comes back
**Restore my budget** on the left, and choose their file. Their approved levy, yellow cells, notes, which income lowers the levy, and which recovery comes off which cost, come back. Then download again if you need a clean sheet.

Do not use **Erase everything** unless you mean to wipe the screen. Download first if you still need the numbers.

### 20. The owner presentation
On **For the meeting**, type the complex name, then **Build the presentation**.

- **Download the presentation** is the full-screen version. Open it in Chrome. Arrow keys move on. F is full screen. The Domus logo comes in at the start and leaves at the end.
- **Download the PowerPoint** is the same story. Open it and press F5. The logo fills the first screen, then zooms out onto the budget.

It uses this complex’s own numbers, in plain words, so owners can see what the levy pays for. Build it again after you change figures, so the slides match the sheet.
            """
        )

    with tabs[1]:
        st.info("Ordinary and Reserve update when you save the cost sections / sidebar.")
        st.session_state.insurance_mode = st.radio(
            "Insurance — how do owners pay it?",
            ["levy", "pq"],
            format_func=lambda x: (
                "Inside the ordinary levy (no extra column on the invoice)"
                if x == "levy"
                else "Extra on the owner invoice (its own column)"
            ),
            index=0 if st.session_state.get("insurance_mode") != "pq" else 1,
            help="Extra = we recover the premium from owners. It is NOT in ordinary levies. Part of levy = insurance expense stays in the admin levy.",
        )
        if st.session_state.insurance_mode == "pq":
            prem = insurance_expense_amount(st.session_state)
            last_rec = 0.0
            for r in st.session_state.sections.get("levy") or []:
                if family(r["desc"]) == "ins_bill" or re.search(r"insurance recovered|levy\s*[-–]\s*insurance", r["desc"], re.I):
                    last_rec = float(r.get("actual") or 0)
                    break
            if "insurance_bill_yearly" not in st.session_state:
                st.session_state.insurance_bill_yearly = 0.0
            st.number_input(
                "Insurance billed to owners this year (0 = use the Insurance recovered line, or the premium if that is 0)",
                min_value=0.0,
                step=100.0,
                key="insurance_bill_yearly",
                help="Type the rand amount you want on the PQ. Save on Levy Income also keeps what you type on Insurance recovered.",
            )
            bill = insurance_bill_amount(st.session_state)
            st.caption(
                f"PQ Insurance column = {money(bill)} a year ({money(bill/12)} / month for the complex). "
                f"Insurance expense {money(prem)} is left out of ordinary levies."
            )
            if last_rec and prem and last_rec + 1 < prem:
                st.warning(
                    f"Last year owners were billed {money(last_rec)} but the premium was {money(prem)}. "
                    f"That is why recovery looks too small. This year we bill {money(bill)} on the PQ."
                )
        st.checkbox(
            "Calculate CSOS from the legal formula: 2% of (monthly admin levy − R500), max R40 per unit per month.",
            key="auto_csos",
            help="2% of (monthly admin levy minus R500), maximum R40 per unit per month. CSOS is its own column, not inside ordinary levies.",
        )
        if st.session_state.auto_csos and st.session_state.pq:
            own = scheme_csos_yearly(st.session_state, approved_ordinary(st.session_state))
            st.caption(f"This scheme’s CSOS = {money(own)} a year ({money(own/12)} / month for the complex). It is a PQ column, not part of ordinary levies.")
        st.divider()
        st.divider()
        st.subheader("Meeting decision — Levies received")
        st.caption(
            "You do not set this before the meeting. Download the sheet. "
            "In the meeting the trustees use one yellow line, **Approved — type a % or a rand**. "
            "They type the % in the % column, or they type a rand over the yearly amount. "
            "They can also change the other yellow cells. When they send the file back, use **Restore my budget**. "
            "The % on the Levies received line itself still will not stick."
        )
        _mode = st.radio(
            "What should owners be charged?",
            ["costs", "pct", "amount"],
            format_func=lambda x: {
                "costs": "What the costs need",
                "pct": "A % on last year’s levies",
                "amount": "A rand amount for the year",
            }[x],
            index=["costs", "pct", "amount"].index(st.session_state.get("levy_approve_mode") or "costs"),
            key="levy_approve_mode",
        )
        _costs_now = ordinary_total(st.session_state)
        _actual_now = ordinary_actual(st.session_state)
        if _mode == "pct":
            st.number_input(
                "Approved % (type 10 for 10%)",
                step=0.5,
                key="levy_approved_pct",
                help="10 means last year’s levies plus 10%.",
            )
        elif _mode == "amount":
            st.number_input(
                "Approved amount for the year",
                min_value=0.0,
                step=1000.0,
                key="levy_approved_amount",
            )
        _approved_now = approved_ordinary(st.session_state)
        _gap_now = _costs_now - _approved_now
        st.write(
            f"Costs need **{money(_costs_now)}** ({money(_costs_now / 12)} a month). "
            f"Approved **{money(_approved_now)}** ({money(_approved_now / 12)} a month)."
        )
        if _actual_now > 0.5 and _mode == "pct":
            st.caption(f"Last year’s levies on the line: {money(_actual_now)}.")
        if _gap_now > 1:
            st.warning(
                f"Gap {money(_gap_now)}. The costs are still higher than the approved levy. "
                "Cut a cost, or the meeting must take this from the reserve."
            )
        elif _gap_now < -1:
            st.info(f"The approved levy is {money(-_gap_now)} more than the costs. That extra stays in the funds.")
        section_form("levy", "Levy Income", "Ordinary, Reserve and CSOS. The ordinary levy line shows what the costs need. The meeting decision above is what the PQ uses. CSOS income = CSOS expense. Add a row if you need another levy, then Save.")
        st.session_state.has_master_hoa = st.checkbox(
            "We still bill an estate levy to owners (they pay it through us).",
            value=bool(st.session_state.get("has_master_hoa")),
            help="Off = owners pay that estate themselves. On = extra columns on our owner schedule.",
        )
        if st.session_state.has_master_hoa:
            st.divider()
            section_form(
                "hoa_income",
                "Estate / HOA recovered from owners",
                "Billed to owners on the PQ as extra columns. Not part of ordinary levies.",
            )
            st.divider()
            section_form(
                "hoa_expense",
                "Estate / HOA paid to the estate",
                "What we pay the master HOA. Not included in ordinary levies.",
            )
        else:
            st.caption("Estate lines are hidden. Owners pay that estate directly — it does not go through this budget.")
        st.divider()
        section_form(
            "other",
            "Other Income",
            "Rental, interest, penalties and the like. Tick Lower the levies only on the lines that should bring this complex’s levy down. Leave the tick off if that income is not used.",
            income=True,
        )
        st.divider()
        section_form(
            "recoveries_other",
            "Recovery income",
            "Legal fees recovered, an insurance recovery, or a repair invoiced to an owner. "
            "The cost stays. This amount comes off the cost you pick. "
            "Type the amount, pick Comes off this cost, then Save.",
            claims=True,
        )
        st.divider()
        section_form(
            "fixed",
            "Fixed monthly charges on the owner invoice",
            "Equal extra on the invoice (same rand each unit): prepaid / Eskom fixed / communal. "
            "Type the YEARLY total for the complex, then Save. Each owner pays total ÷ 12 ÷ units. "
            "Do **not** put garden here. If owners pay the gardener themselves, set **Garden service** on Expenditure to **R0** (omit it). "
            "**Garden Expenses** on Repair & Maintenance stays in the levy (general garden repairs). "
            "Do not put insurance here.",
        )

    with tabs[2]:
        ag, ar = municipal_split(st.session_state, "actual")
        yg, yr = municipal_split(st.session_state, "yearly")
        st.markdown("**Last year — add the Actual column like this**")
        st.caption("City-bill actuals are added. Lines named recovered are subtracted. Do not add those actuals on top.")
        a1, a2, a3 = st.columns(3)
        a1.metric("Actual city bill", money(ag))
        a2.metric("Actual recovered", money(ar))
        a3.metric("Actual net", money(ag - ar))
        for r in st.session_state.sections.get("municipal") or []:
            act = abs(float(r.get("actual") or 0))
            if act < 0.5:
                continue
            kind = "subtract" if is_muni_recovery(r.get("desc") or "", r.get("is_recovery")) else "add"
            st.write(f"- {r.get('desc')}: Actual {money(act)} — **{kind}**")
        st.write(f"**{money(ag)} − {money(ar)} = {money(ag - ar)}**")
        st.markdown("**This year — Budgeted yearly**")
        n1, n2, n3 = st.columns(3)
        n1.metric("Budget city bill", money(yg))
        n2.metric("Budget recovered", money(yr))
        n3.metric("Budget net (in the levy)", money(yg - yr))
        g, rec = yg, yr
        gaps = municipal_gaps(st.session_state)
        under = [x for x in gaps if x["gap"] > 1]
        over = [x for x in gaps if x["gap"] < -1]
        if under:
            total_under = sum(x["gap"] for x in under)
            st.warning(
                f"**Under-recovery {money(total_under)}** — the city bill is more than owners paid back. "
                f"That shortfall is **already in ordinary levies**. Trustees should attend: meters, billing, or raise recoveries."
            )
            for x in under:
                st.write(
                    f"- **{x['name']}**: city {money(x['gross'])} − recovered {money(x['rec'])} = **under {money(x['gap'])}**"
                )
        if over:
            st.info(
                "Over-recovery (owners paid back more than the city bill). That credit reduces the levy."
            )
            for x in over:
                st.write(
                    f"- **{x['name']}**: recovered {money(x['rec'])} − city {money(x['gross'])} = **over {money(-x['gap'])}**"
                )
        if not under and not over and (g > 1 or rec > 1):
            st.success("Municipal recoveries match the city bill. Nothing extra is sitting in the levy.")
        gaps = municipal_gaps(st.session_state)
        st.caption(
            f"City bill {money(g)} − recovered {money(rec)} = **net {money(g - rec)}**. "
            "Every line on this tab is a positive rand. Monthly = Budgeted yearly ÷ 12. "
            "We subtract the lines named recovered. Do not add them to the city bill."
        )
        missing_gross = [
            x["name"] for x in gaps
            if x["rec"] > 1 and x["gross"] < 1
        ]
        if missing_gross:
            st.warning(
                "Recovered with no city-bill line: "
                + ", ".join(missing_gross)
                + ". Add the city bill (for example Water) or those recoveries reduce the levy on their own."
            )
        odd_rec = []
        for r in st.session_state.sections.get("municipal") or []:
            act = abs(float(r.get("actual") or 0))
            y = abs(float(r.get("yearly") or 0))
            if act > 1000 and y < act * 0.25:
                odd_rec.append(f"{r.get('desc')}: last year {money(act)}, this year {money(y)}")
        if odd_rec:
            st.warning(
                "These municipal lines are far below last year. If that was not on purpose, set % to 0 and Save: "
                + "; ".join(odd_rec)
            )
        section_form(
            "municipal",
            "Municipal charges",
            "City bill on its own line. Any line with the word recovered is subtracted, including a sewer-plant rental. Type a positive rand, then Save.",
        )

    with tabs[3]:
        section_form(
            "expenditure",
            "Expenditure",
            "Operating costs except R&M, personnel and tax. "
            "Budgeted yearly is the full bill. Recovered from some owners is what only some owners pay towards that same bill. "
            "It is not income and not a levy column. The levy carries the bill minus that recovery. "
            "CSOS Levies (Expense) is locked to CSOS income — same rand. "
            "If owners pay the gardener themselves, set Garden service to R0.",
            recover=True,
        )

    with tabs[4]:
        over_pay = [
            it for it in (st.session_state.sections.get("rm") or [])
            if float(it.get("insurance") or 0) > float(it.get("yearly") or 0) + 0.5
            and float(it.get("insurance") or 0) > 0.5
        ]
        if over_pay:
            st.error(
                "Insurance payout is bigger than Budgeted yearly on: "
                + ", ".join(it["desc"] for it in over_pay)
                + ". That column is a **deduction**, not the new budget. "
                "Put the fire / equipment amount in **Budgeted yearly**. Net in the levy is R0 on those lines."
            )
        both = []
        taken = claims_against(st.session_state)
        for it in st.session_state.sections.get("rm") or []:
            if float(it.get("insurance") or 0) > 0.5 and taken.get(norm(it.get("desc") or ""), 0) > 0.5:
                both.append(it.get("desc") or "")
        if both:
            st.warning(
                "These repairs have an insurance payout and recovery income: "
                + ", ".join(both)
                + ". Use only one, or the levy is reduced twice."
            )
        section_form(
            "rm",
            "Repair and Maintenance",
            "Budgeted yearly is the full job. Insurance payout is what the insurer pays. "
            "Recovered from some owners is what only some owners pay towards that same job. Neither is income. "
            "The levy carries the job minus those two, never below R0.",
            rm=True,
            recover=True,
        )

    with tabs[5]:
        section_form("personnel", "Personnel", "Salaries, casuals, PAYE/UIF, bonuses.")

    with tabs[6]:
        other_inc, tax_est = estimate_income_tax(st.session_state)
        st.subheader("Income tax")
        st.markdown(
            f"""
**Levies are not taxed.** Interest, investment income and rent can be.

Estimate from Other Income on the Income tab: **{money(other_inc)}**.
The first **R50,000** of that income is exempt. The rest × **27%** = **{money(tax_est)}**.

This is a budget estimate (section 10(1)(e)), not a SARS assessment. If the auditor has a different figure, type that instead.
            """
        )
        if st.button("Put this estimate on the tax line"):
            found = False
            for r in st.session_state.sections["tax"]:
                if "tax" in (r.get("desc") or "").lower():
                    r["yearly"] = tax_est
                    a = float(r.get("actual") or 0)
                    r["pct"] = 0.0 if a == 0 else (tax_est / a) * 100 - 100
                    found = True
            if not found:
                rec = row("Taxation Payable", "Estimate: other income above R50,000 × 27%.")
                rec["yearly"] = tax_est
                st.session_state.sections["tax"].append(rec)
            apply_levy_lines(st.session_state)
            st.success(f"Tax line set to {money(tax_est)}. It is included in ordinary levies.")
            st.rerun()
        st.markdown(
            """
**When it is paid**

- Put **one yearly amount** in this budget. Owners do not pay SARS every month. The monthly column is only yearly ÷ 12, so the levy can fund it.
- If the estimate is **R0** (other income under R50,000), nothing is paid to SARS.
- If tax is payable, SARS usually wants **provisional tax**: one payment **six months** into the financial year, and one at **year-end**. A top-up can be due a few months after year-end if the estimate was short.
- Example for a **February** year-end: 31 August, 28 February, and a top-up by 30 September if needed.
- The tax return is filed **after** year-end, not when you finish this budget.

Confirm the dates and the amount with the auditor or tax practitioner.
            """
        )
        section_form("tax", "Income Tax", "Use the estimate button, or type the auditor’s yearly figure, then Save.")

    with tabs[7]:
        section_form("special", "Special Projects", "Year 1 of the 10-year plan. Tick the sidebar box only if levies must fund it.")

    with tabs[8]:
        st.subheader("PQ / levy schedule")
        st.caption("Choose the unit ratio Excel here. WeConnectU’s PQ column is often 0 — we use Ratio 1. The ratios must add to 1.000.")
        pq_here = st.file_uploader(
            "Upload unit ratios",
            type=["xlsx", "xls", "xlsm", "csv"],
            key="pq_tab_file",
        )
        if pq_here is not None:
            sig = f"{pq_here.name}:{getattr(pq_here, 'size', 0)}"
            if st.session_state.get("_pq_sig") != sig:
                try:
                    records, msg = parse_pq_upload(pq_here)
                    st.session_state.pq = records
                    st.session_state["_pq_sig"] = sig
                    st.session_state.msg = msg
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not read this ratio file: {e}")
        if not st.session_state.pq:
            st.info("No PQs loaded yet. On the left, choose the unit PQs Excel and click Load PQs.")
        else:
            prev = pd.DataFrame(st.session_state.pq)
            if "Unit" not in prev.columns:
                prev = prev.rename(columns={prev.columns[0]: "Unit"})
            bills = pq_bill_lines(st.session_state)
            extra_cols = []
            n_units = max(len(prev), 1)
            for name, yearly, split in bills:
                col = str(name)
                if col in ("Unit", "PQ"):
                    col = f"{name} levy"
                monthly = float(yearly or 0) / 12
                if split == "equal":
                    prev[col] = monthly / n_units
                else:
                    prev[col] = prev["PQ"].astype(float) * monthly
                extra_cols.append(col)
            if extra_cols:
                prev["Total monthly"] = prev[extra_cols].sum(axis=1)
            show = prev.copy()
            if show.columns.duplicated().any():
                show.columns = _dedupe_headers([str(c) for c in show.columns])
            if "PQ" in show.columns:
                show["PQ"] = show["PQ"].astype(float).round(6)
            for c in extra_cols + (["Total monthly"] if extra_cols else []):
                show[c] = show[c].astype(float).round(2)
            st.dataframe(show, use_container_width=True, hide_index=True)
            levy_sum = sum(float(y or 0) for _, y, _s in bills)
            if levy_sum < 1:
                st.warning(
                    "PQ shares are loaded, but levy rands are still 0. "
                    "Load the financial statement and WeConnectU Excel on the left, "
                    "then Save the cost tabs so ordinary / reserve / CSOS fill in."
                )
            else:
                st.caption("Monthly column totals: " + " · ".join(f"{n} {money(y/12)}" for n, y, _s in bills))
            st.caption(f"{len(prev)} units. PQ total {float(prev['PQ'].sum()):.6f} (should be about 1.000).")

    with tabs[9]:
        st.subheader("10-year maintenance plan")
        st.caption(
            f"We are in {NOW_YEAR}. Year 1 of this budget is {NOW_YEAR}, Year 2 is {NOW_YEAR+1}, … Year 10 is {NOW_YEAR+9}. "
            "If you upload last year’s plan (2025–2034), we roll it: last year’s 2026 column becomes this Year 1."
        )
        ymp_up = st.file_uploader("Upload 10-year plan Excel (or last year’s budget workbook)", type=["xlsx", "xls"], key="ymp_xlsx")
        if ymp_up and st.button("Load 10-year plan"):
            try:
                xl = pd.ExcelFile(ymp_up)
                sheet = next((n for n in xl.sheet_names if re.search(r"ymp|10|maintenance", n, re.I)), xl.sheet_names[-1])
                parsed = parse_ymp_sheet(pd.read_excel(xl, sheet_name=sheet, header=None), plan_start_year(st.session_state))
                if parsed:
                    st.session_state.ymp = parsed
                    y0 = plan_start_year(st.session_state)
                    st.success(f"Loaded {len(parsed)} projects. Year 1 is {y0} (this budget year), not last year’s 2025 column.")
                else:
                    st.error("Could not find Year 1–10 amounts on that sheet.")
            except Exception as e:
                st.error(f"Could not read 10-year plan: {e}")
        paste = st.text_area("Or paste from Excel (keep the empty cells / tabs)", height=160)
        if st.button("Paste into plan") and paste.strip():
            parsed = parse_ymp_paste(paste, plan_start_year(st.session_state))
            if parsed:
                st.session_state.ymp = parsed
                y0 = plan_start_year(st.session_state)
                st.success(
                    f"Loaded {len(parsed)} projects. Year 1 is {y0} (this budget). "
                    f"Last year’s 2025 column is dropped. Painting / garage rands sit in {y0} onward."
                )
            else:
                st.error("Could not read that paste. Copy the whole 10-year block from Excel, including the 2025–2034 year headings.")
        ymp_rows = []
        for p in st.session_state.ymp:
            rec = {
                "Project": p.get("desc") or "",
                "First cycle": p.get("first_cycle") or "",
                "Frequency": p.get("freq") or "",
                "Current estimate": float(p.get("estimate") or 0),
            }
            years = p.get("years") or [0] * 10
            for i in range(10):
                rec[f"Year {i+1}"] = float(years[i] if i < len(years) else 0)
            ymp_rows.append(rec)
        with st.form("ymp_form"):
            ed = st.data_editor(pd.DataFrame(ymp_rows), num_rows="dynamic", use_container_width=True, hide_index=True)
            if st.form_submit_button("Save 10-year plan"):
                st.session_state.ymp = [{
                    "desc": str(r.get("Project") or ""),
                    "first_cycle": r.get("First cycle") or "",
                    "freq": r.get("Frequency") or "",
                    "estimate": float(r.get("Current estimate") or 0),
                    "years": [float(r.get(f"Year {i+1}") or 0) for i in range(10)],
                } for _, r in ed.iterrows()]
                st.success("Saved.")
        if st.button("Copy Year 1 into Special Projects"):
            spec = []
            for p in st.session_state.ymp:
                y1 = float((p.get("years") or [0])[0] or 0)
                if y1:
                    rec = row(p["desc"], "From 10-year plan Year 1")
                    rec["yearly"] = y1
                    spec.append(rec)
            if spec:
                st.session_state.sections["special"] = spec
                st.success(f"Copied {len(spec)} projects (Year 1 rands only).")

    with tabs[10]:
        st.subheader("For the owners")
        st.caption(
            "Open the presentation in Chrome. Arrow keys move on. F is full screen. "
            "It explains the levy in plain words, with this complex’s own numbers, so owners can see what they are paying for."
        )
        if not st.session_state.complex_name:
            st.warning("Type the complex name in the sidebar first.")
        elif st.button("Build the presentation", type="primary"):
            st.session_state["meeting_html"] = generate_meeting_html(st.session_state)
            deck = generate_pptx(st.session_state)
            st.session_state["pptx"] = deck.getvalue()
        if st.session_state.get("meeting_html"):
            name = re.sub(r"\s+", "_", st.session_state.complex_name or "budget")
            st.download_button(
                "Download the presentation",
                data=st.session_state["meeting_html"],
                file_name=f"Budget_presentation_{name}.html",
                mime="text/html",
            )
            st.caption("Open that file in Chrome. Arrow keys move on. Press F for full screen.")
        if st.session_state.get("pptx"):
            name = re.sub(r"\s+", "_", st.session_state.complex_name or "budget")
            st.download_button(
                "Download the PowerPoint — same slides",
                data=st.session_state["pptx"],
                file_name=f"Budget_presentation_{name}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
            st.caption("Open it in PowerPoint and press F5. The logo fills the first screen, then zooms out onto the budget. The last page is unchanged. It only moves in the slideshow, not while you scroll in edit view.")

    with tabs[11]:
        st.subheader("Download Excel")
        st.caption("Budget + PQ + 10-year plan. Yellow cells are inputs.")
        if not st.session_state.complex_name:
            st.warning("Type the complex name in the sidebar first.")
        elif st.button("Build Excel file", type="primary"):
            xls = generate_excel(st.session_state)
            st.session_state["xlsx"] = xls.getvalue()
        if st.session_state.get("xlsx"):
            name = re.sub(r"\s+", "_", st.session_state.complex_name)
            st.download_button(
                "Download budget Excel",
                data=st.session_state["xlsx"],
                file_name=f"Budget_{name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


if __name__ == "__main__":
    main()
