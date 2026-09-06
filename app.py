"""Domus Property Management – Body Corporate / HOA Budget (Streamlit)."""

from __future__ import annotations

import re
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Domus Property Management Budget",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

YELLOW = "FFFF99"
NAVY = "1F4E79"
BLUE = "0000FF"
RED = "FFC7CE"
GREEN = "C6EFCE"
TOTAL = "D9E2F3"
SECTION = "2E75B6"
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
MONEY = '#,##0.00;(#,##0.00);"-"'


def uid() -> str:
    import random
    import string
    return "".join(random.choices(string.ascii_lowercase + string.digits, k=8))


def norm(s: str) -> str:
    s = str(s or "").lower()
    s = re.sub(r"[–—-]", " ", s)
    s = re.sub(r"[^a-z0-9/& ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def family(desc: str) -> str | None:
    d = norm(desc)
    if re.search(r"insurance\s*claim", d):
        return "insurance_claims"
    if "eskom" in d:
        return "eskom"
    if "reserve" in d:
        return "reserve"
    if "csos" in d:
        if "collect" in d:
            return "csos_collection"
        if re.search(r"contrib|admin|expense", d):
            return "csos_expense"
        return "csos_income"
    if re.search(r"ordinary\s*lev", d) or d in ("levies", "levy") or re.search(r"levies?\s*received", d):
        return "ordinary"
    if "electricity" in d and "recover" in d and "commun" in d:
        return "elec_rec_communal"
    if "electricity" in d and "recover" in d:
        return "elec_rec"
    if "water" in d and "recover" in d:
        return "water_rec"
    if "sewer" in d and "recover" in d:
        return "sewer_rec"
    if "refuse" in d and "recover" in d:
        return "refuse_rec"
    if "interest" in d and "arrear" in d:
        return "interest_arrears"
    if re.search(r"interest|investment|marketlink", d) and re.search(r"bank|invest|marketlink|earn", d):
        return "investment"
    if "penalty" in d:
        return "penalty"
    if "electricity" in d and "recover" not in d:
        return "elec_gross"
    if re.search(r"^water$|water\s*(charge|expense|municipal)", d):
        return "water_gross"
    if "sewer" in d and "recover" not in d:
        return "sewer_gross"
    if "refuse" in d and "recover" not in d:
        return "refuse_gross"
    if "management" in d and "fee" in d:
        return "management"
    if re.search(r"^insurance$|insurance\s*(premium|expense)", d):
        return "insurance"
    if re.search(r"security|guarding", d):
        return "security"
    if re.search(r"salar|staff wages", d):
        return "salaries"
    return None


def section_for_name(desc: str) -> str:
    d = norm(desc)
    f = family(desc)
    if f in ("ordinary", "reserve", "csos_income"):
        return "levy_income"
    if f == "csos_expense":
        return "expenditure"
    if f in ("csos_collection", "eskom", "interest_arrears", "investment", "penalty"):
        return "other_income"
    if f in ("elec_rec", "elec_rec_communal", "water_rec", "sewer_rec", "refuse_rec"):
        return "muni_recoveries"
    if f in ("elec_gross", "water_gross", "sewer_gross", "refuse_gross"):
        return "municipal"
    if f == "salaries":
        return "personnel"
    if re.search(r"\b(repair|maintenance|plumb|paint|roof|gutter|pool|electrical|fire equipment|gate)\b", d):
        return "rm"
    if re.search(r"\b(wages|salary|salaries|paye|uif|bonus|overtime|casual|relief|wca|coida|staff)\b", d):
        return "personnel"
    if re.search(r"\b(income tax|taxation)\b", d):
        return "tax"
    if re.search(r"\b(special project|improvement|jungle gym)\b", d):
        return "special"
    if re.search(r"\b(recover|rental|interest|penalty|eskom|other income)\b", d):
        return "other_income"
    return "expenditure"


def item(desc: str, pct: float = 10, note: str = "", computed: bool = False, mode: str = "pct") -> dict:
    return {
        "id": uid(),
        "desc": desc,
        "actual": 0.0,
        "budgeted": 0.0,
        "pct": float(pct),
        "mode": mode,
        "note": note,
        "computed": computed,
    }


def default_sections() -> dict:
    return {
        "levy_income": [
            item("Ordinary Levies (Gross required)", 0, "Calculated automatically.", computed=True, mode="amount"),
            item("Reserve Fund Contribution", 0, "Type the yearly rand amount. Billed separately from ordinary levies.", mode="amount"),
            item("CSOS Levy (Income)", 0, "What owners are billed for CSOS.", mode="amount"),
        ],
        "other_income": [
            item("Interest on Arrears", 0),
            item("Investment Income – Bank/Investments", 0),
            item("Penalty Levy", 0),
            item("Eskom fixed charge", 0, "Only if this complex bills a fixed Eskom charge."),
            item("Other Income 1", 0),
            item("Other Income 2", 0),
        ],
        "muni_recoveries": [
            item("Electricity Recovered", 0, "Recovered from units / owners."),
            item("Electricity Recovered – Communal", 0, "Common-area electricity recovered."),
            item("Water Recovered", 0),
            item("Sewerage Recovered", 0),
            item("Refuse Recovered", 0),
            item("Other Recoveries", 0),
        ],
        "municipal": [
            item("Water", 10),
            item("Sewerage", 5),
            item("Electricity (Common / Gross)", 10),
            item("Refuse", 10),
            item("Rates / Property Tax", 5),
            item("Other Municipal", 10),
        ],
        "expenditure": [
            item("Accounting Fees", 10),
            item("Audit Fees", 10),
            item("Bank Charges", 10),
            item("Insurance", 10),
            item("CSOS Levies (Expense)", 0),
            item("Management Fee", 10),
            item("Legal & Professional Fees", 10),
            item("Meeting / Venue Expenses", 10),
            item("Telephone / Communications", 10),
            item("Office / General Expenses", 10),
            item("Property Valuation", 0),
            item("Security / Guarding", 10),
            item("Garden Service (contract)", 10),
            item("Cleaning & Materials", 10),
            item("Health & Safety", 10),
            item("Keys & Remotes", 10),
            item("Computer / IT Expenses", 10),
            item("Printing & Stationery", 10),
            item("Other Expenditure 1", 10),
            item("Other Expenditure 2", 10),
        ],
        "rm": [
            item("Electrical", 5),
            item("Fire Equipment", 5),
            item("General / Buildings Maintenance", 0),
            item("Garden Expenses", 5),
            item("Gate & Intercom", 5),
            item("Plumbing", 5),
            item("Painting / Waterproofing", 5),
            item("Roofs & Gutters", 5),
            item("Pool", 5),
            item("Electric Fence", 5),
            item("CCTV / Cameras", 5),
            item("Paving / Roadways", 5),
            item("Lifts", 5),
            item("Other R&M 1", 5),
            item("Other R&M 2", 5),
        ],
        "personnel": [
            item("Salaries & Wages", 10),
            item("Casual / Relief Wages", 10),
            item("PAYE / UIF / SDL Contributions", 10),
            item("Travel / Allowances", 10),
            item("Bonuses & Overtime", 10),
            item("Staff Welfare / Protective Clothing", 10),
            item("Pension / Provident Fund", 10),
            item("WCA / COIDA", 10),
            item("Other Personnel Costs", 10),
        ],
        "tax": [
            item("Taxation Payable", 0, "Usually based on taxable investment income."),
        ],
        "special": [
            item("Special Project 1 (from 10YMP)", 0, mode="amount"),
            item("Special Project 2", 0, mode="amount"),
            item("Special Project 3", 0, mode="amount"),
            item("Special Project 4", 0, mode="amount"),
            item("Special Project 5", 0, mode="amount"),
        ],
    }


def yearly(it: dict) -> float:
    if it.get("computed"):
        return float(it.get("budgeted") or 0)
    if it.get("mode") == "amount":
        return float(it.get("budgeted") or 0)
    return float(it.get("actual") or 0) * (1 + float(it.get("pct") or 0) / 100)


def implied_pct(it: dict) -> float:
    y = yearly(it)
    a = float(it.get("actual") or 0)
    if a == 0:
        return 0.0
    return (y / a - 1) * 100


def sum_yearly(items: list) -> float:
    return sum(yearly(i) for i in items)


def compute_totals(state: dict) -> dict:
    sec = state["sections"]
    municipal = sum_yearly(sec["municipal"])
    expenditure = sum_yearly(sec["expenditure"])
    rm_gross = sum_yearly(sec["rm"])
    rm_net = rm_gross - float(state.get("insurance_recoveries") or 0)
    personnel = sum_yearly(sec["personnel"])
    tax = sum_yearly(sec["tax"])
    special = sum_yearly(sec["special"])
    other = sum_yearly(sec["other_income"])
    recoveries = sum_yearly(sec["muni_recoveries"])
    reserve_item = next((i for i in sec["levy_income"] if "reserve" in i["desc"].lower()), None)
    csos_item = next((i for i in sec["levy_income"] if "csos" in i["desc"].lower()), None)
    reserve = yearly(reserve_item) if reserve_item else 0
    csos = yearly(csos_item) if csos_item else 0
    expenses = municipal + expenditure + rm_net + personnel + tax + special
    csos_exp = sum(yearly(i) for i in sec["expenditure"] if "csos" in i["desc"].lower())
    to_collect = expenses - csos_exp - other - recoveries
    rate = float(state.get("collection_rate") or 0) / 100
    gross = to_collect / rate if rate else 0
    return {
        "municipal": municipal,
        "expenditure": expenditure,
        "rm_net": rm_net,
        "personnel": personnel,
        "tax": tax,
        "special": special,
        "other": other,
        "recoveries": recoveries,
        "reserve": reserve,
        "csos": csos,
        "expenses": expenses,
        "to_collect": to_collect,
        "gross_ordinary": gross,
        "monthly_ordinary": gross / 12,
        "monthly_reserve": reserve / 12,
        "monthly_csos": csos / 12,
        "collection_rate": float(state.get("collection_rate") or 0),
    }


def apply_ordinary(state: dict) -> None:
    tot = compute_totals(state)
    for it in state["sections"]["levy_income"]:
        if it.get("computed") or "ordinary" in it["desc"].lower():
            it["computed"] = True
            it["mode"] = "amount"
            it["budgeted"] = tot["gross_ordinary"]
            it["pct"] = implied_pct(it)


def num(v) -> float | None:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        n = float(v)
        return 0.0 if abs(n) < 0.01 else n
    s = str(v).strip()
    if not s or s in ("-", "–"):
        return 0.0
    s2 = re.sub(r"[R$\s,]", "", s)
    s2 = s2.replace("(", "-").replace(")", "")
    try:
        n = float(s2)
    except ValueError:
        return None
    return 0.0 if abs(n) < 0.01 else n


def extract_weconnectu(uploaded) -> list:
    xl = pd.ExcelFile(uploaded)
    rows = []
    for sheet in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet, header=None)
        if df.empty or df.shape[1] < 4:
            continue
        header_row = None
        for i in range(min(15, len(df))):
            vals = [str(v).strip().lower() for v in df.iloc[i].tolist()]
            if "actual" in vals and any("budget" in v for v in vals):
                header_row = i
                break
        if header_row is None:
            continue
        prev = [str(v).strip().lower() for v in df.iloc[header_row - 1].tolist()] if header_row else [""] * df.shape[1]
        cur = [str(v).strip().lower() for v in df.iloc[header_row].tolist()]
        merged = [(prev[i] if i < len(prev) else "") + " " + (cur[i] if i < len(cur) else "") for i in range(df.shape[1])]
        ytd_actual = next((i for i, t in enumerate(merged) if "ytd" in t and "actual" in t and "var" not in t), None)
        total_budget = next((i for i, t in enumerate(merged) if "total" in t and "budget" in t), None)
        if ytd_actual is None:
            ytd_actual = next((i for i, t in enumerate(cur) if t == "actual"), None)
        if ytd_actual is None:
            continue
        for r in range(header_row + 1, len(df)):
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
                gl, desc = "", raw
                if len(desc) < 3:
                    continue
            actual = num(df.iloc[r, ytd_actual]) or 0.0
            budget = num(df.iloc[r, total_budget]) if total_budget is not None else None
            if abs(actual) < 0.5 and (budget is None or abs(budget) < 0.5):
                continue
            rows.append({"desc": desc, "gl": gl, "actual": abs(actual), "budget": abs(budget) if budget is not None else None})
    return rows


def match_rows(extracted: list, sections: dict) -> tuple[dict, float, int]:
    next_sec = {k: [dict(x) for x in v] for k, v in sections.items()}
    used = set()
    insurance = 0.0
    added = 0
    all_items = [(k, it) for k, items in next_sec.items() for it in items]

    unmatched = []
    for row in extracted:
        fam = family(row["desc"])
        if fam == "insurance_claims":
            insurance += row["actual"]
            continue
        best = None
        best_s = 0.0
        for key, it in all_items:
            mark = f"{key}:{it['id']}"
            if mark in used:
                continue
            item_fam = family(it["desc"])
            if fam and item_fam and fam != item_fam:
                continue
            n_desc, n_item = norm(row["desc"]), norm(it["desc"])
            score = 0.0
            if fam and item_fam and fam == item_fam:
                score = 0.96
            if n_desc == n_item:
                score = 1.0
            elif n_item in n_desc or n_desc in n_item:
                score = max(score, 0.86)
            else:
                stop = {"levy", "levies", "income", "other", "general", "expense", "expenses", "fee", "fees", "and", "the"}
                a = {w for w in n_desc.split() if len(w) > 3 and w not in stop}
                b = {w for w in n_item.split() if len(w) > 3 and w not in stop}
                if a and b:
                    score = max(score, len(a & b) / max(len(a), len(b)))
            if score > 0.55 and score > best_s:
                best_s, best = score, (key, it)
        if best:
            key, it = best
            used.add(f"{key}:{it['id']}")
            it["actual"] = row["actual"]
            if row.get("budget") is not None:
                it["budgeted"] = row["budget"]
                it["mode"] = "amount"
                it["pct"] = implied_pct(it)
            else:
                it["budgeted"] = yearly(it)
        else:
            unmatched.append(row)

    for row in unmatched:
        fam = family(row["desc"])
        if fam == "csos_collection":
            continue
        if fam == "csos_income":
            existing_inc = next((i for i in next_sec["levy_income"] if family(i["desc"]) == "csos_income" and i["actual"] > 0), None)
            if existing_inc:
                fam = "csos_expense"
        section = "expenditure" if fam == "csos_expense" else section_for_name(row["desc"])
        if fam:
            existing = next((i for i in next_sec[section] if family(i["desc"]) == fam), None)
            if existing:
                existing["actual"] = float(existing["actual"] or 0) + row["actual"]
                if row.get("budget") is not None:
                    existing["budgeted"] = row["budget"]
                    existing["mode"] = "amount"
                continue
        next_sec[section].append(item(row["desc"], 0, "Added from WeConnectU — unique to this complex", mode="amount" if row.get("budget") is not None else "pct"))
        nxt = next_sec[section][-1]
        nxt["actual"] = row["actual"]
        nxt["budgeted"] = row["budget"] if row.get("budget") is not None else row["actual"]
        added += 1
    return next_sec, insurance, added


def parse_ymp_paste(text: str) -> list:
    out = []
    for line in text.splitlines():
        line = line.strip()
        if not line or re.match(r"^(project|description|item)", line, re.I):
            continue
        parts = [p.strip() for p in re.split(r"\t|;|,|\s{2,}", line) if p.strip()]
        if not parts:
            continue
        desc = parts[0].strip("\"'")
        if re.match(r"^total", desc, re.I):
            continue
        years = [0.0] * 10
        yi = 0
        for p in parts[1:]:
            if yi >= 10:
                break
            n = num(p)
            if n is None:
                continue
            years[yi] = n
            yi += 1
        out.append({"desc": desc, "years": years})
    return out


def generate_excel(state: dict, pq_df: pd.DataFrame | None, ymp: list) -> BytesIO:
    apply_ordinary(state)
    tot = compute_totals(state)
    sec = state["sections"]
    wb = Workbook()
    ws = wb.active
    ws.title = "BUDGET"
    widths = [3, 42, 14, 12, 16, 14, 44]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    def fill(cell, color):
        cell.fill = PatternFill("solid", fgColor=color)

    def input_cell(cell, value, fmt=None):
        cell.value = value
        fill(cell, YELLOW)
        cell.font = Font(name="Calibri", color=BLUE, size=10)
        cell.border = THIN
        if fmt:
            cell.num_fmt = fmt

    def formula(cell, f, bg=None):
        cell.value = f"={f}"
        cell.font = Font(name="Calibri", size=10)
        cell.border = THIN
        cell.num_fmt = MONEY
        if bg:
            fill(cell, bg)

    ws.merge_cells("B2:G2")
    ws["B2"] = "BODY CORPORATE / HOA BUDGET"
    ws["B2"].font = Font(name="Calibri", bold=True, size=16, color=NAVY)
    ws["B3"] = "Complex Name:"
    input_cell(ws["C3"], state["complex_name"] or "BODY CORPORATE")
    ws["E3"] = "Financial Year:"
    input_cell(ws["F3"], state["fin_year"])
    ws["B6"] = "Expected Levy Collection Rate %"
    input_cell(ws["C6"], state["collection_rate"] / 100, "0.0%")
    ws["B7"] = "Opening Reserve Fund Balance"
    input_cell(ws["C7"], state["reserve_opening"], MONEY)

    row = 10

    def header(r):
        labels = ["Description", "Actual", "% Increase", "Budgeted Yearly", "Monthly", "Notes"]
        for i, lab in enumerate(labels, 2):
            c = ws.cell(r, i, lab)
            fill(c, NAVY)
            c.font = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
            c.alignment = Alignment(horizontal="center")
            c.border = THIN

    def bar(r, title):
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        c = ws.cell(r, 2, title)
        fill(c, SECTION)
        c.font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
        for col in range(2, 8):
            fill(ws.cell(r, col), SECTION)
            ws.cell(r, col).border = THIN

    def write_items(start, items, computed_idx=None):
        for i, it in enumerate(items):
            r = start + i
            input_cell(ws.cell(r, 2), it["desc"])
            input_cell(ws.cell(r, 3), float(it.get("actual") or 0), MONEY)
            pct = implied_pct(it) / 100
            input_cell(ws.cell(r, 4), pct, "0.0%")
            if it.get("computed") or (computed_idx is not None and i == computed_idx):
                formula(ws.cell(r, 5), f"C{r}*(1+D{r})" if not it.get("computed") else f"C{r}*(1+D{r})", RED)
                # overwritten later for ordinary
                input_cell(ws.cell(r, 5), yearly(it), MONEY)
                fill(ws.cell(r, 5), RED)
                ws.cell(r, 5).font = Font(name="Calibri", bold=True, size=10)
            elif it.get("mode") == "amount":
                input_cell(ws.cell(r, 5), yearly(it), MONEY)
            else:
                formula(ws.cell(r, 5), f"C{r}*(1+D{r})")
            formula(ws.cell(r, 6), f"E{r}/12")
            input_cell(ws.cell(r, 7), it.get("note") or "")
        return start + len(items) - 1

    def total_row(r, label, start, end):
        ws.cell(r, 2, label).font = Font(bold=True)
        formula(ws.cell(r, 5), f"SUM(E{start}:E{end})", TOTAL)
        formula(ws.cell(r, 6), f"E{r}/12", TOTAL)

    bar(row, "INCOME")
    row += 1
    header(row)
    row += 1
    levy_start = row
    end = write_items(row, sec["levy_income"], 0)
    ws.cell(levy_start, 5).value = tot["gross_ordinary"]
    fill(ws.cell(levy_start, 5), RED)
    row = end + 2
    other_start = row
    end = write_items(row, sec["other_income"])
    other_end = end
    row = end + 2
    rec_start = row
    end = write_items(row, sec["muni_recoveries"])
    rec_end = end
    row = end + 3

    bar(row, "MUNICIPAL EXPENSES (GROSS)")
    row += 1
    header(row)
    row += 1
    muni_start = row
    end = write_items(row, sec["municipal"])
    row = end + 1
    muni_tot = row
    total_row(row, "TOTAL MUNICIPAL", muni_start, end)
    row += 3

    bar(row, "EXPENDITURE")
    row += 1
    header(row)
    row += 1
    exp_start = row
    end = write_items(row, sec["expenditure"])
    row = end + 1
    exp_tot = row
    total_row(row, "TOTAL EXPENDITURE", exp_start, end)
    row += 3

    bar(row, "REPAIR AND MAINTENANCE")
    row += 1
    header(row)
    row += 1
    rm_start = row
    end = write_items(row, sec["rm"])
    row = end + 1
    ws.cell(row, 2, "Less: Insurance recoveries")
    input_cell(ws.cell(row, 5), float(state.get("insurance_recoveries") or 0), MONEY)
    ins_row = row
    row += 1
    rm_tot = row
    ws.cell(row, 2, "NET R&M").font = Font(bold=True)
    formula(ws.cell(row, 5), f"SUM(E{rm_start}:E{end})-E{ins_row}", TOTAL)
    formula(ws.cell(row, 6), f"E{row}/12", TOTAL)
    row += 3

    bar(row, "PERSONNEL")
    row += 1
    header(row)
    row += 1
    pers_start = row
    end = write_items(row, sec["personnel"])
    row = end + 1
    pers_tot = row
    total_row(row, "TOTAL PERSONNEL", pers_start, end)
    row += 3

    bar(row, "INCOME TAX")
    row += 1
    header(row)
    row += 1
    tax_start = row
    end = write_items(row, sec["tax"])
    tax_tot = end
    row = end + 3

    bar(row, "SPECIAL PROJECTS (Year 1 of 10-year plan)")
    row += 1
    header(row)
    row += 1
    sp_start = row
    end = write_items(row, sec["special"])
    row = end + 1
    sp_tot = row
    total_row(row, "TOTAL SPECIAL PROJECTS", sp_start, end)
    row += 3

    bar(row, "SUMMARY & LEVY CALCULATION")
    row += 2
    ws.cell(row, 2, "Total running costs (Municipal + Expenditure + Net R&M + Personnel + Tax + Special)")
    formula(ws.cell(row, 5), f"E{muni_tot}+E{exp_tot}+E{rm_tot}+E{pers_tot}+E{tax_tot}+E{sp_tot}")
    exp_sum = row
    row += 1
    ws.cell(row, 2, "Less: CSOS expense (owners pay CSOS on its own column)")
    formula(ws.cell(row, 5), f'SUMIF(B{exp_start}:B{exp_tot},"*CSOS*",E{exp_start}:E{exp_tot})')
    csos_row = row
    row += 1
    ws.cell(row, 2, "Less: Other Income + Municipal Recoveries")
    formula(ws.cell(row, 5), f"SUM(E{other_start}:E{other_end})+SUM(E{rec_start}:E{rec_end})")
    less_row = row
    row += 1
    ws.cell(row, 2, "AMOUNT ORDINARY LEVIES MUST COVER").font = Font(bold=True)
    formula(ws.cell(row, 5), f"E{exp_sum}-E{csos_row}-E{less_row}", GREEN)
    collect_row = row
    row += 2
    ws.cell(row, 2, "GROSS ORDINARY LEVIES (adjusted for unpaid levies)").font = Font(bold=True, color="9C0006", size=12)
    formula(ws.cell(row, 5), f"IF(C6=0,0,E{collect_row}/C6)", RED)
    ws.cell(row, 5).font = Font(bold=True, size=12)
    gross_row = row
    ws.cell(levy_start, 5).value = f"=E{gross_row}"
    fill(ws.cell(levy_start, 5), RED)
    row += 2
    ws.cell(row, 2, "Reserve (billed separately)")
    formula(ws.cell(row, 5), f"E{levy_start+1}")
    row += 1
    ws.cell(row, 2, "CSOS (billed separately)")
    formula(ws.cell(row, 5), f"E{levy_start+2}")

    # PQ
    pq = wb.create_sheet("PQ")
    pq["A1"] = "PARTICIPATION QUOTA / LEVY ALLOCATION"
    pq["A1"].font = Font(bold=True, size=14, color=NAVY)
    pq["B5"] = "Ordinary Levies (Monthly)"
    pq["D5"] = f"=BUDGET!F{gross_row}" if False else f"=BUDGET!E{gross_row}/12"
    pq["B6"] = "Reserve Fund (Monthly)"
    pq["D6"] = f"=BUDGET!E{levy_start+1}/12"
    pq["B7"] = "CSOS (Monthly)"
    pq["D7"] = f"=BUDGET!E{levy_start+2}/12"
    pq["B8"] = "TOTAL MONTHLY PER COMPLEX"
    pq["D8"] = "=D5+D6+D7"
    for c in ("D5", "D6", "D7", "D8"):
        pq[c].number_format = MONEY
    headers = ["#", "Unit / Owner Code", "PQ", "Ordinary Levy", "Reserve Fund", "CSOS", "Total Monthly"]
    for i, h in enumerate(headers, 1):
        cell = pq.cell(10, i, h)
        fill(cell, NAVY)
        cell.font = Font(bold=True, color="FFFFFF")
    units = []
    if pq_df is not None and not pq_df.empty:
        units = pq_df.to_dict("records")
    if not units:
        units = [{"Unit": "UNIT-1", "PQ": 1.0}]
    for i, u in enumerate(units):
        r = 11 + i
        pq.cell(r, 1, i + 1)
        pq.cell(r, 2, str(u.get("Unit", f"UNIT-{i+1}")))
        pq.cell(r, 3, float(u.get("PQ") or 0)).number_format = "0.000000"
        pq.cell(r, 4, f"=$D$5*C{r}").number_format = MONEY
        pq.cell(r, 5, f"=$D$6*C{r}").number_format = MONEY
        pq.cell(r, 6, f"=$D$7*C{r}").number_format = MONEY
        pq.cell(r, 7, f"=D{r}+E{r}+F{r}").number_format = MONEY
    last = 10 + len(units)
    tot_r = last + 1
    pq.cell(tot_r, 2, "TOTAL").font = Font(bold=True)
    pq.cell(tot_r, 3, f"=SUM(C11:C{last})")
    for col, letter in enumerate("DEFG", 4):
        pq.cell(tot_r, col, f"=SUM({letter}11:{letter}{last})").number_format = MONEY

    # 10 YMP
    ymp_ws = wb.create_sheet("10 YMP")
    ymp_ws["A1"] = "10 YEAR MAINTENANCE PLAN"
    ymp_ws["A1"].font = Font(bold=True, size=14, color=NAVY)
    ymp_ws["A2"] = "Paste or type Year 1–10 amounts. Year 1 should match Special Projects."
    ymp_ws.cell(5, 1, "Project")
    fill(ymp_ws.cell(5, 1), NAVY)
    ymp_ws.cell(5, 1).font = Font(bold=True, color="FFFFFF")
    for y in range(10):
        c = ymp_ws.cell(5, 2 + y, f"Year {y+1}")
        fill(c, NAVY)
        c.font = Font(bold=True, color="FFFFFF")
    ymp_ws.column_dimensions["A"].width = 36
    if not ymp:
        ymp = [{"desc": "Project 1", "years": [0] * 10}]
    for i, p in enumerate(ymp):
        r = 6 + i
        input_cell(ymp_ws.cell(r, 1), p.get("desc", ""))
        years = p.get("years") or [0] * 10
        for y in range(10):
            input_cell(ymp_ws.cell(r, 2 + y), float(years[y] if y < len(years) else 0), MONEY)
    last_y = 5 + max(len(ymp), 1)
    ymp_ws.cell(last_y + 1, 1, "TOTAL PER YEAR").font = Font(bold=True)
    for y in range(10):
        letter = get_column_letter(2 + y)
        formula(ymp_ws.cell(last_y + 1, 2 + y), f"SUM({letter}6:{letter}{last_y})", TOTAL)

    note = wb.create_sheet("HOW TO USE")
    note["A1"] = "Yellow cells with blue text are inputs. Black amounts are formulas."
    note["A2"] = "Ordinary levies cover running costs after other income. Reserve and CSOS are billed on their own PQ columns."
    note["A3"] = "95% collection rate means you expect 5% of billed ordinary levies not to be paid."
    note.column_dimensions["A"].width = 120

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


def money(n: float) -> str:
    return f"R {n:,.2f}"


def init_state():
    ss = st.session_state
    ss.setdefault("sections", default_sections())
    ss.setdefault("complex_name", "")
    ss.setdefault("fin_year", "01-03-2026 / 28-02-2027")
    ss.setdefault("collection_rate", 95.0)
    ss.setdefault("reserve_opening", 0.0)
    ss.setdefault("insurance_recoveries", 0.0)
    ss.setdefault("pq_df", None)
    ss.setdefault("ymp", [
        {"desc": "Painting of Buildings", "years": [100000, 0, 0, 0, 0, 100000, 0, 0, 0, 0]},
        {"desc": "Roof Maintenance", "years": [25000, 0, 0, 25000, 0, 0, 25000, 0, 0, 25000]},
    ])
    ss.setdefault("loaded_msg", "")


def records_from_editor(edited: pd.DataFrame, previous: list) -> list:
    prev_by_desc = {p["desc"]: p for p in previous}
    out = []
    for _, row in edited.iterrows():
        desc = str(row.get("Description") or "").strip()
        if not desc:
            continue
        actual = float(row.get("Actual") or 0)
        pct = float(row.get("% Increase") or 0)
        budgeted = float(row.get("Budgeted yearly") or 0)
        note = str(row.get("Notes") or "")
        prev = prev_by_desc.get(desc, {})
        formula_bud = actual * (1 + pct / 100)
        mode = prev.get("mode", "pct")
        if abs(budgeted - formula_bud) > 0.05:
            mode = "amount"
        elif abs(pct - float(prev.get("pct") or 0)) > 0.05:
            mode = "pct"
            budgeted = formula_bud
        rec = {
            "id": prev.get("id") or uid(),
            "desc": desc,
            "actual": actual,
            "pct": pct if mode != "amount" else (0 if actual == 0 else (budgeted / actual - 1) * 100),
            "budgeted": budgeted,
            "mode": mode,
            "note": note,
            "computed": bool(prev.get("computed")),
        }
        out.append(rec)
    return out


def edit_section(key: str, title: str, help_text: str = ""):
    st.subheader(title)
    if help_text:
        st.caption(help_text)
    items = st.session_state.sections[key]
    apply_ordinary(st.session_state)
    rows = []
    for it in items:
        rows.append({
            "Description": it["desc"],
            "Actual": float(it.get("actual") or 0),
            "% Increase": float(implied_pct(it)),
            "Budgeted yearly": float(yearly(it)),
            "Monthly": float(yearly(it) / 12),
            "Notes": it.get("note") or "",
        })
    df = pd.DataFrame(rows)
    disabled = ["Monthly"]
    if key == "levy_income":
        # ordinary yearly is calculated
        pass
    edited = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "Description": st.column_config.TextColumn("Description", width="medium"),
            "Actual": st.column_config.NumberColumn("Actual", format="%.2f"),
            "% Increase": st.column_config.NumberColumn("% Increase", format="%.1f", help="Type this, or type Budgeted yearly"),
            "Budgeted yearly": st.column_config.NumberColumn("Budgeted yearly", format="%.2f", help="Type a quote / contract / trustee amount"),
            "Monthly": st.column_config.NumberColumn("Monthly", format="%.2f", disabled=True),
            "Notes": st.column_config.TextColumn("Notes"),
        },
        key=f"ed_{key}",
        disabled=disabled,
    )
    st.session_state.sections[key] = records_from_editor(edited, items)
    apply_ordinary(st.session_state)
    st.caption(f"Section total: {money(sum_yearly(st.session_state.sections[key]))}")


def main():
    init_state()
    logo = Path(__file__).parent / "domus_logo.jpeg"
    top = st.columns([1, 5])
    with top[0]:
        if logo.exists():
            st.image(str(logo), width=120)
    with top[1]:
        st.title("Domus Property Management Budget")
        st.caption("Load last year’s figures → review in plain sections → download Excel")

    tot = compute_totals(st.session_state)
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Gross ordinary levies", money(tot["gross_ordinary"]))
    m2.metric("Collection rate", f"{tot['collection_rate']:.0f}%")
    m3.metric("R&M net", money(tot["rm_net"]))
    m4.metric("Total expenses", money(tot["expenses"]))

    with st.sidebar:
        st.header("1. Complex")
        st.session_state.complex_name = st.text_input("Complex name", st.session_state.complex_name)
        st.session_state.fin_year = st.text_input("Financial year", st.session_state.fin_year)

        st.header("2. Last year’s figures")
        st.caption("WeConnectU Excel is the reliable source. PDF is a backup only.")
        excel_up = st.file_uploader("WeConnectU Actual vs Budget (Excel)", type=["xlsx", "xls", "xlsm"], key="wcu")
        if excel_up and st.button("Load Excel", type="primary"):
            try:
                rows = extract_weconnectu(excel_up)
                if not rows:
                    st.error("No lines found. In WeConnectU export Options → Budget and Actuals.")
                else:
                    sections, ins, added = match_rows(rows, default_sections())
                    st.session_state.sections = sections
                    if ins:
                        st.session_state.insurance_recoveries = ins
                    st.session_state.loaded_msg = f"Loaded {len(rows)} lines. Extra complex-specific lines: {added}."
                    st.success(st.session_state.loaded_msg)
                    st.rerun()
            except Exception as e:
                st.error(f"Could not read Excel: {e}")

        pdf_up = st.file_uploader("PDF (backup only)", type=["pdf"], key="pdf")
        if pdf_up and st.button("Try PDF extract"):
            try:
                import pdfplumber
                text = ""
                with pdfplumber.open(pdf_up) as pdf:
                    for page in pdf.pages:
                        text += (page.extract_text() or "") + "\n"
                st.info("PDF text was read. Prefer the WeConnectU Excel if numbers look thin.")
                st.text_area("Extracted text (check this)", text[:4000], height=160)
            except Exception as e:
                st.error(f"PDF failed: {e}")

        st.header("3. Assumptions")
        st.session_state.collection_rate = st.slider(
            "Expected collection rate %", 50.0, 100.0, float(st.session_state.collection_rate), 0.5,
            help="95% means you expect 5% of billed ordinary levies not to be paid.",
        )
        st.session_state.reserve_opening = st.number_input("Opening reserve balance", value=float(st.session_state.reserve_opening), step=1000.0)
        st.session_state.insurance_recoveries = st.number_input("Insurance recoveries (taken off R&M)", value=float(st.session_state.insurance_recoveries), step=100.0)

        if st.button("Start over"):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()

    if st.session_state.loaded_msg:
        st.success(st.session_state.loaded_msg)

    tabs = st.tabs([
        "Setup",
        "Income",
        "Municipal",
        "Expenditure",
        "Repair & Maintenance",
        "Personnel",
        "Income Tax",
        "Special Projects",
        "PQ / Levy schedule",
        "10-year plan",
        "Download Excel",
    ])

    with tabs[0]:
        st.markdown("### How the columns work")
        st.markdown(
            """
1. **Actual** — last year’s figure (from the Excel).
2. **% Increase** — type this if you want “last year plus 10%”. Yearly then fills in.
3. **Budgeted yearly** — type this if you already know the rand amount (quote, Eskom, contract). The % then fills in.
4. **Monthly** — always yearly ÷ 12.

Ordinary levies are calculated so that after unpaid levies you still cover running costs.  
**Reserve** and **CSOS** are billed on their own columns so owners can see them.

**95% / 100%:** not a tax. 95% means you expect 5% of ordinary levies not to be paid, so you bill a little more.
            """
        )

    with tabs[1]:
        edit_section("levy_income", "Levy Income", "Type the reserve as a yearly rand amount. Ordinary is calculated.")
        st.divider()
        edit_section("other_income", "Other Income")
        st.divider()
        edit_section("muni_recoveries", "Municipal Recovery Income", "Recoveries sit here. Gross municipal bills are on the next tab.")

    with tabs[2]:
        edit_section("municipal", "Municipal Expenses (Gross)")

    with tabs[3]:
        edit_section("expenditure", "Expenditure (operating, except R&M and Personnel)")

    with tabs[4]:
        edit_section("rm", "Repair & Maintenance")
        st.caption(f"Insurance recoveries {money(st.session_state.insurance_recoveries)} are deducted. Net R&M {money(tot['rm_net'])}.")

    with tabs[5]:
        edit_section("personnel", "Personnel")

    with tabs[6]:
        edit_section("tax", "Income Tax")

    with tabs[7]:
        edit_section("special", "Special Projects (Year 1 of the 10-year plan)")

    with tabs[8]:
        st.subheader("PQ / Levy schedule")
        st.caption("Upload CSV or Excel with Unit and PQ. Each owner pays Ordinary + Reserve + CSOS, each × their PQ.")
        pq_file = st.file_uploader("PQ file", type=["csv", "xlsx", "xls"], key="pq")
        if pq_file:
            try:
                raw = pd.read_csv(pq_file) if pq_file.name.lower().endswith(".csv") else pd.read_excel(pq_file)
                raw.columns = [str(c).strip() for c in raw.columns]
                seen = {}
                cols = []
                for c in raw.columns:
                    if c in seen:
                        seen[c] += 1
                        cols.append(f"{c}_{seen[c]}")
                    else:
                        seen[c] = 0
                        cols.append(c)
                raw.columns = cols
                unit_col = next((c for c in raw.columns if re.search(r"unit|owner|code", str(c), re.I)), None)
                pq_col = next((c for c in raw.columns if re.search(r"pq|quota|ratio|^share$", str(c), re.I)), None)
                if unit_col is None or pq_col is None:
                    st.error("Need Unit and PQ columns. Found: " + ", ".join(raw.columns.astype(str)))
                    st.dataframe(raw.head(8), use_container_width=True)
                else:
                    clean = pd.DataFrame({
                        "Unit": raw[unit_col].astype(str).str.strip(),
                        "PQ": pd.to_numeric(raw[pq_col], errors="coerce").fillna(0),
                    })
                    clean = clean[clean["Unit"].str.lower().ne("nan") & clean["Unit"].ne("")].reset_index(drop=True)
                    s = clean["PQ"].sum()
                    if 50 < s < 150:
                        clean["PQ"] = clean["PQ"] / 100
                        st.info(f"PQ looked like percentages (total {s:.2f}). Divided by 100.")
                    st.session_state.pq_df = clean
                    st.success(f"Loaded {len(clean)} units. PQ total = {clean['PQ'].sum():.6f}")
            except Exception as e:
                st.error(f"Could not read PQ file: {e}")

        if st.session_state.pq_df is not None and not st.session_state.pq_df.empty:
            preview = st.session_state.pq_df.copy()
            preview["Ordinary"] = preview["PQ"] * tot["monthly_ordinary"]
            preview["Reserve"] = preview["PQ"] * tot["monthly_reserve"]
            preview["CSOS"] = preview["PQ"] * tot["monthly_csos"]
            preview["Total monthly"] = preview["Ordinary"] + preview["Reserve"] + preview["CSOS"]
            st.dataframe(preview, use_container_width=True, hide_index=True)

    with tabs[9]:
        st.subheader("10-year maintenance plan")
        st.caption("Copy the plan from Excel (project name + 10 year amounts) and paste below. No file upload needed.")
        paste = st.text_area("Paste from Excel", height=140, placeholder="Painting\t100000\t0\t0\t0\t0\t100000\t0\t0\t0\t0")
        c1, c2 = st.columns(2)
        if c1.button("Paste into plan") and paste.strip():
            parsed = parse_ymp_paste(paste)
            if parsed:
                st.session_state.ymp = parsed
                st.success(f"Loaded {len(parsed)} projects.")
                st.rerun()
            else:
                st.warning("Nothing readable. Copy name + amounts from Excel.")
        if c2.button("Send Year 1 to Special Projects"):
            y1 = [p for p in st.session_state.ymp if (p.get("years") or [0])[0]]
            special = []
            for p in y1:
                special.append(item(p["desc"], 0, "From 10-year plan Year 1", mode="amount"))
                special[-1]["budgeted"] = float(p["years"][0])
            if special:
                st.session_state.sections["special"] = special
                st.success(f"Copied {len(special)} Year-1 projects to Special Projects.")
        ymp_rows = []
        for p in st.session_state.ymp:
            rec = {"Project": p["desc"]}
            years = p.get("years") or [0] * 10
            for i in range(10):
                rec[f"Y{i+1}"] = float(years[i] if i < len(years) else 0)
            ymp_rows.append(rec)
        ymp_df = pd.DataFrame(ymp_rows)
        edited_ymp = st.data_editor(ymp_df, num_rows="dynamic", use_container_width=True, hide_index=True, key="ymp_ed")
        new_ymp = []
        for _, r in edited_ymp.iterrows():
            new_ymp.append({
                "desc": str(r.get("Project") or ""),
                "years": [float(r.get(f"Y{i+1}") or 0) for i in range(10)],
            })
        st.session_state.ymp = new_ymp

    with tabs[10]:
        st.subheader("Download the full Excel pack")
        st.caption("Budget + PQ schedule + 10-year plan. Yellow cells are inputs. Formulas stay live.")
        if not st.session_state.complex_name:
            st.warning("Enter the complex name in the sidebar first.")
        else:
            xls = generate_excel(st.session_state, st.session_state.pq_df, st.session_state.ymp)
            name = re.sub(r"\s+", "_", st.session_state.complex_name)
            st.download_button(
                "Download budget Excel",
                data=xls,
                file_name=f"Budget_{name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


if __name__ == "__main__":
    main()
