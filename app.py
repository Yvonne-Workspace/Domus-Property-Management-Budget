"""Domus budget app — Streamlit. Layout and levy math follow Domus Excel packs."""

from __future__ import annotations

import re
from io import BytesIO
from pathlib import Path

import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Domus Budget", layout="wide", initial_sidebar_state="expanded")

NOW_YEAR = 2026
YELLOW, NAVY, BLUE, RED, TOTAL, SECTION = "FFFF99", "1F4E79", "0000FF", "FFC7CE", "D9E2F3", "2E75B6"
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
MONEY = '#,##0.00;(#,##0.00);"-"'


def uid() -> str:
    import random, string
    return "".join(random.choices(string.ascii_lowercase + string.digits, k=8))


def money(n: float) -> str:
    return f"R {n:,.2f}"


def norm(s: str) -> str:
    s = str(s or "").lower()
    s = re.sub(r"[–—-]", " ", s)
    s = re.sub(r"[^a-z0-9/& ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def row(desc: str, note: str = "") -> dict:
    return {
        "id": uid(), "desc": desc, "actual": 0.0, "pct": 0.0, "yearly": 0.0,
        "insurance": 0.0, "note": note, "is_recovery": False,
    }


def default_sections() -> dict:
    return {
        "levy": [
            row("Ordinary Levies", "Admin levy. Does not include estate/HOA pass-throughs or insurance billed on its own PQ column."),
            row("Reserve Fund Contribution", "Type the yearly amount, or use 15% of ordinary."),
            row("CSOS Levy (Income)", "This complex’s own CSOS. Not the master-estate CSOS."),
            row("Insurance billed to owners", "Thornhill / Mount Kos style. Own PQ column. Leave 0 if insurance stays inside ordinary."),
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
            row("Estate / HOA Levies recovered", "e.g. Xanadu Levies. Billed to owners on the PQ. Not ordinary."),
            row("Estate / HOA CSOS recovered", "e.g. Xanadu HOA CSOS. Own PQ column."),
        ],
        "hoa_expense": [
            row("Estate / HOA Levies paid", "e.g. Xanadu Eco Park. Paid to the estate. Not in ordinary."),
            row("Estate / HOA CSOS paid", "e.g. Xanadu CSOS paid. Not in ordinary."),
        ],
        "recoveries_other": [
            row("Insurance claims recovered", "Claim payouts. Deduct on the R&M line. Do not budget as normal income."),
            row("Legal Fees Recovered"),
            row("Maintenance Recovered"),
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
            row("Insurance billed to owners (monthly fixed)", "Mount Kos-style extra on the levy invoice."),
            row("Prepaid meters estimate (monthly)"),
            row("Eskom fixed charge (monthly)"),
            row("Communal charge (monthly)"),
        ],
    }


def is_muni_recovery(desc: str, flag: bool = False) -> bool:
    d = (desc or "").lower()
    # Insurance recovered / legal recovered are INCOME, not municipal credits.
    if "insurance" in d or "legal recover" in d:
        return False
    if flag:
        return True
    if d.startswith("less:"):
        return True
    if "recover" in d:
        return True
    if "sewer" in d and "plant" in d and "electric" in d:
        return True
    return False


def net_of(r: dict) -> float:
    y = float(r.get("yearly") or 0)
    ins = float(r.get("insurance") or 0)
    if is_muni_recovery(r.get("desc") or "", r.get("is_recovery")):
        return -abs(y)
    return y - ins


def sum_net(items: list) -> float:
    return sum(net_of(r) for r in items)


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


def skip_from_ordinary(r: dict, state: dict) -> bool:
    f = family(r.get("desc") or "")
    if f in ("hoa_levy_exp", "hoa_csos_exp", "hoa_levy_inc", "hoa_csos_inc"):
        return True
    if f == "insurance" and insurance_on_pq(state):
        return True
    return False


def ordinary_total(state: dict) -> float:
    s = state["sections"]
    total = sum_net(s["municipal"]) + sum_net(s["rm"]) + sum_net(s["personnel"]) + sum_net(s["tax"])
    total += sum(net_of(r) for r in s["expenditure"] if not skip_from_ordinary(r, state))
    if state.get("special_in_ordinary"):
        total += sum_net(s["special"])
    return total


def levy_pieces(state: dict) -> list:
    s = state["sections"]
    return [
        ("Net municipal (gross minus recoveries)", sum_net(s["municipal"])),
        ("Expenditure", sum(net_of(r) for r in s["expenditure"] if not skip_from_ordinary(r, state))),
        ("R&M after insurance", sum_net(s["rm"])),
        ("Personnel", sum_net(s["personnel"])),
        ("Tax", sum_net(s["tax"])),
        ("Special (only if ticked)", sum_net(s["special"]) if state.get("special_in_ordinary") else 0.0),
        ("Insurance premium (billed on PQ, not in levy)", insurance_expense_amount(state) if insurance_on_pq(state) else 0.0),
    ]


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
            f"Municipal recoveries {money(abs(rec))} vs electricity/water {money(gross)} — last year Falcon View recovered most of this, so net municipal was ~R116,000 not {money(sum_net(state['sections']['municipal']))}."
        )
    return flags


def apply_levy_lines(state: dict) -> None:
    ord_amt = ordinary_total(state)
    for r in state["sections"]["levy"]:
        f = family(r["desc"])
        if f == "ordinary":
            r["yearly"] = ord_amt
            a = float(r.get("actual") or 0)
            r["pct"] = 0.0 if a == 0 else (ord_amt / a) * 100 - 100
        if f == "reserve":
            if state.get("reserve_mode") == "15pct":
                r["yearly"] = ord_amt * 0.15
            else:
                r["yearly"] = float(state.get("reserve_amount") or r.get("yearly") or 0)
            a = float(r.get("actual") or 0)
            r["pct"] = 0.0 if a == 0 else (float(r["yearly"]) / a) * 100 - 100
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
            add("Levies", r.get("yearly"))
            break
    if not out:
        add("Levies", ordinary_total(state))
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
        if "contrib" in d or "expense" in d or "admin" in d:
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
    if re.search(r"garden(ing)? contract|site cleaning", d):
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
            item["note"] = "Insurance claim. Deduct on the matching R&M line. Do not budget as normal income."
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


def parse_pq_upload(uploaded):
    """WeConnectU unit-pqs: PQ column is often 0, real share is Ratio 1."""
    if uploaded.name.lower().endswith(".csv"):
        raw = pd.read_csv(uploaded)
    else:
        raw = pd.read_excel(uploaded)
    if isinstance(raw.columns, pd.MultiIndex):
        raw.columns = [" ".join(str(x) for x in col if str(x) != "nan").strip() for col in raw.columns]
    raw.columns = [str(c).strip() for c in raw.columns]
    seen, cols = {}, []
    for c in raw.columns:
        if c in seen:
            seen[c] += 1
            cols.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            cols.append(c)
    raw.columns = cols

    unit_candidates = [
        c for c in raw.columns
        if re.search(r"customer code|owner code|account code|unit|owner|code|section", str(c), re.I)
        and "size" not in str(c).lower()
    ]
    unit_col = None
    for c in unit_candidates:
        sample = raw[c].astype(str).head(20)
        if sample.str.contains(r"[A-Za-z]", regex=True).any() and "customer" in str(c).lower():
            unit_col = c
            break
    if unit_col is None:
        for c in unit_candidates:
            sample = raw[c].astype(str).head(20)
            if sample.str.contains(r"[A-Za-z]", regex=True).any():
                unit_col = c
                break
    if unit_col is None and unit_candidates:
        unit_col = unit_candidates[0]

    pq_candidates = [c for c in raw.columns if re.search(r"pq|quota|ratio|^share$", str(c), re.I)]
    pq_col, best = None, -1
    for c in pq_candidates:
        nums = pd.to_numeric(raw[c], errors="coerce").fillna(0.0)
        pos = nums[nums > 0]
        if len(pos) < 2:
            continue
        s = float(pos.sum())
        score = len(pos)
        if abs(s - 1) < 0.2:
            score += 80
        if abs(s - 100) < 20:
            score += 60
        if score > best:
            best, pq_col = score, c
    if pq_col is None and pq_candidates:
        pq_col = pq_candidates[0]
    if not unit_col or not pq_col:
        raise ValueError("Need a Unit / Customer Code column and a PQ or Ratio column. Found: " + ", ".join(map(str, raw.columns)))

    clean = pd.DataFrame({
        "Unit": raw[unit_col].astype(str).str.strip(),
        "PQ": pd.to_numeric(raw[pq_col], errors="coerce").fillna(0),
    })
    clean = clean[clean["Unit"].str.lower().ne("nan") & clean["Unit"].ne("")].reset_index(drop=True)
    total = float(clean["PQ"].sum())
    note = ""
    if 50 < total < 150:
        clean["PQ"] = clean["PQ"] / 100
        note = f" Values looked like percentages (total {total:.2f}) so they were divided by 100."
    if clean.empty:
        raise ValueError(f"No units with a PQ greater than 0. Tried column {pq_col}.")
    msg = f"Loaded {len(clean)} units from column {pq_col} (PQ total {clean['PQ'].sum():.6f}).{note} Open the PQ / Levies tab to see each owner."
    return clean.to_dict("records"), msg


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


def items_to_df(items: list, rm: bool) -> pd.DataFrame:
    recs = []
    for it in items:
        recs.append({
            "Description": it["desc"],
            "Actual": float(it.get("actual") or 0),
            "% Increase": float(it.get("pct") or 0),
            "Budgeted yearly": float(it.get("yearly") or 0),
            "Monthly": net_of(it) / 12,
            "Insurance payout": float(it.get("insurance") or 0),
            "Notes": it.get("note") or "",
        })
    cols = ["Description", "Actual", "% Increase", "Budgeted yearly", "Monthly", "Notes"]
    if rm:
        cols = ["Description", "Actual", "% Increase", "Budgeted yearly", "Monthly", "Insurance payout", "Notes"]
    if not recs:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame(recs)[cols]


def save_editor(edited: pd.DataFrame, previous: list, rm: bool, municipal: bool = False) -> list:
    out = []
    records = edited.to_dict("records")
    for i, rec in enumerate(records):
        desc = str(rec.get("Description") or "").strip()
        if not desc:
            continue
        prev = previous[i] if i < len(previous) else {}
        actual = abs(float(rec.get("Actual") or 0))
        pct = float(rec.get("% Increase") or 0)
        yearly = abs(float(rec.get("Budgeted yearly") or 0))
        ins = float(rec.get("Insurance payout") or 0) if rm else float(prev.get("insurance") or 0)
        old_pct = float(prev.get("pct") or 0)
        old_y = abs(float(prev.get("yearly") or 0))
        pct_changed = abs(pct - old_pct) > 0.05
        y_changed = abs(yearly - old_y) > 0.02
        if pct_changed and not y_changed:
            yearly = actual * (1 + pct / 100)
        elif y_changed:
            pct = 0.0 if actual == 0 else (yearly / actual) * 100 - 100
        recovery = municipal and is_muni_recovery(desc, prev.get("is_recovery"))
        out.append({
            "id": prev.get("id") or uid(),
            "desc": desc,
            "actual": actual,
            "pct": pct,
            "yearly": yearly,
            "insurance": ins,
            "note": str(rec.get("Notes") or ""),
            "is_recovery": recovery,
        })
    return out


def ymp_years_from_nums(nums: list) -> list:
    """Thornhill layout: First cycle (2025) | Frequency (7) | Current estimate | Year 1 … Year 10."""
    n = [float(x) for x in nums if x is not None]
    if n and 2020 <= n[0] <= 2040:
        n = n[1:]
    if n and 1 <= n[0] <= 15 and (len(n) == 1 or n[1] >= 20):
        n = n[1:]
    if len(n) > 10:
        n = n[-10:]
    return (n + [0.0] * 10)[:10]


def parse_ymp_sheet(df, start_year: int | None = None) -> list:
    """Read a 10 YMP sheet. Calendar headers (2025, 2026, …) are aligned so Year 1 = this budget year (2026)."""
    start_year = int(start_year or NOW_YEAR)
    header_row = None
    year_cols = []  # column indexes in Year1..Year10 order (already aligned)
    cal_pairs = []  # (col, calendar_year)
    for i in range(min(12, len(df))):
        raw = [df.iat[i, j] for j in range(df.shape[1])]
        labels = [str(v).strip().lower() for v in raw]
        ylabels = [j for j, v in enumerate(labels) if re.match(r"^year\s*\d+", v) or re.match(r"^y\d+$", v)]
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
        if not desc or re.match(r"^(total|planned maintenance|project|projects|first|cycle|frequency|current|estimate)$", desc, re.I):
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
            years = ymp_years_from_nums(nums)
            first, freq, est = "", "", 0.0
        if not desc or (not any(years) and not est):
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
    # filled after we know reserve row
    ws["B7"] = "Projected reserve at year-end"
    fml(ws["D7"], "D5+D6")
    ws["B8"] = "Yellow cells = type here. Budgeted Yearly = Actual × (1 + %). Monthly = Yearly ÷ 12. Change % or overwrite Yearly in the meeting."
    ws["B8"].font = Font(italic=True, size=9, color="666666")

    r = 10

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

    def write(items, recovery_as_income=False):
        """Old pack formulas: F = D*(1+E), G = F/12. Recoveries shown as positive income when asked."""
        nonlocal r
        if not items:
            start = r
            r += 1
            return start, start
        start = r
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
            if fam == "csos_inc":
                levy_rows["csos"] = r
            if fam == "reserve":
                levy_rows["reserve"] = r
            if fam == "ins_bill" or re.search(r"insurance recovered|levy\s*[-–]\s*insurance", desc, re.I):
                levy_rows["insurance"] = r
            named_rows[desc] = r
            rec = is_muni_recovery(desc, it.get("is_recovery"))
            inp(ws.cell(r, 4), act, MONEY)
            if fam == "ordinary":
                # % follows the levy formula so a meeting change to costs updates the %
                fml(ws.cell(r, 5), f'IF(D{r}=0,0,F{r}/D{r}-1)')
                ws.cell(r, 5).number_format = "0.0%"
            else:
                inp(ws.cell(r, 5), pct / 100.0, "0.0%")
            expected = act * (1 + pct / 100.0)
            if fam == "ordinary":
                pass  # F filled after totals
            elif fam == "reserve" and state.get("reserve_mode") == "15pct" and levy_rows.get("ordinary"):
                fml(ws.cell(r, 6), f"0.15*F{levy_rows['ordinary']}")
            elif rec and not recovery_as_income:
                fml(ws.cell(r, 6), f"-ABS(D{r}*(1+E{r}))")
            elif act < 0.5 and abs(y) > 0.5:
                inp(ws.cell(r, 6), y if not (rec and not recovery_as_income) else -abs(y), MONEY)
            elif abs(y - expected) > 1 and y > 0:
                inp(ws.cell(r, 6), abs(y) if recovery_as_income or not rec else -abs(y), MONEY)
            else:
                fml(ws.cell(r, 6), f"D{r}*(1+E{r})")
            fml(ws.cell(r, 7), f"F{r}/12")
            note = it.get("note") or ""
            ins = float(it.get("insurance") or 0)
            if ins:
                note = (note + " | Insurance payout " + f"{ins:,.2f}").strip(" |")
            inp(ws.cell(r, 8), note)
            r += 1
        return start, r - 1

    def tot(label, start, end):
        nonlocal r
        ws.cell(r, 2, label).font = Font(bold=True)
        fml(ws.cell(r, 6), f"SUM(F{start}:F{end})", TOTAL)
        fml(ws.cell(r, 7), f"F{r}/12", TOTAL)
        row_n = r
        r += 2
        return row_n

    bar("INCOME")
    hdr()
    a, b = write(s["levy"])
    inc_tot = tot("TOTAL INCOME", a, b)
    bar("OTHER INCOME")
    hdr()
    a, b = write(s.get("other") or [])
    tot("TOTAL OTHER INCOME", a, b)
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
    util = list(s.get("recoveries_other") or []) + muni_rec
    bar("Recoveries on Utilities")
    hdr()
    a, b = write(util, recovery_as_income=True)
    util_tot = tot("TOTAL UTILITY RECOVERIES", a, b)
    bar("Municipal Charges")
    hdr()
    a, b = write(muni_gross)
    muni_g_tot = tot("TOTAL", a, b)
    ws.cell(r, 2, "TOTAL NET MUNICIPAL CHARGES").font = Font(bold=True)
    fml(ws.cell(r, 6), f"F{muni_g_tot}-F{util_tot}", RED)
    fml(ws.cell(r, 7), f"F{r}/12", RED)
    net_muni = r
    r += 2
    bar("EXPENDITURE")
    hdr()
    exp_items = list(s.get("expenditure") or [])
    if state.get("has_master_hoa"):
        exp_items = exp_items + list(s.get("hoa_expense") or [])
    a, b = write(exp_items)
    exp_tot = tot("TOTAL EXPENDITURE", a, b)
    bar("REPAIR AND MAINTENANCE")
    hdr()
    a, b = write(s.get("rm") or [])
    rm_tot = tot("Total Repair and Maintenance", a, b)
    bar("PERSONNEL")
    hdr()
    a, b = write(s.get("personnel") or [])
    per_tot = tot("Total Personnel Expenses", a, b)
    bar("INCOME TAX")
    hdr()
    a, b = write(s.get("tax") or [])
    tax_tot = tot("TOTAL TAX", a, b)
    bar("SPECIAL PROJECTS")
    hdr()
    a, b = write(s.get("special") or [])
    sp_tot = tot("Total Special Projects Expenses", a, b)
    charged = [x for x in (s.get("fixed") or []) if float(x.get("yearly") or 0) > 0.5]
    if charged:
        bar("EQUAL CHARGES BILLED TO OWNERS (same rand each unit)")
        hdr()
        a, b = write(charged)
        tot("TOTAL EQUAL CHARGES", a, b)

    # Ordinary levies = costs owners must cover (not Xanadu pass-through, not utility recoveries)
    bits = f"F{net_muni}+F{exp_tot}+F{rm_tot}+F{per_tot}+F{tax_tot}"
    # Subtract Xanadu expense if it sits inside expenditure totals
    hoa_exp_items = s.get("hoa_expense") or []
    hoa_exp_sum = ""
    if state.get("has_master_hoa") and hoa_exp_items:
        # expenditure total includes Xanadu; take it out of ordinary
        names = [it.get("desc") for it in hoa_exp_items if it.get("desc")]
        refs = [f"F{named_rows[n]}" for n in names if n in named_rows]
        if refs:
            bits += "-" + "-".join(refs)
    if state.get("special_in_ordinary"):
        bits += f"+F{sp_tot}"
    if insurance_on_pq(state):
        for it in s.get("expenditure") or []:
            d = it.get("desc") or ""
            if family(d) == "insurance" and d in named_rows:
                bits += f"-F{named_rows[d]}"
    bar("ORDINARY LEVY (what we charge)")
    ws.cell(r, 2, "Ordinary levies = net municipal + expenditure (not estate pass-through) + R&M + personnel + tax")
    fml(ws.cell(r, 6), bits, RED)
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

    pq = wb.create_sheet("PQ")
    pq["A1"] = "PQ / LEVY SCHEDULE"
    pq["A1"].font = Font(bold=True, size=14, color=NAVY)
    pq["A2"] = state.get("complex_name") or ""
    bills = pq_bill_lines(state)
    pq["B3"] = "Monthly"
    name_to_budget = {
        "Levies": f"BUDGET!G{levy_rows['ordinary']}" if levy_rows.get("ordinary") else None,
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

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


RESTORE_BARS = [
    ("other income", "other"),
    ("recoveries on hoa", "hoa_income"),
    ("hoa / estate recovered", "hoa_income"),
    ("hoa / estate paid", "hoa_expense"),
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


def restore_from_app_excel(uploaded) -> dict:
    """Reload a file this app previously downloaded so work is not lost."""
    xl = pd.ExcelFile(uploaded)
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
        current = None
        # New pack: B desc, C GL, D actual, E %, F yearly. Old pack: B desc, C actual, D %, E yearly.
        i_act, i_pct, i_year, i_note = 2, 3, 4, 6
        for r in range(len(df)):
            desc = str(df.iat[r, 1] if df.shape[1] > 1 else "").strip()
            if not desc:
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
            if re.match(r"^(description|total |net |ordinary levy)", desc, re.I):
                continue
            actual = abs(num(df.iat[r, i_act] if df.shape[1] > i_act else None) or 0.0)
            pct_raw = num(df.iat[r, i_pct] if df.shape[1] > i_pct else None) or 0.0
            pct = pct_raw * 100 if abs(pct_raw) <= 2 else pct_raw
            yearly = abs(num(df.iat[r, i_year] if df.shape[1] > i_year else None) or 0.0)
            note = str(df.iat[r, i_note] if df.shape[1] > i_note else "") or ""
            item = row(desc, "" if note in ("nan", "None") else note)
            item["actual"] = actual
            item["pct"] = pct
            item["yearly"] = yearly if yearly else actual * (1 + pct / 100)
            item["is_recovery"] = "recover" in norm(desc) and current == "municipal"
            out["sections"][current].append(item)
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


def init():
    ss = st.session_state
    ss.setdefault("sections", default_sections())
    ss.setdefault("complex_name", "")
    ss.setdefault("fin_year", "01-03-2026 / 28-02-2027")
    ss.setdefault("reserve_mode", "amount")
    ss.setdefault("reserve_amount", 0.0)
    ss.setdefault("reserve_balance", 0.0)
    ss.setdefault("special_in_ordinary", False)
    ss.setdefault("afs_sections", None)
    ss.setdefault("wcu_rows", None)
    ss.setdefault("has_master_hoa", False)
    ss.setdefault("insurance_mode", "levy")
    ss.setdefault("insurance_bill_yearly", 0.0)
    ss.setdefault("pq", None)
    ss.setdefault("ymp", [{"desc": "", "years": [0.0] * 10}])
    ss.setdefault("msg", "")
    ss.setdefault("current_monthly_levy", 0.0)
    ss.setdefault("actual_months", 12)
    ss.setdefault("estate_levy_yearly", 0.0)
    ss.setdefault("estate_levy_name", "Estate / master HOA levy")
    ss.setdefault("estate_levy_mode", "separate")
    ss.setdefault("estate_split", "equal")


def section_form(key: str, title: str, help_text: str, rm: bool = False):
    st.subheader(title)
    if help_text:
        st.caption(help_text)
    items = st.session_state.sections.get(key) or []
    df = items_to_df(items, rm)
    with st.form(f"form_{key}"):
        edited = st.data_editor(
            df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "Description": st.column_config.TextColumn("Description", width="medium"),
                "Actual": st.column_config.NumberColumn("Actual", format="%.2f"),
                "% Increase": st.column_config.NumberColumn("% Increase", format="%.1f", help="Type % then click Save"),
                "Budgeted yearly": st.column_config.NumberColumn("Budgeted yearly", format="%.2f", help="Or type the rand amount then Save"),
                "Monthly": st.column_config.NumberColumn("Monthly", format="%.2f", disabled=True, help="Yearly ÷ 12. Updates when you Save."),
                "Insurance payout": st.column_config.NumberColumn("Insurance payout", format="%.2f"),
                "Notes": st.column_config.TextColumn("Notes"),
            },
            disabled=["Monthly"],
        )
        saved = st.form_submit_button("Save this section", type="primary")
    if saved:
        st.session_state.sections[key] = save_editor(edited, items, rm, municipal=(key == "municipal"))
        if key == "levy":
            for r in st.session_state.sections["levy"]:
                if is_ins_bill_line(r.get("desc") or "") and float(r.get("yearly") or 0) > 0.5:
                    st.session_state["_pending_insurance_bill"] = float(r["yearly"])
                    break
        apply_levy_lines(st.session_state)
        st.success("Saved. Monthly = yearly ÷ 12. % = (yearly ÷ actual) × 100 − 100.")
        st.rerun()
    st.caption(f"Section net total: {money(sum_net(st.session_state.sections[key]))}  ·  Monthly total: {money(sum_net(st.session_state.sections[key]) / 12)}")


def main():
    init()
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
    new_m = ord_amt / 12
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
    m3.metric("Ordinary levies for the year", money(ord_amt))
    m4.metric("Reserve for the year", money(float(reserve["yearly"]) if reserve else 0))

    with st.expander("Why is the levy this amount? (plain English)", expanded=abs(levy_pct) > 25):
        st.write(
            "Ordinary levy is **not** last year’s levy plus a %. "
            "It is the **cost the owners must cover**: municipal after recoveries + running costs + repairs + staff + tax."
        )
        for label, amt in levy_pieces(st.session_state):
            if amt or "Special" not in label:
                st.write(f"- {label}: **{money(amt)}**")
        st.write(f"- **Ordinary levies for the year: {money(ord_amt)}**")
        flags = odd_budget_lines(st.session_state)
        if flags:
            st.warning("These lines are making the levy jump. Fix them on the tabs, then Save — you do not need to start over.")
            for f in flags:
                st.write(f"- {f}")
        st.caption("Last year Falcon View ordinary levies were about R1.36 million (about R1,113 per full PQ unit per month).")

    with st.sidebar:
        st.header("Complex")
        st.session_state.complex_name = st.text_input("Complex name", st.session_state.complex_name)
        st.session_state.fin_year = st.text_input("Financial year", st.session_state.fin_year)

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

        pq_up = st.file_uploader("PQ / unit ratios (Excel or CSV)", type=["csv", "xlsx", "xls"], key="pq_sidebar")
        if pq_up and st.button("Load PQs"):
            try:
                records, msg = parse_pq_upload(pq_up)
                st.session_state.pq = records
                st.session_state.msg = msg
                st.rerun()
            except Exception as e:
                st.error(f"Could not read PQ file: {e}")

        st.header("Keep your work")
        st.caption("Do not click Start over. Download Excel, then you can restore it here.")
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
                apply_levy_lines(st.session_state)
                st.session_state.msg = "Budget restored from Excel. Nothing was started over."
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
        st.session_state.reserve_balance = st.number_input(
            "How much is already in the reserve fund (bank / 15% account)?",
            value=float(st.session_state.get("reserve_balance") or 0),
            min_value=0.0,
            step=1000.0,
            help="The money sitting there now. Not this year’s contribution.",
        )
        st.session_state.reserve_mode = st.radio(
            "How is this year’s contribution calculated?",
            ["amount", "15pct"],
            format_func=lambda x: "I will type the yearly amount" if x == "amount" else "15% of ordinary levies",
            index=0 if st.session_state.reserve_mode == "amount" else 1,
        )
        if st.session_state.reserve_mode == "amount":
            st.session_state.reserve_amount = st.number_input(
                "Reserve fund contribution (yearly rands)",
                value=float(st.session_state.reserve_amount),
                step=1000.0,
                min_value=0.0,
            )
        contrib = (
            ordinary_total(st.session_state) * 0.15
            if st.session_state.reserve_mode == "15pct"
            else float(st.session_state.reserve_amount or 0)
        )
        st.caption(
            f"Already in reserve {money(st.session_state.reserve_balance)} + this year {money(contrib)} "
            f"= projected {money(float(st.session_state.reserve_balance or 0) + contrib)}."
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
        "PQ / Levies", "10-year plan", "Download",
    ])

    with tabs[0]:
        st.markdown(
            """
### One rule for typing
Change the numbers, then click **Save this section** once. Nothing is stored until Save.
That stops the amount jumping back to 0.

- Type **% Increase** and Save → Budgeted yearly = Actual × (1 + %).
- Type **Budgeted yearly** and Save → % = (Yearly ÷ Actual) × 100 − 100.
- **Monthly** is always yearly ÷ 12.

### Ordinary levies (same as LTP / Depotel / Mount Kos)
**Ordinary = Net municipal + Expenditure + R&M after insurance + Personnel + Tax**

Net municipal means gross electricity/water/sewer/refuse minus recoveries from owners.

Reserve and CSOS are billed on their own PQ columns. They are not folded into ordinary.

### Second levy (inside another estate) — Thornhill / Xanadu
Xanadu invoices **Thornhill BC**. Thornhill bills owners on the PQ as extra columns:

- Levies (ordinary admin)
- Xanadu Levies
- Xanadu HOA CSOS
- Insurance
- Reserve Fund
- CSOS (this complex)

Those estate lines are **not** folded into ordinary. Leave them 0 on a standalone complex.

### Reserve
Type the yearly rand amount, **or** choose 15% of ordinary (Matte Court style).

### Insurance claims
On **Repair & Maintenance**, type the payout against the **specific line** (roof, plumbing, …).
That line’s net = budgeted yearly − insurance payout.
            """
        )

    with tabs[1]:
        st.info("Ordinary and Reserve update when you save the cost sections / sidebar.")
        st.session_state.insurance_mode = st.radio(
            "Insurance — how do owners pay it?",
            ["levy", "pq"],
            format_func=lambda x: (
                "Part of the ordinary levy (no extra PQ column)"
                if x == "levy"
                else "Extra on the owner invoice (own PQ column — Thornhill style)"
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
        st.divider()
        section_form("levy", "Levy Income", "Add boathouse / boatport / extra levy types with a new row, then Save.")
        st.session_state.has_master_hoa = st.checkbox(
            "We still bill an estate / Xanadu levy to owners (they pay it through us).",
            value=bool(st.session_state.get("has_master_hoa")),
            help="OFF = Xanadu (or the estate) bills owners themselves. PQ will only show our levies, reserve and CSOS. ON = Thornhill-style extra PQ columns.",
        )
        if st.session_state.has_master_hoa:
            st.divider()
            section_form(
                "hoa_income",
                "Estate / HOA recovered from owners",
                "Billed to owners on the PQ as extra columns (Thornhill / Xanadu style).",
            )
            st.divider()
            section_form(
                "hoa_expense",
                "Estate / HOA paid to the estate",
                "What we pay the master HOA. Not included in ordinary levies.",
            )
        else:
            st.caption("Estate / Xanadu lines are hidden. Owners pay that estate directly — it does not go through this budget.")
        st.divider()
        section_form("other", "Other Income", "Fixed Eskom / rental / interest live here. Leave unused lines at 0.")
        st.divider()
        section_form("recoveries_other", "Other recoveries", "Insurance / legal recoveries. Utility recoveries sit under Municipal.")
        st.divider()
        section_form(
            "fixed",
            "Fixed monthly charges on the owner invoice",
            "Use this when EVERY owner pays the SAME extra rand (garden service, prepaid estimate, Eskom fixed, communal). "
            "Type the YEARLY total for the whole complex, add a row if you need Garden services, then Save. "
            "Each owner pays that total ÷ 12 ÷ number of units. It is NOT split by PQ. "
            "Do NOT put Thornhill insurance here — that is the PQ Insurance column. "
            "If garden is already in Expenditure / R&M and paid from the levy, leave this at 0 or you bill twice.",
        )

    with tabs[2]:
        section_form(
            "municipal",
            "Municipal charges",
            "Gross on its own line. Recoveries: add a new row, name it ‘… recovered’, type a POSITIVE rand (no minus). Save. We subtract it. Example: Less: Sewer plant electricity recovered  335079.",
        )

    with tabs[3]:
        section_form("expenditure", "Expenditure", "Operating costs except R&M, personnel and tax. Add or delete rows as needed.")

    with tabs[4]:
        section_form(
            "rm",
            "Repair and Maintenance",
            "Type an insurance payout on the line it belongs to, then Save. Net = yearly − payout.",
            rm=True,
        )

    with tabs[5]:
        section_form("personnel", "Personnel", "Salaries, casuals, PAYE/UIF, bonuses.")

    with tabs[6]:
        section_form("tax", "Income Tax", "Most packs forget this. If the financial statements show tax, budget it here.")

    with tabs[7]:
        section_form("special", "Special Projects", "Year 1 of the 10-year plan. Tick the sidebar box only if levies must fund it.")

    with tabs[8]:
        st.subheader("PQ / levy schedule")
        st.caption("Upload the WeConnectU unit PQs on the left. Each owner = their PQ × that column’s monthly total.")
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
        paste = st.text_area("Or paste from Excel (keep the empty cells)", height=120)
        if st.button("Paste into plan") and paste.strip():
            parsed = []
            for line in paste.splitlines():
                line = line.rstrip()
                if not line.strip() or re.match(r"^(project|description|planned)", line, re.I):
                    continue
                if "\t" in line:
                    bits = line.split("\t")
                    desc = (bits[0] or "").strip()
                    nums = []
                    for p in bits[1:]:
                        if not str(p).strip():
                            nums.append(0.0)
                        else:
                            n = num(p)
                            if n is not None:
                                nums.append(n)
                else:
                    parts = [p.strip() for p in re.split(r";|,|\s{2,}", line) if p.strip()]
                    if not parts:
                        continue
                    desc = parts[0]
                    nums = [num(p) for p in parts[1:]]
                    nums = [n for n in nums if n is not None]
                if not desc or re.match(r"^(total|first|cycle|frequency|current|estimate|projects)$", desc, re.I):
                    continue
                years = ymp_years_from_nums(nums)
                first = int(nums[0]) if nums and 2020 <= nums[0] <= 2040 else ""
                freq = ""
                if nums and 2020 <= (nums[0] if nums else 0) <= 2040 and len(nums) > 1 and 1 <= nums[1] <= 15:
                    freq = int(nums[1])
                parsed.append({"desc": desc, "years": years, "first_cycle": first, "freq": freq, "estimate": 0.0})
            if parsed:
                st.session_state.ymp = parsed
                st.success(f"Loaded {len(parsed)} projects.")
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
