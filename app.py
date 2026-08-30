"""
Domus Property Management – Body Corporate / HOA Budget App
Upload financial statement PDF → extract actuals → review in clear sections → download Excel
"""

import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from copy import copy

st.set_page_config(
    page_title="Domus Property Management Budget",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

TEMPLATE_PATH = Path(__file__).parent / "Body_Corporate_Budget_Template_v2.xlsx"

# ---------------------------------------------------------------------------
# PDF EXTRACTION (kept from previous version, improved mapping)
# ---------------------------------------------------------------------------

def extract_text_from_pdf(uploaded_file) -> str:
    text = ""
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text += (page.extract_text() or "") + "\n"
    return text


def parse_sa_numbers(line: str):
    comma = re.findall(r"\d{1,3}(?:,\d{3})+(?:\.\d+)?", line)
    if comma:
        return [float(n.replace(",", "")) for n in comma]
    tokens = re.findall(r"\d+", line)
    results = []
    i = 0
    while i < len(tokens):
        if i + 2 < len(tokens) and len(tokens[i]) <= 3 and len(tokens[i+1]) == 3 and len(tokens[i+2]) == 3:
            results.append(float(tokens[i] + tokens[i+1] + tokens[i+2]))
            i += 3
        elif i + 1 < len(tokens) and len(tokens[i]) <= 3 and len(tokens[i+1]) == 3:
            results.append(float(tokens[i] + tokens[i+1]))
            i += 2
        else:
            results.append(float(tokens[i]))
            i += 1
    return results


def get_detail_section(text: str) -> str:
    matches = list(re.finditer(r"Detailed\s+Income\s+Statement", text, re.I))
    if matches:
        return text[matches[-1].start():]
    for marker in ["Statement of Comprehensive Income", "Income Statement", "Operating expenses"]:
        m = re.search(marker, text, re.I)
        if m:
            return text[m.start():]
    return text


LABEL_MAP = [
    (r"Ordinary\s+levies|Levies\s+received|Levy\s+income", "ordinary_levies"),
    (r"Levy\s*[-–]?\s*Csos|Csos\s+levies|CSOS\s+levies|Csos\s+Levy", "csos_income"),
    (r"Reserve\s+fund\s+levies|Levy\s*[-–]?\s*Reserve", "reserve_levies"),
    (r"Accounting\s+fees", "accounting_fees"),
    (r"Auditors?\s*[\'’]?s?\s*remuneration|Audit\s+fees|Auditors?\s+remuneration", "audit_fees"),
    (r"Bank\s+charges", "bank_charges"),
    (r"Insurance(?!\s+(Premium|claims|recovered|recover|Claims))", "insurance"),
    (r"Management\s+fees?", "management_fee"),
    (r"Repairs\s+and\s+maintenance|Repair\s+and\s+Maintenance", "repairs_total"),
    (r"Salaries\s*&?\s*Wages|Salaries", "salaries"),
    (r"Security(?:\s*:\s*Guarding\s+service)?|Guarding\s+Service", "security"),
    (r"Legal\s+(?:expense|fees?)", "legal"),
    (r"Insurance\s+Premium\s+Recovered", "ins_premium_recovered"),
    (r"Insurance\s+claims\s+received|Insurance\s+Claims", "ins_claims_received"),
    (r"Garden\s+(?:service|expenses)", "garden"),
    (r"Electricity\s+(?:recovered|Recovered)", "electricity_recovered"),
    (r"Water\s+(?:recovered|Recovered)", "water_recovered"),
    (r"Sewerage\s+(?:recovered|Recovered)", "sewerage_recovered"),
    (r"Refuse\s+(?:recovered|Recovered)", "refuse_recovered"),
    (r"Electricity(?!\s+[Rr]ecover)", "electricity_gross"),
    (r"Water(?!\s+[Rr]ecover)", "water_gross"),
    (r"Sewerage(?!\s+[Rr]ecover)", "sewerage_gross"),
    (r"Refuse(?!\s+[Rr]ecover)", "refuse_gross"),
    (r"Utilities", "utilities"),
    (r"Property\s+valuation", "property_valuation"),
    (r"Venue\s+hire|Meeting\s+(?:expenses|refreshments)", "venue_hire"),
    (r"Telephone|Communication", "telephone"),
    (r"Office\s+expenses|General\s+[Oo]ffice", "office_expenses"),
    (r"Interest\s+on\s+arrears|Interest\s+received", "interest_arrears"),
    (r"Investment\s+[Ii]ncome|Interest\s+earned", "investment_income"),
    (r"Casual\s*/?\s*Relief\s+Wages|Casual\s+Wages", "casual_wages"),
    (r"Bonuss?es?\s*&?\s*Overtime|Bonuses", "bonuses"),
    (r"PAYE|UIF", "paye_uif"),
    (r"Cleaning\s*&?\s*Materials", "cleaning"),
    (r"Fire\s+[Ee]quipment", "fire_equipment"),
    (r"Plumbing", "plumbing"),
    (r"Gate\s*&?\s*Intercom", "gate_intercom"),
    (r"Electrical(?:\s+[Rr]epair)?", "electrical"),
]


def extract_actuals_generic(text: str) -> dict:
    detail = get_detail_section(text)
    actuals = {}
    for pattern, key in LABEL_MAP:
        if key in actuals and actuals[key] != 0:
            continue
        for line in detail.splitlines():
            if re.search(pattern, line, re.I):
                if re.search(r"total\s+net|figures\s+in|note\(s\)|page\s+\d", line, re.I):
                    continue
                nums = parse_sa_numbers(line)
                nums = [n for n in nums if 20 < n < 50_000_000]
                if nums:
                    actuals[key] = nums[0]
                    break
    for line in text.splitlines():
        if re.search(r"Reserves\s+\d", line, re.I) and "reserve_balance" not in actuals:
            nums = parse_sa_numbers(line)
            nums = [n for n in nums if n > 1000]
            if nums:
                actuals["reserve_balance"] = nums[0]
    return actuals


# ---------------------------------------------------------------------------
# DEFAULT DATA STRUCTURE (rich defaults from real examples)
# ---------------------------------------------------------------------------

def default_sections():
    return {
        "levy_income": [
            {"desc": "Ordinary Levies (Gross required)", "gl": "1000/001", "actual": 0.0, "pct": 0.0, "note": "Calculated automatically"},
            {"desc": "Reserve Fund Contribution", "gl": "RFI", "actual": 0.0, "pct": 0.0, "note": "Set by trustees / 10YMP"},
            {"desc": "CSOS Levy (Income)", "gl": "1000/017", "actual": 0.0, "pct": 0.0, "note": ""},
        ],
        "other_income": [
            {"desc": "Interest on Arrears", "gl": "1000/004", "actual": 0.0, "pct": 0.0, "note": ""},
            {"desc": "Investment Income – Bank/Investments", "gl": "1000/003", "actual": 0.0, "pct": 0.0, "note": ""},
            {"desc": "Penalty Levy", "gl": "", "actual": 0.0, "pct": 0.0, "note": ""},
            {"desc": "Other Income 1", "gl": "", "actual": 0.0, "pct": 0.0, "note": ""},
            {"desc": "Other Income 2", "gl": "", "actual": 0.0, "pct": 0.0, "note": ""},
        ],
        "muni_recoveries": [
            {"desc": "Electricity Recovered", "gl": "1000/013", "actual": 0.0, "pct": 0.0, "note": ""},
            {"desc": "Water Recovered", "gl": "1000/011", "actual": 0.0, "pct": 0.0, "note": ""},
            {"desc": "Sewerage Recovered", "gl": "1000/012", "actual": 0.0, "pct": 0.0, "note": ""},
            {"desc": "Refuse Recovered", "gl": "1000/014", "actual": 0.0, "pct": 0.0, "note": ""},
            {"desc": "Other Recoveries", "gl": "", "actual": 0.0, "pct": 0.0, "note": ""},
        ],
        "municipal": [
            {"desc": "Water", "gl": "2100/001", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Sewerage", "gl": "2100/003", "actual": 0.0, "pct": 5.0, "note": ""},
            {"desc": "Electricity (Common / Gross)", "gl": "2100/004", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Refuse", "gl": "2100/002", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Rates / Property Tax", "gl": "2100/005", "actual": 0.0, "pct": 5.0, "note": ""},
            {"desc": "Other Municipal", "gl": "", "actual": 0.0, "pct": 10.0, "note": ""},
        ],
        "expenditure": [
            {"desc": "Accounting Fees", "gl": "3000/001", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Audit Fees", "gl": "2000/011", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Bank Charges", "gl": "2000/001", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Insurance", "gl": "2000/016", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "CSOS Levies (Expense)", "gl": "2000/002", "actual": 0.0, "pct": 0.0, "note": ""},
            {"desc": "Management Fee", "gl": "2000/002", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Legal & Professional Fees", "gl": "2000/007", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Meeting / Venue Expenses", "gl": "", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Telephone / Communications", "gl": "", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Office / General Expenses", "gl": "", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Property Valuation", "gl": "2000/023", "actual": 0.0, "pct": 0.0, "note": ""},
            {"desc": "Security / Guarding", "gl": "2000/015", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Garden Service (contract)", "gl": "", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Cleaning & Materials", "gl": "", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Health & Safety", "gl": "", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Keys & Remotes", "gl": "", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Computer / IT Expenses", "gl": "", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Printing & Stationery", "gl": "", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Other Expenditure 1", "gl": "", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Other Expenditure 2", "gl": "", "actual": 0.0, "pct": 10.0, "note": ""},
        ],
        "rm": [
            {"desc": "Electrical", "gl": "2200/005", "actual": 0.0, "pct": 5.0, "note": ""},
            {"desc": "Fire Equipment", "gl": "2200/001", "actual": 0.0, "pct": 5.0, "note": ""},
            {"desc": "General / Buildings Maintenance", "gl": "2200/002", "actual": 0.0, "pct": 0.0, "note": ""},
            {"desc": "Garden Expenses", "gl": "2200/009", "actual": 0.0, "pct": 5.0, "note": ""},
            {"desc": "Gate & Intercom", "gl": "2200/004", "actual": 0.0, "pct": 5.0, "note": ""},
            {"desc": "Plumbing", "gl": "2200/003", "actual": 0.0, "pct": 5.0, "note": ""},
            {"desc": "Painting / Waterproofing", "gl": "", "actual": 0.0, "pct": 5.0, "note": ""},
            {"desc": "Roofs & Gutters", "gl": "", "actual": 0.0, "pct": 5.0, "note": ""},
            {"desc": "Pool", "gl": "2200/012", "actual": 0.0, "pct": 5.0, "note": ""},
            {"desc": "Electric Fence", "gl": "2200/006", "actual": 0.0, "pct": 5.0, "note": ""},
            {"desc": "CCTV / Cameras", "gl": "", "actual": 0.0, "pct": 5.0, "note": ""},
            {"desc": "Paving / Roadways", "gl": "", "actual": 0.0, "pct": 5.0, "note": ""},
            {"desc": "Lifts", "gl": "", "actual": 0.0, "pct": 5.0, "note": ""},
            {"desc": "Other R&M 1", "gl": "", "actual": 0.0, "pct": 5.0, "note": ""},
            {"desc": "Other R&M 2", "gl": "", "actual": 0.0, "pct": 5.0, "note": ""},
        ],
        "personnel": [
            {"desc": "Salaries & Wages", "gl": "4000/001", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Casual / Relief Wages", "gl": "4000/003", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "PAYE / UIF / SDL Contributions", "gl": "4000/005", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Travel / Allowances", "gl": "", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Bonuses & Overtime", "gl": "4000/007", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Staff Welfare / Protective Clothing", "gl": "", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Pension / Provident Fund", "gl": "", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "WCA / COIDA", "gl": "", "actual": 0.0, "pct": 10.0, "note": ""},
            {"desc": "Other Personnel Costs", "gl": "", "actual": 0.0, "pct": 10.0, "note": ""},
        ],
        "tax": [
            {"desc": "Taxation Payable", "gl": "4700/001", "actual": 0.0, "pct": 0.0, "note": "Usually based on taxable investment income"},
        ],
        "special": [
            {"desc": "Special Project 1 (from 10YMP)", "gl": "", "actual": 0.0, "pct": 0.0, "note": ""},
            {"desc": "Special Project 2", "gl": "", "actual": 0.0, "pct": 0.0, "note": ""},
            {"desc": "Special Project 3", "gl": "", "actual": 0.0, "pct": 0.0, "note": ""},
            {"desc": "Special Project 4", "gl": "", "actual": 0.0, "pct": 0.0, "note": ""},
            {"desc": "Special Project 5", "gl": "", "actual": 0.0, "pct": 0.0, "note": ""},
        ],
    }


def apply_actuals(sections: dict, actuals: dict) -> dict:
    """Map extracted keys into the section lists."""
    # Levy / Income
    for item in sections["levy_income"]:
        if "Ordinary" in item["desc"] and actuals.get("ordinary_levies"):
            item["actual"] = actuals["ordinary_levies"]
        if "CSOS" in item["desc"] and actuals.get("csos_income"):
            item["actual"] = actuals["csos_income"]
        if "Reserve" in item["desc"] and actuals.get("reserve_levies"):
            item["actual"] = actuals["reserve_levies"]

    for item in sections["other_income"]:
        if "Interest on Arrears" in item["desc"] and actuals.get("interest_arrears"):
            item["actual"] = actuals["interest_arrears"]
        if "Investment" in item["desc"] and actuals.get("investment_income"):
            item["actual"] = actuals["investment_income"]

    # Municipal recoveries
    for item in sections["muni_recoveries"]:
        if "Electricity" in item["desc"] and actuals.get("electricity_recovered"):
            item["actual"] = actuals["electricity_recovered"]
        if "Water" in item["desc"] and actuals.get("water_recovered"):
            item["actual"] = actuals["water_recovered"]
        if "Sewerage" in item["desc"] and actuals.get("sewerage_recovered"):
            item["actual"] = actuals["sewerage_recovered"]
        if "Refuse" in item["desc"] and actuals.get("refuse_recovered"):
            item["actual"] = actuals["refuse_recovered"]

    # Municipal expenses (gross)
    for item in sections["municipal"]:
        if "Water" in item["desc"] and actuals.get("water_gross"):
            item["actual"] = actuals["water_gross"]
        if "Electricity" in item["desc"] and actuals.get("electricity_gross"):
            item["actual"] = actuals["electricity_gross"]
        if "Sewerage" in item["desc"] and actuals.get("sewerage_gross"):
            item["actual"] = actuals["sewerage_gross"]
        if "Refuse" in item["desc"] and actuals.get("refuse_gross"):
            item["actual"] = actuals["refuse_gross"]

    # Expenditure
    op_map = {
        "Accounting Fees": "accounting_fees",
        "Audit Fees": "audit_fees",
        "Bank Charges": "bank_charges",
        "Insurance": "insurance",
        "CSOS Levies (Expense)": "csos_income",
        "Management Fee": "management_fee",
        "Legal & Professional Fees": "legal",
        "Meeting / Venue Expenses": "venue_hire",
        "Telephone / Communications": "telephone",
        "Office / General Expenses": "office_expenses",
        "Property Valuation": "property_valuation",
        "Security / Guarding": "security",
        "Garden Service (contract)": "garden",
        "Cleaning & Materials": "cleaning",
    }
    for item in sections["expenditure"]:
        k = op_map.get(item["desc"])
        if k and actuals.get(k):
            item["actual"] = actuals[k]

    # Personnel
    for item in sections["personnel"]:
        if item["desc"].startswith("Salaries") and actuals.get("salaries"):
            item["actual"] = actuals["salaries"]
        if "Casual" in item["desc"] and actuals.get("casual_wages"):
            item["actual"] = actuals["casual_wages"]
        if "Bonuses" in item["desc"] and actuals.get("bonuses"):
            item["actual"] = actuals["bonuses"]
        if "PAYE" in item["desc"] and actuals.get("paye_uif"):
            item["actual"] = actuals["paye_uif"]

    # R&M – prefer specific lines, otherwise total into General
    rm_map = {
        "Electrical": "electrical",
        "Fire Equipment": "fire_equipment",
        "Plumbing": "plumbing",
        "Gate & Intercom": "gate_intercom",
        "Garden Expenses": "garden",
    }
    for item in sections["rm"]:
        k = rm_map.get(item["desc"])
        if k and actuals.get(k):
            item["actual"] = actuals[k]
    if actuals.get("repairs_total"):
        for item in sections["rm"]:
            if "General" in item["desc"] and not item.get("actual"):
                item["actual"] = actuals["repairs_total"]
                break

    return sections


# ---------------------------------------------------------------------------
# EXCEL GENERATION
# ---------------------------------------------------------------------------

def generate_excel(
    complex_name, fin_year, collection_rate, reserve_opening,
    sections, insurance_recoveries, pq_df=None, ymp_df=None
) -> BytesIO:
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb["BUDGET"]

    # Header
    ws["C3"] = complex_name
    ws["F3"] = fin_year
    ws["D7"] = collection_rate
    ws["D8"] = reserve_opening

    # Helper to write a list of items into a block of rows
    def write_block(start_row, items, max_rows):
        for i, item in enumerate(items):
            if i >= max_rows:
                break
            r = start_row + i
            ws.cell(row=r, column=2, value=item.get("desc", ""))
            ws.cell(row=r, column=3, value=item.get("gl", ""))
            ws.cell(row=r, column=4, value=float(item.get("actual") or 0))
            ws.cell(row=r, column=5, value=float(item.get("pct") or 0))
            # note
            if item.get("note"):
                ws.cell(row=r, column=8, value=item["note"])

    # Levy income (rows 14-16) – special handling
    levy = sections.get("levy_income", [])
    # Ordinary stays formula-driven; we only set Actual for reference
    if len(levy) > 0:
        ws["D14"] = float(levy[0].get("actual") or 0)
    if len(levy) > 1:
        ws["D15"] = float(levy[1].get("actual") or 0)
        ws["F15"] = float(levy[1].get("actual") or 0)  # budgeted reserve contribution
    if len(levy) > 2:
        ws["D16"] = float(levy[2].get("actual") or 0)
        ws["F16"] = float(levy[2].get("actual") or 0)  # or apply % if needed

    # Other Income rows 19-23
    write_block(19, sections.get("other_income", []), 5)

    # Municipal Recoveries rows 26-30
    write_block(26, sections.get("muni_recoveries", []), 5)

    # Municipal Expenses rows 36-41
    write_block(36, sections.get("municipal", []), 6)

    # Expenditure rows 46-65
    write_block(46, sections.get("expenditure", []), 20)

    # R&M rows 70-84
    write_block(70, sections.get("rm", []), 15)

    # Insurance recoveries (row 86)
    ws["D86"] = float(insurance_recoveries or 0)

    # Personnel rows 91-99
    write_block(91, sections.get("personnel", []), 9)

    # Tax row 104
    tax = sections.get("tax", [])
    if tax:
        ws["D104"] = float(tax[0].get("actual") or 0)
        ws["E104"] = float(tax[0].get("pct") or 0)
        if tax[0].get("note"):
            ws["H104"] = tax[0]["note"]

    # Special Projects rows 109-113
    write_block(109, sections.get("special", []), 5)

    # ----- PQ sheet -----
    if pq_df is not None and not pq_df.empty:
        ws_pq = wb["PQ"]
        # Clear example rows 11-30
        for r in range(11, 31):
            for c in range(1, 9):
                ws_pq.cell(row=r, column=c).value = None

        for i, row in pq_df.iterrows():
            r = 11 + i
            if r > 200:  # safety
                break
            ws_pq.cell(row=r, column=1, value=i + 1)
            ws_pq.cell(row=r, column=2, value=str(row.get("Unit", row.get("unit", f"UNIT-{i+1}"))))
            pq_val = float(row.get("PQ", row.get("pq", 0)) or 0)
            ws_pq.cell(row=r, column=3, value=pq_val)
            ws_pq.cell(row=r, column=3).number_format = "0.000000"
            # Formulas
            ws_pq.cell(row=r, column=4, value=f"=$D$5*C{r}")
            ws_pq.cell(row=r, column=5, value=f"=$D$6*C{r}")
            ws_pq.cell(row=r, column=6, value=f"=$D$7*C{r}")
            ws_pq.cell(row=r, column=7, value=f"=D{r}+E{r}+F{r}")
            note = row.get("Notes", row.get("notes", ""))
            if note:
                ws_pq.cell(row=r, column=8, value=str(note))

        last_data_row = 10 + len(pq_df)
        # Update TOTAL formulas
        total_row = last_data_row + 1
        ws_pq.cell(row=total_row, column=2, value="TOTAL")
        ws_pq.cell(row=total_row, column=3, value=f"=SUM(C11:C{last_data_row})")
        ws_pq.cell(row=total_row, column=4, value=f"=SUM(D11:D{last_data_row})")
        ws_pq.cell(row=total_row, column=5, value=f"=SUM(E11:E{last_data_row})")
        ws_pq.cell(row=total_row, column=6, value=f"=SUM(F11:F{last_data_row})")
        ws_pq.cell(row=total_row, column=7, value=f"=SUM(G11:G{last_data_row})")

        # PQ check
        ws_pq["C33"] = f"=C{total_row}"
        ws_pq["D33"] = f'=IF(ABS(C33-1)<0.00001,"OK - PQ totals 1","CHECK - PQ does not total 1")'

    # ----- 10 YMP sheet (optional) -----
    if ymp_df is not None and not ymp_df.empty:
        ws_ymp = wb["10 YMP"]
        for i, row in ymp_df.iterrows():
            r = 7 + i
            if r > 16:
                break
            ws_ymp.cell(row=r, column=1, value=str(row.get("Description", row.get("desc", ""))))
            if "First Cycle" in row:
                ws_ymp.cell(row=r, column=2, value=row.get("First Cycle", 1))
            if "Frequency (yrs)" in row or "Frequency" in row:
                ws_ymp.cell(row=r, column=3, value=row.get("Frequency (yrs)", row.get("Frequency", 1)))
            if "Current Estimate" in row or "Estimate" in row:
                ws_ymp.cell(row=r, column=4, value=float(row.get("Current Estimate", row.get("Estimate", 0)) or 0))
            # Year columns if present
            for y_idx, y_col in enumerate([f"Year {j+1}" for j in range(10)]):
                if y_col in row:
                    ws_ymp.cell(row=r, column=5 + y_idx, value=float(row.get(y_col, 0) or 0))

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


# ---------------------------------------------------------------------------
# STREAMLIT UI
# ---------------------------------------------------------------------------

def main():
    logo_path = Path(__file__).parent / "domus_logo.jpeg"
    if logo_path.exists():
        col_logo, col_title = st.columns([1, 4])
        with col_logo:
            st.image(str(logo_path), width=140)
        with col_title:
            st.title("Domus Property Management Budget")
            st.caption("Upload financial statement PDF → extract actuals → review → download Excel")
    else:
        st.title("Domus Property Management Budget")
        st.caption("Upload financial statement PDF → extract actuals → review → download Excel")

    # Session state defaults
    defaults = {
        "sections": default_sections(),
        "actuals": {},
        "complex_name": "",
        "fin_year": "01-03-2026 / 28-02-2027",
        "collection_rate": 95.0,
        "reserve_opening": 0.0,
        "insurance_recoveries": 0.0,
        "pq_df": None,
        "ymp_df": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

    # Sidebar
    with st.sidebar:
        st.header("1. Complex")
        st.session_state.complex_name = st.text_input("Complex Name", st.session_state.complex_name)
        st.session_state.fin_year = st.text_input("Financial Year", st.session_state.fin_year)

        st.header("2. Upload PDF")
        uploaded = st.file_uploader("Financial statement (PDF)", type=["pdf"])

        if uploaded and st.button("Extract Actuals from PDF", type="primary"):
            with st.spinner("Reading PDF…"):
                text = extract_text_from_pdf(uploaded)
                actuals = extract_actuals_generic(text)
                st.session_state.actuals = actuals
                st.session_state.sections = apply_actuals(default_sections(), actuals)

                if actuals.get("ordinary_levies"):
                    for item in st.session_state.sections["levy_income"]:
                        if "Ordinary" in item["desc"]:
                            item["actual"] = actuals["ordinary_levies"]
                if actuals.get("csos_income"):
                    for item in st.session_state.sections["levy_income"]:
                        if "CSOS" in item["desc"]:
                            item["actual"] = actuals["csos_income"]
                if actuals.get("reserve_balance"):
                    st.session_state.reserve_opening = actuals["reserve_balance"]

                rec = (actuals.get("ins_premium_recovered") or 0) + (actuals.get("ins_claims_received") or 0)
                if rec:
                    st.session_state.insurance_recoveries = abs(rec)

                # Clear cached dataframes so the editor shows the newly extracted values
                for k in list(st.session_state.keys()):
                    if k.startswith("df_"):
                        del st.session_state[k]

                st.success("Extraction complete. Review the numbers in the tabs.")
                st.rerun()

        st.header("3. Key Assumptions")
        st.session_state.collection_rate = st.number_input(
            "Expected Collection Rate %", 50.0, 100.0,
            float(st.session_state.collection_rate), 0.5,
            help="95 = budget for 5% under-recovery"
        )
        st.session_state.reserve_opening = st.number_input(
            "Opening Reserve Balance", 0.0,
            value=float(st.session_state.reserve_opening), step=1000.0, format="%.2f"
        )
        st.session_state.insurance_recoveries = st.number_input(
            "Insurance Recoveries (deduct from R&M)", 0.0,
            value=float(st.session_state.insurance_recoveries), step=100.0, format="%.2f"
        )

    # Main tabs
    tabs = st.tabs([
        "💰 Income",
        "🏢 Municipal Expenses",
        "📋 Expenditure",
        "🔧 Repair & Maintenance",
        "👥 Personnel",
        "🧾 Income Tax",
        "🏗️ Special Projects",
        "📐 PQ / Levy Schedule",
        "📅 10-Year Plan",
        "📥 Download Excel",
    ])

    def edit_section(key, title, help_text=None, show_budgeted=True):
        st.subheader(title)
        if help_text:
            st.caption(help_text)

        df_key = f"df_{key}"
        if df_key not in st.session_state:
            base = pd.DataFrame(st.session_state.sections[key])
            # Ensure budgeted column exists
            if "budgeted" not in base.columns:
                base["budgeted"] = base.apply(
                    lambda r: float(r.get("actual") or 0) * (1 + float(r.get("pct") or 0) / 100), axis=1
                )
            st.session_state[df_key] = base

        df = st.session_state[df_key].copy()

        # Ensure required columns
        for col, default in [("desc", ""), ("gl", ""), ("actual", 0.0), ("pct", 0.0), ("budgeted", 0.0), ("note", "")]:
            if col not in df.columns:
                df[col] = default

        # Recalculate budgeted from actual+% where budgeted was never manually set
        # We store a flag column _manual_budgeted if user overrode
        if "_manual_budgeted" not in df.columns:
            df["_manual_budgeted"] = False

        # Live calculated columns for display
        calc_budgeted = df.apply(
            lambda r: float(r.get("actual") or 0) * (1 + float(r.get("pct") or 0) / 100), axis=1
        )
        # Where user has not manually overridden, keep budgeted in sync with formula
        mask = ~df["_manual_budgeted"].astype(bool)
        df.loc[mask, "budgeted"] = calc_budgeted[mask]

        # Implied % when budgeted was set manually
        def implied_pct(row):
            act = float(row.get("actual") or 0)
            bud = float(row.get("budgeted") or 0)
            if act == 0:
                return 0.0 if bud == 0 else 100.0
            return (bud / act - 1) * 100

        df["implied_pct"] = df.apply(implied_pct, axis=1)
        df["monthly"] = df["budgeted"].astype(float) / 12

        # Column order for editor
        display_cols = ["desc", "gl", "actual", "pct", "budgeted", "monthly", "implied_pct", "note"]
        df_display = df[display_cols].copy()

        edited = st.data_editor(
            df_display,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "desc": st.column_config.TextColumn("Description", width="medium"),
                "gl": st.column_config.TextColumn("GL Code", width="small"),
                "actual": st.column_config.NumberColumn("Actual", format="%.2f"),
                "pct": st.column_config.NumberColumn("% Increase", format="%.1f",
                    help="Change this to recalculate Budgeted Yearly"),
                "budgeted": st.column_config.NumberColumn("Budgeted Yearly", format="%.2f",
                    help="You can type an amount directly. Implied % will update."),
                "monthly": st.column_config.NumberColumn("Monthly", format="%.2f", disabled=True),
                "implied_pct": st.column_config.NumberColumn("Implied %", format="%.1f", disabled=True,
                    help="Shows what % increase the Budgeted Yearly represents vs Actual"),
                "note": st.column_config.TextColumn("Notes"),
            },
            key=f"ed_{key}",
            disabled=["monthly", "implied_pct"],
        )

        # Detect manual budgeted overrides: if budgeted differs from formula result, mark as manual
        records = []
        manual_flags = []
        for i, row in edited.iterrows():
            act = float(row.get("actual") or 0)
            pct = float(row.get("pct") or 0)
            bud = float(row.get("budgeted") or 0)
            formula_bud = act * (1 + pct / 100)
            # If user changed budgeted away from formula, treat as manual
            is_manual = abs(bud - formula_bud) > 0.02
            # If they also changed pct to match, clear manual
            if is_manual and abs(implied_pct({"actual": act, "budgeted": bud}) - pct) < 0.05:
                is_manual = False
            records.append({
                "desc": str(row.get("desc") or ""),
                "gl": str(row.get("gl") or ""),
                "actual": act,
                "pct": pct if not is_manual else implied_pct({"actual": act, "budgeted": bud}),
                "budgeted": bud,
                "note": str(row.get("note") or ""),
                "_manual_budgeted": is_manual,
            })
            manual_flags.append(is_manual)

        out_df = pd.DataFrame(records)
        st.session_state[df_key] = out_df
        # Store clean records (without internal flag) into sections for Excel export
        clean_records = []
        for r in records:
            clean_records.append({
                "desc": r["desc"],
                "gl": r["gl"],
                "actual": r["actual"],
                "pct": r["pct"],
                "budgeted": r["budgeted"],
                "note": r["note"],
            })
        st.session_state.sections[key] = clean_records

    # ---- Tab 0: Income ----
    with tabs[0]:
        st.markdown("### Levy Income")
        st.caption("Ordinary Levies are calculated automatically so that after under-recovery you still cover everything. You can still enter the Actual for reference.")
        edit_section("levy_income", "Levy Income")
        st.divider()
        edit_section("other_income", "Other Income", "Interest, investment income, penalties, etc.")
        st.divider()
        edit_section("muni_recoveries", "Municipal Recovery Income", "Recoveries shown here under Income (gross municipal expenses are on the next tab).")

    # ---- Tab 1: Municipal Expenses ----
    with tabs[1]:
        edit_section("municipal", "Municipal Expenses (Gross)",
                     "Enter the full municipal charges. Recoveries are recorded separately under Income.")

    # ---- Tab 2: Expenditure ----
    with tabs[2]:
        edit_section("expenditure", "Expenditure (Operating)",
                     "All operating costs except Repair & Maintenance and Personnel. Add or delete rows as needed.")

    # ---- Tab 3: R&M ----
    with tabs[3]:
        edit_section("rm", "Repair & Maintenance")
        st.caption(f"Insurance recoveries currently set to **R {st.session_state.insurance_recoveries:,.2f}** (change in sidebar). This amount is deducted from the R&M total.")

    # ---- Tab 4: Personnel ----
    with tabs[4]:
        edit_section("personnel", "Personnel",
                     "Salaries, casual wages, UIF, bonuses, travel, etc. Add or delete rows as needed.")

    # ---- Tab 5: Income Tax ----
    with tabs[5]:
        edit_section("tax", "Income Tax",
                     "Usually based on taxable investment income.")

    # ---- Tab 6: Special Projects ----
    with tabs[6]:
        edit_section("special", "Special Projects (Year 1 from 10 YMP)",
                     "Enter the Year 1 amounts from the 10-Year Maintenance Plan. These are usually funded from the reserve or special levies.")

    # ---- Tab 7: PQ / Levy Schedule ----
    with tabs[7]:
        st.subheader("PQ / Levy Schedule")
        st.caption("Upload a CSV or Excel file with columns: Unit (or Unit / Owner Code) and PQ. The app will calculate the monthly levy for each unit.")

        pq_file = st.file_uploader("Upload PQ / Ratio file (CSV or Excel)", type=["csv", "xlsx", "xls"], key="pq_upload")
        if pq_file:
            try:
                if pq_file.name.lower().endswith(".csv"):
                    df = pd.read_csv(pq_file)
                else:
                    df = pd.read_excel(pq_file)

                # Flatten multi-index columns if present
                if isinstance(df.columns, pd.MultiIndex):
                    df.columns = [" ".join(str(x) for x in col if str(x) != "nan").strip() for col in df.columns]

                # Make column names unique and clean
                df.columns = [str(c).strip() for c in df.columns]
                # Deduplicate column names
                seen = {}
                new_cols = []
                for c in df.columns:
                    if c in seen:
                        seen[c] += 1
                        new_cols.append(f"{c}_{seen[c]}")
                    else:
                        seen[c] = 0
                        new_cols.append(c)
                df.columns = new_cols

                # Map to standard names – prefer meaningful unit codes over pure numbers
                unit_candidates = []
                pq_col = None
                notes_col = None
                for c in df.columns:
                    cl = str(c).lower().strip()
                    if "unit" in cl or "owner" in cl or "code" in cl or cl in ("description", "name"):
                        unit_candidates.append(c)
                    elif pq_col is None and (cl in ("pq", "participation quota", "ratio", "share") or "quota" in cl):
                        pq_col = c
                    elif notes_col is None and "note" in cl:
                        notes_col = c

                unit_col = None
                if unit_candidates:
                    # Prefer a column whose values contain letters (real unit codes)
                    for c in unit_candidates:
                        sample = df[c].astype(str).head(20)
                        if sample.str.contains(r"[A-Za-z]", regex=True).any():
                            unit_col = c
                            break
                    if unit_col is None:
                        unit_col = unit_candidates[0]

                if unit_col is None or pq_col is None:
                    st.error("File must contain columns for Unit (or Owner/Code) and PQ (or Ratio). Found columns: " + ", ".join(df.columns.astype(str)))
                    st.write("Raw preview of uploaded file:")
                    st.dataframe(df.head(10), use_container_width=True)
                else:
                    clean = pd.DataFrame({
                        "Unit": df[unit_col].astype(str).str.strip(),
                        "PQ": pd.to_numeric(df[pq_col], errors="coerce").fillna(0.0),
                    })
                    if notes_col is not None:
                        clean["Notes"] = df[notes_col].astype(str)
                    # Drop empty / nan unit rows
                    clean = clean[clean["Unit"].str.len() > 0]
                    clean = clean[clean["Unit"].str.lower() != "nan"].reset_index(drop=True)

                    pq_sum = clean["PQ"].sum()
                    # Auto-detect percentage style PQ (sums to ~100) vs decimal (sums to ~1)
                    if 50 < pq_sum < 150:
                        clean["PQ"] = clean["PQ"] / 100.0
                        st.info(f"PQ values looked like percentages (total was {pq_sum:.2f}). Divided by 100 so they now total ~1.000.")
                    elif pq_sum > 150:
                        st.warning(f"PQ total is {pq_sum:.4f} — this is unusually high. Check that the correct PQ column was detected.")

                    st.session_state.pq_df = clean
                    st.success(f"Loaded {len(clean)} units. PQ total = {clean['PQ'].sum():.6f}")
                    st.caption(f"Detected columns → Unit: **{unit_col}**  |  PQ: **{pq_col}**")
                    st.dataframe(clean, use_container_width=True)
            except Exception as e:
                st.error(f"Could not read file: {e}")

        if st.session_state.pq_df is not None and not st.session_state.pq_df.empty:
            st.markdown("#### Preview of calculated monthly levies (approximate)")

            def budgeted(items):
                total = 0.0
                for i in items:
                    if i.get("budgeted") is not None and i.get("budgeted") != "":
                        total += float(i.get("budgeted") or 0)
                    else:
                        total += float(i.get("actual") or 0) * (1 + float(i.get("pct") or 0) / 100)
                return total

            total_exp = (
                budgeted(st.session_state.sections["municipal"])
                + budgeted(st.session_state.sections["expenditure"])
                + budgeted(st.session_state.sections["rm"]) - st.session_state.insurance_recoveries
                + budgeted(st.session_state.sections["personnel"])
                + budgeted(st.session_state.sections["tax"])
                + budgeted(st.session_state.sections["special"])
            )
            reserve = next((i.get("actual") or 0 for i in st.session_state.sections["levy_income"] if "Reserve" in i["desc"]), 0)
            other = budgeted(st.session_state.sections["other_income"]) + budgeted(st.session_state.sections["muni_recoveries"])
            to_collect = total_exp + reserve - other
            rate = st.session_state.collection_rate / 100 or 1
            gross_ordinary = to_collect / rate if rate else 0
            monthly_ord = gross_ordinary / 12
            monthly_res = reserve / 12
            csos = next((i.get("actual") or 0 for i in st.session_state.sections["levy_income"] if "CSOS" in i["desc"]), 0)
            monthly_csos = csos / 12

            base = st.session_state.pq_df[["Unit", "PQ"]].copy()
            preview = pd.DataFrame({
                "Unit": base["Unit"].values,
                "PQ": base["PQ"].values,
                "Ordinary Levy": (base["PQ"] * monthly_ord).values,
                "Reserve Fund": (base["PQ"] * monthly_res).values,
                "CSOS": (base["PQ"] * monthly_csos).values,
            })
            preview["Total Monthly"] = preview["Ordinary Levy"] + preview["Reserve Fund"] + preview["CSOS"]

            # Format as strings to avoid styler / arrow issues
            display = preview.copy()
            display["PQ"] = display["PQ"].map(lambda x: f"{x:.6f}")
            display["Ordinary Levy"] = display["Ordinary Levy"].map(lambda x: f"R {x:,.2f}")
            display["Reserve Fund"] = display["Reserve Fund"].map(lambda x: f"R {x:,.2f}")
            display["CSOS"] = display["CSOS"].map(lambda x: f"R {x:,.2f}")
            display["Total Monthly"] = display["Total Monthly"].map(lambda x: f"R {x:,.2f}")
            st.dataframe(display, use_container_width=True)
            st.info("Exact figures will be calculated with live Excel formulas when you download the file.")

    # ---- Tab 8: 10-Year Plan ----
    with tabs[8]:
        st.subheader("10-Year Maintenance Plan")
        st.caption("Edit the plan here or upload a file. Year 1 totals should be entered into the Special Projects tab.")

        ymp_file = st.file_uploader("Upload 10 YMP (Excel)", type=["xlsx", "xls"], key="ymp_upload")
        if ymp_file:
            try:
                df = pd.read_excel(ymp_file)
                st.session_state.ymp_df = df
                st.success("10 YMP loaded.")
                st.dataframe(df, use_container_width=True)
            except Exception as e:
                st.error(f"Could not read file: {e}")

        if st.session_state.ymp_df is not None:
            st.dataframe(st.session_state.ymp_df, use_container_width=True)
        else:
            st.info("No 10 YMP uploaded yet. You can still enter Year 1 projects manually on the Special Projects tab. The downloaded Excel will contain a blank 10 YMP sheet you can fill in.")

    # ---- Tab 9: Download ----
    with tabs[9]:
        st.subheader("Generate Excel budget pack")

        def budgeted(items):
            total = 0.0
            for i in items:
                if i.get("budgeted") is not None and i.get("budgeted") != "":
                    total += float(i.get("budgeted") or 0)
                else:
                    total += float(i.get("actual") or 0) * (1 + float(i.get("pct") or 0) / 100)
            return total

        total_muni = budgeted(st.session_state.sections["municipal"])
        total_exp = budgeted(st.session_state.sections["expenditure"])
        total_rm = budgeted(st.session_state.sections["rm"]) - st.session_state.insurance_recoveries
        total_pers = budgeted(st.session_state.sections["personnel"])
        total_tax = budgeted(st.session_state.sections["tax"])
        total_sp = budgeted(st.session_state.sections["special"])
        other_inc = budgeted(st.session_state.sections["other_income"]) + budgeted(st.session_state.sections["muni_recoveries"])
        reserve = next((i.get("actual") or 0 for i in st.session_state.sections["levy_income"] if "Reserve" in i["desc"]), 0)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Municipal (budgeted)", f"R {total_muni:,.0f}")
        c2.metric("Expenditure (budgeted)", f"R {total_exp:,.0f}")
        c3.metric("R&M Net (budgeted)", f"R {total_rm:,.0f}")
        c4.metric("Personnel (budgeted)", f"R {total_pers:,.0f}")

        c1.metric("Income Tax", f"R {total_tax:,.0f}")
        c2.metric("Special Projects", f"R {total_sp:,.0f}")
        c3.metric("Other + Recoveries Income", f"R {other_inc:,.0f}")
        c4.metric("Collection Rate", f"{st.session_state.collection_rate}%")

        to_collect = total_muni + total_exp + total_rm + total_pers + total_tax + total_sp + reserve - other_inc
        rate = st.session_state.collection_rate / 100 or 1
        gross_ordinary = to_collect / rate if rate else 0

        st.markdown("---")
        st.metric("Estimated Gross Ordinary Levies Required (yearly)", f"R {gross_ordinary:,.0f}")
        st.caption("This figure will be calculated exactly by the Excel formulas (including all live % increases and links).")

        if st.button("📥 Generate Excel File", type="primary"):
            if not st.session_state.complex_name:
                st.error("Please enter a Complex Name first.")
            else:
                with st.spinner("Building Excel…"):
                    excel = generate_excel(
                        st.session_state.complex_name,
                        st.session_state.fin_year,
                        st.session_state.collection_rate,
                        st.session_state.reserve_opening,
                        st.session_state.sections,
                        st.session_state.insurance_recoveries,
                        st.session_state.pq_df,
                        st.session_state.ymp_df,
                    )
                    st.download_button(
                        "Download Budget Excel",
                        data=excel,
                        file_name=f"Budget_{st.session_state.complex_name.replace(' ', '_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                    st.success("Ready — click the download button. The file contains live formulas so you can still edit any yellow cell in Excel.")

    st.markdown("---")
    st.caption("Yellow cells in the downloaded Excel are inputs. All totals and the Gross Ordinary Levy are calculated with live Excel formulas. Always review the numbers before finalising.")


if __name__ == "__main__":
    main()
