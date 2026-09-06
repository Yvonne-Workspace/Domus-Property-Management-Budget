# Domus Property Management – Body Corporate Budget

Streamlit app: upload **WeConnectU Actual vs Budget Excel** (preferred) or a financial-statement PDF, review sections, download a full Excel budget pack.

## Streamlit Cloud files

Upload these to GitHub (repo root):

- `app.py`
- `requirements.txt`
- `Body_Corporate_Budget_Template_v2.xlsx`
- `domus_logo.jpeg`

## Tabs

1. Income (levies, other, municipal recoveries)
2. Municipal expenses (gross)
3. Expenditure
4. Repair & Maintenance (insurance recoveries deducted)
5. Personnel
6. Income Tax
7. Special Projects
8. PQ / Levy schedule
9. 10-year plan
10. Download Excel

## WeConnectU

In WeConnectU: Budgeting and Actuals → Options → **Budget and Actuals**.  
The app reads **YTD Actual** and **TOTAL Budget**, matches GL codes, and adds extra lines unique to that complex (e.g. Eskom fixed charge).

## Collection rate (95% / 100%)

Not a tax. 95% means you expect 5% of billed levies not to be paid, so you bill a little more so the people who *do* pay still cover the bills.

## Columns

Actual, % Increase, Budgeted yearly, Monthly.  
Type either % or the yearly rand amount (quotes / Eskom). The other updates.
