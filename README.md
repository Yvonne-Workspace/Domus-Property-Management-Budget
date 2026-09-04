# Body Corporate Budget Tool – Domus Property Management

Upload a WeConnectU **Actual vs Budget** Excel export → extract actuals → review in clear sections → download a complete Excel budget pack.

PDF upload remains available as a fallback, but Excel is the recommended and more reliable path (especially when using the app from another computer / Streamlit Cloud).

## Workflow

1. Export **Actual vs Budget** from WeConnectU for the complex (full year / YTD).
2. Open the app and upload that Excel file in the sidebar.
3. Click **Extract Actuals from Excel**.
4. Review and adjust % increases / budgeted amounts in the tabs.
5. Optionally upload PQ schedule and 10-Year Maintenance Plan.
6. Download the completed budget workbook (live Excel formulas).

## Tabs

1. **Income** – Levy Income, Other Income, Municipal Recovery Income  
2. **Municipal Expenses** – Gross municipal charges  
3. **Expenditure** – Operating costs (excl. R&M and Personnel)  
4. **Repair & Maintenance** – Full list + Insurance recoveries deduction  
5. **Personnel** – Salaries, casual, UIF, bonuses, etc.  
6. **Income Tax**  
7. **Special Projects** – Year 1 from the 10-Year Plan  
8. **PQ / Levy Schedule** – Upload CSV/Excel of units + PQs  
9. **10-Year Plan** – Upload or review the maintenance plan  
10. **Download Excel** – Full workbook with live formulas

## Files needed for Streamlit Cloud

- `app.py`
- `requirements.txt`
- `Body_Corporate_Budget_Template_v2.xlsx`
- `domus_logo.jpeg` (optional)

## Notes

- Yellow cells with blue text in the downloaded Excel = inputs  
- All totals and the Gross Ordinary Levy use live Excel formulas  
- Collection Rate % handles under-recovery  
- Insurance recoveries are deducted from R&M  
- PQ upload accepts CSV or Excel (columns: Unit + PQ)  
- Unmatched lines from the Excel (e.g. unique complex-specific accounts) are listed after extraction so you can add them manually
