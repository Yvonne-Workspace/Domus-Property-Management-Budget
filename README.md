# Body Corporate Budget Tool – Domus Property Management

Upload a financial statement PDF → extract actuals → review in clear sections → download a complete Excel budget pack.

## New structure (tabs)

1. **Income** – Levy Income, Other Income, Municipal Recovery Income  
2. **Municipal Expenses** – Gross municipal charges  
3. **Expenditure** – All operating costs except R&M and Personnel  
4. **Repair & Maintenance** – Full list + Insurance recoveries deduction  
5. **Personnel** – Salaries, casual, UIF, bonuses, etc.  
6. **Income Tax**  
7. **Special Projects** – Year 1 from the 10-Year Plan  
8. **PQ / Levy Schedule** – Upload CSV/Excel of units + PQs, see calculated monthly levies  
9. **10-Year Plan** – Upload or review the maintenance plan  
10. **Download Excel** – Full workbook with live formulas

## Files needed for Streamlit Cloud

- `app.py`
- `requirements.txt`
- `Body_Corporate_Budget_Template_v2.xlsx`
- `domus_logo.jpeg` (optional)

## Notes

- Yellow cells with blue text = inputs  
- All totals and the Gross Ordinary Levy use live Excel formulas  
- Collection Rate % handles under-recovery  
- Insurance recoveries are deducted from R&M  
- PQ upload accepts CSV or Excel (columns: Unit + PQ)
