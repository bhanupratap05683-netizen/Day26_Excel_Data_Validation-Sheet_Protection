# Day 26 — Excel Data Validation & Sheet Protection

## What I Built
A loan application Excel template with programmatic data validation rules and formula protection using openpyxl.

## Files
| File | Description |
|------|-------------|
| `Day26_Practice.xlsx` | Input file with 3 sheets — Loan form, Expense data, Reference lists |
| `Day26_Excel_Data_Validation_Protection.py` | Python script that adds all validation rules and sheet protection |
| `Day26_Output.xlsx` | Final output with validation and protection applied |

## What the Script Does

### Data Validation (Loan_Application sheet)
- **Department** → dropdown: Finance, HR, Operations, IT, Marketing, Legal
- **Employment Type** → dropdown: Full-Time, Part-Time, Contract, Intern
- **Annual Income** → whole number, must be > 0
- **Credit Score** → whole number, 300 to 900 only
- **Loan Type** → dropdown: 5 loan type options
- **Loan Amount** → whole number, ₹10,000 to ₹50,00,000
- **Tenure** → whole number, 6 to 360 months
- **Interest Rate** → decimal, 1.0% to 30.0%

### Data Validation (Expenses_Raw sheet)
- **Department column** → dropdown (same 6 options)
- **Category column** → dropdown: Travel, Food, Software, Hardware, Training, Utilities, Other
- **Amount column** → decimal, must be > 0
- **Approved column** → dropdown: Yes, No, Pending

### Sheet Protection (Loan_Application sheet)
- All input cells → unlocked (freely editable)
- Formula cells E19–E23 (EMI, Total Payment, Total Interest, DTI Ratio, Eligibility) → locked
- Protection password: `finance2026`

## Key Concepts Used
- `DataValidation(type, operator, formula1)` — defines the rule
- `dv.add("C6")` — attaches rule to a single cell
- `dv.sqref = "C3:C12"` — attaches rule to a cell range
- `ws.add_data_validation(dv)` — registers validation on the sheet
- `Protection(locked=False/True)` — marks cells as locked or unlocked
- `ws.protection.sheet = True` — activates protection

## How to Run
```bash
pip install openpyxl
python Day26_validation_protection.py
```
Open `Day26_Output.xlsx` in Excel to test dropdowns and protection.
