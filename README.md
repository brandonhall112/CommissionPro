# Commissioning Budget Tool

This is a Windows desktop GUI app that estimates commissioning labor + expenses and generates a customer-facing PDF quote.

**Excel structure expected (3 tabs):**
- `Instal days by Model`: install-only days per machine (`Item`, `Technician Days Required`, `Field Engineer Days Required`)
- `Service Rates`: cost lines + labor rates (`Description`, `Unit Price`, `Application Notes`)
- `Requirements and Assumptions`: used in the PDF Terms section

**Key rules implemented:**
- Training = `ceil(qty / 3)` day(s) per machine model (only if "Training Required" is checked for that line)
- Dedicated skills: personnel are **not shared** across different machine types
- Customer Install Window: selectable 3–14 days; error if install + training for a single machine exceeds the window
- Load balancing: balanced split to minimize the maximum assignment
- Expenses: calculated using **person-days** including travel-in + travel-out per person

**Overrides:**
- Airfare = $1,500 per person
- Baggage = $150 per day per person

## Build the EXE
GitHub → Actions → Build Windows GUI EXE → Run workflow

Download artifact `CommissioningBudgetTool-Windows` → `dist/` contains the EXE.
