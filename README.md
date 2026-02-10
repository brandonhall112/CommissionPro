# Pearson Commissioning Pro

Pearson Commissioning Pro estimates onsite commissioning labor, expenses, and total project cost based on machine quantities and a rate sheet stored in Excel. It also generates a customer-ready printable quote.

---

## What this tool does

Given one or more machine types and quantities, the tool calculates:

- **Required technicians and engineers**
- **Onsite days per person** (balanced)
- **Training days** (optional per machine type via checkbox)
- **Labor costs** by role and total
- **Estimated travel & project expenses** (with subtotals)
- **A printable quote** (preview + print dialog)

This is not a basic multiplier — it enforces constraints like customer install window limits and load balancing.

---

## Getting started (normal use)

1. Launch **PearsonCommissioningPro.exe**
2. Add machines using **Add Machine**
3. Select the **Machine Model** and enter the **Quantity**
4. Set **Customer Install Window** (days)
5. Review calculated outputs:
   - Machine breakdown
   - Personnel assignments
   - Labor costs
   - Estimated expenses
   - Workload distribution graph
6. Click **Print Quote** to preview and print/save a customer quote

---

## Machine configuration

### Adding machines
- By default, **no machines are loaded** on startup.
- Click **Add Machine** to add a line.
- Machine selection starts **blank** until you pick a model.

### Training Required checkbox
Each machine line includes a **Training Required** checkbox.

- **Checked**: training days are added for that machine type (applies to both tech and engineer when applicable)
- **Unchecked**: training is excluded and the breakdown will show “(training excluded)”
- **Important:** unchecking training should only be done **by customer request**.

### Training day logic
Training is applied per machine type using this rule:

- **1 training day per 3 machines** of the same type  
  (e.g., qty 1–3 → 1 day, qty 4–6 → 2 days, etc.)

If more than one training day is required, training days are **balanced across personnel** to minimize peak duration.

---

## Customer Install Window (days)

The **Customer Install Window** controls the maximum onsite days allowed per person for this project.

- Minimum: **3 days**
- Maximum: **14 days**
- Work is automatically balanced so no person exceeds this limit.

### Constraint / validation
If a *single machine type* requires more onsite days (install + training) than the selected window allows, the tool will throw an error. This prevents quotes that can’t be executed within the customer’s requested schedule.

---

## Allocation rules (how staffing is calculated)

### Roles
Some models require:
- **Technicians only**
- **Engineers only**
- **Both technicians and engineers**

The tool calculates each role separately using the same rules:
- training inclusion/exclusion
- max window constraint
- load balancing

### Load balancing (critical behavior)
When more than one person is required, the tool does **not** “fill up Tech 1 to the limit” and push overflow onto Tech 2.

Instead it rebalances work to:
- **minimize the longest assigned duration**
- keep each person’s onsite days as even as possible
- respect the Customer Install Window limit

---

## Outputs explained

### Machine Breakdown
Shows each machine type with:
- quantity
- install days per unit
- training indicator:
  - “(incl. X Train)” when included
  - “(training excluded)” when unchecked

### Personnel Assignments
Shows the number of techs/engineers required and how days are distributed.

### Labor Costs
Shows cost by role and line item and includes:
- per-day rates pulled from Excel
- totals/subtotals at the bottom

### Estimated Expenses
Shows expense items and subtotals. It is sized to avoid scrolling for the full list.

Travel day logic:
- Expenses can include travel in/out depending on the rate rules in the Excel sheet.

### Workload Distribution Graph
Bar chart showing each person’s:
- onsite days
- travel days (stacked)
Techs and engineers use different colors so they’re easy to distinguish.

---

## Printing a quote

1. Click **Print Quote**
2. A **preview window** opens
3. From the preview, use the standard **Print** dialog to:
   - select printer
   - print to PDF
   - set page options

### Quote formatting notes
- Logo appears in the header (right-justified)
- Sections are spaced to avoid “crammed” formatting
- Payment terms are **not** included (handled case-by-case)

---

## Excel data source

The tool reads from:

`assets/Tech days and quote rates.xlsx`

This workbook contains the tables used for:
- machine day requirements (install and role requirements)
- service rates
- requirements/assumptions text used in the printed quote
- expense rules and application notes (if included in your version)

### “Open Excel” button
This opens the Excel file so you can view/edit the tables.
- If you change values and want the tool to use the updates:
  - save Excel
  - close and re-open Pearson Commissioning Pro (or use any built-in refresh/reload if present)

**Important:** Keep the file name and location the same unless you know your build is set up to detect alternate filenames.

---

## Common issues / troubleshooting

### “It keeps asking me to select the Excel file”
This occurs when the Excel workbook cannot be found in the app assets at runtime.

Confirm:
- The file exists at: `assets/Tech days and quote rates.xlsx` in the repo
- Your packaged build includes it in the final EXE assets

If you are using a folder-based deployment build (onedir), it may appear under:
- `_internal/assets/`

### “Print Quote does nothing”
- Confirm you’re running the latest EXE from the latest build artifact
- Ensure the preview window is not opening behind the main window

### “Numbers look wrong”
- Verify the machine model row in Excel is correct
- Confirm training checkbox state
- Confirm Customer Install Window isn’t forcing extra staffing

---

## Versioning / builds (internal)

This repository is intended to be packaged as a Windows executable using PyInstaller via GitHub Actions.
- Keep `assets/` intact
- Do not rename the Excel file unless you also update the app logic/spec accordingly

---

## Support / ownership

Publisher: Brandon T. Hall  
Tool name: Pearson Commissioning Pro


## Git safety guard (prevent accidental branch commits)

To prevent accidental commits on non-`main` branches, install the repository hooks once:

```bash
./scripts/install-git-hooks.sh
```

After installation, commits on branches other than `main` are blocked by default.
If you intentionally need a non-`main` commit, you can bypass once with:

```bash
ALLOW_NON_MAIN_COMMIT=1 git commit -m "..."
```
