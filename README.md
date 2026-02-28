# Pearson Commissioning Pro

Pearson Commissioning Pro is a desktop quoting tool used to estimate labor, expenses, staffing, and total project cost for commissioning work.

> **Important:** This README is also the in-app **Help** guide. If a user clicks **Help** in the app, this is what they see.

---

## Quick start (for non-technical users)

If you only remember one section, use this one:

1. Open the app.
2. Add each machine type your customer purchased.
3. Enter quantities.
4. Check **Training Required** for lines that need training.
5. Set the **Customer Install Window** (3 to 14 days).
6. Click **Customer Details** and fill in quote header info.
7. Click **Calculate/Generate Quote** (or equivalent quote action).
8. Review machine breakdown, labor, expenses, and total.
9. Click **Print Quote…** to print or save as PDF.

If something does not calculate, check the install window first. A window that is too short is the most common reason for a blocked quote.

---

## What the app calculates

From your selected machine lines, the app produces:

- Machine day totals (technician and engineer)
- Recommended staffing counts
- Per-person workload distribution
- Labor totals (regular + overtime)
- Expense totals (travel-related + prep/travel-time)
- Two-week workload calendar (Sun–Sat)
- Final quote preview suitable for print/PDF

---

## Inputs you must provide

### 1) Machine lines

Each line includes:

- **Model**
- **Quantity**
- **Training Required** (if applicable)

Tips:

- Do not add the same model twice; increase quantity on the existing line.
- If training is not needed for a model, uncheck **Training Required** on that line.

### 2) Customer Install Window

- Hard limit: **3 to 14 days**.
- This limit controls how many onsite days any one person can be assigned.
- If required work cannot fit this window, the app blocks the quote and shows an error.

### 3) Excel workbook(s)

Required:

Required:

---

Optional (recommended for smarter tech grouping):

- **Customer Details**: enter Customer Name, Reference, Submitted to, Prepared By.
- **Load Excel…**: load a different quote workbook for the current session.
- **Open Bundled Excel**: open the default packaged workbook.
- **Help**: open this document in-app.
- **Print Quote…**: open print preview and print/save PDF.

---

## Top-row buttons

- **Customer Details**: enter Customer Name, Reference, Submitted to, Prepared By.
- **Load Excel…**: load a different quote workbook for the current session.
- **Open Bundled Excel**: open the default packaged workbook.
- **Help**: open this document in-app.
- **Print Quote…**: open print preview and print/save PDF.

---

## How day totals are built

For each machine model:

- `tech_install_days = technician_days_per_machine × quantity`
- `eng_install_days = engineer_days_per_machine × quantity`
- `training_days = ceil(quantity / 3)` when training is enabled for that model

Training is calculated per model line and only included when **Training Required** is checked.

---

## Technician grouping and re-grouping (factoring / refactoring)

This is the part many users ask about.

When multiple machine families are present, the app does more than “divide total days by number of people.” It uses grouping rules so assignments are realistic and safe.

### Why grouping exists

Different technicians may be qualified for different model families. Grouping prevents invalid assignments and helps keep teams focused.

### How grouping works

When the optional skills matrix workbook is available, tech-only work is partitioned into skill-compatible pools:

- RPC models (`RPC-C`, `RPC-DF`, `RPC-PH`, `RPC-OU`) are grouped together in a dedicated RPC pool.
- RPC technicians are kept separate from non-RPC technicians for normal model-specific work.
- **Conveyor / Production Support / Training Day** technician lines are treated as shared support and can be factored across **all technicians** (including RPC techs) to balance load.
- Models missing from the matrix are treated as supplemental and assigned conservatively.

### What “refactoring” means in practice

As you add/remove lines or change quantities, the app re-balances technician pools:

- It recalculates required headcount.
- It redistributes days to reduce overloaded individuals.
- It redistributes install and training days using the same shared-support refactoring path (for both technicians and engineers), only adding headcount when the install-window cap is actually reached.
- It keeps RPC isolation for model-specific work while still spreading shared support days across all technician crews.
- For RPC models, model-selected training days are first refactored onto existing RPC technician/engineer assignments before new headcount is added.

So if your scope changes, staffing can be “re-factored” automatically without manual reshuffling.

### If the skills matrix is missing

- Quote still works.
- App falls back to baseline allocation logic.
- You receive a warning, not a fatal error.

---

## Engineer and travel scheduling behavior

### Workload calendar

The quote includes a **2-week (14-day) Sun–Sat** calendar:

- Rows = personnel (`T1`, `T2`, … and `E1`, `E2`, …)
- Columns = day slots
- Light bar = travel day
- Solid bar = onsite day

Color key:

- Technician: `#e04426`
- Engineer: `#6790a0`

### RPC travel defaults

- Technician travel-in defaults to **Sunday**.
- Engineer travel-in defaults to **Monday** for RPC jobs.
- Engineer travel-in shifts to **Tuesday** when `RPC-PH` or `RPC-OU` is in scope.

### Engineer production-support factoring

- Engineer **Production Support** days are treated as shared engineer-support work.
- The app first fills existing engineer capacity up to the install-window limit (adjusted for engineer start stagger on RPC jobs).
- Only when existing engineer loads are full does it add another engineer.
- If a quote has production-support engineer days and no engineer is present yet, the app creates the first engineer assignment.

---

## Labor and overtime rules

Labor is shown by role and time class:

- Tech Regular Time
- Tech Overtime
- Eng Regular Time
- Eng Overtime

Rates come from the workbook **Service Rates** section using keys:

- `tech. regular time`
- `eng. regular time`
- `tech. overtime`
- `eng. overtime`

Overtime rule:

- Onsite days that land on **Saturday/Sunday** count as overtime.
- If OT rate keys are missing, overtime falls back to regular role rates.

---

## Expense rules (high level)

Estimated expenses are calculated from person-days/trip-days and workbook rates:

- Airfare
- Baggage
- Car rental
- Parking
- Hotel
- Per diem
- Pre/Post trip prep
- Travel time

Travel-day assumptions are applied consistently alongside onsite days.

---

## Quote header and output details

The shaded quote summary header uses two columns:

- Left: Customer Name, Reference, Submitted to
- Right: Prepared By, Quote Validity, Total Personnel, Estimated Duration

Auto-populated values:

- Quote Validity
- Total Personnel
- Estimated Duration

User-entered values come from **Customer Details** and are reused in-session.

Quote output includes:

- Machine breakdown
- Labor (regular + overtime)
- Expenses
- Total estimate
- Workload calendar
- Workbook requirements/assumptions text (when present)

---

## Troubleshooting

### “Calculation failed” or blocked quote

Most likely causes:

1. Install window too short for the required per-person workload.
2. Required workbook data is missing or changed unexpectedly.

Try:

- Increase install window.
- Re-open default workbook.
- Recheck quantities/training checkboxes.

### Workbook not found

Confirm this file exists:

- `assets/Tech days and quote rates.xlsx`

### Help button shows fallback text

`README.md` is missing from runtime location. Include it with the packaged app.

### Syntax/runtime check (for maintainers)

```bash
python -m py_compile app.py
```

---

## Internal repo safety (maintainers)

Install git hooks once:

```bash
./scripts/install-git-hooks.sh
```

---

Publisher: Brandon T. Hall  
Tool: Pearson Commissioning Pro
