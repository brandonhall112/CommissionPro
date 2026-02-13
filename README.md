# Pearson Commissioning Pro

Pearson Commissioning Pro is a PySide6 desktop quoting tool that estimates commissioning labor, travel-related expenses, staffing distribution, and total project cost for selected machine models.

---

## What the application does

Given a set of machine lines (model + quantity + training selection), the app calculates and displays:

- Machine breakdown (tech/eng day totals by model)
- Personnel assignments (balanced by role)
- Labor costs (regular + overtime)
- Estimated expenses (airfare, baggage, car, parking, hotel, per diem, prep, travel time)
- Workload calendar (2-week Sun–Sat Gantt view)
- Print-ready quote preview

---

## Primary inputs

### 1) Machine lines
Each line has:

- **Model**
- **Quantity**
- **Training Required** checkbox (when applicable)

Duplicate model selection is prevented across lines; increase quantity on the existing line instead.

### 2) Customer Install Window
Install window is a hard limit (3–14 days) for maximum onsite duration per person.

### 3) Excel workbook(s)
The core workbook is loaded from:

- `assets/Tech days and quote rates.xlsx`

Optional skills matrix workbook:

- `assets/Machine Qualifications for PCP Quoting.xlsx`

---

## Buttons and top-row controls

- **Header**: opens a quote-header form for `Customer Name`, `Reference`, `Submitted to`, and `Prepared By` fields used in the shaded top quote summary area.
- **Load Excel…**: load a different `.xlsx` workbook for the session.
- **Open Bundled Excel**: opens the packaged/default workbook in your default spreadsheet app.
- **Help**: opens this README inside the app so users can review logic/assumptions directly.

---

## Core principles of operation

## 1) Day construction per model
For each selected model:

- `tech_install_days = technician_days_per_machine × qty`
- `eng_install_days = engineer_days_per_machine × qty`
- `training_days = ceil(qty / 3)` when training is applicable and enabled

Training is model-scoped and can be excluded by unchecking **Training Required**.

## 2) Install-window enforcement
The app validates that per-person allocations do not exceed the customer install window.
If a model’s required install + training cannot fit the selected window, the app blocks calculation and shows a clear error.

## 3) Personnel allocation and load balancing
Allocation targets minimum feasible headcount while keeping workloads balanced.
When multiple people are required, days are distributed to minimize peak individual load.

## 4) Technician skills-group partitioning
Tech-only lines are partitioned into crew pools using the optional skills matrix.
Rules include:

- RPC models (`RPC-C`, `RPC-DF`, `RPC-PH`, `RPC-OU`) are always grouped together in a dedicated RPC pool
- generic/support lines treated as non-blocking
- matrix-missing models treated as supplemental and distributed safely

If skills matrix is missing/unreadable, the app falls back to baseline allocation and shows a non-fatal warning.
- RPC staffing is isolated from non-RPC staffing (no cross-pollination across pools).

---

## Workload calendar behavior (left column)

The quote’s shaded top summary area is two columns:

- Left: Customer Name, Reference, Submitted to
- Right: Prepared By, Quote Validity, Total Personnel, Estimated Duration

Quote Validity, Total Personnel, and Estimated Duration are auto-populated from calculations; the remaining fields come from the **Header** form.
Header form values are stored in-app for the session and used each time a quote is generated.

The workload view is a **2-week (14-day) Sun–Sat Gantt calendar**:

- Rows = personnel (T1, T2, … / E1, E2, …)
- Columns = day slots
- Light bar = travel day
- Solid bar = onsite day
- Printed quote includes the same legend directly below the calendar table.

Colors:

- Technician: `#e04426`
- Engineer: `#6790a0`

RPC rule:

- Tech travel-in defaults to Sunday.
- Engineer travel-in defaults to Monday for RPC jobs, and shifts to Tuesday when `RPC-PH` or `RPC-OU` is in scope.

Training is model-scoped and can be excluded by unchecking **Training Required**.

## Labor cost model

Labor costs are shown as separate lines by role:

- Tech Regular Time (daily rate)
- Tech Overtime (Sat/Sun, daily-equivalent rate)
- Eng Regular Time (daily rate)
- Eng Overtime (Sat/Sun, daily-equivalent rate)

### Rate sourcing
Rates are read from **Service Rates** in Excel:

- `tech. regular time`
- `eng. regular time`
- `tech. overtime`
- `eng. overtime`

If OT keys are missing, OT falls back to regular role rates.

### OT triggering
Overtime applies when onsite days land on Saturday or Sunday.
The UI and printed quote display OT using day-based units to stay consistent with the labor table headers.

Colors:

## Expense model

Estimated expenses are calculated from person-days/trip-days and workbook rates:

- Airfare
- Baggage
- Car rental
- Parking
- Hotel nights
- Per diem
- Pre/Post trip prep
- Travel time

Travel day assumptions (in addition to onsite days) are applied consistently to expense calculations.

---

## Quote preview / print

**Print Quote…** opens a preview window and supports normal print/PDF flows.
The quote includes:

- machine breakdown
- labor breakdown (regular + OT)
- expenses
- total estimate
- assumptions/requirements text from workbook where provided

---

## Assumptions and constraints summary

- Training day rule: **1 day per 3 machines** (when applicable)
- Install window is a hard per-person onsite cap
- Allocations are balanced, not greedy
- Skills matrix affects technician grouping (engineer grouping is not matrix-split in this scope)
- Missing skills matrix is non-fatal (fallback path)
- Weekend onsite days trigger overtime calculation

---

## Troubleshooting

### Build/runtime syntax issues
Run:

```bash
python -m py_compile app.py
```

### Workbook not found
Confirm `assets/Tech days and quote rates.xlsx` exists in packaged assets.

### Help button shows fallback text
`README.md` was not found in the runtime location. Include it beside the executable (or packaged internal location) if you want full in-app guide text.

---

## Git safety guard (internal)

Install repository hooks once:

```bash
./scripts/install-git-hooks.sh
```

This blocks accidental non-`main` commits unless explicitly bypassed.

---

Publisher: Brandon T. Hall  
Tool: Pearson Commissioning Pro

---

## Recent quote-form updates

- Generated quote now includes a right-aligned Pearson logo pinned at page top so it repeats on subsequent printed pages.
- Generated quote now includes the workload calendar directly after Machine Breakdown.
- Overtime lines are presented as day-based rates/quantities to match the labor table semantics.
