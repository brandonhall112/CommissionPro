# Commissioning Budget Tool (Desktop)

Fixes in this build:
- Window scales correctly (right side is scrollable; no fixed-size layout)
- Print Quote opens Print Preview and uses standard printer selection from the preview dialog
- Removed the Windows-invalid date format string (was causing invalid format errors)
- Default start state: no machines (click **Add Machine**)

Build:
GitHub → Actions → Build Windows GUI EXE → Run workflow
Artifact contains `dist/CommissioningBudgetTool.exe`
