# Forflowmatic — Reporting Portal

Static dashboard hosting all Forflowmatic operational reports in one place.
Live at: https://flomaticauto.github.io/reporting/

## Reports
- **Petrol Slip Tracker** (`petrol-slips.html`) — fuel-station receipts.
- **Other Expense Slips** (`other-slips.html`) — non-petrol receipts.
- **Account Transactions** (`account-transactions.html`) — FNB Gold Business Current view.

## How it updates
The Friday autonomous slip-processing run (Windows Task Scheduler →
`agents/accountant/scripts/weekly-slip-processor.ps1` in the Flomatic
workspace) writes new rows to the trackers in `Reports/`, then runs
`scripts/generate-viewers.py` to regenerate the HTML viewers, then commits
and pushes — which triggers the GitHub Pages redeploy here.

## Brand
Brighter alert colours for visual triage:
- Amber alert (`--flo-amber-alert`): `#FFB020`
- Fault red (`--flo-fault-red`): `#FF3B3B`

Standard Forflowmatic teal (`#2A8C8C`), navy (`#1A2530`), and steel (`#3D5A72`)
remain primary.
