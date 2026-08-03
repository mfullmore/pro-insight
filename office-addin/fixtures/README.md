# Fixtures — none live here

This add-in doesn't bundle its own test workbooks, unlike `browser-converter/fixtures/`. Every other option in this project reads a `.xlsx` file — uploaded, attached, or referenced by path — so each one has its own copy (or a shared one) to test against. This one is different: it reads whatever workbook is **already open in Excel** through Excel's own JS API (`Excel.run`), so "picking a fixture" just means opening one of the existing test workbooks in Excel (desktop or web) before clicking **Export as Interactive Plan**.

## Where to get a test workbook

- **`../../browser-converter/fixtures/budget-test-3.xlsx`** — the canonical demo/edge-case workbook (INDIRECT + a structured table reference, both intentionally unsupported). See `browser-converter/fixtures/README.md` for the full set and what each one exercises.
- **`../../support/enterprise-budget-fy26.xlsx`** — the large, realistic "everything works cleanly" workbook (9 sheets, 426 formulas, zero unsupported formulas). See `support/README.md`.

## Using one

1. Open the chosen `.xlsx` in Excel — desktop, or Excel for the web via a document library / OneDrive.
2. Sideload this add-in if you haven't already (see `../support/README.md`).
3. Click **Export as Interactive Plan** in the ribbon, then the button inside the task pane.

Whatever you get should match what that workbook's own README says to expect — e.g. `budget-test-3.xlsx` should flag exactly two unsupported formulas.
