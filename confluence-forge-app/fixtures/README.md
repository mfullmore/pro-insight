# Fixtures — none live here

This app doesn't bundle its own test workbooks, unlike `browser-converter/fixtures/`. It reads whatever `.xlsx` is already **attached to the Confluence page** the macro is placed on (via the resolver's `getAttachment`, see `src/resolvers/index.js`), so "picking a fixture" just means attaching one of the existing test workbooks to a page before adding the macro — same situation as the Office Add-in (see `office-addin/fixtures/README.md`).

## Where to get a test workbook

- **`../../browser-converter/fixtures/budget-test-3.xlsx`** — the canonical demo/edge-case workbook (INDIRECT + a structured table reference, both intentionally unsupported). See `browser-converter/fixtures/README.md` for the full set and what each one exercises.
- **`../../support/enterprise-budget-fy26.xlsx`** — the large, realistic "everything works cleanly" workbook (9 sheets, 426 formulas, zero unsupported formulas). See `support/README.md`.

## Using one

1. Attach the chosen `.xlsx` to a Confluence page (drag it on, or use the attachment picker).
2. Add the **Interactive Plan** macro to that same page.
3. The macro automatically reads the first `.xlsx` attached to the page — no configuration needed.

Whatever you get should match what that workbook's own README says to expect — e.g. `budget-test-3.xlsx` should flag exactly two unsupported formulas (`Marketing!F6`, `Summary!B12`), both shown with the frozen-value fallback.
