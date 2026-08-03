# `enterprise-budget-fy26.xlsx` — Test Fixture Guide

**What this is:** a realistic FY26 annual operating budget workbook for a fictional large enterprise bank ("Northbridge Financial Group"), used to test all four Interactive Plan implementations against a workbook that **converts cleanly end to end** — no unsupported formulas, no known edge cases. It's the "everything works" companion to `browser-converter/fixtures/budget-test-3.xlsx`, which exists specifically to exercise the two known unsupported-formula cases (`INDIRECT()` and Excel Table structured references).

Lives at the repo root (`enterprise-budget-fy26.xlsx`) since it's used to test all four sibling projects, not owned by any one of them.

---

## Structure

**9 sheets**, **426 formulas**, **0 unsupported formulas**:

| Sheet | Contents |
|---|---|
| **Assumptions** | Global rate assumptions — benefits load %, merit increase %, blended salary rates (tech vs. non-tech), and per-FTE quarterly allowances for travel, facilities, software, and telecom. Editable. |
| **Headcount Plan** | Quarterly FTE counts (editable) by department, plus a computed quarterly salary $ column driving every department's Salaries line via a cross-sheet formula. |
| **Technology, Retail Banking Ops, Risk & Compliance, Marketing, Human Resources, Corporate Services** | Six department budget sheets, each with the same 11 line items (Salaries, Benefits, Contractors, Software, Travel, Training, Facilities, Telecom, Office Supplies, a department-specific program line, Other), an OpEx subtotal, CapEx, and a department total — across Q1–Q4, with FY25 Actual and YoY % columns. |
| **Summary** | Rolls up all six departments' totals, a Grand Total, and an OpEx-vs-CapEx breakdown. |

**Total FY26 budget: ~$250.4M** across the six departments.

### The live cross-sheet cascade

This is the most useful thing about this file for demos: editing a single **Headcount Plan** FTE cell ripples through the whole model in one recalculation:

```
Headcount Plan (FTE) → Salaries & Wages → Payroll Taxes & Benefits
                     → Software, Travel, Facilities, Telecom (each = FTE × a per-FTE rate)
                     → Department OpEx subtotal → Department Total
                     → Summary sheet's rollup for that department → Grand Total
```

Confirmed in testing: bumping Technology's Q1 headcount from 142 to 192 correctly cascaded through all of the above in a single HyperFormula recalculation, with the Summary sheet's Technology row updating from $38.4M to $41.4M.

---

## Regenerating the file

```bash
cd support
node generate-enterprise-budget.js
```

Writes `enterprise-budget-fy26.xlsx` to the repo root, overwriting the existing one. The script builds the entire workbook programmatically — all department figures, assumption rates, and FY25 "actual" comparisons are defined as plain JS data at the top of `generate-enterprise-budget.js`, so editing the file means editing that script and regenerating, not hand-editing the `.xlsx`.

The script computes a full "shadow" calculation in JS alongside every formula it writes, so the generated `.xlsx`'s cached values always exactly match what the formulas would compute — this matters because Excel files need a cached value alongside each formula to display correctly before a recalculation ever runs.

---

## Using it to test each option

Same file works for all four — pick whichever you're testing:

- **Browser Converter** — drag it onto `converter.html`, or run `node tools/generate-html.js` after pointing `SRC_FILE` at it (currently defaults to `fixtures/budget-test-3.xlsx` — swap the path for a one-off test).
- **Confluence app** — attach it to any Confluence page, add the Interactive Plan macro.
- **SharePoint web part** — upload it to a document library, point the web part's **File server-relative URL** property at it.
- **Excel Add-in** — open it in Excel (desktop or web) with the add-in sideloaded, click **Export as Interactive Plan**.

### What to expect

- **No warning banner** — this file has zero unsupported formulas, unlike `budget-test-3.xlsx`. If you see one, something's actually broken; it's not an expected edge case for this file.
- All 9 sheet tabs render with live, correctly cross-referenced numbers.
- Editing any Headcount Plan FTE cell (or any Assumptions rate) should visibly ripple through multiple sheets — good for demonstrating live recalculation to someone who hasn't seen the tool before.
