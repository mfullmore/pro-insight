# Browser Converter — Setup & Support Guide

**What this is:** a single self-contained HTML page (`converter.html`) that turns any budget `.xlsx` into an interactive, editable web page. No server, no account, no install — it runs entirely in the browser tab that opens it.

---

## For end users (no technical setup needed)

1. Open `converter.html` in any modern browser (Chrome, Edge, Safari, Firefox).
2. Drag a budget `.xlsx` onto the page, or click the drop zone to browse for one.
3. A **Download** button appears once conversion finishes — click it to save the generated `<workbook-name>-interactive.html` file.
4. Attach or embed that downloaded file wherever your team already works — Confluence, SharePoint, email, a network share. It's fully self-contained (one file, no dependencies), so it works anywhere you can open an HTML file in a browser.

Nothing is uploaded anywhere. The conversion — parsing the spreadsheet, checking which formulas are supported, building the interactive page — happens entirely inside that browser tab.

---

## For IT: making this available to your team

`converter.html` is a static file with zero server-side component, so "deploying" it is just "hosting a file":

- **Simplest option:** share the file directly — email it, post it to Teams/Slack, drop it on a network share. Anyone opens it locally, no server involved.
- **Or host it as a static file** on any internal web server, intranet page, or SharePoint document library. No special server configuration is needed — plain HTTP or HTTPS both work, since the page makes no network calls of its own.
- No backend, no database, no user accounts to provision, no ongoing hosting cost beyond wherever you park the file.

### Requirements
- Any modern browser. That's it, for end users.
- Node.js + npm are only needed if you're **rebuilding** the tool from source (see below) — not for using it.

---

## Rebuilding from source (for developers)

Only needed if you're changing the underlying conversion or rendering logic. If you just want to use the tool, skip this section.

### First-time setup

```bash
cd browser-converter
npm install
```

### Build order matters

Two vendored libraries (SheetJS for parsing, HyperFormula for recalculation) need to be built into browser-safe bundles *before* `converter.html` is assembled, since that assembly step embeds their output directly:

```bash
npm run build:xlsx      # rebuilds core/xlsx.core.min.js from the installed xlsx package
npm run build:hf        # rebuilds core/hyperformula.bundle.js from core/hf-entry.js
npm run build:converter # assembles converter.html from core/ + converter/
npm run build:demo      # optional — generates a demo interactive page from a fixture workbook, for quick local testing
```

You only need `build:xlsx` / `build:hf` again if you've upgraded the `xlsx` or `hyperformula` npm packages. Day-to-day, after editing files in `core/` or `converter/`, just re-run `npm run build:converter`.

> **Why `build:xlsx` and `build:hf` use `--define` flags:** both SheetJS and HyperFormula ship UMD-style module wrappers that check for `module`/`exports`/`define` globals at runtime to decide how to export themselves. That check is meant to distinguish "am I in Node" vs "am I in a browser" — but a **webpack bundle** (like the SharePoint web part's) also provides fake `module`/`exports`/`define` globals even though it's building for the browser, so the libraries would silently take the wrong branch and never actually attach to `window`. The `--define` flags force those checks to resolve as `undefined` at build time, so the libraries always take the plain-browser-global path, everywhere. This was a real bug found by testing the SharePoint web part live — see that project's support doc.

### Running the fidelity tests

```bash
npm run test:fidelity
```

Confirms the recalculation engine (HyperFormula) matches Excel's own computed values against two test workbooks. Expect **18 passed, 2 failed** — the two failures are `INDIRECT()` and a structured Excel Table reference, both intentionally unsupported (see "Known limitations" below), not bugs.

---

## Directory structure

```
browser-converter/
  converter.html         ← the actual deliverable — a self-contained authoring tool
  package.json

  core/                   shared engine used by ALL FOUR options in this project
                          (this one, plus the Confluence app, SharePoint web part, and Office Add-in)
    convert-core.js          parses a workbook (SheetJS) into a plain JSON payload
    render-core.js           builds the HyperFormula recalculation engine and renders
                              the interactive tabbed grid from that payload
    app.css, app.js          styling + the thin per-page bootstrap that wires the two above together
    xlsx.core.min.js         vendored SheetJS, patched for bundler-safety (see note above)
    hyperformula.bundle.js   vendored HyperFormula, built from hf-entry.js

  converter/              source for converter.html
    converter.js             drag-drop/upload UI + calls convert-core.js
    converter.css
    build-converter.js       assembles converter.html by embedding core/ + converter/ files

  tools/
    generate-html.js        Node CLI convenience tool — generates a demo interactive
                             page from a fixture workbook without opening a browser

  fixtures/                sample budget workbooks used for testing, plus the scripts
                            that generated them (build-sheet*.js)

  tests/                   fidelity tests — compare HyperFormula's recalculation
                            against Excel's own cached values

  dist/                    generated demo output (regenerate with `npm run build:demo`)
```

**Important:** `core/` is the shared source of truth for all four implementations in this project (browser converter, Confluence app, SharePoint web part, Office Add-in). Each of the other three *copies* (vendors) these files at their own build time — editing `core/` alone does not update anything already deployed. If you fix something here, you need to rebuild **and redeploy** the other three projects too. Each has its own `support/README.md` covering that.

---

## Known limitations (by design, not bugs)

The recalculation engine (HyperFormula) cannot evaluate two kinds of formulas:

- **`INDIRECT()`** — builds a cell reference dynamically from text, which HyperFormula's architecture has no path to support (confirmed by reading its source — not a missing feature that might get added later).
- **Excel Table structured references** — e.g. `Table1[Column]` — a hard parser-level limitation.

Rather than silently computing and showing a plausible-but-wrong number, the tool:

1. Flags every such formula explicitly in a warning banner when the page loads, naming the exact cell and why.
2. Shows that flagged cell's **last-known value from Excel itself**, marked with a **✱** and labeled "frozen" (static, not live) — so the page still shows something real, not a blank or an error, for the one cell Excel already computed correctly.
3. Shows a genuine **live error** (never a guessed number) for any *other* cell that depends on a flagged one, since a frozen value would silently go stale the moment someone edits a related input.

---

## Troubleshooting

- **Build error: "`<file>` contains a literal `</script` — cannot embed safely as-is"** — one of the files being embedded into `converter.html` or a generated page contains the literal string `</script`, which would prematurely close the `<script>` tag it's embedded in and corrupt the page. Find and fix that string in the named file.
- **`converter.html` looks stale after editing source files** — you edited something in `core/` or `converter/` but haven't re-run `npm run build:converter` yet.
- **A cell shows a number that looks wrong** — check the warning banner first. Only the exact flagged cell (✱) and cells that depend on it are affected by the two known limitations above; everything else recalculates normally and correctly.
- **File won't drag-and-drop / nothing happens on drop** — confirm the file has a `.xlsx` extension; other formats (`.xls`, `.csv`) aren't supported by this tool.
