# Excel Add-in — Setup & Support Guide

**What this is:** an Office Add-in that adds an **Export as Interactive Plan** button to Excel's ribbon. Click it while a budget workbook is open, and it reads the workbook's own already-computed values and formulas directly through Excel's JavaScript API, then downloads a self-contained interactive HTML page — no upload, no separate conversion tool.

---

## For end users

1. With the add-in installed (see below) and a budget workbook open, go to the **Home** tab and find the **Interactive Plan** group.
2. Click **Export as Interactive Plan**.
3. A task pane opens on the right — click the **Export as Interactive Plan** button inside it.
4. A file downloads: `<workbook-name>-interactive.html`. Attach or embed it wherever your team already works (Confluence, SharePoint, email) — it's fully self-contained.

Nothing is uploaded anywhere. Excel's own API is read locally; the export happens entirely in the task pane.

---

## Local testing setup

Before this add-in can be sideloaded anywhere, it needs to be served over **HTTPS** — Office requires this even for `localhost`. This project vendors the tooling for that.

### 1. Install dependencies

```bash
cd office-addin
npm install
```

### 2. Install a trusted local dev certificate (one-time)

```bash
npm run dev-certs
```

This generates a certificate at `~/.office-addin-dev-certs/` and attempts to add it to your system's trust store automatically.

⚠️ **This step commonly needs a manual follow-up.** Adding a certificate to the trust store requires OS-level authorization (a Keychain password/Touch ID prompt on macOS) that only works when a human is present to approve it — running this non-interactively (e.g. through an automation script) can report success while silently failing to actually grant trust. If your browser still shows a certificate warning for `https://localhost:3000` afterward:

- **Simplest fix:** visit `https://localhost:3000/taskpane.html` once, and click through the browser's one-time warning (in Chrome: **Advanced → Proceed to localhost (unsafe)**). This is a one-time trust decision for your own local dev server, not a real security risk.
- **Cleaner fix:** open **Keychain Access** (macOS) and manually mark the "Developer CA for Microsoft Office Add-ins" certificate as trusted, or re-run `npm run dev-certs` directly in an interactive terminal (not through a script or CI) so the OS prompt can actually appear.

### 3. Build the task pane

```bash
npm run build
```

Runs `make-icons.js` (generates the ribbon icons from a source image) and `build-taskpane.js` (assembles `taskpane.html` from the shared `../browser-converter/core/` files — same pattern as the other two platform integrations).

### 4. Start the local HTTPS server

```bash
npm run serve
```

Serves this directory at `https://localhost:3000/`. Leave this running while testing — `manifest.xml`'s `SourceLocation` and icon URLs all point at `https://localhost:3000`.

### 5. Sideload

**Desktop Excel (Windows/Mac):**

```bash
npm run sideload
```

This uses `office-addin-debugging`, which registers the manifest with the desktop app directly and launches Excel with the add-in loaded.

**Excel for the web:** there's no CLI path for this — sideload manually:

1. Open a workbook in Excel for the web (`https://<yourtenant>.sharepoint.com/...` or via OneDrive)
2. **Insert → Add-ins** (or the **Add-ins** button in the Home ribbon) → **Upload My Add-in**
3. Select `office-addin/manifest.xml`

Once sideloaded this way, the add-in stays available across different workbooks you open in that browser session — you don't need to re-upload it for every file.

> **Note for anyone scripting/automating this:** Excel for the web renders inside a cross-origin iframe (`WacFrame_Excel_0`) that browser automation tools generally can't read via the accessibility tree or reliably click via DOM selectors — raw pixel-coordinate clicks work for ribbon buttons and the resulting task pane, but the sideload dialog itself still needs a real file picker, which is not automatable at all. Budget for a manual step here.

### 6. Validate the manifest (optional, but useful after edits)

```bash
npm run validate-manifest
```

Checks `manifest.xml` against Microsoft's own schema — catches malformed XML or missing required fields before you try to sideload.

---

## Moving beyond local testing

Everything above uses `https://localhost:3000` — fine for development, but **not something you can hand to end users**, since it only works on the machine running the local server. For a real rollout, two things need to change:

1. **Host the built files somewhere with real, publicly-reachable HTTPS** — an Azure Storage static site, an internal web server with a valid cert, etc. Copy `taskpane.html`, `taskpane.css`, and `assets/` there.
2. **Update `manifest.xml`** — replace every `https://localhost:3000` reference (`SourceLocation`, `IconUrl`, `HighResolutionIconUrl`, `AppDomains`, and the `bt:Url`/`bt:Image` entries under `Resources`) with the real hosted URLs.

Then distribute the updated `manifest.xml` through one of Microsoft's standard centralized deployment paths (not something this project automates, since it depends on your organization's own M365 setup):

- **Microsoft 365 admin center → Integrated apps** — upload the manifest, target specific users/groups/the whole org. This is the most common path for org-wide rollout.
- **A SharePoint App Catalog** can also host Office Add-in manifests alongside SPFx solutions, if your organization already manages deployment that way.

---

## Directory structure

```
office-addin/
  manifest.xml              add-in manifest — ribbon button definition, task pane URL,
                             icon URLs, permissions (currently ReadDocument only —
                             this add-in never writes back to the open workbook)
  build-taskpane.js          assembles taskpane.html from ../browser-converter/core/
                              (same embedding pattern as browser-converter's own converter.html)
  taskpane.js                 the task pane's own logic — reads the open workbook via
                               Excel's JS API (Excel.run, range.formulas/values/numberFormat),
                               then hands off to the shared render-core.js/convert pipeline
  taskpane.html                built output — do not hand-edit, regenerate with `npm run build`
  make-icons.js                generates assets/icon-*.png from a source image
  serve.js                     minimal local HTTPS static server (reads the dev cert)
  assets/                      ribbon icons (generated)
```

---

## Known drift risk: `taskpane.js` hand-duplicates two small pieces of `convert-core.js`

Unlike `app.css`/`app.js`/`render-core.js` (all read fresh from `../browser-converter/core/` at build time — see `build-taskpane.js`, no copy ever gets stored here), `taskpane.js` **cannot** use the shared `convert-core.js` directly: that file parses a SheetJS workbook object, and this add-in reads the open workbook through Excel's own JS API instead (`Excel.run`, `range.formulas`/`values`/`numberFormat`) — a genuinely different extraction path, not just a different data source.

Because of that, two small pieces of logic are hand-copied into `taskpane.js` rather than shared:

- **`UNSUPPORTED_PATTERNS`** — the regex list that flags `INDIRECT()` and Excel Table structured references. Currently identical to `convert-core.js`'s copy.
- **`colLetter()`** — converts a column index to its letter (e.g. `5` → `F`) for building addresses like `F6` in the warning list. Currently identical to `render-core.js`'s copy.

**Both are in sync as of this writing** (verified by direct comparison), but nothing enforces that automatically — if either regex or `colLetter` changes in `browser-converter/core/`, these two spots in `taskpane.js` need a manual, matching edit. Worth a quick `diff`-by-eye check here whenever you touch `UNSUPPORTED_PATTERNS` or `colLetter` anywhere else in the project. This is the same class of risk as `confluence-forge-app/static/main/app.css` (see that project's `static/README.md`) — a necessarily hand-maintained fork, not an auto-vendored copy.

---

## Known limitations

Same recalculation engine as the other three options in this project — `INDIRECT()` and Excel Table structured references aren't supported, and are flagged explicitly rather than silently computed wrong. See `browser-converter/support/README.md` for the full explanation.

**One real difference specific to this option, confirmed in testing:** because this add-in reads formulas through Excel's own JS API rather than parsing the raw file with SheetJS, Excel sometimes hands back an already-resolved formula rather than the original text. An Excel Table structured reference (e.g. `RatesTable[Overhead Rate]`) came back from the API as `AVERAGE(#REF!)` — meaning the upfront pattern-based warning banner doesn't catch it (the literal `Table[Column]` syntax is gone by the time this code sees it). The underlying safety guarantee still held: HyperFormula evaluated `#REF!` as a genuine error and displayed it as a live error, never a plausible-looking wrong number — it just didn't get the "frozen ✱ value" treatment the other three options give the equivalent cell, since it was never flagged as unsupported to begin with.

---

## Troubleshooting

- **Certificate warning in the browser for `https://localhost:3000`** — see step 2 above; this is the most common snag and almost always needs a one-time manual click-through or a Keychain fix, even after `npm run dev-certs` reports success.
- **Task pane shows "This add-in only runs inside Excel."** — you're viewing `taskpane.html` directly in a plain browser tab rather than through Excel/Excel Online. This is expected and correct behavior (`Office.onReady` checks the host and disables the export button outside Excel) — not a bug, just a reminder you're not inside the actual host.
- **Sideload dialog doesn't show "Upload My Add-in" in Excel for the web** — look under the **Add-ins** dropdown in the Home ribbon, not just the Insert tab; the exact label and location shift slightly between Microsoft 365 UI updates.
- **Add-in sideloaded, but the ribbon button doesn't appear** — confirm you're looking at the **Home** tab specifically (that's where `manifest.xml` places the `ExportGroup`), and that the workbook was reopened/refreshed after sideloading.
- **Export button in the task pane stays disabled** — the task pane hasn't received `Office.onReady` yet, or isn't actually running inside an Excel host frame (see the "This add-in only runs inside Excel" note above). Reload the task pane.
- **A cell shows a value or error you didn't expect** — check whether it's the Excel Table structured-reference case described above under **Known limitations** before assuming it's a bug; it's a known, safety-preserving difference in how this option specifically extracts formulas.
