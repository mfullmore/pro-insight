# `static/` — Confluence Custom UI Resource Guide

**What this is:** the Confluence macro's actual front-end — the HTML/CSS/JS that runs inside the macro's iframe when someone views a page with the Interactive Plan macro on it. Registered in `manifest.yml` as the `main` resource:

```yaml
resources:
  - key: main
    path: static/main
```

---

## Directory contents

```
static/main/
  index.html      the macro's entry point — loads every script below, in order
  app.js            thin bootstrap: invoke()s the resolver for load/save,
                     calls the shared render-core.js for everything else
  app.css            styling — a manually-maintained fork of
                       ../browser-converter/core/app.css (see below — this
                       one is NOT auto-vendored, unlike everything in vendor/)
  vendor/            vendored + built libraries — regenerate, don't hand-edit
```

### `index.html`

Loads five scripts in a specific order — each sets a global the next one depends on:

```html
<script src="vendor/forge-bridge.bundle.js"></script>  <!-- sets window.ForgeBridge -->
<script src="vendor/xlsx.core.min.js"></script>          <!-- sets window.XLSX -->
<script src="vendor/hyperformula.bundle.js"></script>    <!-- sets window.HyperFormula -->
<script src="vendor/convert-core.js"></script>            <!-- sets window.InteractivePlanConvert -->
<script src="vendor/render-core.js"></script>              <!-- sets window.InteractivePlanRender -->
<script src="app.js"></script>                              <!-- uses all of the above -->
```

Don't reorder these — `app.js` and the shared `convert-core.js`/`render-core.js` assume `window.XLSX` and `window.HyperFormula` already exist by the time they run.

### `app.js`

The one file in this folder that's genuinely specific to Confluence — not shared with the other three options. It's a thin bootstrap: on load, calls `invoke('getAttachment')` to fetch the page's attached workbook (base64-encoded) from the resolver, converts it with the shared `InteractivePlanConvert.convertWorkbook`, renders it with `InteractivePlanRender.renderApp`, and wires the "Save as new version" button to `invoke('saveNewVersion', ...)`. See `src/resolvers/index.js` for the server-side half of those two `invoke()` calls.

### `app.css` — ⚠️ not automatically kept in sync

Unlike everything in `vendor/`, **`build-static.js` does not copy this file** — it's a manually-maintained fork of `../browser-converter/core/app.css`, with one intentional addition on top: the `.status`/`.status.visible`/`.status.good`/`.status.bad` rules at the bottom, which style this macro's own "Loading…" / "Saved" / error banner (`app.js`'s `setStatus()`). Those rules don't exist in the shared `app.css` because no other option needs them.

**This means a fix or new style added to `browser-converter/core/app.css` does not reach this file automatically.** It has to be manually re-applied here, on top of the `.status` additions. This was confirmed as a real gap during testing: this file was missing the `.frozen` class (added to the shared `app.css` for the frozen-value fix) until it was manually patched in — the class was still being *applied* correctly by the shared `render-core.js`, it just had no CSS rule behind it here, so frozen cells rendered without their distinctive amber background.

**When you change `browser-converter/core/app.css`, check this file for drift too** (`diff confluence-forge-app/static/main/app.css browser-converter/core/app.css` — expect only the `.status` block to differ) and manually port over anything relevant.

### `vendor/`

Everything in here is generated — **never edit these files directly**, changes will be silently overwritten by the next build:

| File | Source | Rebuilt by |
|---|---|---|
| `xlsx.core.min.js` | `../browser-converter/core/xlsx.core.min.js` | `node build-static.js` (plain copy) |
| `hyperformula.bundle.js` | `../browser-converter/core/hyperformula.bundle.js` | `node build-static.js` (plain copy) |
| `convert-core.js` | `../browser-converter/core/convert-core.js` | `node build-static.js` (plain copy) |
| `render-core.js` | `../browser-converter/core/render-core.js` | `node build-static.js` (plain copy) |
| `forge-bridge.bundle.js` | `forge-bridge-entry.js` (this project) | `esbuild`, via `npm run build` |

`npm run build` (from the project root) runs `build-static.js` and the `esbuild` step together. See `confluence-forge-app/support/README.md` for the full build/deploy process.

---

## Making a change here

- **Changing the macro's own behavior** (load/save flow, status messages) → edit `app.js` directly, then `npm run build` — no vendoring involved, `app.js` isn't a generated file.
- **Changing rendering/recalculation behavior** (grid layout, formula handling, the frozen-value fix, etc.) → these live in `../browser-converter/core/`, shared by all four options. Edit there, then run `npm run build` here to re-vendor and rebuild, then redeploy. See `browser-converter/support/README.md` for what changing shared code actually entails.
- **Changing styling** → edit `browser-converter/core/app.css` for anything that should apply everywhere, **then manually port the change into `static/main/app.css` too** (see the warning above) — this is the one file in this folder that build tooling won't do for you.
