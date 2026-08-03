# SharePoint Web Part — Setup & Support Guide

**What this is:** an SPFx (SharePoint Framework) web part. It reads a budget `.xlsx` file from a document library on the site it's placed on, renders it as an interactive, editable plan directly on the page, and writes edits back as a new version of that same file — no download, no re-upload, no separate tool.

---

## For end users

1. Edit the SharePoint page you want the plan on, and add the **Interactive Plan** web part (search for it in the web part picker).
2. Open the web part's property pane (the pencil/edit icon) and set **File server-relative URL** to the workbook's path in a document library, e.g. `/sites/Finance/Shared Documents/Q3 Budget.xlsx`.
3. Publish the page. The web part reads that file and renders it — no further configuration needed.
4. Edit any plain-number cell — formulas recalculate live as you type.
5. Click **Save as new version…** to write your edits back. It shows up in the file's normal SharePoint version history, same as a manual re-upload would.

---

## For SharePoint/Microsoft 365 admins: installing this on your tenant

This app deploys through a **SharePoint App Catalog**. One-time setup, then the web part is available to add to any page on any site in the tenant.

### Prerequisites

- **Node.js 22.x specifically** — `engines` in `package.json` pins `>=22.14.0 <23.0.0`. This differs from the Confluence app (which needs Node 24+) — if you have both projects on the same machine, keep them on separate Node versions (nvm, or the Confluence app's own vendored `.local-node/` — see that project's support doc — keeps them from colliding).
- A Microsoft 365 account that's a SharePoint admin for the target tenant
- A modern browser (the upload step below has no CLI path — see **A note on uploading**)

### 1. Install dependencies and vendor the shared libraries

```bash
cd sharepoint-webpart
npm install
node vendor.js
```

`vendor.js` copies SheetJS, HyperFormula, and the shared `convert-core.js` / `render-core.js` files from `../browser-converter/core/` into `src/webparts/interactivePlan/vendor/` — the web part's own build has no visibility into that sibling folder otherwise.

### 2. Build and package

```bash
npx heft build --clean
npx heft package-solution --production
```

(or just `npm run build`, which runs both together with `heft test` first)

This produces `sharepoint/solution/interactive-plan-webpart.sppkg` — the single file you upload in the next step.

### 3. Create a tenant App Catalog (one-time, if your tenant doesn't have one yet)

1. Go to the **SharePoint admin center** (`https://<yourtenant>-admin.sharepoint.com`)
2. **More features → Apps → Open** — if no catalog exists yet, this auto-provisions one (takes a minute or two the first time)

### 4. Upload the package

1. In the App Catalog's **Apps for SharePoint** library, click **Upload**
2. Select `sharepoint-webpart/sharepoint/solution/interactive-plan-webpart.sppkg`
3. When prompted, confirm **Deploy** to trust the client-side solution

> **A note on uploading:** SharePoint's upload dialogs (both the modern "Manage apps" page and the classic document library ribbon) use the browser's native file picker with no accessible HTML `<input type="file">` behind them — there's no way to script this step. It's a manual, one-click action every time you deploy a new build.

### 5. Add the web part to a page

Once uploaded, the web part becomes available tenant-wide (this solution is configured with `skipFeatureDeployment: true`, meaning no separate per-site "Add an app" step is *required* — see the caching note below if it doesn't show up immediately). Edit any page, search for **Interactive Plan** in the web part picker, add it, and configure the file path as described in **For end users** above.

---

## Rebuilding after changing shared code

This project vendors (copies) SheetJS, HyperFormula, and `convert-core.js`/`render-core.js` from `../browser-converter/core/` at build time. After changing anything in `browser-converter/core/`:

```bash
node vendor.js
npx heft build --clean
npx heft package-solution --production
```

Then **re-upload** the resulting `.sppkg` to the App Catalog (step 4 above) — every code change requires a fresh upload, there's no "just redeploy" shortcut like the Confluence app has.

---

## Directory structure

```
sharepoint-webpart/
  config/
    package-solution.json      solution ID, version, skipFeatureDeployment flag
  src/webparts/interactivePlan/
    InteractivePlanWebPart.ts    thin bootstrap — SharePoint REST calls for load/save,
                                  calls the shared render-core.js for everything else
    vendor/                       vendored copies of the shared libraries (regenerate
                                   with `node vendor.js` — don't hand-edit)
  sharepoint/solution/
    interactive-plan-webpart.sppkg   the actual deployable package — upload this
  vendor.js                      copies shared libraries from ../browser-converter/core/
```

---

## Vendor drift check

Unlike the other two platform integrations, this project currently has **no hand-maintained-fork risk**:

- `vendor.js` mechanically copies `xlsx.core.min.js`, `hyperformula.bundle.js`, `convert-core.js`, and `render-core.js` verbatim (`fs.copyFileSync`, no transformation), and wraps `app.css` as a TypeScript string constant (`appStyles.ts`) via a straight `JSON.stringify` — nothing here is hand-edited or hand-ported.
- `InteractivePlanWebPart.ts` calls `InteractivePlanConvert.convertWorkbook()` directly rather than re-implementing any part of it, because this web part reads the file via SheetJS (`XLSX.read` on bytes fetched over the SharePoint REST API) — the same extraction path `convert-core.js` itself expects. (Contrast with the Office Add-in, which reads through Excel's own JS API instead and — for that reason alone — has to hand-duplicate a couple of small pieces of extraction logic; see `office-addin/support/README.md`'s "Known drift risk" section.)

Verified as of this writing: every file in `src/webparts/interactivePlan/vendor/` is byte-for-byte identical to its `browser-converter/core/` source, including `appStyles.ts` (confirmed to contain the current `.frozen` CSS rule). The only real risk here is forgetting to **re-run `node vendor.js` and rebuild** after a `browser-converter/core/` change — a staleness risk, not a drift-while-editing risk — already covered under "Rebuilding after changing shared code" above.

---

## Known limitations

Same recalculation engine as the other three options in this project — `INDIRECT()` and Excel Table structured references aren't supported, and are flagged explicitly rather than silently computed wrong. The flagged cell shows Excel's last-known value marked **✱** (frozen, non-live); anything that depends on it shows a live error instead. See `browser-converter/support/README.md` for the full explanation.

---

## Troubleshooting

This project surfaced more real, environment-specific bugs during testing than the other three combined — worth reading even if nothing looks broken yet, since some of these fail *silently* in the browser console rather than obviously.

### "Something went wrong" error box where the web part should render

Open the browser console. Two distinct root causes have been found and fixed, both from the same underlying pattern: **a library's UMD module-detection code gets confused by webpack's bundling environment** and silently exports to the wrong place.

- **`TypeError: ... RegExpParser is not a constructor`** — a HyperFormula dependency (`regexp-to-ast`, pulled in via chevrotain) detects SharePoint's own AMD script loader (`window.define`, present on real SharePoint pages, absent everywhere else this code runs) and wrongly registers itself as an AMD module instead of a plain value. Fixed in `browser-converter/core/hf-entry.js`'s build command with `--define:define=undefined` — if you see this, the `browser-converter` project's `hyperformula.bundle.js` predates that fix; rebuild it there (`npm run build:hf`), re-vendor (`node vendor.js`), rebuild, and re-upload.
- **`XLSX.read is not a function`** (or similar — `window.XLSX` exists but has no methods) — SheetJS's own UMD wrapper sees webpack's `exports` object and writes its API onto that instead of the real `XLSX` variable, leaving `window.XLSX` empty. Fixed the same way, with `browser-converter`'s `npm run build:xlsx`. Same remediation: rebuild there, re-vendor, rebuild, re-upload.

Both of these are specific to being bundled by webpack (as this project does) — they never occur in the other three options, which all load these libraries as plain `<script>` tags.

### The web part loads, but every sheet after the first is blank when you click its tab

HyperFormula throws a genuine JS exception (rather than returning a normal error value) for at least one known formula pattern (`INDIRECT`), which used to abort the whole grid-rendering loop partway through. Fixed in `browser-converter/core/render-core.js`'s `getVal()` — it now catches any exception from the engine and displays it as a live error cell instead of crashing. If you see a blank grid after switching tabs, check the console for an uncaught exception; the fix should already be in current `render-core.js`, so re-vendoring + rebuilding + re-uploading resolves it if you're on an older build.

### You uploaded a fixed build, but the browser still shows the old broken behavior

SharePoint serves this web part's JS bundle with a **very aggressive cache header** (`Cache-Control: public, max-age=31536000` — one year), and this project's build doesn't add a content hash to the bundle's filename, so the URL never changes between uploads. A normal page reload — even a hard reload (Cmd+Shift+R / Ctrl+Shift+R) — often isn't enough, because that only forces revalidation for resources fetched during the page's own navigation, not ones a script loads afterward (which is how SharePoint's own component loader fetches this bundle).

**What reliably works:** clear "Cached images and files" in your browser's settings, then reload. (This will likely also sign you out of the site — sign back in.) You can verify the *server* already has the fix, independent of your browser cache, with:

```js
fetch("https://<site>/sites/appcatalog/ClientSideAssets/<solution-id>/interactive-plan-web-part.js", { cache: "no-store" })
  .then(r => r.text())
  .then(t => console.log(t.length, t.slice(0, 200)))
```
(run in the browser console on any page on your tenant) — if this returns your latest code, the deploy worked and the problem is purely local browser caching.

### The web part doesn't appear in the web part picker right after uploading

Even with `skipFeatureDeployment: true`, there can be a propagation delay before a newly uploaded app shows up in a page editor's web part search. If it's been more than a few minutes, force it explicitly:

1. Go to the site's **Site Contents → New → App** (or `/_layouts/15/AddAnApp.aspx`)
2. Find **interactive-plan-webpart-client-side-...** under "Apps you can add" and click through to add it
3. Return to the page editor — the web part should now be searchable

### Deploy/build fails with a Node version-related error

Check `node -v` — this project needs Node 22.x (`>=22.14.0 <23.0.0`), not 24+. If you also work with the Confluence app on the same machine (which needs Node 24+), use separate Node versions per project (nvm, or run each project's own preferred Node binary explicitly) rather than a single global Node install.
