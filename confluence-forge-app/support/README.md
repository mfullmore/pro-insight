# Confluence App — Setup & Support Guide

**What this is:** a Forge Custom UI macro for Confluence. It reads a budget `.xlsx` already attached to the page it's placed on, renders it as an interactive, editable plan in place, and writes edits back as a new version of that same attachment — no download, no re-upload, no separate tool.

---

## For end users

1. On any Confluence page, attach a budget `.xlsx` file (drag it onto the page, or use the attachment picker).
2. Add the **Interactive Plan** macro to the page — type `/Interactive Plan` or find it in the macro browser.
3. The macro automatically reads the first `.xlsx` attached to the page and renders it. No configuration needed.
4. Edit any plain-number cell — formulas recalculate live as you type.
5. Click **Save as new version…** to write your edits back as a new version of the attachment. It shows up in the page's normal Confluence attachment version history, same as any manual re-upload would.

---

## For Confluence/Atlassian admins: installing this on your site

This app deploys through Atlassian Forge. One-time setup, then it's available to add to any page on the site you install it on.

### Prerequisites

- Node.js 24+ (see **Node version note** below — this project also vendors its own copy, so your system Node doesn't strictly need to be this new)
- An Atlassian account with access to the target Confluence site
- The Forge CLI (installed automatically via `npm install`, below)

### 1. Install dependencies

```bash
cd confluence-forge-app
npm install
```

### 2. Sign in to Forge

```bash
npx forge login
```

Follow the prompt — it opens a browser window to authenticate. **This is the one step that genuinely requires a human at the keyboard** — never share or script your Atlassian credentials.

### 3. Create a Developer Space (one-time, tenant-level)

Every Forge app must belong to a "Developer Space" — a billing/admin container. There is **no CLI-only way to create one** — you'll get `You are not currently a member of a Developer Space` from `forge register` until this exists:

1. Go to `https://developer.atlassian.com/console/`
2. Create a new Developer Space

⚠️ The console will tell you that creating a Developer Space makes you responsible for any applicable usage charges — confirm with whoever owns billing for your Atlassian org before doing this on a real (non-trial) tenant.

### 4. Register the app

```bash
npx forge register "Interactive Plan" -s <your-developer-space-id> --accept-terms
```

This assigns a real, server-issued app ID and writes it into `manifest.yml` automatically. Only ever needs to be done once for this app — don't re-run it for routine updates.

### 5. Build and deploy

```bash
npm run build
npx forge deploy -e development
```

(swap `development` for `staging` or `production` once you're past initial testing — see **Environments** below)

### 6. Install on your site

```bash
npx forge install
```

Follow the prompts to pick the target Confluence site. Once installed, the **Interactive Plan** macro becomes available to add to any page on that site.

---

## Node version note

⚠️ Deploying this app needs **Node 24+**. If your system Node is older, don't upgrade it just for this — it could disturb other Node-dependent projects. This project vendors its own local copy at `.local-node/`, so you can run every `forge` command through that instead of your system Node:

```bash
./.local-node/bin/node ./node_modules/@forge/cli/out/bin/cli.js deploy -e development
./.local-node/bin/node ./node_modules/@forge/cli/out/bin/cli.js install
./.local-node/bin/node ./node_modules/@forge/cli/out/bin/cli.js logs -e development
```

**Why Node 24 specifically:** several of `@forge/cli`'s own dependencies (`archiver`, and packages pulled in by `chalk`/`cli-table3`) ship as pure ESM with no CommonJS fallback. Loading them only works reliably through Node's native `require(esm)` support, which is mature starting at Node 24 — on Node 22 it either needs experimental flags (which then collide with another bundled dependency, `v8-compile-cache`) or fails outright. Node 24 just works, with zero code-level workarounds.

If your system Node is already 24+, ignore `.local-node/` entirely and use plain `npx forge ...`.

---

## Rebuilding after changing shared code

This app **vendors** (copies) SheetJS, HyperFormula, and the shared `convert-core.js` / `render-core.js` files from `../browser-converter/core/` at build time. `npm run build` re-copies whatever is currently in that folder and rebuilds the Forge bridge bundle. After changing anything in `browser-converter/core/`:

```bash
npm run build
npx forge deploy -e development   # or your target environment
```

You do **not** need to run `forge install` again unless `manifest.yml` itself changed (new scopes, new modules, etc.) — a plain re-deploy updates the app for everyone who already has it installed, automatically.

---

## Directory structure

```
confluence-forge-app/
  manifest.yml               Forge app manifest — modules, permission scopes, app ID
  build-static.js            vendors SheetJS/HyperFormula/convert-core.js/render-core.js
                              from ../browser-converter/core/ into static/main/vendor/
  forge-bridge-entry.js       thin entry point bundled into forge-bridge.bundle.js
                              (bridges the Custom UI iframe to Forge's invoke() API)
  src/resolvers/index.js      server-side resolver — reads/writes the page's attachment
                              via Confluence's REST API (runs in Atlassian's cloud, not the browser)
  static/main/                 the Custom UI macro itself (runs in the browser, inside
                                the macro's iframe)
    index.html
    app.js                      thin bootstrap — invoke()s the resolver for load/save,
                                 calls the shared render-core.js for everything else
    app.css
    vendor/                      vendored copies of the shared libraries (see above)
  .local-node/                 project-local Node 24 binary (see Node version note)
```

---

## Known limitations

Same recalculation engine as the other three options in this project — `INDIRECT()` and Excel Table structured references aren't supported, and are flagged explicitly rather than silently computed wrong. The flagged cell shows Excel's last-known value marked **✱** (frozen, non-live); anything that depends on it shows a live error instead. See `browser-converter/support/README.md` for the full explanation — it's identical here since the rendering logic is shared.

---

## Troubleshooting

- **`You are not currently a member of a Developer Space`** on `forge register` — you haven't created/joined one yet. See step 3 above.
- **`401 Unauthorized; scope does not match`** on attachment read/write — check `manifest.yml`'s `permissions.scopes` includes `read:attachment:confluence` and `write:attachment:confluence` (the granular `verb:resource:confluence` form — a different naming pattern than the older `verb:confluence-resource` scopes, easy to mix up).
- **Deploy fails with an ESM-related error** (e.g. mentioning `archiver` or `require() of ES Module`) — your Node is older than 24. Use the vendored `.local-node/` binary (see **Node version note**) or upgrade your system Node.
- **Macro appears to vanish from a page after an edit** — usually a stale view-mode render cache, not a real bug. Switch to Edit mode to confirm the macro is still there and working, then republish the page to refresh the cached view.
- **`forge deploy` succeeds but the macro still shows old behavior** — Confluence's Custom UI resources aren't nearly as aggressively cached as, say, SharePoint's (see the SharePoint web part's support doc for a much worse version of this problem), but a hard-refresh of the page clears it if you do hit stale content.
- **Debugging resolver issues** — `./.local-node/bin/node ./node_modules/@forge/cli/out/bin/cli.js logs -e development` shows the resolver's own `console.log`/`console.error` output, since it runs server-side and you can't just open browser DevTools for it.

---

## Environments

Forge supports `development`, `staging`, and `production` environments, deployed and installed independently:

```bash
npx forge deploy -e staging
npx forge install -e staging
```

Use `development` for day-to-day testing (what this project has been verified against). Promote to `staging`/`production` once ready for a broader rollout — each is a separate install, so existing `development` users are unaffected.
