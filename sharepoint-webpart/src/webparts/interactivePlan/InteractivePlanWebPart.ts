// Verified at runtime against a real Microsoft 365 tenant/App Catalog: read
// a workbook from a document library, rendered all sheets, live edits and
// "Save as new version" confirmed via the file's actual SharePoint version
// history. One real bug found and fixed this way that compiling/bundling
// cleanly never would have caught — see ensureHyperFormulaLoaded() below.
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, type SPHttpClientResponse } from '@microsoft/sp-http';

import './vendor/xlsx.core.min.js';
import './vendor/convert-core.js';
import { APP_CSS } from './vendor/appStyles';
import { HYPERFORMULA_SOURCE } from './vendor/hyperformulaSource';
import { RENDER_CORE_SOURCE } from './vendor/renderCoreSource';

declare const XLSX: any;
declare const InteractivePlanConvert: any;
declare const InteractivePlanRender: any;

// HyperFormula and render-core.js specifically can't be loaded via a
// side-effect `import` like xlsx.core.min.js/convert-core.js above.
//
// tsc's allowJs pipeline (target: es5) silently corrupts HyperFormula's
// evaluator when it recompiles the already-built bundle: parsing still
// works (HyperFormula.version, getCellFormula, getCellType all report
// correctly) but getCellValue throws "Value of the formula cell is not
// computed" for every formula, even a trivial =A1+B1 in a brand-new engine
// — reproduced in complete isolation outside of SPFx/webpack, so this is
// tsc's ES5 downlevel of this specific minified bundle at fault, not
// anything SharePoint-specific.
//
// render-core.js has to move with it: its outer IIFE captures HyperFormula
// from `root.HyperFormula` exactly once, at the moment it executes
// (`factory(root.XLSX, root.HyperFormula)`) — a static `import` runs at
// module load time, before HyperFormula has been injected, so
// createEngine would close over `undefined`.
//
// Both instead live as inert TS string constants (see vendor.js) and get
// executed as real, untransformed JS at runtime, in order, so tsc never
// parses them as code.
//
// This tenant enforces a CSP that silently blocks inline <script> execution
// (confirmed: appending a <script> with .textContent set does nothing, no
// exception, no console error — this tenant's script-src has no
// 'unsafe-inline') — the first version of this fix used that approach and
// it never ran. `new Function(source)()` was confirmed to still be
// permitted (this tenant's CSP allows 'unsafe-eval') and produces identical
// results to a real <script> tag for this code (both were verified against
// the exact same isolated =A1+B1 reproduction). `window.HyperFormula =
// HyperFormula;` inside the executed source still sets a true global
// either way, since `window` isn't affected by the enclosing function's
// scope — only that file's own top-level `var`/`let` declarations stay
// local to it instead of leaking into global scope, which doesn't matter
// here since nothing depends on those.
let engineLibsInjected = false;
function ensureEngineLibsLoaded(): void {
  if (engineLibsInjected) return;
  for (const source of [HYPERFORMULA_SOURCE, RENDER_CORE_SOURCE]) {
    // eslint-disable-next-line no-new-func -- deliberate: see the comment above this function for why.
    new Function(source)();
  }
  engineLibsInjected = true;
}

export interface IInteractivePlanWebPartProps {
  fileServerRelativeUrl: string;
}

const EXTRA_CSS = `
.spfx-status { display: none; margin: 12px 0; padding: 10px 14px; border-radius: 6px; font-size: 0.85rem; background: var(--card); border: 1px solid var(--rule); color: var(--ink-muted); }
.spfx-status.visible { display: block; }
.spfx-status.good { background: var(--status-good-bg); color: var(--status-good-fg); border-color: var(--status-good-fg); }
.spfx-status.bad { background: var(--status-bad-bg); color: var(--status-bad-fg); border-color: var(--status-bad-fg); }
`;

let stylesInjected = false;
function ensureStylesInjected(): void {
  if (stylesInjected) return;
  const style = document.createElement('style');
  style.textContent = APP_CSS + EXTRA_CSS;
  document.head.appendChild(style);
  stylesInjected = true;
}

function el(tag: string, attrs?: Record<string, any>, children?: (HTMLElement | null | undefined)[]): HTMLElement {
  const node = document.createElement(tag);
  if (attrs) {
    for (const k of Object.keys(attrs)) {
      const v = attrs[k];
      if (k === 'text') node.textContent = v;
      else if (k.startsWith('on')) node.addEventListener(k.slice(2), v);
      else node.setAttribute(k, v);
    }
  }
  (children || []).forEach((c) => c && node.appendChild(c));
  return node;
}

export default class InteractivePlanWebPart extends BaseClientSideWebPart<IInteractivePlanWebPartProps> {
  private payload: any = null;
  private engine: any = null;
  private rootEl: HTMLElement | null = null;
  private statusEl: HTMLElement | null = null;

  public render(): void {
    ensureStylesInjected();

    if (!this.properties.fileServerRelativeUrl) {
      this.domElement.innerHTML = '';
      this.domElement.appendChild(el('div', {
        class: 'spfx-status visible bad',
        text: 'Configure this web part (the pencil/edit icon) with the server-relative path to a budget .xlsx in a document library, e.g. /sites/Finance/Shared Documents/Q3 Budget.xlsx',
      }));
      return;
    }

    this.domElement.innerHTML = '';
    this.statusEl = el('div', { class: 'spfx-status' }) as HTMLElement;
    this.rootEl = el('div') as HTMLElement;
    this.domElement.appendChild(this.statusEl);
    this.domElement.appendChild(this.rootEl);

    this.loadWorkbook().catch((err) => this.setStatus('Failed to load: ' + err.message, true));
  }

  private setStatus(msg: string, isError?: boolean): void {
    if (!this.statusEl) return;
    this.statusEl.textContent = msg;
    this.statusEl.className = msg ? `spfx-status visible ${isError ? 'bad' : 'good'}` : 'spfx-status';
  }

  // ---- Read: fetches the configured file directly from the document
  // library via the SharePoint REST API. Auth is ambient — this.context
  // .spHttpClient runs in the page's own security context, no separate
  // OAuth flow needed, the same "native to the platform" advantage the
  // Confluence Forge app gets from api.asApp(). ----
  private async loadWorkbook(): Promise<void> {
    this.setStatus('Loading workbook…');
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const filePath = this.properties.fileServerRelativeUrl;
    const endpoint = `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(filePath)}')/$value`;

    const res: SPHttpClientResponse = await this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
    if (!res.ok) throw new Error(`Could not read file (HTTP ${res.status})`);
    const buf = await res.arrayBuffer();

    const fileName = filePath.split('/').pop() || 'workbook.xlsx';
    const wb = XLSX.read(buf, { type: 'array', cellFormula: true, cellNF: true });
    this.payload = InteractivePlanConvert.convertWorkbook(wb, fileName);
    ensureEngineLibsLoaded();
    this.engine = InteractivePlanRender.createEngine(this.payload);
    this.setStatus('');
    this.renderApp();
  }

  private renderApp(): void {
    if (!this.rootEl) return;
    InteractivePlanRender.renderApp({
      root: this.rootEl,
      payload: this.payload,
      engine: this.engine,
      sourceLine: `Source: ${this.payload.sourceFile} · read from this site · recalculated live in your browser`,
      footerText: `Read from ${this.payload.sourceFile}. Formulas recalculate client-side via HyperFormula.`,
      onSave: () => this.saveNewVersion(),
    });
  }

  // ---- Save: writes the new version directly back to the same document
  // library file via the REST API's "POST + X-HTTP-Method: PUT" override
  // (the documented pattern for replacing file content through this
  // endpoint). If version history is enabled on the library, this becomes a
  // new version automatically — no separate download/re-upload step. ----
  private async saveNewVersion(): Promise<void> {
    this.setStatus('Saving new version…');
    try {
      const arrayBuffer = InteractivePlanRender.buildWorkbookBytes(this.payload, this.engine);

      const siteUrl = this.context.pageContext.web.absoluteUrl;
      const filePath = this.properties.fileServerRelativeUrl;
      const endpoint = `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(filePath)}')/$value`;

      const res: SPHttpClientResponse = await this.context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
        headers: { 'X-HTTP-Method': 'PUT' },
        body: arrayBuffer,
      });
      if (!res.ok) throw new Error(`Upload failed (HTTP ${res.status})`);
      this.setStatus('Saved as a new version.');
    } catch (err) {
      this.setStatus('Save failed: ' + (err as Error).message, true);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Point this web part at a budget workbook already in a document library on this site.' },
          groups: [
            {
              groupName: 'Source file',
              groupFields: [
                PropertyPaneTextField('fileServerRelativeUrl', {
                  label: 'File server-relative URL',
                  description: "e.g. /sites/Finance/Shared Documents/Q3 Budget.xlsx",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
