// UNVERIFIED AT RUNTIME: this compiles and bundles successfully against the
// real SPFx toolchain (see below), but has never run inside an actual
// SharePoint site — there's no Microsoft 365 tenant available in this
// environment to add the web part to a page and test it. The REST API calls
// (GetFileByServerRelativeUrl, the X-HTTP-Method: PUT override for writing
// file content) are written against Microsoft's documented SharePoint REST
// API, but that's a different confidence level than something click-tested.
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, type SPHttpClientResponse } from '@microsoft/sp-http';

import './vendor/xlsx.core.min.js';
import './vendor/hyperformula.bundle.js';
import './vendor/convert-core.js';
import './vendor/render-core.js';
import { APP_CSS } from './vendor/appStyles';

declare const XLSX: any;
declare const InteractivePlanConvert: any;
declare const InteractivePlanRender: any;

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
