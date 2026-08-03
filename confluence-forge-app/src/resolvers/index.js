// UNVERIFIED: written against documented Confluence Cloud REST API (v1
// content/attachment endpoints) and the Forge resolver/api patterns, but
// never run against a live site — there's no Atlassian account available in
// this environment to install a real Forge app and test it. Most likely
// things to double-check first: whether PUT vs POST is right for versioning
// an attachment by matching filename, and the exact shape of api.asApp()
// requestConfluence's Response object in this runtime.
import Resolver from '@forge/resolver';
import api, { route } from '@forge/api';

const resolver = new Resolver();

const XLSX_MEDIA_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

// Finds the first .xlsx attached to the page this macro is placed on, and
// returns its bytes as base64 (Custom UI <-> resolver messages are JSON,
// so binary has to be encoded).
resolver.define('getAttachment', async (req) => {
  const pageId = req.context.extension.content.id;
  console.log('getAttachment: pageId =', pageId);

  // v2 API — the v1 `/wiki/rest/api/content/{id}/child/attachment` endpoint
  // returned "401 Unauthorized; scope does not match" against every scope
  // combination tried (read:confluence-content.all, read:attachment:confluence),
  // confirmed via `forge logs`. v2 is the modern, Forge-granular-scope-aligned
  // surface, so trying that instead rather than continuing to guess scopes.
  const listRes = await api.asApp().requestConfluence(route`/wiki/api/v2/pages/${pageId}/attachments`);
  console.log('getAttachment: list status =', listRes.status, listRes.statusText);
  if (!listRes.ok) {
    const text = await listRes.text().catch(() => '');
    console.error('getAttachment: list body =', text.slice(0, 500));
    return { found: false, error: `Could not list attachments (${listRes.status})` };
  }
  const list = await listRes.json();
  console.log('getAttachment: found', list.results ? list.results.length : 0, 'attachments');
  const attachment = (list.results || []).find((a) => a.title && a.title.toLowerCase().endsWith('.xlsx'));
  if (!attachment) {
    return { found: false };
  }

  console.log('getAttachment: downloadLink =', attachment.downloadLink, 'id =', attachment.id);
  // requestConfluence requires every call to go through the `route` tag
  // (confirmed via forge logs: a plain string is rejected outright), but
  // `route` also requires actual static literal segments in the template —
  // passing an entire pre-built path as the whole interpolation ("no static
  // anchor") was rejected too. Reconstructing the known REST shape
  // (confirmed from attachment.downloadLink's own structure) with `route`
  // properly wrapping just the two dynamic IDs satisfies both rules.
  const fileRes = await api.asApp().requestConfluence(
    route`/wiki/rest/api/content/${pageId}/child/attachment/${attachment.id}/download`
  );
  console.log('getAttachment: download status =', fileRes.status);
  if (!fileRes.ok) {
    return { found: false, error: `Could not download attachment (${fileRes.status})` };
  }
  const buf = await fileRes.arrayBuffer();
  const base64 = Buffer.from(buf).toString('base64');

  return { found: true, filename: attachment.title, base64 };
});

// Uploads a new version of the budget workbook. Confluence's attachment
// endpoint treats an upload with a matching existing filename as a new
// version of that attachment (not a duplicate) — this is what lets "Save as
// new version" behave like real page history instead of a fresh download.
resolver.define('saveNewVersion', async (req) => {
  const pageId = req.context.extension.content.id;
  const { filename, base64 } = req.payload;
  if (!filename || !base64) {
    return { ok: false, error: 'Missing filename or file content.' };
  }

  const buffer = Buffer.from(base64, 'base64');
  const form = new FormData();
  form.append('file', new Blob([buffer], { type: XLSX_MEDIA_TYPE }), filename);
  form.append('minorEdit', 'true');
  form.append('comment', 'Updated via Interactive Plan macro');

  const res = await api.asApp().requestConfluence(route`/wiki/rest/api/content/${pageId}/child/attachment`, {
    method: 'PUT',
    headers: { 'X-Atlassian-Token': 'no-check' },
    body: form,
  });

  if (!res.ok) {
    const text = await res.text().catch(() => '');
    return { ok: false, error: `Upload failed (${res.status}): ${text.slice(0, 300)}` };
  }
  return { ok: true };
});

export const handler = resolver.getDefinitions();
