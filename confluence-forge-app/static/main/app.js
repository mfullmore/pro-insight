(function () {
  'use strict';

  const invoke = window.ForgeBridge.invoke;

  function base64ToArrayBuffer(base64) {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    return bytes.buffer;
  }
  function arrayBufferToBase64(buffer) {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    return btoa(binary);
  }

  const root = document.getElementById('root');
  let payload = null;
  let engine = null;

  // ---- Save: unlike browser-converter/office-addin (which trigger a browser
  // download), this writes the new version directly back to the page as an
  // attachment, via the resolver — the one capability genuinely unique to
  // living inside Confluence rather than being opened as a standalone file. ----
  async function saveNewVersion() {
    setStatus('Saving new version…');
    try {
      const arrayBuffer = InteractivePlanRender.buildWorkbookBytes(payload, engine);
      const base64 = arrayBufferToBase64(arrayBuffer);

      const result = await invoke('saveNewVersion', { filename: payload.sourceFile, base64 });
      if (result && result.ok) {
        setStatus('Saved as a new attachment version.', 'good');
      } else {
        setStatus('Save failed: ' + ((result && result.error) || 'unknown error'), 'bad');
      }
    } catch (err) {
      console.error(err);
      setStatus('Save failed: ' + err.message, 'bad');
    }
  }

  function setStatus(msg, kind) {
    let statusEl = document.getElementById('macro-status');
    if (!statusEl) {
      statusEl = document.createElement('div');
      statusEl.id = 'macro-status';
      statusEl.className = 'status';
      document.body.insertBefore(statusEl, root);
    }
    statusEl.textContent = msg;
    statusEl.className = 'status visible' + (kind ? ' ' + kind : '');
  }

  // ---- Bootstrap: read whatever .xlsx is already attached to this page ----
  async function init() {
    setStatus('Loading attached workbook…');
    try {
      const result = await invoke('getAttachment');
      if (!result || !result.found) {
        setStatus(
          (result && result.error) || 'No .xlsx attached to this page yet. Attach a budget workbook and reload the macro.',
          'bad'
        );
        return;
      }
      const buf = base64ToArrayBuffer(result.base64);
      const wb = XLSX.read(buf, { type: 'array', cellFormula: true, cellNF: true });
      payload = InteractivePlanConvert.convertWorkbook(wb, result.filename);
      engine = InteractivePlanRender.createEngine(payload);
      setStatus('', null);
      document.getElementById('macro-status').className = 'status';

      InteractivePlanRender.renderApp({
        root,
        payload,
        engine,
        sourceLine: `Source: ${payload.sourceFile} · attached to this page · recalculated live in your browser`,
        footerText: `Read from ${payload.sourceFile}, the .xlsx attached to this page. Formulas recalculate client-side via HyperFormula.`,
        onSave: saveNewVersion,
      });
    } catch (err) {
      console.error(err);
      setStatus('Failed to load workbook: ' + err.message, 'bad');
    }
  }

  init();
})();
