(function () {
  'use strict';

  const payload = JSON.parse(document.getElementById('budget-data').textContent);
  const engine = InteractivePlanRender.createEngine(payload);
  const root = document.getElementById('root');

  InteractivePlanRender.renderApp({
    root,
    payload,
    engine,
    sourceLine: `Source: ${payload.sourceFile} · editable prototype, recalculated live in your browser`,
    onSave: openSaveDialog,
  });

  // ---- Save as new version: writes a real .xlsx with live formulas + updated cached values ----
  const dialog = InteractivePlanRender.el('dialog', { class: 'save-dialog' });
  const form = InteractivePlanRender.el('form', { method: 'dialog' });
  form.appendChild(InteractivePlanRender.el('label', { for: 'author-name', text: 'Your name or initials' }));
  const authorInput = InteractivePlanRender.el('input', { type: 'text', id: 'author-name', required: 'required', placeholder: 'e.g. J. Rivera' });
  form.appendChild(authorInput);
  const actions = InteractivePlanRender.el('div', { class: 'actions' });
  actions.appendChild(InteractivePlanRender.el('button', { type: 'button', class: 'btn cancel', text: 'Cancel', onclick: () => dialog.close() }));
  actions.appendChild(InteractivePlanRender.el('button', { type: 'submit', class: 'btn', text: 'Download .xlsx' }));
  form.appendChild(actions);
  form.addEventListener('submit', (e) => {
    e.preventDefault();
    downloadVersion(authorInput.value.trim() || 'Unknown');
    dialog.close();
  });
  dialog.appendChild(form);
  document.body.appendChild(dialog);

  function openSaveDialog() {
    dialog.showModal();
    authorInput.focus();
  }

  function downloadVersion(author) {
    const now = new Date();
    const stamp = now.toISOString().slice(0, 16).replace('T', ' ');
    const arrayBuffer = InteractivePlanRender.buildWorkbookBytes(payload, engine, {
      versionInfo: { author, stamp, basedOn: payload.sourceFile },
    });

    const safeAuthor = author.replace(/[^a-z0-9]+/gi, '-').replace(/^-+|-+$/g, '') || 'user';
    const dateTag = now.toISOString().slice(0, 10).replace(/-/g, '');
    const timeTag = now.toTimeString().slice(0, 5).replace(':', '');
    const base = payload.sourceFile.replace(/\.xlsx$/i, '');
    const filename = `${base}_v-${safeAuthor}-${dateTag}-${timeTag}.xlsx`;

    const blob = new Blob([arrayBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }
})();
