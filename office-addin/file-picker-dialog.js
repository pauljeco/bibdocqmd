/* The dialog uses an unmodified native file input in its own Office window. */
(() => {
  'use strict';
  const settings = {
    'bib-file': { label: 'Select bibliography files', accept: '.bib,.bibtex,text/plain', multiple: true },
    'csl-file': { label: 'Select a CSL style', accept: '.csl,application/xml,text/xml' },
    'export-docx': { label: 'Select your saved Word document', accept: '.docx' },
    'export-preparation': { label: 'Select the preparation ZIP', accept: '.zip' },
    'export-markdown': { label: 'Select Word’s Markdown export', accept: '.md,.zip' },
  };
  const query = new URLSearchParams(location.search);
  const configuration = settings[query.get('input')];
  const token = query.get('token');
  const input = document.getElementById('dialog-files');
  const status = document.getElementById('dialog-status');
  const cancel = document.getElementById('dialog-cancel');
  if (!configuration || !/^[0-9a-f-]{36}$/i.test(token || '')) {
    status.textContent = 'Invalid file selection request. Close this window and try again.';
    return;
  }
  const send = (message) => Office.context.ui.messageParent(JSON.stringify({ version: 1, token, ...message }));
  Office.onReady(() => {
    document.getElementById('dialog-heading').textContent = configuration.label;
    document.getElementById('dialog-label').textContent = configuration.label;
    input.accept = configuration.accept;
    input.multiple = Boolean(configuration.multiple);
    input.disabled = false;
    cancel.disabled = false;
    status.textContent = 'Choose files to return them to Word.';
    cancel.addEventListener('click', () => send({ type: 'cancel' }));
    input.addEventListener('change', async () => {
      const files = [...input.files];
      if (!files.length) return;
      input.disabled = true;
      try {
        if (files.length > 100 || files.reduce((sum, file) => sum + file.size, 0) > 25_000_000) {
          throw new Error('Select at most 25 MB of files at a time.');
        }
        send({ type: 'start', files: files.map(({ name, size, type, lastModified }) => ({ name, size, type, lastModified })) });
        for (let file = 0; file < files.length; file += 1) {
          status.textContent = `Returning file ${file + 1} of ${files.length} to Word…`;
          let index = 0;
          for (let offset = 0; offset < files[file].size; offset += 49152) {
            const bytes = new Uint8Array(await files[file].slice(offset, offset + 49152).arrayBuffer());
            send({ type: 'chunk', file, index, data: btoa(String.fromCharCode(...bytes)) });
            index += 1;
            // Give the Office message bridge time to process bounded chunks.
            await new Promise((done) => setTimeout(done, 0));
          }
        }
        send({ type: 'finish' });
        status.textContent = 'Files returned. Closing this window…';
      } catch (error) {
        status.textContent = error.message || String(error);
        send({ type: 'error', error: status.textContent });
      }
    });
  });
})();
