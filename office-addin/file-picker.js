/* Local file selection helpers; usable in Office webviews and normal browsers. */
(function exposeFilePicker(root, factory) {
  const helpers = factory(root);
  if (typeof module === 'object' && module.exports) module.exports = helpers;
  else root.BibWordFilePicker = helpers;
}(typeof globalThis !== 'undefined' ? globalThis : this, (root) => {
  'use strict';
  const maximumBytes = 25_000_000;
  let activeDialog = null;

  function createFileReceiver(token, multiple) {
    let files = null;
    return function receive(message) {
      if (message.token !== token || message.version !== 1) throw new Error('Invalid file chooser message.');
      if (message.type === 'cancel') return { cancelled: true };
      if (message.type === 'error') throw new Error(String(message.error || 'File selection failed.').slice(0, 500));
      if (message.type === 'start') {
        if (files || !Array.isArray(message.files) || !message.files.length
            || message.files.length > (multiple ? 100 : 1)) throw new Error('Invalid file selection.');
        let total = 0;
        files = message.files.map((file) => {
          if (typeof file.name !== 'string' || !file.name || file.name.length > 1024
              || !Number.isSafeInteger(file.size) || file.size < 0) throw new Error('Invalid file metadata.');
          total += file.size;
          return { ...file, chunks: [], received: 0, next: 0 };
        });
        if (total > maximumBytes) throw new Error('Select at most 25 MB of files at a time.');
      } else if (message.type === 'chunk') {
        const file = Number.isInteger(message.file) && message.file >= 0 ? files?.[message.file] : null;
        if (!file || message.index !== file.next || typeof message.data !== 'string' || message.data.length > 65536) {
          throw new Error('Invalid or out-of-order file data.');
        }
        const bytes = Uint8Array.from(root.atob(message.data), (char) => char.charCodeAt(0));
        file.received += bytes.length;
        if (file.received > file.size) throw new Error('The received file is larger than declared.');
        file.chunks.push(bytes);
        file.next += 1;
      } else if (message.type === 'finish') {
        if (!files || files.some((file) => file.received !== file.size)) throw new Error('Incomplete file transfer; select the files again.');
        return { files: files.map((file) => new root.File(file.chunks, file.name, {
          type: typeof file.type === 'string' ? file.type : '',
          lastModified: Number.isFinite(file.lastModified) ? file.lastModified : Date.now(),
        })) };
      } else throw new Error('Unsupported file chooser message.');
      return null;
    };
  }

  function openWordFileDialog(input, document, office, showError) {
    if (activeDialog) throw new Error('Close the current file selection window first.');
    const token = root.crypto.randomUUID();
    const url = new root.URL('file-picker-dialog.html', document.baseURI);
    url.searchParams.set('input', input.id);
    url.searchParams.set('token', token);
    const receiver = createFileReceiver(token, input.multiple);
    const attempt = { dialog: null };
    activeDialog = attempt;
    function close() {
      if (activeDialog === attempt) activeDialog = null;
      attempt.dialog?.close();
      attempt.dialog = null;
    }
    try {
      office.context.ui.displayDialogAsync(url.href, { height: 42, width: 45, displayInIframe: false }, (result) => {
        if (result.status !== office.AsyncResultStatus.Succeeded) {
          close();
          showError(new Error(`Word could not open the file selection window (${result.error?.code || 'unknown'}). ${result.error?.message || ''}`));
          return;
        }
        attempt.dialog = result.value;
        result.value.addEventHandler(office.EventType.DialogMessageReceived, (event) => {
          if (activeDialog !== attempt) return;
          try {
            if (event.origin && event.origin !== url.origin) throw new Error('Unexpected file chooser origin.');
            if (typeof event.message !== 'string' || event.message.length > 100_000) throw new Error('Invalid file chooser message size.');
            const transfer = receiver(JSON.parse(event.message));
            if (!transfer) return;
            if (transfer.files) {
              const data = new root.DataTransfer();
              transfer.files.forEach((file) => data.items.add(file));
              input.files = data.files;
            }
            close();
            if (transfer.files) input.dispatchEvent(new root.Event('change', { bubbles: true }));
          } catch (error) { close(); showError(error); }
        });
        result.value.addEventHandler(office.EventType.DialogEventReceived, (event) => {
          if (activeDialog === attempt) activeDialog = null;
          attempt.dialog = null;
          if (event.error !== 12006) showError(new Error(`Word file selection window closed with error ${event.error}.`));
        });
      });
    } catch (error) { close(); throw error; }
  }

  function bindFilePickers(document, { office = root.Office } = {}) {
    document.querySelectorAll('input[type="file"]').forEach((input) => {
      const feedback = document.getElementById(`${input.id}-picker-status`);
      const clearFeedback = () => {
        if (feedback) { feedback.hidden = true; feedback.textContent = ''; }
      };
      const showError = (error) => {
        if (feedback) {
          feedback.textContent = `${error.message || 'File selection failed.'} Use Open local tools in browser if needed.`;
          feedback.hidden = false;
        }
      };
      input.addEventListener('click', (event) => {
        clearFeedback();
        if (input.disabled) return;
        if (office?.context?.host === 'Word' && office.context.platform === 'PC'
            && typeof office.context.ui?.displayDialogAsync === 'function') {
          // Desktop Word can silently drop the embedded pane's native chooser.
          // Its separate Office dialog has a working native file-input route.
          event.preventDefault();
          try { openWordFileDialog(input, document, office, showError); }
          catch (error) { showError(error); }
          return;
        }
        if (typeof input.showPicker !== 'function') return;
        try {
          // Keep this synchronous in the actual input click. Awaiting Office
          // readiness or other work first can lose transient user activation.
          input.showPicker();
          event.preventDefault(); // The picker is already open; avoid a second.
        } catch (error) {
          // Preserve the native input default as the fallback. A programmatic
          // input.click() here would recurse through this same handler.
          if (feedback) {
            feedback.textContent = `File chooser could not be opened (${error.name || 'Error'}). `
              + 'Try the native Choose file control again, or use Open local tools in browser.';
            feedback.hidden = false;
          }
        }
      });
      input.addEventListener('change', clearFeedback);
      input.addEventListener('cancel', clearFeedback);
    });
  }

  return { bindFilePickers, createFileReceiver };
}));
