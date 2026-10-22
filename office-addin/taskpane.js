/* BibWord 0.1 prototype: intentionally dependency-free UI and Office.js integration.
 * The .bib scanner extracts keys, NOT full BibLaTeX metadata; it is preliminary.
 */
(() => {
  'use strict';
  const $ = (id) => document.getElementById(id);
  const { citationSyntax, parseKeys } = globalThis.BibWordCitation;
  let knownKeys = null;
  let officeReady = false;

  function enteredKeys() {
    return $('citation-keys').value.split(';').map((x) => x.trim()).filter(Boolean);
  }

  function setStatus(msg, error = false) {
    $('status').textContent = msg;
    $('status').className = error ? 'error' : '';
  }

  function updatePreview() {
    try {
      const keys = enteredKeys();
      $('citation-preview').textContent = citationSyntax(keys.length ? keys : ['key'], $('mode').value, $('locator').value);
    } catch (e) {
      $('citation-preview').textContent = e.message;
    }
  }

  async function loadBibFile(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const contents = await file.text();
      knownKeys = parseKeys(contents);
      const options = $('known-keys');
      options.replaceChildren();
      [...knownKeys].sort().forEach((key) => {
        const option = document.createElement('option');
        option.value = key;
        options.appendChild(option);
      });
      $('library-status').textContent = `${file.name}: ${knownKeys.size} keys discovered (prototype parser).`;
      setStatus('Bibliography opened locally; no file data uploaded.');
    } catch (error) {
      setStatus(error.message, true);
    }
  }

  async function insertIntoWord() {
    if (!officeReady) return setStatus('Office.js is not ready. Open this pane inside Word.', true);
    let syntax;
    try {
      const keys = enteredKeys();
      if (knownKeys) {
        const missing = keys.filter((key) => !knownKeys.has(key));
        if (missing.length) throw new Error('Keys missing in current library: ' + missing.join(', '));
      }
      syntax = citationSyntax(keys, $('mode').value, $('locator').value);
    } catch (error) {
      return setStatus(error.message, true);
    }
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        const inserted = selection.insertText(syntax, Word.InsertLocation.replace);
        const control = inserted.insertContentControl();
        control.tag = 'bibword:cite:v1:' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 8);
        control.title = 'BibWord citation';
        inserted.font.name = 'Consolas';
        inserted.font.color = '#165789';
        await context.sync();
      });
      setStatus('Citation inserted with editable Quarto syntax.');
    } catch (error) {
      setStatus('Word insertion failed: ' + (error.message || String(error)), true);
    }
  }

  $('bib-file').addEventListener('change', loadBibFile);
  ['citation-keys', 'locator'].forEach((id) => $(id).addEventListener('input', updatePreview));
  $('mode').addEventListener('change', updatePreview);
  $('insert').addEventListener('click', insertIntoWord);
  updatePreview();
  if (typeof Office !== 'undefined') {
    Office.onReady((info) => {
      if (info.host !== Office.HostType.Word) {
        setStatus('This prototype requires Microsoft Word.', true);
        return;
      }
      if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
        setStatus('This Word host does not support the required WordApi 1.1 APIs.', true);
        return;
      }
      officeReady = true;
      $('insert').disabled = false;
      setStatus('Ready to insert citations.');
    });
  } else {
    setStatus('Office.js unavailable. Load the add-in within Microsoft Word.', true);
  }
})();
