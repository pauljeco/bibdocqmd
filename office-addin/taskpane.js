/* BibWord 0.1 prototype: intentionally dependency-free UI and Office.js integration.
 * The .bib scanner extracts keys, NOT full BibLaTeX metadata; it is preliminary.
 */
(() => {
  'use strict';
  const $ = (id) => document.getElementById(id);
  const {
    citationSyntax, formatCitationPreview, formatReference, parseBibEntries,
    parseCitationSyntax, replaceActiveKey, suggestKeys,
  } = globalThis.BibWordCitation;
  let knownKeys = null;
  let bibliography = new Map();
  let officeReady = false;
  let activeSuggestion = -1;
  let selectedControlTag = null;

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
    renderFormattedPreviews();
  }

  function renderFormattedPreviews() {
    const keys = enteredKeys();
    const references = $('reference-preview');
    references.replaceChildren();
    if (!bibliography.size) {
      $('formatted-preview').textContent = 'Load a bibliography to preview formatting.';
      const item = document.createElement('li');
      item.textContent = 'Load a bibliography and choose a citation key.';
      references.appendChild(item);
      return;
    }
    if (!keys.length) {
      $('formatted-preview').textContent = 'Choose at least one citation key.';
      const item = document.createElement('li');
      item.textContent = 'No citation selected.';
      references.appendChild(item);
      return;
    }
    try {
      $('formatted-preview').textContent = formatCitationPreview(
        keys, $('mode').value, $('locator').value, bibliography,
      );
    } catch (error) {
      $('formatted-preview').textContent = error.message;
    }
    keys.forEach((key) => {
      const item = document.createElement('li');
      item.textContent = bibliography.has(key)
        ? formatReference(bibliography.get(key))
        : `${key}: not found in the current library.`;
      references.appendChild(item);
    });
  }

  function hideSuggestions() {
    const suggestions = $('key-suggestions');
    suggestions.hidden = true;
    suggestions.replaceChildren();
    activeSuggestion = -1;
    $('citation-keys').setAttribute('aria-expanded', 'false');
    $('citation-keys').removeAttribute('aria-activedescendant');
  }

  function chooseSuggestion(key) {
    const input = $('citation-keys');
    input.value = replaceActiveKey(input.value, key);
    hideSuggestions();
    updatePreview();
    input.focus();
  }

  function highlightSuggestion(index) {
    const options = [...$('key-suggestions').querySelectorAll('[role="option"]')];
    if (!options.length) return;
    activeSuggestion = (index + options.length) % options.length;
    options.forEach((option, position) => {
      option.setAttribute('aria-selected', position === activeSuggestion ? 'true' : 'false');
    });
    const selected = options[activeSuggestion];
    $('citation-keys').setAttribute('aria-activedescendant', selected.id);
    selected.scrollIntoView({ block: 'nearest' });
  }

  function renderSuggestions() {
    if (!knownKeys) return hideSuggestions();
    const matches = suggestKeys(knownKeys, $('citation-keys').value);
    if (!matches.length) return hideSuggestions();
    const suggestions = $('key-suggestions');
    suggestions.replaceChildren();
    matches.forEach((key, index) => {
      const option = document.createElement('button');
      option.type = 'button';
      option.id = `key-suggestion-${index}`;
      option.className = 'key-suggestion';
      option.setAttribute('role', 'option');
      option.setAttribute('aria-selected', 'false');
      option.textContent = key;
      option.addEventListener('click', () => chooseSuggestion(key));
      suggestions.appendChild(option);
    });
    suggestions.hidden = false;
    activeSuggestion = -1;
    $('citation-keys').setAttribute('aria-expanded', 'true');
  }

  function handleKeyPickerKeyboard(event) {
    const suggestions = $('key-suggestions');
    const options = [...suggestions.querySelectorAll('[role="option"]')];
    if (event.key === 'Escape') return hideSuggestions();
    if (suggestions.hidden || !options.length) return;
    if (event.key === 'ArrowDown') {
      event.preventDefault();
      highlightSuggestion(activeSuggestion + 1);
    } else if (event.key === 'ArrowUp') {
      event.preventDefault();
      highlightSuggestion(activeSuggestion - 1);
    } else if (event.key === 'Enter' && activeSuggestion >= 0) {
      event.preventDefault();
      chooseSuggestion(options[activeSuggestion].textContent);
    }
  }

  async function loadBibFile(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const contents = await file.text();
      bibliography = parseBibEntries(contents);
      knownKeys = new Set(bibliography.keys());
      $('library-status').textContent = `${file.name}: ${knownKeys.size} keys discovered (prototype parser).`;
      setStatus('Bibliography opened locally; no file data uploaded.');
      renderSuggestions();
      updatePreview();
    } catch (error) {
      setStatus(error.message, true);
    }
  }

  async function insertIntoWord() {
    if (!officeReady) return setStatus('Office.js is not ready. Open this pane inside Word.', true);
    let syntax;
    try {
      syntax = validatedSyntax();
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

  async function loadSelectedCitation() {
    if (!officeReady) return setStatus('Office.js is not ready. Open this pane inside Word.', true);
    try {
      let selected;
      await Word.run(async (context) => {
        const control = context.document.getSelection().parentContentControl;
        control.load('tag,text');
        await context.sync();
        selected = { tag: control.tag, text: control.text };
      });
      if (!selected.tag.startsWith('bibword:cite:v1:')) {
        throw new Error('The selected content control is not a BibWord citation.');
      }
      const parsed = parseCitationSyntax(selected.text);
      $('citation-keys').value = parsed.keys.join('; ');
      $('mode').value = parsed.mode;
      $('locator').value = parsed.locator;
      selectedControlTag = selected.tag;
      $('update-selected').disabled = false;
      $('editing-status').textContent = `Editing ${selected.text}`;
      updatePreview();
      setStatus('Selected citation loaded. Change its fields, then choose Update selected.');
    } catch (error) {
      selectedControlTag = null;
      $('update-selected').disabled = true;
      $('editing-status').textContent = 'To edit, place the cursor inside a BibWord citation and load it.';
      setStatus('Could not load a BibWord citation. Place the cursor inside its blue citation text and try again.', true);
    }
  }

  function validatedSyntax() {
    const keys = enteredKeys();
    if (knownKeys) {
      const missing = keys.filter((key) => !knownKeys.has(key));
      if (missing.length) throw new Error('Keys missing in current library: ' + missing.join(', '));
    }
    return citationSyntax(keys, $('mode').value, $('locator').value);
  }

  async function updateSelectedCitation() {
    if (!selectedControlTag) return setStatus('Load a selected BibWord citation first.', true);
    let syntax;
    try {
      syntax = validatedSyntax();
    } catch (error) {
      return setStatus(error.message, true);
    }
    try {
      await Word.run(async (context) => {
        const matches = context.document.contentControls.getByTag(selectedControlTag);
        matches.load('items');
        await context.sync();
        if (!matches.items.length) throw new Error('The selected citation no longer exists.');
        const updated = matches.items[0].insertText(syntax, Word.InsertLocation.replace);
        updated.font.name = 'Consolas';
        updated.font.color = '#165789';
        await context.sync();
      });
      $('editing-status').textContent = `Editing ${syntax}`;
      setStatus('Selected citation updated without changing its BibWord identity.');
    } catch (error) {
      setStatus('Word update failed: ' + (error.message || String(error)), true);
    }
  }

  async function insertBibliographyMarker() {
    if (!officeReady) return setStatus('Office.js is not ready. Open this pane inside Word.', true);
    try {
      await Word.run(async (context) => {
        const controls = context.document.contentControls;
        controls.load('items/tag');
        await context.sync();
        if (controls.items.some((control) => control.tag.startsWith('bibword:refs:v1:'))) {
          throw new Error('This document already contains a BibWord bibliography marker.');
        }
        const selection = context.document.getSelection();
        const paragraph = selection.insertParagraph('References', Word.InsertLocation.after);
        paragraph.styleBuiltIn = 'Heading1';
        const control = paragraph.insertContentControl();
        control.tag = 'bibword:refs:v1:' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 8);
        control.title = 'BibWord bibliography position';
        await context.sync();
      });
      setStatus('Bibliography marker inserted. Quarto will place formatted references here.');
    } catch (error) {
      setStatus('Bibliography marker failed: ' + (error.message || String(error)), true);
    }
  }

  $('bib-file').addEventListener('change', loadBibFile);
  $('citation-keys').addEventListener('input', () => {
    updatePreview();
    renderSuggestions();
  });
  $('citation-keys').addEventListener('focus', renderSuggestions);
  $('citation-keys').addEventListener('keydown', handleKeyPickerKeyboard);
  $('locator').addEventListener('input', updatePreview);
  $('mode').addEventListener('change', updatePreview);
  $('insert').addEventListener('click', insertIntoWord);
  $('load-selected').addEventListener('click', loadSelectedCitation);
  $('update-selected').addEventListener('click', updateSelectedCitation);
  $('insert-bibliography').addEventListener('click', insertBibliographyMarker);
  document.addEventListener('click', (event) => {
    if (!$('citation-keys').closest('.key-picker').contains(event.target)) hideSuggestions();
  });
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
      $('load-selected').disabled = false;
      $('insert-bibliography').disabled = false;
      setStatus('Ready to insert citations.');
    });
  } else {
    setStatus('Office.js unavailable. Load the add-in within Microsoft Word.', true);
  }
})();
