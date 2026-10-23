/* BibWord 0.1 prototype: local CSL previews plus Office.js integration. */
(() => {
  'use strict';
  const $ = (id) => document.getElementById(id);
  const {
    bibliographyTag, CITATION_TITLE, citationKeysFromControls, citationSourceFromTitle,
    citationSourceTitle, citationSyntax, citationSyntaxFromControl, parseBibliographyTag,
    parseCitationSyntax, replaceActiveKey, suggestKeys,
  } = globalThis.BibWordCitation;
  const {
    createLibrary, CSL_STYLES, customStyleIdentity, registerCslStyle, validateCslStyle,
  } = globalThis.BibWordCsl;
  const styleIds = new Set(CSL_STYLES.map(({ id }) => id));
  const customStyleNamespace = 'urn:bibword:csl-style:v1';
  const customStyleIdPattern = /^custom-[0-9a-f]{12}$/;
  const customStyles = new Map();
  let knownKeys = null;
  let cslLibrary = null;
  let officeReady = false;
  let activeSuggestion = -1;
  let selectedControlId = null;
  let documentCitationKeys = [];

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

  function selectedStyle() {
    return $('citation-style').value;
  }

  function escapeHtml(value) {
    return String(value).replace(/[&<>"']/g, (character) => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
    })[character]);
  }

  function wordBibliographyHtml(title, bibliography) {
    const styled = bibliography
      .replace('class="csl-bib-body"', 'class="csl-bib-body" style="margin:0;"')
      .replace(/class="csl-entry"/g,
        'class="csl-entry" style="margin-left:0.5in;text-indent:-0.5in;margin-bottom:0.12in;"');
    return `<h1>${escapeHtml(title)}</h1>${styled}`;
  }

  function markWordBibliographyStale(message = 'The Word bibliography may be out of date. Choose Update in Word.') {
    $('bibliography-render-status').textContent = message;
  }

  function markFormattedCitationsStale(message = 'Formatted citations may be out of date. Choose Show formatted citations to refresh them.') {
    $('citation-display-status').textContent = message;
  }

  function applyKeyFormatting(range) {
    // Keep this to WordApi 1.1 setters. Word for the web returned ApiNotFound
    // for Font.reset after applying only the first citation in a batch.
    range.font.name = 'Consolas';
    range.font.color = '#165789';
  }

  function bytesToBase64(value) {
    const bytes = new TextEncoder().encode(value);
    let binary = '';
    for (let offset = 0; offset < bytes.length; offset += 0x8000) {
      binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
    }
    return btoa(binary);
  }

  function base64ToText(value) {
    const binary = atob(value);
    const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
    return new TextDecoder().decode(bytes);
  }

  function customStyleStoreXml(style) {
    return `<bibwordStyle xmlns="${customStyleNamespace}" id="${style.id}" hash="${style.hash}">`
      + `<label encoding="base64">${bytesToBase64(style.label)}</label>`
      + `<csl encoding="base64">${bytesToBase64(style.xml)}</csl>`
      + '</bibwordStyle>';
  }

  function parseCustomStyleStore(xml) {
    const parsed = new DOMParser().parseFromString(xml, 'application/xml');
    if (parsed.querySelector('parsererror')) throw new Error('The embedded BibWord CSL part is invalid XML.');
    const root = parsed.documentElement;
    if (root.localName !== 'bibwordStyle' || root.namespaceURI !== customStyleNamespace) return null;
    const id = root.getAttribute('id') || '';
    const hash = root.getAttribute('hash') || '';
    const labelNode = root.getElementsByTagNameNS(customStyleNamespace, 'label')[0];
    const cslNode = root.getElementsByTagNameNS(customStyleNamespace, 'csl')[0];
    if (!customStyleIdPattern.test(id) || !/^[0-9a-f]{64}$/.test(hash) || !labelNode || !cslNode) {
      throw new Error('The embedded BibWord CSL part has invalid metadata.');
    }
    return {
      id, hash,
      label: base64ToText(labelNode.textContent || ''),
      xml: base64ToText(cslNode.textContent || ''),
    };
  }

  function assertWellFormedCsl(xml) {
    const parsed = new DOMParser().parseFromString(xml, 'application/xml');
    if (parsed.querySelector('parsererror')) throw new Error('The selected CSL style is not well-formed XML.');
    if (parsed.documentElement.localName !== 'style'
        || parsed.documentElement.namespaceURI !== 'http://purl.org/net/xbiblio/csl') {
      throw new Error('The selected file is not a CSL style document.');
    }
  }

  function addCustomStyleOption(style) {
    let option = [...$('citation-style').options].find(({ value }) => value === style.id);
    if (!option) {
      option = document.createElement('option');
      option.value = style.id;
      $('citation-style').appendChild(option);
    }
    option.textContent = `Custom: ${style.label}`;
    styleIds.add(style.id);
  }

  function uniqueId() {
    return Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 8);
  }

  function replaceListContents(list, messages) {
    list.replaceChildren();
    messages.forEach((message) => {
      const item = document.createElement('li');
      item.textContent = message;
      list.appendChild(item);
    });
  }

  function renderFormattedPreviews() {
    const keys = enteredKeys();
    const references = $('reference-preview');
    if (!cslLibrary) {
      $('formatted-preview').textContent = 'Load a bibliography to preview formatting.';
      replaceListContents(references, ['Load a bibliography and choose a citation key.']);
      return;
    }
    if (!keys.length) {
      $('formatted-preview').textContent = 'Choose at least one citation key.';
      replaceListContents(references, ['No citation selected.']);
      return;
    }
    try {
      $('formatted-preview').textContent = cslLibrary.formatCitation(
        keys, $('mode').value, $('locator').value, selectedStyle(),
      );
      replaceListContents(references,
        cslLibrary.formatBibliography(keys, selectedStyle()).map(({ text }) => text));
    } catch (error) {
      $('formatted-preview').textContent = error.message;
      replaceListContents(references, [error.message]);
    }
  }

  function renderDocumentBibliography() {
    const preview = $('document-bibliography-preview');
    if (!cslLibrary) {
      replaceListContents(preview, ['Load a bibliography first.']);
      return;
    }
    if (!documentCitationKeys.length) {
      replaceListContents(preview, ['No managed citations found in the document.']);
      return;
    }
    try {
      const entries = cslLibrary.formatBibliography(documentCitationKeys, selectedStyle());
      replaceListContents(preview, entries.map(({ text }) => text));
      $('document-preview-status').textContent = `${entries.length} unique references formatted with CSL.`;
    } catch (error) {
      replaceListContents(preview, [error.message]);
      $('document-preview-status').textContent = error.message;
    }
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
      cslLibrary = createLibrary(contents);
      knownKeys = new Set(cslLibrary.keys);
      $('library-status').textContent = `${file.name}: ${cslLibrary.size} entries parsed as BibLaTeX/CSL.`;
      $('refresh-bibliography').disabled = !officeReady;
      $('format-bibliography').disabled = !officeReady;
      $('show-formatted-citations').disabled = !officeReady;
      setStatus('Bibliography parsed locally with Citation.js; no file data uploaded.');
      markWordBibliographyStale();
      markFormattedCitationsStale();
      renderSuggestions();
      updatePreview();
      if (documentCitationKeys.length) renderDocumentBibliography();
    } catch (error) {
      cslLibrary = null;
      knownKeys = null;
      $('refresh-bibliography').disabled = true;
      $('format-bibliography').disabled = true;
      $('show-formatted-citations').disabled = true;
      setStatus(error.message, true);
    }
  }

  function supportsCustomStylePersistence() {
    return typeof Office !== 'undefined'
      && Office.context.requirements.isSetSupported('WordApi', '1.4');
  }

  async function readCustomStylesFromDocument(context) {
    if (!supportsCustomStylePersistence()) {
      throw new Error('Embedding custom CSL styles requires WordApi 1.4 or newer.');
    }
    const collection = context.document.customXmlParts.getByNamespace(customStyleNamespace);
    collection.load('items/id');
    await context.sync();
    const blobs = collection.items.map((part) => part.getXml());
    await context.sync();
    const styles = [];
    for (const blob of blobs) {
      const style = parseCustomStyleStore(blob.value);
      if (!style) continue;
      assertWellFormedCsl(style.xml);
      const identity = await customStyleIdentity(style.xml);
      if (identity.id !== style.id || identity.hash !== style.hash) {
        throw new Error(`Embedded CSL style ${style.id} failed its integrity check.`);
      }
      styles.push(style);
    }
    return styles;
  }

  async function ensureCustomStylePart(context, style) {
    const storedStyles = await readCustomStylesFromDocument(context);
    const existing = storedStyles.find(({ id }) => id === style.id);
    if (existing) {
      if (existing.hash !== style.hash || existing.xml !== style.xml) {
        throw new Error(`The document already contains different CSL data for ${style.id}.`);
      }
      return;
    }
    context.document.customXmlParts.add(customStyleStoreXml(style));
    await context.sync();
  }

  async function loadCustomStyleFromDocument(id) {
    let style = null;
    await Word.run(async (context) => {
      const matches = (await readCustomStylesFromDocument(context)).filter((item) => item.id === id);
      if (!matches.length) throw new Error(`The embedded CSL style ${id} is missing.`);
      if (matches.length > 1 && matches.some((item) => item.xml !== matches[0].xml)) {
        throw new Error(`The document contains conflicting copies of CSL style ${id}.`);
      }
      [style] = matches;
    });
    registerCslStyle(style.id, style.xml);
    customStyles.set(style.id, style);
    addCustomStyleOption(style);
    return style;
  }

  async function loadCslFile(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const xml = await file.text();
      validateCslStyle(xml);
      assertWellFormedCsl(xml);
      const identity = await customStyleIdentity(xml);
      const style = { ...identity, xml, label: file.name };
      registerCslStyle(style.id, style.xml);
      customStyles.set(style.id, style);
      addCustomStyleOption(style);
      $('citation-style').value = style.id;
      $('style-status').textContent = `${file.name} loaded as ${style.id}.`;
      updatePreview();
      renderDocumentBibliography();
      markWordBibliographyStale();
      markFormattedCitationsStale();
      if (officeReady && await persistSelectedStyle()) {
        setStatus('Custom CSL style embedded and saved with the bibliography marker.');
      } else {
        setStatus('Custom CSL style loaded locally. Insert a bibliography marker to embed it in the document.');
      }
    } catch (error) {
      $('style-status').textContent = error.message || String(error);
      setStatus('Could not load the CSL style: ' + (error.message || String(error)), true);
    }
  }

  async function refreshDocumentBibliography() {
    if (!officeReady) return setStatus('Office.js is not ready. Open this pane inside Word.', true);
    if (!cslLibrary) return setStatus('Load a bibliography before refreshing the document preview.', true);
    try {
      let controls;
      await Word.run(async (context) => {
        const collection = context.document.contentControls;
        collection.load('items/tag,text,title');
        await context.sync();
        controls = collection.items.map(({ tag, text, title }) => ({ tag, text, title }));
      });
      documentCitationKeys = citationKeysFromControls(controls);
      renderDocumentBibliography();
      setStatus(`Document bibliography preview refreshed from ${documentCitationKeys.length} unique keys.`);
    } catch (error) {
      $('document-preview-status').textContent = error.message || String(error);
      setStatus('Could not refresh document bibliography: ' + (error.message || String(error)), true);
    }
  }

  async function persistSelectedStyle() {
    let saved = false;
    await Word.run(async (context) => {
      const controls = context.document.contentControls;
      controls.load('items/tag');
      await context.sync();
      const markers = controls.items.filter(({ tag }) => tag.startsWith('bibword:refs:'));
      if (markers.length > 1) throw new Error('The document contains multiple bibliography markers.');
      if (!markers.length) return;
      const metadata = parseBibliographyTag(markers[0].tag);
      if (!metadata) return;
      const styleId = selectedStyle();
      if (customStyleIdPattern.test(styleId)) {
        const style = customStyles.get(styleId);
        if (!style) throw new Error(`Custom CSL style ${styleId} is not loaded.`);
        await ensureCustomStylePart(context, style);
      }
      markers[0].tag = bibliographyTag(metadata.id, styleId);
      await context.sync();
      saved = true;
    });
    return saved;
  }

  async function restoreDocumentStyle() {
    try {
      let markerTag = null;
      await Word.run(async (context) => {
        const controls = context.document.contentControls;
        controls.load('items/tag');
        await context.sync();
        const markers = controls.items.filter(({ tag }) => tag.startsWith('bibword:refs:'));
        if (markers.length > 1) throw new Error('The document contains multiple bibliography markers.');
        markerTag = markers[0]?.tag || null;
      });
      if (!markerTag) return;
      const metadata = parseBibliographyTag(markerTag);
      if (metadata?.style) {
        if (customStyleIdPattern.test(metadata.style) && !styleIds.has(metadata.style)) {
          const style = await loadCustomStyleFromDocument(metadata.style);
          $('style-status').textContent = `${style.label} restored from this Word document.`;
        }
        if (!styleIds.has(metadata.style)) {
          throw new Error(`The document uses an unsupported citation style: ${metadata.style}`);
        }
        $('citation-style').value = metadata.style;
        updatePreview();
        setStatus('Citation style restored from the document bibliography marker.');
      }
    } catch (error) {
      setStatus('Could not restore the document citation style: ' + (error.message || String(error)), true);
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
        control.tag = 'bibword:cite:v1:' + uniqueId();
        control.title = CITATION_TITLE;
        applyKeyFormatting(inserted);
        await context.sync();
      });
      setStatus('Citation inserted with editable Quarto syntax.');
      markWordBibliographyStale();
      markFormattedCitationsStale('A new key-view citation was inserted. Choose Show formatted citations to refresh the reading view.');
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
        control.load('id,tag,text,title');
        await context.sync();
        selected = { id: control.id, tag: control.tag, text: control.text, title: control.title };
      });
      if (!selected.tag.startsWith('bibword:cite:v1:')) {
        throw new Error('The selected content control is not a BibWord citation.');
      }
      const source = citationSyntaxFromControl(selected);
      const parsed = parseCitationSyntax(source);
      $('citation-keys').value = parsed.keys.join('; ');
      $('mode').value = parsed.mode;
      $('locator').value = parsed.locator;
      selectedControlId = selected.id;
      $('update-selected').disabled = false;
      $('editing-status').textContent = `Editing ${source}`;
      updatePreview();
      setStatus('Selected citation loaded. Change its fields, then choose Update selected.');
    } catch (error) {
      selectedControlId = null;
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
    if (selectedControlId === null) return setStatus('Load a selected BibWord citation first.', true);
    let syntax;
    try {
      syntax = validatedSyntax();
    } catch (error) {
      return setStatus(error.message, true);
    }
    try {
      await Word.run(async (context) => {
        const control = context.document.contentControls.getById(selectedControlId);
        control.load('tag,title');
        await context.sync();
        if (!control.tag.startsWith('bibword:cite:v1:')) {
          throw new Error('The selected citation no longer exists.');
        }
        const formatted = citationSourceFromTitle(control.title) !== null;
        if (formatted && !cslLibrary) {
          throw new Error('Reload the bibliography before updating a formatted citation.');
        }
        const display = formatted
          ? cslLibrary.formatCitation(enteredKeys(), $('mode').value, $('locator').value, selectedStyle())
          : syntax;
        const updated = control.insertText(display, Word.InsertLocation.replace);
        control.title = formatted ? citationSourceTitle(syntax) : CITATION_TITLE;
        // Replacing text in formatted view already restores inherited run
        // formatting in Word for the web; do not call the unavailable Font.reset.
        if (!formatted) applyKeyFormatting(updated);
        await context.sync();
      });
      $('editing-status').textContent = `Editing ${syntax}`;
      setStatus('Selected citation updated without changing its BibWord identity.');
      markWordBibliographyStale();
    } catch (error) {
      setStatus('Word update failed: ' + (error.message || String(error)), true);
    }
  }

  async function setCitationDisplay(mode) {
    if (!officeReady) return setStatus('Office.js is not ready. Open this pane inside Word.', true);
    if (mode === 'formatted' && !cslLibrary) {
      return setStatus('Load the bibliography before showing formatted citations.', true);
    }
    try {
      let changed = 0;
      await Word.run(async (context) => {
        const collection = context.document.contentControls;
        collection.load('items/tag,text,title');
        await context.sync();
        const controls = collection.items.filter(({ tag }) => tag.startsWith('bibword:cite:v1:'));
        if (!controls.length) throw new Error('The document has no managed citations.');
        const replacements = controls.map((control) => {
          const source = citationSyntaxFromControl(control);
          const parsed = parseCitationSyntax(source);
          const text = mode === 'formatted'
            ? cslLibrary.formatCitation(parsed.keys, parsed.mode, parsed.locator, selectedStyle())
            : source;
          return { control, source, text };
        });
        replacements.forEach(({ control, source, text }) => {
          const updated = control.insertText(text, Word.InsertLocation.replace);
          if (mode === 'formatted') {
            control.title = citationSourceTitle(source);
          } else {
            control.title = CITATION_TITLE;
            applyKeyFormatting(updated);
          }
        });
        await context.sync();
        documentCitationKeys = citationKeysFromControls(replacements.map(({ control, source }) => ({
          tag: control.tag, text: source, title: CITATION_TITLE,
        })));
        changed = replacements.length;
      });
      const label = mode === 'formatted' ? 'formatted citations' : 'editable citation keys';
      $('citation-display-status').textContent = `${changed} managed citations now show ${label}.`;
      setStatus(`Document now shows ${label}; canonical Quarto syntax remains exportable.`);
    } catch (error) {
      $('citation-display-status').textContent = error.message || String(error);
      setStatus('Could not change citation display: ' + (error.message || String(error)), true);
    }
  }

  async function insertBibliographyMarker() {
    if (!officeReady) return setStatus('Office.js is not ready. Open this pane inside Word.', true);
    try {
      await Word.run(async (context) => {
        const styleId = selectedStyle();
        if (customStyleIdPattern.test(styleId)) {
          const style = customStyles.get(styleId);
          if (!style) throw new Error(`Custom CSL style ${styleId} is not loaded.`);
          await ensureCustomStylePart(context, style);
        }
        const controls = context.document.contentControls;
        controls.load('items/tag');
        await context.sync();
        if (controls.items.some((control) => control.tag.startsWith('bibword:refs:'))) {
          throw new Error('This document already contains a BibWord bibliography marker.');
        }
        const selection = context.document.getSelection();
        const paragraph = selection.insertParagraph('References', Word.InsertLocation.after);
        paragraph.styleBuiltIn = 'Heading1';
        const control = paragraph.insertContentControl();
        control.tag = bibliographyTag(uniqueId(), styleId);
        control.title = 'BibWord bibliography position';
        await context.sync();
      });
      setStatus('Bibliography marker inserted with the selected CSL style.');
      markWordBibliographyStale('Bibliography marker inserted. Choose Update in Word to generate references.');
    } catch (error) {
      setStatus('Bibliography marker failed: ' + (error.message || String(error)), true);
    }
  }

  async function formatBibliographyInWord() {
    if (!officeReady) return setStatus('Office.js is not ready. Open this pane inside Word.', true);
    if (!cslLibrary) return setStatus('Load a bibliography before formatting it in Word.', true);
    try {
      let referenceCount = 0;
      await Word.run(async (context) => {
        const controls = context.document.contentControls;
        controls.load('items/tag,text,title');
        await context.sync();
        const markers = controls.items.filter(({ tag }) => tag.startsWith('bibword:refs:'));
        if (markers.length !== 1) {
          throw new Error(markers.length
            ? 'The document contains multiple bibliography markers.'
            : 'Insert a bibliography marker before updating the Word bibliography.');
        }
        const keys = citationKeysFromControls(
          controls.items.map(({ tag, text, title }) => ({ tag, text, title })),
        );
        if (!keys.length) throw new Error('The document has no managed citations to format.');
        const marker = markers[0];
        const metadata = parseBibliographyTag(marker.tag);
        const styleId = selectedStyle();
        if (customStyleIdPattern.test(styleId)) {
          const style = customStyles.get(styleId);
          if (!style) throw new Error(`Custom CSL style ${styleId} is not loaded.`);
          await ensureCustomStylePart(context, style);
        }
        const title = marker.text.split(/[\r\n]+/)[0].trim() || 'References';
        const bibliography = cslLibrary.formatBibliographyHtml(keys, styleId);
        marker.tag = bibliographyTag(metadata.id, styleId);
        marker.title = 'BibWord bibliography position and generated references';
        marker.insertHtml(wordBibliographyHtml(title, bibliography), Word.InsertLocation.replace);
        await context.sync();
        documentCitationKeys = keys;
        referenceCount = keys.length;
      });
      renderDocumentBibliography();
      const styleLabel = $('citation-style').selectedOptions[0].textContent;
      $('bibliography-render-status').textContent = `${referenceCount} references formatted in Word with ${styleLabel}.`;
      setStatus('Word bibliography updated. Export will rebuild it from the citation keys.');
    } catch (error) {
      $('bibliography-render-status').textContent = error.message || String(error);
      setStatus('Could not update the Word bibliography: ' + (error.message || String(error)), true);
    }
  }

  $('bib-file').addEventListener('change', loadBibFile);
  $('csl-file').addEventListener('change', loadCslFile);
  $('citation-keys').addEventListener('input', () => {
    updatePreview();
    renderSuggestions();
  });
  $('citation-keys').addEventListener('focus', renderSuggestions);
  $('citation-keys').addEventListener('keydown', handleKeyPickerKeyboard);
  $('locator').addEventListener('input', updatePreview);
  $('mode').addEventListener('change', updatePreview);
  $('citation-style').addEventListener('change', async () => {
    updatePreview();
    renderDocumentBibliography();
    markWordBibliographyStale();
    markFormattedCitationsStale();
    if (!officeReady) return;
    try {
      if (await persistSelectedStyle()) {
        setStatus('Citation style saved with the document bibliography marker.');
      }
    } catch (error) {
      setStatus('Could not save the citation style: ' + (error.message || String(error)), true);
    }
  });
  $('insert').addEventListener('click', insertIntoWord);
  $('load-selected').addEventListener('click', loadSelectedCitation);
  $('update-selected').addEventListener('click', updateSelectedCitation);
  $('show-formatted-citations').addEventListener('click', () => setCitationDisplay('formatted'));
  $('show-citation-keys').addEventListener('click', () => setCitationDisplay('keys'));
  $('insert-bibliography').addEventListener('click', insertBibliographyMarker);
  $('format-bibliography').addEventListener('click', formatBibliographyInWord);
  $('refresh-bibliography').addEventListener('click', refreshDocumentBibliography);
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
      $('show-citation-keys').disabled = false;
      $('show-formatted-citations').disabled = !cslLibrary;
      $('refresh-bibliography').disabled = !cslLibrary;
      $('format-bibliography').disabled = !cslLibrary;
      setStatus('Ready to insert citations.');
      restoreDocumentStyle();
    });
  } else {
    setStatus('Office.js unavailable. Load the add-in within Microsoft Word.', true);
  }
})();
