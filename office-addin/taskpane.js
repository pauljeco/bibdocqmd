/* bibdocqmd 0.1 prototype: local CSL previews plus Office.js integration. */
(() => {
  'use strict';
  const $ = (id) => document.getElementById(id);
  const {
    bibliographyTag, CITATION_TITLE, citationKeysForBibliographyScope,
    citationKeysFromControls, citationSourceFromTitle,
    citationSourceTitle, citationSyntax, parseBibliographyTag, parseCitationControl,
    replaceActiveKey, suggestKeys,
  } = globalThis.BibWordCitation;
  const {
    createLibraryFromSources, CSL_STYLES, customStyleIdentity, registerCslStyle, sha256Hex,
    validateCslStyle,
  } = globalThis.BibWordCsl;
  const { fetchGithubBibliography, githubBibliographyUrl } = globalThis.BibWordGithubLibrary;
  const styleIds = new Set(CSL_STYLES.map(({ id }) => id));
  const customStyleNamespace = 'urn:bibword:csl-style:v1';
  const librarySourceNamespace = 'urn:bibword:library-sources:v1';
  const customStyleIdPattern = /^custom-[0-9a-f]{12}$/;
  const customStyles = new Map();
  let knownKeys = null;
  let cslLibrary = null;
  let officeReady = false;
  let activeSuggestion = -1;
  let selectedControlId = null;
  let documentCitationKeys = [];
  let rememberedLibrarySources = [];
  let activeLibrarySources = new Map();

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

  function wordBibliographyHtml(title, bibliography, headingLevel = 1) {
    // Flatten the known citeproc HTML div structure. A surrounding bibliography
    // div with indented entry divs can give Word's first entry an extra offset.
    // Each entry becomes a sibling paragraph; preserve inline CSL markup.
    const closingTags = [];
    const styled = bibliography.replace(/<\/?div\b[^>]*>/gi, (tag) => {
      if (/^<\//.test(tag)) return closingTags.pop();
      if (/\bclass="csl-bib-body"/.test(tag)) {
        closingTags.push('');
        return '';
      }
      if (/\bclass="csl-entry"/.test(tag)) {
        closingTags.push('</p>');
        return '<p class="csl-entry" style="margin:0;text-indent:0;margin-bottom:0.12in;">';
      }
      closingTags.push('</span> ');
      return tag.replace(/^<div\b/i, '<span');
    });
    return `<h${headingLevel}>${escapeHtml(title)}</h${headingLevel}>${styled}`;
  }

  function bibliographyIndentPoints() {
    const value = $('bibliography-indent').value.trim();
    const centimetres = Number(value);
    if (!value || !Number.isFinite(centimetres) || centimetres < 0 || centimetres > 5) {
      throw new Error('Enter a hanging indent between 0 and 5 cm.');
    }
    return centimetres * 72 / 2.54;
  }

  const wordNamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

  function wordValue(element) {
    return element?.getAttributeNS(wordNamespace, 'val')
      || element?.getAttribute('w:val') || element?.getAttribute('val') || '';
  }

  function firstDescendant(element, localName) {
    return element?.getElementsByTagNameNS(wordNamespace, localName)[0] || null;
  }

  function controlTag(control) {
    return wordValue(firstDescendant(firstDescendant(control, 'sdtPr'), 'tag'));
  }

  function controlSummary(control) {
    const properties = firstDescendant(control, 'sdtPr');
    const content = firstDescendant(control, 'sdtContent');
    const text = content
      ? [...content.getElementsByTagNameNS(wordNamespace, 't')]
        .map((node) => node.textContent || '').join('')
      : '';
    return {
      tag: controlTag(control),
      title: wordValue(firstDescendant(properties, 'alias')),
      text,
    };
  }

  function documentStructureFromOoxml(ooxml) {
    const parsed = new DOMParser().parseFromString(ooxml, 'application/xml');
    if (parsed.querySelector('parsererror')) throw new Error('Word returned invalid document XML.');
    const body = parsed.getElementsByTagNameNS(wordNamespace, 'body')[0];
    if (!body) throw new Error('Word document XML has no main document body.');
    const paragraphs = [];

    function addMarker(summary) {
      paragraphs.push({
        headingLevel: 0,
        citationKeys: [],
        marker: { ...summary, metadata: parseBibliographyTag(summary.tag) },
      });
    }

    function addParagraph(paragraph) {
      const style = wordValue(firstDescendant(firstDescendant(paragraph, 'pPr'), 'pStyle'));
      const heading = /^Heading([1-9])$/i.exec(style.replace(/\s+/g, ''));
      const citationKeys = [];
      const controls = [...paragraph.getElementsByTagNameNS(wordNamespace, 'sdt')];
      const summaries = controls.map(controlSummary);
      const markers = summaries.filter(({ tag }) => tag.startsWith('bibword:refs:'));
      if (markers.length) {
        if (markers.length > 1) {
          throw new Error('A paragraph contains more than one bibliography marker.');
        }
        // Word may serialize a live paragraph control as w:p/w:sdt, while a
        // downloaded DOCX may serialize the same control as w:sdt/w:p. Its
        // References heading must not become a boundary for its own scope.
        addMarker(markers[0]);
        return;
      }
      summaries.forEach((summary) => {
        if (summary.tag.startsWith('bibword:cite:v1:')) {
          citationKeys.push(...parseCitationControl(summary).parsed.keys);
        }
      });
      paragraphs.push({
        headingLevel: heading ? Number(heading[1]) : 0,
        citationKeys,
        marker: null,
      });
    }

    function walk(container) {
      [...container.children].forEach((child) => {
        if (child.namespaceURI !== wordNamespace) return;
        if (child.localName === 'p') {
          addParagraph(child);
          return;
        }
        if (child.localName === 'sdt') {
          const summary = controlSummary(child);
          if (summary.tag.startsWith('bibword:refs:')) {
            const content = firstDescendant(child, 'sdtContent');
            const titleParagraph = firstDescendant(content, 'p');
            summary.text = titleParagraph
              ? [...titleParagraph.getElementsByTagNameNS(wordNamespace, 't')]
                .map((node) => node.textContent || '').join('')
              : '';
            addMarker(summary);
            return;
          }
          const content = firstDescendant(child, 'sdtContent');
          if (content) walk(content);
          return;
        }
        walk(child);
      });
    }
    walk(body);
    return paragraphs;
  }

  function bibliographyPlansFromParagraphs(paragraphs) {
    const plans = new Map();
    paragraphs.forEach((paragraph, index) => {
      if (!paragraph.marker) return;
      const { metadata, text, tag } = paragraph.marker;
      if (plans.has(metadata.id)) {
        throw new Error(`Bibliography marker ${metadata.id} was copied; insert a new marker instead.`);
      }
      const keys = citationKeysForBibliographyScope(paragraphs, index, metadata.scope);
      const title = text.split(/[\r\n]+/)[0].trim() || 'References';
      plans.set(metadata.id, { keys, metadata, tag, title });
    });
    return plans;
  }

  function bibliographyPlansFromOoxml(ooxml) {
    return bibliographyPlansFromParagraphs(documentStructureFromOoxml(ooxml));
  }

  async function bibliographyPlansFromWordParagraphs(context) {
    const paragraphCollection = context.document.body.paragraphs;
    paragraphCollection.load('items/styleBuiltIn');
    await context.sync();

    const details = paragraphCollection.items.map((paragraph) => {
      const controls = paragraph.contentControls;
      const parent = paragraph.parentContentControlOrNullObject;
      controls.load('items/id,tag,text,title');
      parent.load('id,tag,text,title');
      return { paragraph, controls, parent };
    });
    await context.sync();

    const paragraphs = [];
    const consumedMarkerControlIds = new Set();
    const markerLastParagraphIndexes = new Map();
    details.forEach(({ paragraph, controls, parent }, detailIndex) => {
      const candidates = [...controls.items];
      if (!parent.isNullObject && !candidates.some(({ id }) => id === parent.id)) {
        candidates.push(parent);
      }
      const markers = candidates.filter(({ tag }) => tag.startsWith('bibword:refs:'));
      if (markers.length) {
        const identities = [...new Set(markers.map(({ id }) => id))];
        if (identities.length > 1) {
          throw new Error('A paragraph contains more than one bibliography marker.');
        }
        const marker = markers[0];
        const metadata = parseBibliographyTag(marker.tag);
        markerLastParagraphIndexes.set(metadata.id, detailIndex);
        if (!consumedMarkerControlIds.has(marker.id)) {
          consumedMarkerControlIds.add(marker.id);
          paragraphs.push({
            headingLevel: 0,
            citationKeys: [],
            marker: {
              tag: marker.tag,
              text: marker.text,
              title: marker.title,
              metadata,
            },
          });
        }
        // Generated bibliography paragraphs all have the marker as their
        // parent. None of them are part of the surrounding heading scope.
        return;
      }
      const citationKeys = [];
      controls.items.forEach((control) => {
        if (control.tag.startsWith('bibword:cite:v1:')) {
          citationKeys.push(...parseCitationControl(control).parsed.keys);
        }
      });
      const heading = /^Heading([1-9])$/i.exec(
        String(paragraph.styleBuiltIn || '').replace(/\s+/g, ''),
      );
      paragraphs.push({
        headingLevel: heading ? Number(heading[1]) : 0,
        citationKeys,
        marker: null,
      });
    });
    return {
      plans: bibliographyPlansFromParagraphs(paragraphs),
      markerHasFollowingParagraph: new Map(
        [...markerLastParagraphIndexes].map(([id, index]) => [id, index < details.length - 1]),
      ),
    };
  }

  function bibliographyHeadingLevel(scope) {
    return scope === 'document' ? 1 : Math.min(Number(scope.slice('heading'.length)) + 1, 9);
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

  function librarySourceStoreXml(sources) {
    const items = sources.map(({ name, hash }) => (
      `<source hash="${hash}"><name encoding="base64">${bytesToBase64(name)}</name></source>`
    )).join('');
    return `<bibwordLibraries xmlns="${librarySourceNamespace}" version="1">${items}</bibwordLibraries>`;
  }

  function parseLibrarySourceStore(xml) {
    const parsed = new DOMParser().parseFromString(xml, 'application/xml');
    if (parsed.querySelector('parsererror')) throw new Error('The embedded library-source list is invalid XML.');
    const root = parsed.documentElement;
    if (root.localName !== 'bibwordLibraries' || root.namespaceURI !== librarySourceNamespace) return null;
    if (root.getAttribute('version') !== '1') throw new Error('The embedded library-source list has an unsupported version.');
    return [...root.getElementsByTagNameNS(librarySourceNamespace, 'source')].map((node) => {
      const hash = node.getAttribute('hash') || '';
      const nameNode = node.getElementsByTagNameNS(librarySourceNamespace, 'name')[0];
      if (!/^[0-9a-f]{64}$/.test(hash) || !nameNode) {
        throw new Error('The embedded library-source list has invalid metadata.');
      }
      return { name: base64ToText(nameNode.textContent || ''), hash };
    });
  }

  function parseCustomStyleStore(xml) {
    const parsed = new DOMParser().parseFromString(xml, 'application/xml');
    if (parsed.querySelector('parsererror')) throw new Error('The embedded bibdocqmd CSL part is invalid XML.');
    const root = parsed.documentElement;
    if (root.localName !== 'bibwordStyle' || root.namespaceURI !== customStyleNamespace) return null;
    const id = root.getAttribute('id') || '';
    const hash = root.getAttribute('hash') || '';
    const labelNode = root.getElementsByTagNameNS(customStyleNamespace, 'label')[0];
    const cslNode = root.getElementsByTagNameNS(customStyleNamespace, 'csl')[0];
    if (!customStyleIdPattern.test(id) || !/^[0-9a-f]{64}$/.test(hash) || !labelNode || !cslNode) {
      throw new Error('The embedded bibdocqmd CSL part has invalid metadata.');
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
      option.dataset.key = key;
      const source = cslLibrary.sourceForKey(key);
      option.textContent = source ? `${key} — ${source}` : key;
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
      chooseSuggestion(options[activeSuggestion].dataset.key);
    }
  }

  async function persistLibrarySources(sources) {
    if (!supportsCustomStylePersistence()) return false;
    await Word.run(async (context) => {
      const existing = context.document.customXmlParts.getByNamespace(librarySourceNamespace);
      existing.load('items/id');
      await context.sync();
      existing.items.forEach((part) => part.delete());
      if (sources.length) context.document.customXmlParts.add(librarySourceStoreXml(sources));
      await context.sync();
    });
    return true;
  }

  function savedGithubUrls() {
    return rememberedLibrarySources.map(({ name }) => name).filter((name) => {
      try { return githubBibliographyUrl(name) === name; } catch { return false; }
    });
  }

  async function restoreDocumentLibrarySources() {
    if (!supportsCustomStylePersistence()) return;
    try {
      await Word.run(async (context) => {
        const parts = context.document.customXmlParts.getByNamespace(librarySourceNamespace);
        parts.load('items/id');
        await context.sync();
        const blobs = parts.items.map((part) => part.getXml());
        await context.sync();
        const stores = blobs.map(({ value }) => parseLibrarySourceStore(value)).filter(Boolean);
        if (stores.length > 1) throw new Error('The document contains more than one library-source list.');
        rememberedLibrarySources = stores[0] || [];
      });
      if (rememberedLibrarySources.length) {
        const names = rememberedLibrarySources.map(({ name }) => name).join(', ');
        $('library-status').textContent = `This document last used: ${names}. Reload saved GitHub files or choose local files again to query them.`;
        const githubUrls = savedGithubUrls();
        $('reload-github-libraries').hidden = githubUrls.length === 0;
        if (githubUrls.length === 1) $('github-bib-url').value = githubUrls[0];
      }
    } catch (error) {
      $('library-status').textContent = 'Could not restore the saved library filenames: '
        + (error.message || String(error));
    }
  }

  async function addBibliographySources(addedSources, originMessage) {
    const combinedSources = new Map(activeLibrarySources);
    addedSources.forEach((source) => combinedSources.set(source.name, source));
    const sources = [...combinedSources.values()];
    cslLibrary = createLibraryFromSources(sources);
    activeLibrarySources = combinedSources;
    knownKeys = new Set(cslLibrary.keys);
    const rememberedHashes = new Map(rememberedLibrarySources.map(({ name, hash }) => [name, hash]));
    const changedSources = sources
      .filter(({ name, hash }) => rememberedHashes.has(name) && rememberedHashes.get(name) !== hash)
      .map(({ name }) => name);
    const names = sources.map(({ name }) => name).join(', ');
    $('library-status').textContent = `${sources.length} file${sources.length === 1 ? '' : 's'}: ${names}. ${cslLibrary.size} unique entries parsed as BibLaTeX/CSL.`;
    $('refresh-bibliography').disabled = !officeReady;
    $('format-bibliography').disabled = !officeReady;
    $('show-formatted-citations').disabled = !officeReady;
    $('clear-bibliographies').disabled = false;
    let persisted = false;
    let persistenceWarning = '';
    if (officeReady) {
      try {
        persisted = await persistLibrarySources(sources);
        if (persisted) {
          rememberedLibrarySources = sources.map(({ name, hash }) => ({ name, hash }));
          $('reload-github-libraries').hidden = savedGithubUrls().length === 0;
        }
      } catch (error) {
        persistenceWarning = ` Filenames could not be saved with the document: ${error.message || String(error)}`;
      }
    }
    const changeWarning = changedSources.length
      ? ` Warning: file contents changed since the document last used ${changedSources.join(', ')}.` : '';
    setStatus(`${originMessage} Combined bibliography parsed in the add-in; no bibliography content uploaded.${persisted ? ' Source names and hashes were saved with the document.' : ''}${changeWarning}${persistenceWarning}`);
    markWordBibliographyStale();
    markFormattedCitationsStale();
    renderSuggestions();
    updatePreview();
    if (documentCitationKeys.length) renderDocumentBibliography();
  }

  async function loadBibFiles(event) {
    const files = [...(event.target.files || [])];
    if (!files.length) return;
    try {
      const addedSources = await Promise.all(files.map(async (file) => {
        const contents = await file.text();
        return { name: file.name, contents, hash: await sha256Hex(contents) };
      }));
      await addBibliographySources(addedSources, 'Selected files loaded.');
    } catch (error) {
      if (!cslLibrary) {
        knownKeys = null;
        $('refresh-bibliography').disabled = true;
        $('format-bibliography').disabled = true;
        $('show-formatted-citations').disabled = true;
      }
      setStatus('Could not combine the selected bibliographies: ' + (error.message || String(error)), true);
    } finally {
      event.target.value = '';
    }
  }

  async function loadGithubBibliographies(urls, originMessage) {
    const button = $('load-github-library');
    const reloadButton = $('reload-github-libraries');
    button.disabled = true;
    reloadButton.disabled = true;
    $('github-library-status').textContent = 'Reading public GitHub bibliography…';
    try {
      const addedSources = await Promise.all(urls.map(async (url) => {
        const source = await fetchGithubBibliography(url);
        return { ...source, hash: await sha256Hex(source.contents) };
      }));
      await addBibliographySources(addedSources, originMessage);
      $('github-library-status').textContent = `${addedSources.length} GitHub file${addedSources.length === 1 ? '' : 's'} loaded. Use Load or refresh again after the repository changes.`;
    } catch (error) {
      const message = error.message || String(error);
      $('github-library-status').textContent = `Could not load GitHub bibliography: ${message}`;
      setStatus(`Could not load GitHub bibliography: ${message}`, true);
    } finally {
      button.disabled = false;
      reloadButton.disabled = false;
    }
  }

  function loadGithubBibliography() {
    const value = $('github-bib-url').value.trim();
    if (value) loadGithubBibliographies([value], 'GitHub bibliography loaded.');
    else $('github-library-status').textContent = 'Enter a public GitHub .bib or .bibtex file link.';
  }

  function reloadSavedGithubBibliographies() {
    const urls = savedGithubUrls();
    if (urls.length) loadGithubBibliographies(urls, 'Saved GitHub bibliographies reloaded.');
  }

  function formatFileSize(bytes) {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }

  async function refreshSharedBibliographyList() {
    const refreshButton = $('refresh-shared-libraries');
    const selector = $('shared-bib-files');
    const loadButton = $('load-shared-libraries');
    refreshButton.disabled = true;
    $('shared-library-status').textContent = 'Checking the configured synced folder…';
    try {
      const response = await fetch('/api/shared-bibliographies', { cache: 'no-store' });
      if (!response.ok) throw new Error(`The local server returned HTTP ${response.status}.`);
      const result = await response.json();
      selector.replaceChildren();
      selector.hidden = true;
      selector.disabled = true;
      loadButton.disabled = true;
      if (!result.configured) {
        $('shared-library-status').textContent = 'No synced folder is configured. Restart the server with BIBDOCQMD_LIBRARY_DIR set to the local SharePoint folder.';
        return;
      }
      result.files.forEach((file) => {
        const option = document.createElement('option');
        option.value = file.name;
        option.textContent = `${file.name} — ${formatFileSize(file.size)}`;
        selector.appendChild(option);
      });
      if (!result.files.length) {
        $('shared-library-status').textContent = `No .bib files were found in the synced folder “${result.folder}”.`;
        return;
      }
      selector.hidden = false;
      selector.disabled = false;
      if (result.files.length === 1) selector.options[0].selected = true;
      loadButton.disabled = selector.selectedOptions.length === 0;
      $('shared-library-status').textContent = `${result.files.length} bibliography file${result.files.length === 1 ? '' : 's'} found in “${result.folder}”. Select one or more, then load them.`;
    } catch (error) {
      $('shared-library-status').textContent = 'Could not list the synced bibliographies: '
        + (error.message || String(error));
    } finally {
      refreshButton.disabled = false;
    }
  }

  async function loadSharedBibliographies() {
    const selector = $('shared-bib-files');
    const names = [...selector.selectedOptions].map(({ value }) => value);
    if (!names.length) return;
    const loadButton = $('load-shared-libraries');
    loadButton.disabled = true;
    $('shared-library-status').textContent = 'Reading the selected synced bibliographies…';
    try {
      const addedSources = await Promise.all(names.map(async (name) => {
        const query = new URLSearchParams({ name });
        const response = await fetch(`/api/shared-bibliographies/file?${query}`, { cache: 'no-store' });
        if (!response.ok) throw new Error(`${name}: the local server returned HTTP ${response.status}.`);
        const contents = await response.text();
        return { name, contents, hash: await sha256Hex(contents) };
      }));
      await addBibliographySources(addedSources, 'Synced SharePoint bibliographies loaded.');
      $('shared-library-status').textContent = `${names.length} selected file${names.length === 1 ? '' : 's'} loaded. Choose Load or refresh selected again after the synced files change.`;
    } catch (error) {
      $('shared-library-status').textContent = 'Could not load the selected synced bibliographies: '
        + (error.message || String(error));
      setStatus('Could not combine the selected bibliographies: ' + (error.message || String(error)), true);
    } finally {
      loadButton.disabled = selector.selectedOptions.length === 0;
    }
  }

  async function clearBibliographies() {
    activeLibrarySources = new Map();
    cslLibrary = null;
    knownKeys = null;
    $('bib-file').value = '';
    $('clear-bibliographies').disabled = true;
    $('refresh-bibliography').disabled = true;
    $('format-bibliography').disabled = true;
    $('show-formatted-citations').disabled = true;
    $('library-status').textContent = 'No libraries selected. You can still enter a citation key manually.';
    rememberedLibrarySources = [];
    $('reload-github-libraries').hidden = true;
    let warning = '';
    if (officeReady) {
      try {
        await persistLibrarySources([]);
      } catch (error) {
        warning = ` Saved filenames could not be cleared: ${error.message || String(error)}`;
      }
    }
    hideSuggestions();
    updatePreview();
    renderDocumentBibliography();
    setStatus(`Combined bibliography cleared.${warning}`, Boolean(warning));
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
      if (!markers.length) return;
      const styleId = selectedStyle();
      if (customStyleIdPattern.test(styleId)) {
        const style = customStyles.get(styleId);
        if (!style) throw new Error(`Custom CSL style ${styleId} is not loaded.`);
        await ensureCustomStylePart(context, style);
      }
      markers.forEach((marker) => {
        const metadata = parseBibliographyTag(marker.tag);
        marker.tag = bibliographyTag(metadata.id, styleId, metadata.scope);
      });
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
        const styled = markers.map(({ tag }) => parseBibliographyTag(tag)?.style).filter(Boolean);
        const styles = [...new Set(styled)];
        if (styles.length > 1) {
          throw new Error('Bibliography markers use different citation styles. Choose a style to make them consistent.');
        }
        markerTag = markers.find(({ tag }) => parseBibliographyTag(tag)?.style)?.tag || null;
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
        if (Office.context.requirements.isSetSupported('WordApi', '1.3')) {
          control.getRange(Word.RangeLocation.after).select();
        } else {
          inserted.select(Word.SelectionMode.end);
        }
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
        throw new Error('The selected content control is not a bibdocqmd citation.');
      }
      const { source, parsed } = parseCitationControl(selected);
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
      $('editing-status').textContent = 'To edit, place the cursor inside a bibdocqmd citation and load it.';
      setStatus('Could not load the selected bibdocqmd citation: '
        + (error.message || String(error)), true);
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
    if (selectedControlId === null) return setStatus('Load a selected bibdocqmd citation first.', true);
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
        if (Office.context.requirements.isSetSupported('WordApi', '1.3')) {
          control.getRange(Word.RangeLocation.after).select();
        } else {
          updated.select(Word.SelectionMode.end);
        }
        await context.sync();
      });
      $('editing-status').textContent = `Editing ${syntax}`;
      setStatus('Selected citation updated without changing its bibdocqmd identity.');
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
        const replacements = controls.map((control, index) => {
          const { source, parsed } = parseCitationControl(control, index + 1);
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
        const scope = $('bibliography-scope').value;
        const selection = context.document.getSelection();
        const paragraph = selection.insertParagraph('References', Word.InsertLocation.after);
        paragraph.styleBuiltIn = `Heading${bibliographyHeadingLevel(scope)}`;
        const control = paragraph.insertContentControl();
        control.tag = bibliographyTag(uniqueId(), styleId, scope);
        control.title = 'bibdocqmd bibliography position';
        const exitParagraph = control.insertParagraph('', Word.InsertLocation.after);
        exitParagraph.styleBuiltIn = 'Normal';
        exitParagraph.select(Word.SelectionMode.start);
        await context.sync();
      });
      setStatus('Bibliography marker inserted. The cursor is in a normal paragraph below it.');
      markWordBibliographyStale('Bibliography marker inserted. Choose Update in Word to generate references.');
    } catch (error) {
      setStatus('Bibliography marker failed: ' + (error.message || String(error)), true);
    }
  }

  async function formatBibliographyInWord() {
    if (!officeReady) return setStatus('Office.js is not ready. Open this pane inside Word.', true);
    if (!cslLibrary) return setStatus('Load a bibliography before formatting it in Word.', true);
    try {
      const indentPoints = bibliographyIndentPoints();
      let referenceCount = 0;
      let bibliographyCount = 0;
      let addedExitParagraph = false;
      await Word.run(async (context) => {
        const controls = context.document.contentControls;
        controls.load('items/tag,text,title');
        let plans;
        let paragraphState = null;
        if (Office.context.requirements.isSetSupported('WordApi', '1.3')) {
          paragraphState = await bibliographyPlansFromWordParagraphs(context);
          plans = paragraphState.plans;
        } else {
          const bodyOoxml = context.document.body.getOoxml();
          await context.sync();
          plans = bibliographyPlansFromOoxml(bodyOoxml.value);
        }
        // The paragraph path performs its own synchronizations; this loads the
        // document-wide control collection in both branches.
        await context.sync();
        const markers = controls.items.filter(({ tag }) => tag.startsWith('bibword:refs:'));
        if (!markers.length) throw new Error('Insert a bibliography marker before updating the Word bibliography.');
        plans.forEach((plan) => {
          if (!plan.keys.length) {
            const label = plan.metadata.scope === 'document'
              ? 'the document' : `its ${plan.metadata.scope.replace('heading', 'Heading ')} section`;
            throw new Error(`The bibliography marker “${plan.title}” has no managed citations in ${label}.`);
          }
        });
        const styleId = selectedStyle();
        if (customStyleIdPattern.test(styleId)) {
          const style = customStyles.get(styleId);
          if (!style) throw new Error(`Custom CSL style ${styleId} is not loaded.`);
          await ensureCustomStylePart(context, style);
        }
        const formattedBibliographies = [];
        markers.forEach((marker) => {
          const metadata = parseBibliographyTag(marker.tag);
          const plan = plans.get(metadata.id);
          if (!plan) throw new Error(`Could not locate bibliography marker ${metadata.id} in the document body.`);
          const title = marker.text.split(/[\r\n]+/)[0].trim() || plan.title;
          const bibliography = cslLibrary.formatBibliographyHtml(plan.keys, styleId);
          marker.tag = bibliographyTag(metadata.id, styleId, metadata.scope);
          marker.title = 'bibdocqmd bibliography position and generated references';
          marker.insertHtml(wordBibliographyHtml(
            title, bibliography, bibliographyHeadingLevel(metadata.scope),
          ), Word.InsertLocation.replace);
          const paragraphs = marker.paragraphs;
          paragraphs.load('items');
          formattedBibliographies.push({ paragraphs, scope: metadata.scope });
          if (paragraphState && !paragraphState.markerHasFollowingParagraph.get(metadata.id)) {
            const exitParagraph = marker.insertParagraph('', Word.InsertLocation.after);
            exitParagraph.styleBuiltIn = 'Normal';
            exitParagraph.select(Word.SelectionMode.start);
            addedExitParagraph = true;
          }
          referenceCount += plan.keys.length;
        });
        await context.sync();
        // Guard against HTML replacement retaining the marker's old heading
        // paragraph style on reference entries. Finish applying styles before
        // setting direct paragraph formatting, which a style change can clear.
        formattedBibliographies.forEach(({ paragraphs, scope }) => {
          paragraphs.items.forEach((paragraph, index) => {
            paragraph.styleBuiltIn = index === 0
              ? `Heading${bibliographyHeadingLevel(scope)}` : 'Normal';
          });
        });
        await context.sync();
        formattedBibliographies.forEach(({ paragraphs }) => {
          paragraphs.items.forEach((paragraph, index) => {
            paragraph.leftIndent = index === 0 ? 0 : indentPoints;
            paragraph.firstLineIndent = index === 0 || indentPoints === 0 ? 0 : -indentPoints;
            paragraph.rightIndent = 0;
            if (index > 0) paragraph.alignment = 'Left';
          });
        });
        await context.sync();
        documentCitationKeys = citationKeysFromControls(
          controls.items.map(({ tag, text, title }) => ({ tag, text, title })),
        );
        bibliographyCount = markers.length;
      });
      renderDocumentBibliography();
      const styleLabel = $('citation-style').selectedOptions[0].textContent;
      $('bibliography-render-status').textContent = `${bibliographyCount} bibliograph${bibliographyCount === 1 ? 'y' : 'ies'} updated with ${referenceCount} scoped reference entries in ${styleLabel}.`;
      setStatus(addedExitParagraph
        ? 'Word bibliographies updated. A normal paragraph was added after the final bibliography for continued typing.'
        : 'Word bibliographies updated. Export will rebuild each scope from the citation keys.');
    } catch (error) {
      $('bibliography-render-status').textContent = error.message || String(error);
      setStatus('Could not update the Word bibliography: ' + (error.message || String(error)), true);
    }
  }

  function downloadExport(bytes, filename) {
    const url = URL.createObjectURL(new Blob([bytes], { type: 'application/zip' }));
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60_000);
  }

  async function runMarkdownExport(action) {
    const status = $('quarto-export-status');
    const prepareButton = $('prepare-markdown-export');
    const finishButton = $('finish-markdown-export');
    prepareButton.disabled = true;
    finishButton.disabled = true;
    status.className = '';
    status.textContent = 'Processing locally…';
    try {
      if (action === 'prepare') {
        const file = $('export-docx').files[0];
        if (!file || !/\.docx$/i.test(file.name)) throw new Error('Choose your saved original DOCX file.');
        if (!activeLibrarySources.size) throw new Error('Load the matching bibliography files first.');
        if (file.size > 25_000_000) throw new Error('Choose a DOCX smaller than 25 MB.');
        const bytes = await globalThis.BibWordMarkdown.prepareMarkdownExport(
          new Uint8Array(await file.arrayBuffer()), {
            sources: [...activeLibrarySources.values()], title: file.name.replace(/\.docx$/i, ''),
            fallbackStyle: selectedStyle(),
          });
        downloadExport(bytes, 'bibdocqmd-preparation.zip');
        status.textContent = 'Preparation ZIP ready. Extract it, open word-markdown-copy.docx in Word for the web, and export that copy as Markdown. Then select both files below.';
      } else {
        const preparation = $('export-preparation').files[0];
        const markdown = $('export-markdown').files[0];
        if (!preparation || !markdown || !/\.zip$/i.test(preparation.name) || !/\.(md|zip)$/i.test(markdown.name)) {
          throw new Error('Choose the preparation ZIP and Word’s exported .md or media ZIP.');
        }
        if (preparation.size > 25_000_000 || markdown.size > 25_000_000) throw new Error('Choose files smaller than 25 MB.');
        const bytes = await globalThis.BibWordMarkdown.finishMarkdownExport(
          new Uint8Array(await preparation.arrayBuffer()), new Uint8Array(await markdown.arrayBuffer()),
          { markdownZip: /\.zip$/i.test(markdown.name) });
        downloadExport(bytes, 'bibdocqmd-quarto.zip');
        status.textContent = 'QMD package ready. Extract it and render document.qmd with Quarto. Check formatting and keep any separately exported images alongside it.';
      }
    } catch (error) {
      status.className = 'error';
      status.textContent = 'Export failed: ' + (error.message || String(error));
    } finally {
      prepareButton.disabled = false;
      finishButton.disabled = false;
    }
  }

  function openLocalTools(event) {
    if (typeof Office === 'undefined' || !Office.context?.ui?.openBrowserWindow) return;
    try {
      Office.context.ui.openBrowserWindow(event.currentTarget.href);
      event.preventDefault();
    } catch (error) {
      // Keep the ordinary link usable when the Office API is unavailable.
      setStatus('Could not open the browser through Word. Open or copy the local tools link manually. '
        + (error.message || String(error)), true);
    }
  }

  $('bib-file').addEventListener('change', loadBibFiles);
  $('prepare-markdown-export').addEventListener('click', () => runMarkdownExport('prepare'));
  $('finish-markdown-export').addEventListener('click', () => runMarkdownExport('finish'));
  $('refresh-shared-libraries').addEventListener('click', refreshSharedBibliographyList);
  $('load-github-library').addEventListener('click', loadGithubBibliography);
  $('reload-github-libraries').addEventListener('click', reloadSavedGithubBibliographies);
  $('shared-bib-files').addEventListener('change', () => {
    $('load-shared-libraries').disabled = $('shared-bib-files').selectedOptions.length === 0;
  });
  $('load-shared-libraries').addEventListener('click', loadSharedBibliographies);
  $('clear-bibliographies').addEventListener('click', clearBibliographies);
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
  $('bibliography-indent').addEventListener('input', () => markWordBibliographyStale());
  $('refresh-bibliography').addEventListener('click', refreshDocumentBibliography);
  globalThis.BibWordFilePicker.bindFilePickers(document);
  document.querySelectorAll('[data-open-local-tools]').forEach((link) => {
    link.addEventListener('click', openLocalTools);
  });
  document.addEventListener('click', (event) => {
    // Let file controls finish their default action without redrawing the pane.
    if (event.target.closest('input[type="file"], [data-open-local-tools]')) return;
    if (!$('citation-keys').closest('.key-picker').contains(event.target)) hideSuggestions();
  });
  updatePreview();
  const browserMode = document.documentElement.dataset.localTools === 'true'
    || (typeof location !== 'undefined' && /(?:[?&])browser=1(?:&|$)/.test(location.search));
  if (browserMode) {
    $('browser-mode-note').hidden = false;
    setStatus('Local browser tools ready. Load bibliography files and use the export controls below. Return to Word for document editing.');
  } else if (typeof Office !== 'undefined') {
    Office.onReady(async (info) => {
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
      await restoreDocumentStyle();
      await restoreDocumentLibrarySources();
    });
  } else {
    setStatus('Office.js unavailable. Load the add-in within Microsoft Word.', true);
  }
})();
