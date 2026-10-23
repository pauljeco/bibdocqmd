/* Pure citation helpers shared by the task pane and the Node test suite. */
(function exposeCitationHelpers(root, factory) {
  const helpers = factory();
  if (typeof module === 'object' && module.exports) module.exports = helpers;
  else root.BibWordCitation = helpers;
}(typeof globalThis !== 'undefined' ? globalThis : this, () => {
  'use strict';

  function parseKeys(bibtex) {
    const found = new Set();
    const entry = /@([a-zA-Z]+)\s*[({]\s*([^,\s]+)\s*,/g;
    let match;
    while ((match = entry.exec(bibtex)) !== null) {
      if (!['string', 'comment', 'preamble'].includes(match[1].toLowerCase())) {
        found.add(match[2]);
      }
    }
    return found;
  }

  function activeKeyFragment(value) {
    return value.slice(value.lastIndexOf(';') + 1).trim();
  }

  function suggestKeys(keys, value, limit = 8) {
    const fragment = activeKeyFragment(value).toLowerCase();
    if (!fragment) return [];
    const completed = new Set(
      value.split(';').slice(0, -1).map((key) => key.trim()).filter(Boolean),
    );
    return [...keys]
      .filter((key) => !completed.has(key) && key.toLowerCase().includes(fragment))
      .sort((left, right) => {
        const leftStarts = left.toLowerCase().startsWith(fragment);
        const rightStarts = right.toLowerCase().startsWith(fragment);
        if (leftStarts !== rightStarts) return leftStarts ? -1 : 1;
        return left.localeCompare(right);
      })
      .slice(0, limit);
  }

  function replaceActiveKey(value, key) {
    const separator = value.lastIndexOf(';');
    if (separator < 0) return key;
    return `${value.slice(0, separator).trimEnd()}; ${key}`;
  }

  function citationSyntax(keys, mode, locator) {
    if (!keys.length) throw new Error('Enter at least one citation key.');
    if (keys.some((key) => !/^[a-zA-Z0-9_][a-zA-Z0-9_:.$#%&+?<>~/-]*$/.test(key))) {
      throw new Error('Invalid citation key. Please choose a key from the .bib file.');
    }
    if (mode !== 'parenthetical' && keys.length !== 1) {
      throw new Error('Narrative and suppressed-author citations currently require one key.');
    }
    const suffix = locator.trim();
    if (suffix.includes(']') || suffix.includes('[') || suffix.includes('\n') || suffix.includes('\r')) {
      throw new Error('Locator must not contain brackets or line breaks.');
    }
    if (mode === 'narrative') return `@${keys[0]}${suffix ? ` [${suffix}]` : ''}`;
    if (mode === 'suppress') return `[-@${keys[0]}${suffix ? `, ${suffix}` : ''}]`;
    if (mode !== 'parenthetical') throw new Error(`Unknown citation mode: ${mode}`);
    const entries = keys.map((key) => `@${key}`);
    if (suffix) entries[entries.length - 1] += `, ${suffix}`;
    return `[${entries.join('; ')}]`;
  }

  return { activeKeyFragment, citationSyntax, parseKeys, replaceActiveKey, suggestKeys };
}));
