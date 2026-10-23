/* Pure citation helpers shared by the task pane and the Node test suite. */
(function exposeCitationHelpers(root, factory) {
  const helpers = factory();
  if (typeof module === 'object' && module.exports) module.exports = helpers;
  else root.BibWordCitation = helpers;
}(typeof globalThis !== 'undefined' ? globalThis : this, () => {
  'use strict';

  function parseKeys(bibtex) {
    return new Set(parseBibEntries(bibtex).keys());
  }

  function matchingDelimiter(opening) {
    return opening === '{' ? '}' : ')';
  }

  function findEntryEnd(source, start, opening) {
    const closing = matchingDelimiter(opening);
    let depth = 1;
    let quoted = false;
    for (let index = start; index < source.length; index += 1) {
      const character = source[index];
      if (character === '"' && source[index - 1] !== '\\') quoted = !quoted;
      if (quoted) continue;
      if (character === opening) depth += 1;
      if (character === closing) depth -= 1;
      if (depth === 0) return index;
    }
    return source.length;
  }

  function splitEntryHead(body) {
    let braceDepth = 0;
    let quoted = false;
    for (let index = 0; index < body.length; index += 1) {
      const character = body[index];
      if (character === '"' && body[index - 1] !== '\\') quoted = !quoted;
      if (quoted) continue;
      if (character === '{') braceDepth += 1;
      if (character === '}') braceDepth -= 1;
      if (character === ',' && braceDepth === 0) {
        return [body.slice(0, index).trim(), body.slice(index + 1)];
      }
    }
    return [body.trim(), ''];
  }

  function cleanBibValue(value) {
    let clean = value.trim();
    if ((clean.startsWith('{') && clean.endsWith('}'))
        || (clean.startsWith('"') && clean.endsWith('"'))) {
      clean = clean.slice(1, -1);
    }
    const accentMaps = [
      [/\\'\{?([A-Za-z])\}?/g, { a: 'á', e: 'é', i: 'í', o: 'ó', u: 'ú', y: 'ý', A: 'Á', E: 'É', I: 'Í', O: 'Ó', U: 'Ú', Y: 'Ý' }],
      [/\\`\{?([A-Za-z])\}?/g, { a: 'à', e: 'è', i: 'ì', o: 'ò', u: 'ù', A: 'À', E: 'È', I: 'Ì', O: 'Ò', U: 'Ù' }],
      [/\\\"\{?([A-Za-z])\}?/g, { a: 'ä', e: 'ë', i: 'ï', o: 'ö', u: 'ü', y: 'ÿ', A: 'Ä', E: 'Ë', I: 'Ï', O: 'Ö', U: 'Ü' }],
      [/\\\^\{?([A-Za-z])\}?/g, { a: 'â', e: 'ê', i: 'î', o: 'ô', u: 'û', A: 'Â', E: 'Ê', I: 'Î', O: 'Ô', U: 'Û' }],
      [/\\~\{?([A-Za-z])\}?/g, { a: 'ã', n: 'ñ', o: 'õ', A: 'Ã', N: 'Ñ', O: 'Õ' }],
      [/\\c\{?([Cc])\}?/g, { c: 'ç', C: 'Ç' }],
    ];
    accentMaps.forEach(([pattern, replacements]) => {
      clean = clean.replace(pattern, (match, letter) => replacements[letter] || letter);
    });
    return clean
      .replace(/[{}]/g, '')
      .replace(/~/g, ' ')
      .replace(/\\&/g, '&')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function parseFields(source) {
    const fields = {};
    let index = 0;
    while (index < source.length) {
      while (index < source.length && /[\s,]/.test(source[index])) index += 1;
      const nameMatch = /^[a-zA-Z][\w-]*/.exec(source.slice(index));
      if (!nameMatch) break;
      const name = nameMatch[0].toLowerCase();
      index += nameMatch[0].length;
      while (index < source.length && /\s/.test(source[index])) index += 1;
      if (source[index] !== '=') {
        while (index < source.length && source[index] !== ',') index += 1;
        continue;
      }
      index += 1;
      while (index < source.length && /\s/.test(source[index])) index += 1;
      const start = index;
      if (source[index] === '{') {
        index = findEntryEnd(source, index + 1, '{') + 1;
      } else if (source[index] === '"') {
        index += 1;
        while (index < source.length) {
          if (source[index] === '"' && source[index - 1] !== '\\') {
            index += 1;
            break;
          }
          index += 1;
        }
      } else {
        while (index < source.length && source[index] !== ',') index += 1;
      }
      fields[name] = cleanBibValue(source.slice(start, index));
    }
    return fields;
  }

  function parseBibEntries(bibtex) {
    const entries = new Map();
    const header = /@([a-zA-Z]+)\s*([({])/g;
    let match;
    while ((match = header.exec(bibtex)) !== null) {
      const type = match[1].toLowerCase();
      const opening = match[2];
      const bodyStart = header.lastIndex;
      const bodyEnd = findEntryEnd(bibtex, bodyStart, opening);
      header.lastIndex = Math.max(bodyEnd + 1, header.lastIndex);
      if (['string', 'comment', 'preamble'].includes(type)) continue;
      const [key, fieldsSource] = splitEntryHead(bibtex.slice(bodyStart, bodyEnd));
      if (!key || /\s/.test(key)) continue;
      entries.set(key, { key, type, fields: parseFields(fieldsSource) });
    }
    return entries;
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

  const KEY_SOURCE = '[a-zA-Z0-9_][a-zA-Z0-9_:.$#%&+?<>~/-]*';

  function parseCitationSyntax(syntax) {
    const text = syntax.trim();
    const narrative = new RegExp(`^@(${KEY_SOURCE})(?:\\s+\\[([^\\]\\r\\n]*)\\])?$`).exec(text);
    if (narrative) {
      return { keys: [narrative[1]], mode: 'narrative', locator: (narrative[2] || '').trim() };
    }
    if (!text.startsWith('[') || !text.endsWith(']')) {
      throw new Error('Selected control does not contain supported BibWord citation syntax.');
    }
    const parts = text.slice(1, -1).split(';').map((part) => part.trim()).filter(Boolean);
    if (!parts.length) throw new Error('Selected citation has no keys.');
    const suppress = parts[0].startsWith('-@');
    if (suppress && parts.length !== 1) {
      throw new Error('Suppressed-author citations currently require one key.');
    }
    const keys = [];
    let locator = '';
    parts.forEach((part, index) => {
      const prefix = suppress && index === 0 ? '-@' : '@';
      if (!part.startsWith(prefix)) throw new Error('Selected control has invalid citation syntax.');
      const value = part.slice(prefix.length);
      const comma = value.indexOf(',');
      const key = (comma < 0 ? value : value.slice(0, comma)).trim();
      if (!new RegExp(`^${KEY_SOURCE}$`).test(key)) {
        throw new Error('Selected control has an invalid citation key.');
      }
      if (comma >= 0) {
        if (index !== parts.length - 1) {
          throw new Error('Only the final key can carry a locator.');
        }
        locator = value.slice(comma + 1).trim();
      }
      keys.push(key);
    });
    return { keys, mode: suppress ? 'suppress' : 'parenthetical', locator };
  }

  function familyName(person) {
    const clean = person.replace(/[{}]/g, '').trim();
    if (!clean) return '';
    if (clean.includes(',')) return clean.split(',')[0].trim();
    const words = clean.split(/\s+/);
    return words[words.length - 1];
  }

  function authorLabel(author) {
    if (!author) return '';
    const names = author.split(/\s+and\s+/i).map(familyName).filter(Boolean);
    if (names.length === 1) return names[0];
    if (names.length === 2) return `${names[0]} and ${names[1]}`;
    return `${names[0]} et al.`;
  }

  function entryYear(entry) {
    if (!entry) return 'n.d.';
    return entry.fields.year || (entry.fields.date || '').slice(0, 4) || 'n.d.';
  }

  function formatCitationPreview(keys, mode, locator, bibliography) {
    const labels = keys.map((key) => {
      const entry = bibliography.get(key);
      const author = entry ? authorLabel(entry.fields.author || entry.fields.editor) : key;
      return `${author || key}, ${entryYear(entry)}`;
    });
    const suffix = locator.trim();
    if (mode === 'narrative') {
      const entry = bibliography.get(keys[0]);
      const author = entry ? authorLabel(entry.fields.author || entry.fields.editor) : keys[0];
      return `${author || keys[0]} (${entryYear(entry)}${suffix ? `, ${suffix}` : ''})`;
    }
    if (mode === 'suppress') {
      const entry = bibliography.get(keys[0]);
      return `(${entryYear(entry)}${suffix ? `, ${suffix}` : ''})`;
    }
    if (suffix && labels.length) labels[labels.length - 1] += `, ${suffix}`;
    return `(${labels.join('; ')})`;
  }

  function formatReference(entry) {
    if (!entry) return '';
    const fields = entry.fields;
    const authors = cleanBibValue(fields.author || fields.editor || entry.key)
      .replace(/\s+and\s+/gi, ', ');
    const year = entryYear(entry);
    const title = fields.title || '[Untitled]';
    const container = fields.journaltitle || fields.journal || fields.booktitle || fields.publisher || '';
    const details = [container, fields.volume, fields.number ? `(${fields.number})` : '', fields.pages]
      .filter(Boolean).join(' ');
    return `${authors}. ${year}. ${title}.${details ? ` ${details}.` : ''}`;
  }

  return {
    activeKeyFragment,
    citationSyntax,
    formatCitationPreview,
    formatReference,
    parseBibEntries,
    parseCitationSyntax,
    parseKeys,
    replaceActiveKey,
    suggestKeys,
  };
}));
