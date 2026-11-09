/* Pure citation helpers shared by the task pane and the Node test suite. */
(function exposeCitationHelpers(root, factory) {
  const helpers = factory();
  if (typeof module === 'object' && module.exports) module.exports = helpers;
  else root.BibWordCitation = helpers;
}(typeof globalThis !== 'undefined' ? globalThis : this, () => {
  'use strict';

  const CITATION_TITLE = 'bibdocqmd citation';
  const CITATION_SOURCE_TITLE_PREFIX = 'bibdocqmd citation source v1: ';
  const LEGACY_CITATION_SOURCE_TITLE_PREFIX = 'BibWord citation source v1: ';
  const CITATION_SOURCE_TITLE_PREFIXES = [
    CITATION_SOURCE_TITLE_PREFIX, LEGACY_CITATION_SOURCE_TITLE_PREFIX,
  ];

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
    for (let pass = 0; pass < 3; pass += 1) {
      clean = clean.replace(/\\(?:textit|textsf|textrm|textbf|emph)\s*\{([^{}]*)\}/g, '$1');
    }
    clean = clean
      .replace(/\\i(?=[^A-Za-z]|$)/g, 'i')
      .replace(/\\j(?=[^A-Za-z]|$)/g, 'j');
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
    const letterCommands = {
      ae: 'æ', AE: 'Æ', oe: 'œ', OE: 'Œ', aa: 'å', AA: 'Å',
      ss: 'ß', l: 'ł', L: 'Ł', o: 'ø', O: 'Ø',
    };
    clean = clean.replace(/\\(ae|AE|oe|OE|aa|AA|ss|l|L|o|O)(?=[^A-Za-z]|$)/g,
      (match, command) => letterCommands[command]);
    return clean
      .replace(/[{}]/g, '')
      .replace(/~/g, ' ')
      .replace(/\\&/g, '&')
      .replace(/\\([_%#$])/g, '$1')
      .replace(/--/g, '–')
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
      throw new Error('Selected control does not contain supported bibdocqmd citation syntax.');
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

  function citationControlError(control, position, error) {
    const title = String(control.title || '');
    const prefix = CITATION_SOURCE_TITLE_PREFIXES.find((candidate) => title.startsWith(candidate));
    const source = prefix
      ? title.slice(prefix.length).trim()
      : String(control.text || '').trim();
    const compact = source.replace(/\s+/g, ' ');
    const preview = compact.length > 100 ? `${compact.slice(0, 97)}...` : compact;
    const label = position ? `Citation ${position}` : 'Selected citation';
    const hint = compact.startsWith('[') && compact.includes(']') && !compact.endsWith(']')
      ? ' Keep sentence punctuation outside the blue citation control.'
      : '';
    return new Error(`${label}${preview ? ` ${JSON.stringify(preview)}` : ''} is invalid: `
      + `${error.message || String(error)}${hint}`);
  }

  function parseCitationControl(control = {}, position = null) {
    try {
      const source = citationSyntaxFromControl(control);
      return { source, parsed: parseCitationSyntax(source) };
    } catch (error) {
      throw citationControlError(control, position, error);
    }
  }

  function citationKeysFromControls(controls) {
    const keys = [];
    const seen = new Set();
    let position = 0;
    controls.forEach((control = {}) => {
      if (!(control.tag || '').startsWith('bibword:cite:v1:')) return;
      position += 1;
      parseCitationControl(control, position).parsed.keys.forEach((key) => {
        if (!seen.has(key)) {
          seen.add(key);
          keys.push(key);
        }
      });
    });
    return keys;
  }

  function citationSourceTitle(syntax) {
    const source = String(syntax || '').trim();
    parseCitationSyntax(source);
    return CITATION_SOURCE_TITLE_PREFIX + source;
  }

  function citationSourceFromTitle(title) {
    const value = String(title || '');
    const prefix = CITATION_SOURCE_TITLE_PREFIXES.find((candidate) => value.startsWith(candidate));
    if (!prefix) return null;
    const source = value.slice(prefix.length).trim();
    parseCitationSyntax(source);
    return source;
  }

  function citationSyntaxFromControl({ text = '', title = '' } = {}) {
    return citationSourceFromTitle(title) || String(text).trim();
  }

  const BIBLIOGRAPHY_SCOPES = new Set(['document', 'heading1', 'heading2', 'heading3']);

  function bibliographyTag(id, style, scope = 'document') {
    if (!/^[a-zA-Z0-9-]+$/.test(id)) throw new Error('Invalid bibliography marker identity.');
    if (!/^[a-z0-9-]+$/.test(style)) throw new Error('Invalid bibliography style identifier.');
    if (!BIBLIOGRAPHY_SCOPES.has(scope)) throw new Error('Invalid bibliography scope.');
    return `bibword:refs:v3:${id}:${style}:${scope}`;
  }

  function parseBibliographyTag(tag) {
    const legacyPrefix = 'bibword:refs:v1:';
    const stylePrefix = 'bibword:refs:v2:';
    const scopePrefix = 'bibword:refs:v3:';
    if (tag.startsWith(legacyPrefix)) {
      const id = tag.slice(legacyPrefix.length);
      if (!id) throw new Error('Bibliography marker has no identity.');
      return { version: 1, id, style: null, scope: 'document' };
    }
    if (tag.startsWith(scopePrefix)) {
      const metadata = tag.slice(scopePrefix.length).split(':');
      if (metadata.length !== 3 || !metadata.every(Boolean)) {
        throw new Error('Bibliography marker has invalid scoped metadata.');
      }
      const [id, style, scope] = metadata;
      if (!BIBLIOGRAPHY_SCOPES.has(scope)) {
        throw new Error(`Unsupported bibliography scope: ${scope}`);
      }
      return { version: 3, id, style, scope };
    }
    if (!tag.startsWith(stylePrefix)) {
      if (tag.startsWith('bibword:refs:')) throw new Error(`Unsupported bibliography marker: ${tag}`);
      return null;
    }
    const metadata = tag.slice(stylePrefix.length);
    const separator = metadata.lastIndexOf(':');
    if (separator < 1 || separator === metadata.length - 1) {
      throw new Error('Bibliography marker has invalid style metadata.');
    }
    return {
      version: 2,
      id: metadata.slice(0, separator),
      style: metadata.slice(separator + 1),
      scope: 'document',
    };
  }

  function citationKeysForBibliographyScope(paragraphs, markerIndex, scope) {
    if (!BIBLIOGRAPHY_SCOPES.has(scope)) throw new Error(`Unsupported bibliography scope: ${scope}`);
    if (!Number.isInteger(markerIndex) || markerIndex < 0 || markerIndex >= paragraphs.length) {
      throw new Error('Bibliography marker position is invalid.');
    }
    let start = 0;
    let end = paragraphs.length;
    if (scope !== 'document') {
      const level = Number(scope.slice('heading'.length));
      start = -1;
      for (let index = markerIndex - 1; index >= 0; index -= 1) {
        const headingLevel = paragraphs[index].headingLevel || 0;
        if (headingLevel && headingLevel <= level) {
          if (headingLevel !== level) {
            throw new Error(`This marker is not inside a Heading ${level} section.`);
          }
          start = index;
          break;
        }
      }
      if (start < 0) throw new Error(`This marker has no preceding Heading ${level}.`);
      for (let index = start + 1; index < paragraphs.length; index += 1) {
        const headingLevel = paragraphs[index].headingLevel || 0;
        if (headingLevel && headingLevel <= level) {
          end = index;
          break;
        }
      }
      if (markerIndex < start || markerIndex >= end) {
        throw new Error(`This marker is outside its Heading ${level} section.`);
      }
    }
    const seen = new Set();
    const keys = [];
    for (let index = start; index < end; index += 1) {
      for (const key of paragraphs[index].citationKeys || []) {
        if (!seen.has(key)) {
          seen.add(key);
          keys.push(key);
        }
      }
    }
    return keys;
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
    const people = (fields.author || fields.editor || entry.key).split(/\s+and\s+/i)
      .map((person) => {
        const parts = person.split(',').map((part) => part.trim()).filter(Boolean);
        const given = (value) => value.replace(/\b([A-Z])(?=\s|$)/g, '$1.');
        if (parts.length === 2) return `${given(parts[1])} ${parts[0]}`;
        if (parts.length >= 3) return `${given(parts.slice(2).join(' '))} ${parts[0]}, ${parts[1]}`;
        return person.trim();
      });
    let authors = people.join(', ');
    if (people.length === 2) authors = `${people[0]} and ${people[1]}`;
    if (people.length > 2) authors = `${people.slice(0, -1).join(', ')}, and ${people.at(-1)}`;
    const year = entryYear(entry);
    const title = fields.title || '[Untitled]';
    const container = fields.journaltitle || fields.journal || fields.booktitle || fields.publisher || '';
    let publication = container;
    if (fields.volume) publication += `${publication ? ' ' : ''}${fields.volume}`;
    if (fields.number) publication += `(${fields.number})`;
    if (fields.pages) publication += `${publication ? ': ' : ''}${fields.pages}`;
    return `${authors}. ${year}. ${title}.${publication ? ` ${publication}.` : ''}`;
  }

  return {
    activeKeyFragment,
    bibliographyTag,
    citationKeysForBibliographyScope,
    CITATION_SOURCE_TITLE_PREFIX,
    CITATION_TITLE,
    citationKeysFromControls,
    citationSourceFromTitle,
    citationSourceTitle,
    citationSyntax,
    citationSyntaxFromControl,
    formatCitationPreview,
    formatReference,
    parseBibEntries,
    parseBibliographyTag,
    parseCitationControl,
    parseCitationSyntax,
    parseKeys,
    replaceActiveKey,
    suggestKeys,
  };
}));
