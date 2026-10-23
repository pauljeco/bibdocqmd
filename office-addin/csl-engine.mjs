import { Cite, plugins } from '@citation-js/core';
import '@citation-js/plugin-bibtex';
import '@citation-js/plugin-csl';

export const CSL_STYLES = Object.freeze([
  Object.freeze({ id: 'apa', label: 'APA 7th edition' }),
  Object.freeze({ id: 'harvard1', label: 'Harvard author-date' }),
]);

const STYLE_IDS = new Set(CSL_STYLES.map(({ id }) => id));
const CUSTOM_STYLE_ID = /^custom-[0-9a-f]{12}$/;
const CSL_NAMESPACE = 'http://purl.org/net/xbiblio/csl';

export function validateCslStyle(xml) {
  if (typeof xml !== 'string' || !xml.trim()) throw new Error('The selected CSL style is empty.');
  if (new TextEncoder().encode(xml).length > 2_000_000) {
    throw new Error('The selected CSL style is larger than the 2 MB prototype limit.');
  }
  if (/<!DOCTYPE\b|<!ENTITY\b/i.test(xml)) {
    throw new Error('CSL styles containing document type or entity declarations are not supported.');
  }
  const root = /<style\b[^>]*\bxmlns=["']http:\/\/purl\.org\/net\/xbiblio\/csl["'][^>]*>/i;
  if (!root.test(xml)) throw new Error(`The file is not an independent CSL style in ${CSL_NAMESPACE}.`);
  if (!/<info\b/i.test(xml) || !/<title\b/i.test(xml)) {
    throw new Error('The CSL style has no title metadata.');
  }
  if (/\brel=["']independent-parent["']/i.test(xml)) {
    throw new Error('Dependent CSL styles are not supported; choose the independent parent style.');
  }
  if (!/<citation\b/i.test(xml) || !/<bibliography\b/i.test(xml)) {
    throw new Error('The CSL style must define both citation and bibliography layouts.');
  }
  return xml;
}

export async function sha256Hex(value) {
  const bytes = new TextEncoder().encode(value);
  const digest = await globalThis.crypto.subtle.digest('SHA-256', bytes);
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, '0')).join('');
}

export async function customStyleIdentity(xml) {
  validateCslStyle(xml);
  const hash = await sha256Hex(xml);
  return { id: `custom-${hash.slice(0, 12)}`, hash };
}

export function registerCslStyle(id, xml) {
  if (!CUSTOM_STYLE_ID.test(id)) throw new Error(`Invalid custom CSL style identifier: ${id}`);
  validateCslStyle(xml);
  plugins.config.get('@csl').styles.add(id, xml);
  STYLE_IDS.add(id);
  return id;
}

function assertStyle(style) {
  if (!STYLE_IDS.has(style)) throw new Error(`Unsupported citation style: ${style}`);
}

function normalizeText(value) {
  return String(value).replace(/\s+/g, ' ').trim();
}

export function parseLocator(value) {
  const raw = normalizeText(value || '');
  if (!raw) return {};
  const labels = [
    [/^(?:pp\.?|p\.?)\s*(.+)$/i, 'page'],
    [/^(?:chapter|chap\.?|ch\.?)\s*(.+)$/i, 'chapter'],
    [/^(?:section|sec\.?|§)\s*(.+)$/i, 'section'],
    [/^(?:figure|fig\.?)\s*(.+)$/i, 'figure'],
    [/^(?:paragraph|para\.?|¶)\s*(.+)$/i, 'paragraph'],
    [/^(?:volume|vol\.?)\s*(.+)$/i, 'volume'],
  ];
  for (const [pattern, label] of labels) {
    const match = pattern.exec(raw);
    if (match) return { locator: match[1].trim(), label };
  }
  return { locator: raw, label: 'page' };
}

function citationEntries(keys, locator = '') {
  const entries = keys.map((id) => ({ id }));
  if (entries.length && locator.trim()) Object.assign(entries.at(-1), parseLocator(locator));
  return entries;
}

export class CslLibrary {
  constructor(bibliography, sourceById = new Map(), sourceNames = []) {
    if ((typeof bibliography === 'string' && !bibliography.trim())
        || (Array.isArray(bibliography) && !bibliography.length)) {
      throw new Error('The selected bibliography is empty.');
    }
    this.cite = new Cite(bibliography, { generateGraph: false });
    this.items = this.cite.data;
    if (!this.items.length) throw new Error('No bibliography entries could be parsed.');
    this.itemById = new Map(this.items.map((item) => [String(item.id), item]));
    if (this.itemById.size !== this.items.length) {
      throw new Error('The bibliography contains duplicate citation keys.');
    }
    this.sourceById = new Map(sourceById);
    this.sourceNames = [...sourceNames];
  }

  get size() {
    return this.items.length;
  }

  get keys() {
    return [...this.itemById.keys()];
  }

  sourceForKey(key) {
    return this.sourceById.get(key) || '';
  }

  missingKeys(keys) {
    return keys.filter((key) => !this.itemById.has(key));
  }

  assertKeys(keys) {
    const missing = this.missingKeys(keys);
    if (missing.length) throw new Error(`Keys missing in current library: ${missing.join(', ')}`);
  }

  formatCitation(keys, mode, locator = '', style = 'apa') {
    assertStyle(style);
    if (!keys.length) throw new Error('Choose at least one citation key.');
    this.assertKeys(keys);
    const entries = citationEntries(keys, locator);
    const options = { format: 'text', style, lang: 'en-US' };
    if (mode === 'parenthetical') {
      return normalizeText(this.cite.format('citation', { ...options, entry: entries }));
    }
    if (keys.length !== 1) {
      throw new Error('Narrative and suppressed-author citations require one key.');
    }
    if (mode === 'suppress') {
      entries[0]['suppress-author'] = true;
      return normalizeText(this.cite.format('citation', { ...options, entry: entries }));
    }
    if (mode === 'narrative') {
      const author = normalizeText(this.cite.format('citation', {
        ...options, entry: [{ id: keys[0], 'author-only': true }],
      }));
      entries[0]['suppress-author'] = true;
      const remainder = normalizeText(this.cite.format('citation', { ...options, entry: entries }));
      return `${author} ${remainder}`.trim();
    }
    throw new Error(`Unknown citation mode: ${mode}`);
  }

  formatBibliography(keys, style = 'apa') {
    assertStyle(style);
    if (!keys.length) return [];
    this.assertKeys(keys);
    return this.cite.format('bibliography', {
      format: 'text', style, lang: 'en-US', entry: keys, asEntryArray: true,
    }).map(([id, text]) => ({ id: String(id), text: normalizeText(text) }));
  }

  formatBibliographyHtml(keys, style = 'apa') {
    assertStyle(style);
    if (!keys.length) return '';
    this.assertKeys(keys);
    return this.cite.format('bibliography', {
      format: 'html', style, lang: 'en-US', entry: keys,
    });
  }
}

export function createLibrary(bibtex) {
  return new CslLibrary(bibtex);
}

export function createLibraryFromSources(sources) {
  if (!Array.isArray(sources) || !sources.length) {
    throw new Error('Choose at least one bibliography file.');
  }
  const sourceById = new Map();
  const sourceNames = [];
  for (const source of sources) {
    const name = String(source?.name || '').trim() || 'Unnamed bibliography';
    const contents = source?.contents;
    if (typeof contents !== 'string' || !contents.trim()) {
      throw new Error(`${name} is empty.`);
    }
    let parsed;
    try {
      parsed = new Cite(contents, { generateGraph: false }).data;
    } catch (error) {
      throw new Error(`${name}: ${error.message || String(error)}`);
    }
    const fileKeys = new Set();
    for (const item of parsed) {
      const key = String(item.id);
      if (fileKeys.has(key)) {
        throw new Error(`Duplicate citation key “${key}” occurs more than once in ${name}.`);
      }
      fileKeys.add(key);
      if (sourceById.has(key)) {
        throw new Error(
          `Duplicate citation key “${key}” occurs in ${sourceById.get(key)} and ${name}. `
          + 'Rename one key before combining these libraries.',
        );
      }
      sourceById.set(key, name);
    }
    sourceNames.push(name);
  }
  // Parse the union once more so BibTeX strings and crossref relationships can
  // resolve across file boundaries while source ownership stays explicit.
  const combined = sources.map(({ contents }) => contents).join('\n\n');
  return new CslLibrary(combined, sourceById, sourceNames);
}
