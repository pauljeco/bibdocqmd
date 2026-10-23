import { Cite } from '@citation-js/core';
import '@citation-js/plugin-bibtex';
import '@citation-js/plugin-csl';

export const CSL_STYLES = Object.freeze([
  Object.freeze({ id: 'apa', label: 'APA 7th edition' }),
  Object.freeze({ id: 'harvard1', label: 'Harvard author-date' }),
]);

const STYLE_IDS = new Set(CSL_STYLES.map(({ id }) => id));

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
  constructor(bibtex) {
    if (typeof bibtex !== 'string' || !bibtex.trim()) {
      throw new Error('The selected bibliography is empty.');
    }
    this.cite = new Cite(bibtex, { generateGraph: false });
    this.items = this.cite.data;
    if (!this.items.length) throw new Error('No bibliography entries could be parsed.');
    this.itemById = new Map(this.items.map((item) => [String(item.id), item]));
    if (this.itemById.size !== this.items.length) {
      throw new Error('The bibliography contains duplicate citation keys.');
    }
  }

  get size() {
    return this.items.length;
  }

  get keys() {
    return [...this.itemById.keys()];
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
}

export function createLibrary(bibtex) {
  return new CslLibrary(bibtex);
}
