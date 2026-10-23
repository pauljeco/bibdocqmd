import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import {
  CSL_STYLES, createLibrary, createLibraryFromSources, customStyleIdentity, parseLocator,
  registerCslStyle, validateCslStyle,
} from './csl-engine.mjs';

const BIBLIOGRAPHY = String.raw`
@article{smith2020,
  author={Ada Smith and B. Jones and C. Brown},
  year={2020},
  title={A Useful Test},
  journaltitle={Journal of Examples},
  volume={12},
  pages={1--9}
}
@article{geomorphons,
  author={Jasiewicz, Jaros{\l}aw and Stepinski, Tomasz F},
  date={2013},
  title={Geomorphons: a pattern recognition approach to classification and mapping of landforms},
  journaltitle={Geomorphology}, volume={182}, pages={147--156}
}`;

test('parses BibLaTeX into CSL data without changing citation keys', () => {
  const library = createLibrary(BIBLIOGRAPHY);
  assert.equal(library.size, 2);
  assert.deepEqual(library.keys, ['smith2020', 'geomorphons']);
  assert.equal(library.itemById.get('geomorphons').author[0].given, 'Jarosław');
});

test('combines several named bibliography files into one searchable library', () => {
  const library = createLibraryFromSources([
    { name: 'articles.bib', contents: '@article{article-key, title={Article}}' },
    { name: 'reports.bib', contents: '@report{report-key, title={Report}}' },
  ]);
  assert.equal(library.size, 2);
  assert.deepEqual(library.keys, ['article-key', 'report-key']);
  assert.deepEqual(library.sourceNames, ['articles.bib', 'reports.bib']);
  assert.equal(library.sourceForKey('article-key'), 'articles.bib');
  assert.equal(library.sourceForKey('report-key'), 'reports.bib');
});

test('resolves BibTeX strings across selected file boundaries', () => {
  const library = createLibraryFromSources([
    {
      name: 'strings-and-articles.bib',
      contents: '@string{sharedjournal="Combined Journal"}\n@article{first, title={First}}',
    },
    {
      name: 'more-articles.bib',
      contents: '@article{second, title={Second}, journaltitle=sharedjournal, year={2024}}',
    },
  ]);
  assert.equal(library.itemById.get('second')['container-title'], 'Combined Journal');
});

test('rejects duplicate keys within or across bibliography files', () => {
  assert.throws(() => createLibraryFromSources([
    { name: 'first.bib', contents: '@article{shared, title={First}}' },
    { name: 'second.bib', contents: '@article{shared, title={Second}}' },
  ]), /Duplicate citation key.*shared.*first\.bib.*second\.bib/);
  assert.throws(() => createLibraryFromSources([{
    name: 'duplicate.bib',
    contents: '@article{same, title={First}}\n@article{same, title={Second}}',
  }]), /Duplicate citation key.*same.*duplicate\.bib/);
});

test('formats citation modes and locators with citeproc-js', () => {
  const library = createLibrary(BIBLIOGRAPHY);
  assert.equal(library.formatCitation(['smith2020'], 'parenthetical', 'p. 5', 'apa'),
    '(Smith et al., 2020, p. 5)');
  assert.equal(library.formatCitation(['smith2020'], 'narrative', 'p. 5', 'apa'),
    'Smith et al. (2020, p. 5)');
  assert.equal(library.formatCitation(['smith2020'], 'suppress', 'p. 5', 'apa'),
    '(2020, p. 5)');
});

test('formats and style-sorts bibliography entries', () => {
  const library = createLibrary(BIBLIOGRAPHY);
  const entries = library.formatBibliography(['smith2020', 'geomorphons'], 'apa');
  assert.deepEqual(entries.map(({ id }) => id), ['geomorphons', 'smith2020']);
  assert.match(entries[0].text, /^Jasiewicz, J\., & Stepinski, T\. F\. \(2013\)\./);
  assert.match(entries[0].text, /Geomorphology, 182, 147–156/);
  const html = library.formatBibliographyHtml(['geomorphons'], 'apa');
  assert.match(html, /class="csl-bib-body"/);
  assert.match(html, /<i>Geomorphology<\/i>/);
});

test('exposes supported author-date styles and parses common locator labels', () => {
  assert.deepEqual(CSL_STYLES.map(({ id }) => id), ['apa', 'harvard1']);
  assert.deepEqual(parseLocator('pp. 45–47'), { locator: '45–47', label: 'page' });
  assert.deepEqual(parseLocator('ch. 2'), { locator: '2', label: 'chapter' });
  assert.deepEqual(parseLocator('figure 3'), { locator: '3', label: 'figure' });
});

test('fails explicitly for unresolved keys and unsupported styles', () => {
  const library = createLibrary(BIBLIOGRAPHY);
  assert.throws(() => library.formatCitation(['missing'], 'parenthetical'), /Keys missing/);
  assert.throws(() => library.formatBibliography(['smith2020'], 'unknown'), /Unsupported/);
});

test('registers an independent custom CSL style under a stable content identity', async () => {
  const xml = await readFile(new URL('../src/bibword/styles/apa.csl', import.meta.url), 'utf8');
  const first = await customStyleIdentity(xml);
  const second = await customStyleIdentity(xml);
  assert.match(first.id, /^custom-[0-9a-f]{12}$/);
  assert.equal(first.hash.length, 64);
  assert.deepEqual(first, second);
  assert.equal(registerCslStyle(first.id, xml), first.id);

  const library = createLibrary(BIBLIOGRAPHY);
  assert.equal(library.formatCitation(['smith2020'], 'parenthetical', 'p. 5', first.id),
    '(Smith et al., 2020, p. 5)');
});

test('rejects malformed and dependent CSL styles before registration', () => {
  assert.throws(() => validateCslStyle('<style/>'), /not an independent CSL style/);
  const dependent = `
    <style xmlns="http://purl.org/net/xbiblio/csl" version="1.0">
      <info><title>Dependent</title><link rel="independent-parent" href="parent"/></info>
      <citation><layout><text variable="title"/></layout></citation>
      <bibliography><layout><text variable="title"/></layout></bibliography>
    </style>`;
  assert.throws(() => validateCslStyle(dependent), /Dependent CSL styles/);
  assert.throws(() => registerCslStyle('custom-bad', dependent), /Invalid custom/);
});
