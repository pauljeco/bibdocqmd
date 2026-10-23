'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const {
  activeKeyFragment, bibliographyTag, citationKeysFromControls, citationSyntax,
  formatCitationPreview, formatReference, parseBibEntries, parseBibliographyTag,
  parseCitationSyntax, parseKeys, replaceActiveKey, suggestKeys,
} = require('./citation.js');

test('extracts entry keys and ignores BibTeX directives', () => {
  const keys = parseKeys(`
    @string{journal = "Example"}
    @article{smith2020, title = {Test}}
    @Book ( jones-2022, title = {Another test} )
    @comment{ignored, note = {Not a record}}
  `);
  assert.deepEqual([...keys], ['smith2020', 'jones-2022']);
});

test('builds supported Pandoc citation forms', () => {
  assert.equal(citationSyntax(['smith2020'], 'parenthetical', ''), '[@smith2020]');
  assert.equal(citationSyntax(['smith2020', 'jones2022'], 'parenthetical', ''),
    '[@smith2020; @jones2022]');
  assert.equal(citationSyntax(['smith2020'], 'narrative', 'p. 4'), '@smith2020 [p. 4]');
  assert.equal(citationSyntax(['smith2020'], 'suppress', 'ch. 2'), '[-@smith2020, ch. 2]');
});

test('parses editable citation syntax back into task-pane fields', () => {
  assert.deepEqual(parseCitationSyntax('[@smith2020; @jones2022, pp. 4-6]'), {
    keys: ['smith2020', 'jones2022'], mode: 'parenthetical', locator: 'pp. 4-6',
  });
  assert.deepEqual(parseCitationSyntax('@smith2020 [p. 4]'), {
    keys: ['smith2020'], mode: 'narrative', locator: 'p. 4',
  });
  assert.deepEqual(parseCitationSyntax('[-@jones2022, ch. 2]'), {
    keys: ['jones2022'], mode: 'suppress', locator: 'ch. 2',
  });
  assert.throws(() => parseCitationSyntax('ordinary text'), /supported BibWord/);
});

test('builds a local author-year preview and formatted reference', () => {
  const entries = parseBibEntries(`
    @article{smith2020,
      author = {Ada Smith and B. Jones and C. Brown},
      year = {2020},
      title = {A {Useful} Test},
      journaltitle = {Journal of Examples},
      volume = {12},
      pages = {1--9}
    }
    @book{jones2022, author={Brian Jones}, date={2022-03-01}, title={A Book}, publisher={Example Press}}
  `);
  assert.equal(entries.size, 2);
  assert.equal(formatCitationPreview(['smith2020', 'jones2022'], 'parenthetical', 'p. 5', entries),
    '(Smith et al., 2020; Jones, 2022, p. 5)');
  assert.equal(formatCitationPreview(['smith2020'], 'narrative', '', entries),
    'Smith et al. (2020)');
  assert.match(formatReference(entries.get('smith2020')),
    /^Ada Smith, B\. Jones, and C\. Brown\. 2020\. A Useful Test\. Journal of Examples 12: 1–9\.$/);
});

test('decodes common BibTeX accents in the local preview', () => {
  const entries = parseBibEntries(String.raw`
    @article{accented, author={H\'{e}ctor Abarca and Alejandra Mor\'{a}n-Ordo\~{n}ez}, year={2022}, title={Test}}
  `);
  assert.equal(formatReference(entries.get('accented')),
    'Héctor Abarca and Alejandra Morán-Ordoñez. 2022. Test.');
});

test('decodes BibTeX letter commands and Family, Given author names', () => {
  const entries = parseBibEntries(String.raw`
    @article{geomorphons,
      author={Jasiewicz, Jaros{\l}aw and Stepinski, Tomasz F},
      date={2013},
      title={Geomorphons: a pattern recognition approach to classification and mapping of landforms},
      journaltitle={Geomorphology}, volume={182}, pages={147--156}
    }
  `);
  assert.equal(formatReference(entries.get('geomorphons')),
    'Jarosław Jasiewicz and Tomasz F. Stepinski. 2013. Geomorphons: a pattern recognition approach to classification and mapping of landforms. Geomorphology 182: 147–156.');
});

test('normalizes common BibTeX formatting commands and escaped symbols', () => {
  const entries = parseBibEntries(String.raw`
    @article{commands,
      author={Sonja Kr{\"u}ger and Dud{\'\i}k, Miroslav},
      title={A \textit{formatted} 30\% result},
      url={example\_file}
    }
  `);
  const entry = entries.get('commands');
  assert.equal(entry.fields.author, 'Sonja Krüger and Dudík, Miroslav');
  assert.equal(entry.fields.title, 'A formatted 30% result');
  assert.equal(entry.fields.url, 'example_file');
});

test('suggests keys for the active segment of a grouped citation', () => {
  const keys = new Set(['abarca2022spo-a', 'beier2010uol-a', 'brost2012', 'smith2020']);
  assert.equal(activeKeyFragment('abarca2022spo-a; bei'), 'bei');
  assert.deepEqual(suggestKeys(keys, 'abarca2022spo-a; bei'), ['beier2010uol-a']);
  assert.deepEqual(suggestKeys(keys, 'abarca2022spo-a; 2012'), ['brost2012']);
  assert.deepEqual(suggestKeys(keys, 'abarca2022spo-a; abar'), []);
  assert.equal(replaceActiveKey('abarca2022spo-a; bei', 'beier2010uol-a'),
    'abarca2022spo-a; beier2010uol-a');
});

test('collects unique citation keys from managed Word controls in document order', () => {
  const keys = citationKeysFromControls([
    { tag: 'ordinary', text: '[@ignored]' },
    { tag: 'bibword:cite:v1:first', text: '[@smith2020; @jones2022]' },
    { tag: 'bibword:refs:v1:marker', text: 'References' },
    { tag: 'bibword:cite:v1:second', text: '@smith2020' },
    { tag: 'bibword:cite:v1:third', text: '[-@brown2024, p. 7]' },
  ]);
  assert.deepEqual(keys, ['smith2020', 'jones2022', 'brown2024']);
});

test('versions bibliography tags and preserves legacy marker identities', () => {
  assert.equal(bibliographyTag('marker-123', 'apa'), 'bibword:refs:v2:marker-123:apa');
  assert.deepEqual(parseBibliographyTag('bibword:refs:v2:marker-123:harvard1'), {
    version: 2, id: 'marker-123', style: 'harvard1',
  });
  assert.deepEqual(parseBibliographyTag('bibword:refs:v2:marker-123:custom-0123456789ab'), {
    version: 2, id: 'marker-123', style: 'custom-0123456789ab',
  });
  assert.deepEqual(parseBibliographyTag('bibword:refs:v1:legacy-marker'), {
    version: 1, id: 'legacy-marker', style: null,
  });
  assert.equal(parseBibliographyTag('bibword:cite:v1:not-a-marker'), null);
  assert.throws(() => parseBibliographyTag('bibword:refs:v2:broken'), /invalid style metadata/);
  assert.throws(() => parseBibliographyTag('bibword:refs:v3:future'), /Unsupported/);
});

test('rejects invalid input instead of creating ambiguous syntax', () => {
  assert.throws(() => citationSyntax([], 'parenthetical', ''), /at least one/);
  assert.throws(() => citationSyntax(['not a key'], 'parenthetical', ''), /Invalid citation key/);
  assert.throws(() => citationSyntax(['a', 'b'], 'narrative', ''), /require one key/);
  assert.throws(() => citationSyntax(['a'], 'parenthetical', 'p. 3\rnext'), /line breaks/);
  assert.throws(() => citationSyntax(['a'], 'unknown', ''), /Unknown citation mode/);
});
