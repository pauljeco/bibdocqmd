'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const {
  activeKeyFragment, citationSyntax, parseKeys, replaceActiveKey, suggestKeys,
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

test('suggests keys for the active segment of a grouped citation', () => {
  const keys = new Set(['abarca2022spo-a', 'beier2010uol-a', 'brost2012', 'smith2020']);
  assert.equal(activeKeyFragment('abarca2022spo-a; bei'), 'bei');
  assert.deepEqual(suggestKeys(keys, 'abarca2022spo-a; bei'), ['beier2010uol-a']);
  assert.deepEqual(suggestKeys(keys, 'abarca2022spo-a; 2012'), ['brost2012']);
  assert.deepEqual(suggestKeys(keys, 'abarca2022spo-a; abar'), []);
  assert.equal(replaceActiveKey('abarca2022spo-a; bei', 'beier2010uol-a'),
    'abarca2022spo-a; beier2010uol-a');
});

test('rejects invalid input instead of creating ambiguous syntax', () => {
  assert.throws(() => citationSyntax([], 'parenthetical', ''), /at least one/);
  assert.throws(() => citationSyntax(['not a key'], 'parenthetical', ''), /Invalid citation key/);
  assert.throws(() => citationSyntax(['a', 'b'], 'narrative', ''), /require one key/);
  assert.throws(() => citationSyntax(['a'], 'parenthetical', 'p. 3\rnext'), /line breaks/);
  assert.throws(() => citationSyntax(['a'], 'unknown', ''), /Unknown citation mode/);
});
