'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { citationSyntax, parseKeys } = require('./citation.js');

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

test('rejects invalid input instead of creating ambiguous syntax', () => {
  assert.throws(() => citationSyntax([], 'parenthetical', ''), /at least one/);
  assert.throws(() => citationSyntax(['not a key'], 'parenthetical', ''), /Invalid citation key/);
  assert.throws(() => citationSyntax(['a', 'b'], 'narrative', ''), /require one key/);
  assert.throws(() => citationSyntax(['a'], 'parenthetical', 'p. 3\rnext'), /line breaks/);
  assert.throws(() => citationSyntax(['a'], 'unknown', ''), /Unknown citation mode/);
});
