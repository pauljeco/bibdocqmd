import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

test('the Word web task pane avoids the unavailable Font.reset API', async () => {
  const source = await readFile(new URL('taskpane.js', import.meta.url), 'utf8');
  assert.doesNotMatch(source, /\.font\.reset\s*\(/);
  assert.match(source, /contentControls\.getById\(selectedControlId\)/);
});

test('scoped bibliographies use paragraph controls with a gated OOXML fallback', async () => {
  const source = await readFile(new URL('taskpane.js', import.meta.url), 'utf8');
  assert.match(source, /isSetSupported\('WordApi', '1\.3'\)/);
  assert.match(source, /paragraph\.parentContentControlOrNullObject/);
  assert.match(source, /paragraph\.contentControls/);
  assert.match(source, /bibliographyPlansFromOoxml\(bodyOoxml\.value\)/);
});
