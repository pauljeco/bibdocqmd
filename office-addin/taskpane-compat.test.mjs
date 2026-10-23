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

test('insertion leaves the Word cursor outside managed controls', async () => {
  const source = await readFile(new URL('taskpane.js', import.meta.url), 'utf8');
  assert.match(source, /control\.getRange\(Word\.RangeLocation\.after\)\.select\(\)/);
  assert.match(source, /control\.insertParagraph\('', Word\.InsertLocation\.after\)/);
  assert.match(source, /exitParagraph\.select\(Word\.SelectionMode\.start\)/);
});

test('bibliography updates repair a missing final exit paragraph once', async () => {
  const source = await readFile(new URL('taskpane.js', import.meta.url), 'utf8');
  assert.match(source, /markerLastParagraphIndexes/);
  assert.match(source, /markerHasFollowingParagraph/);
  assert.match(source, /!paragraphState\.markerHasFollowingParagraph\.get\(metadata\.id\)/);
  assert.match(source, /marker\.insertParagraph\('', Word\.InsertLocation\.after\)/);
});

test('the picker combines several local bibliography files and labels key sources', async () => {
  const html = await readFile(new URL('taskpane.html', import.meta.url), 'utf8');
  const source = await readFile(new URL('taskpane.js', import.meta.url), 'utf8');
  assert.match(html, /id="bib-file"[^>]*\bmultiple\b/);
  assert.match(html, /id="clear-bibliographies"/);
  assert.match(source, /createLibraryFromSources\(sources\)/);
  assert.match(source, /const combinedSources = new Map\(activeLibrarySources\)/);
  assert.match(source, /cslLibrary\.sourceForKey\(key\)/);
  assert.match(source, /librarySourceNamespace/);
  assert.match(source, /customXmlParts\.getByNamespace\(librarySourceNamespace\)/);
});
