import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { runInNewContext } from 'node:vm';

test('the generated CSL bundle exposes a browser-safe global API', async () => {
  const source = await readFile(new URL('dist/csl-engine.js', import.meta.url), 'utf8');
  const browserConsole = { log() {}, warn() {}, error() {} };
  const context = {
    console: browserConsole,
    fetch: globalThis.fetch,
    Headers: globalThis.Headers,
    navigator: { userAgent: 'BibWord bundle test' },
    location: { href: 'https://localhost:3000/taskpane.html' },
    setTimeout,
    clearTimeout,
  };
  context.globalThis = context;
  runInNewContext(source, context);
  const library = context.BibWordCsl.createLibrary(
    '@article{x, author={Doe, Jane}, year={2024}, title={Test}}',
  );
  assert.equal(library.formatCitation(['x'], 'narrative', 'p. 2', 'apa'),
    'Doe (2024, p. 2)');
  assert.match(library.formatBibliographyHtml(['x'], 'apa'), /class="csl-entry"/);
});
