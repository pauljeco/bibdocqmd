import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { runInNewContext } from 'node:vm';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { unzipSync, strFromU8, strToU8 } from 'fflate';

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
  const combined = context.BibWordCsl.createLibraryFromSources([
    { name: 'first.bib', contents: '@article{first, title={First}}' },
    { name: 'second.bib', contents: '@article{second, title={Second}}' },
  ]);
  assert.deepEqual([...combined.keys], ['first', 'second']);
  assert.equal(combined.sourceForKey('second'), 'second.bib');
});

test('the generated Markdown export bundle prepares and finalizes locally in a browser-like context', async () => {
  const source = await readFile(new URL('dist/markdown-export.js', import.meta.url), 'utf8');
  assert.match(source, /fflate 0\.8\.3/);
  assert.match(source, /MIT License/);
  const context = {
    console: { log() {}, warn() {}, error() {} },
    crypto: globalThis.crypto, TextEncoder, TextDecoder, Uint8Array, atob,
    DOMParser, XMLSerializer, setTimeout, clearTimeout,
    fetch() { assert.fail('Export must not upload documents or fetch bibliography content.'); },
  };
  runInNewContext(source, context);
  const bytes = new Uint8Array(await readFile(new URL('../examples/bibword_demo.docx', import.meta.url)));
  const contents = await readFile(new URL('../examples/references.bib', import.meta.url), 'utf8');
  const preparation = await context.BibWordMarkdown.prepareMarkdownExport(bytes,
    { sources: [{ name: 'references.bib', contents }] });
  const plan = JSON.parse(strFromU8(unzipSync(preparation)['bibdocqmd-export-plan.json']));
  const markdown = plan.replacements.map((item) => item.token).join('\n\n');
  const result = await context.BibWordMarkdown.finishMarkdownExport(preparation, strToU8(markdown));
  const files = unzipSync(result);
  assert.match(strFromU8(files['document.qmd']), /@smith2020/);
  assert.equal(strFromU8(files['bibliography-1.bib']), contents);
});
