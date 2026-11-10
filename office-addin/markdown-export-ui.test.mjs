import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { runInNewContext } from 'node:vm';
import { createRequire } from 'node:module';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { strToU8, strFromU8 } from 'fflate';
import { prepareMarkdownExport, finishMarkdownExport, readArchive } from './markdown-export.mjs';

globalThis.DOMParser = DOMParser;
globalThis.XMLSerializer = XMLSerializer;
const citation = createRequire(import.meta.url)('./citation.js');
const script = await readFile(new URL('taskpane.js', import.meta.url), 'utf8');
const html = await readFile(new URL('taskpane.html', import.meta.url), 'utf8');
const assets = { styles: { apa: await readFile(new URL('../src/bibdocqmd/styles/apa.csl', import.meta.url), 'utf8') } };
const docx = new Uint8Array(await readFile(new URL('../examples/bibword_demo.docx', import.meta.url)));
const bibliography = { name: 'references.bib', contents: await readFile(new URL('../examples/references.bib', import.meta.url), 'utf8') };
const file = (name, bytes) => ({ name, size: bytes.length,
  arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) });

function harness() {
  const downloads = [];
  const nodes = new Map();
  const sandbox = {
    BibWordCitation: citation,
    BibWordCsl: { CSL_STYLES: [{ id: 'apa' }] },
    BibWordGithubLibrary: {},
    BibWordMarkdown: {
      prepareMarkdownExport: (bytes, options) => prepareMarkdownExport(bytes, { ...options, assets }),
      finishMarkdownExport: (preparation, markdown, options) => finishMarkdownExport(preparation, markdown, { ...options, assets }),
    },
    document: { getElementById(id) {
      if (!nodes.has(id)) nodes.set(id, { value: 'apa', files: [] });
      return nodes.get(id);
    } },
    fixtureSources: [bibliography],
    collectDownload: (bytes, name) => downloads.push({ bytes, name }),
    Word: { run: () => assert.fail('Export must not change the open Word document.') },
    Uint8Array,
  };
  runInNewContext(script.slice(0, script.indexOf("  $('bib-file').addEventListener")) + `
    activeLibrarySources = new Map(fixtureSources.map(source => [source.name, source]));
    downloadExport = collectDownload;
    globalThis.runExport = runMarkdownExport;
  })();`, sandbox);
  return { nodes, sandbox, downloads, get: (id) => sandbox.document.getElementById(id) };
}

test('pane preparation and finalization download valid local packages without editing Word', async () => {
  const { sandbox, downloads, get } = harness();
  get('export-docx').files = [file('My report.docx', docx)];
  await sandbox.runExport('prepare');
  assert.equal(downloads.length, 1, get('quarto-export-status').textContent);
  assert.equal(downloads[0].name, 'bibdocqmd-preparation.zip');
  const preparation = downloads[0].bytes;
  const plan = JSON.parse(strFromU8(readArchive(preparation)['bibdocqmd-export-plan.json']));
  get('export-preparation').files = [file('preparation.zip', preparation)];
  get('export-markdown').files = [file('report.md', strToU8(plan.replacements.map((item) => item.token).join('\n\n')))];
  await sandbox.runExport('finish');
  assert.equal(downloads.length, 2, get('quarto-export-status').textContent);
  assert.equal(downloads[1].name, 'bibdocqmd-quarto.zip');
  assert.ok(readArchive(downloads[1].bytes)['document.qmd']);
  assert.equal(get('prepare-markdown-export').disabled, false);
  assert.equal(get('finish-markdown-export').disabled, false);
});

test('pane reports incomplete inputs without downloading or leaving buttons disabled', async () => {
  const { sandbox, downloads, get } = harness();
  await sandbox.runExport('prepare');
  assert.match(get('quarto-export-status').textContent, /Choose your saved original DOCX/);
  await sandbox.runExport('finish');
  assert.match(get('quarto-export-status').textContent, /Choose the preparation ZIP/);
  assert.equal(downloads.length, 0);
  assert.equal(get('finish-markdown-export').disabled, false);
  for (const id of ['export-docx', 'export-preparation', 'export-markdown', 'prepare-markdown-export',
    'finish-markdown-export', 'quarto-export-status']) assert.ok(html.includes(`id="${id}"`));
});
