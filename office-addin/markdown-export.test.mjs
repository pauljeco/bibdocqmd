import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile, mkdtemp, writeFile, rm, mkdir } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { spawnSync } from 'node:child_process';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { zipSync, strToU8, strFromU8 } from 'fflate';
import { prepareMarkdownExport, finishMarkdownExport, readArchive } from './markdown-export.mjs';
import { customStyleIdentity } from './csl-engine.mjs';

globalThis.DOMParser = DOMParser;
globalThis.XMLSerializer = XMLSerializer;
const assets = {
  styles: { apa: await readFile(new URL('../src/bibdocqmd/styles/apa.csl', import.meta.url), 'utf8') },
  filter: await readFile(new URL('../src/bibdocqmd/filters/multiple_bibliographies.lua', import.meta.url), 'utf8'),
};
const fixture = new Uint8Array(await readFile(new URL('../examples/bibword_demo.docx', import.meta.url)));
const fixtureSources = [{ name: 'references.bib', contents: await readFile(new URL('../examples/references.bib', import.meta.url), 'utf8') }];
const sources = [
  { name: 'one.bib', contents: '@book{Exact-Key, author={Alpha, Ada}, year={2020}, title={First}}' },
  { name: 'two.bib', contents: '@book{Other:Key, author={Zulu, Zoe}, year={2021}, title={Second}}' },
];
const PLAN = 'bibdocqmd-export-plan.json';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xmlEscape = (value) => value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
const p = (value, heading = 0) => `<w:p>${heading ? `<w:pPr><w:pStyle w:val="Heading${heading}"/></w:pPr>` : ''}<w:r><w:t>${xmlEscape(value)}</w:t></w:r></w:p>`;
const cite = (syntax, alias = false) => `<w:sdt><w:sdtPr><w:tag w:val="bibword:cite:v1:reused-id"/>${alias
  ? `<w:alias w:val="bibdocqmd citation source v1: ${xmlEscape(syntax)}"/>` : ''}</w:sdtPr><w:sdtContent><w:r><w:t>${xmlEscape(alias ? '(Formatted author, 2020)' : syntax)}</w:t></w:r></w:sdtContent></w:sdt>`;
const refs = (id = 'refs', scope = 'document', style = 'apa') => `<w:sdt><w:sdtPr><w:tag w:val="bibword:refs:v3:${id}:${style}:${scope}"/></w:sdtPr><w:sdtContent>${p('References')}${p('STALE GENERATED REFERENCES')}</w:sdtContent></w:sdt>`;
function docx(body, extra = {}) {
  return zipSync({ ...readArchive(fixture), ...extra, 'word/document.xml': strToU8(
    `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`) });
}
const planOf = (preparation) => JSON.parse(strFromU8(readArchive(preparation)[PLAN]));
const placeholderMarkdown = (preparation) => planOf(preparation).replacements.map((item) => item.token).join('\n\n');
const prepare = (bytes, overrides = {}) => prepareMarkdownExport(bytes, { sources, assets, ...overrides });
const finish = (preparation, markdown, options = {}) => finishMarkdownExport(preparation,
  typeof markdown === 'string' ? strToU8(markdown) : markdown, { assets, ...options });

test('preparation keeps original bytes, exact citation identities, source files, and unrelated DOCX parts', async () => {
  const before = new Uint8Array(fixture);
  const preparation = await prepare(fixture, { sources: fixtureSources });
  assert.deepEqual(fixture, before);
  const archive = readArchive(preparation);
  const copy = readArchive(archive['word-markdown-copy.docx']);
  for (const [name, bytes] of Object.entries(readArchive(fixture))) {
    if (name !== 'word/document.xml') assert.deepEqual(copy[name], bytes, name);
  }
  assert.equal(planOf(preparation).replacements.length, 5);
  assert.equal(strFromU8(archive['bibliography-1.bib']), fixtureSources[0].contents);
  const xml = strFromU8(copy['word/document.xml']);
  assert.doesNotMatch(xml, /bibword:cite:/);
  assert.equal((xml.match(/BIBDOCQMD/g) || []).length, 5);
});

test('finalization preserves exact grouped, narrative, suppressed and locator syntax from formatted aliases', async () => {
  const syntax = ['[@Exact-Key; @Other:Key, pp. 2–4]', '@Exact-Key [p. 5]', '[-@Other:Key]'];
  const preparation = await prepare(docx(syntax.map((item) => `<w:p>${cite(item, true)}</w:p>`).join('') + refs()));
  const copy = strFromU8(readArchive(readArchive(preparation)['word-markdown-copy.docx'])['word/document.xml']);
  assert.doesNotMatch(copy, /STALE GENERATED REFERENCES|Formatted author/);
  const output = readArchive(await finish(preparation, placeholderMarkdown(preparation)));
  const qmd = strFromU8(output['document.qmd']);
  syntax.forEach((item) => assert.ok(qmd.includes(item), item));
  assert.match(qmd, /::: \{#refs\}/);
  assert.doesNotMatch(qmd, /BIBDOCQMD|STALE|Formatted author/);
  assert.equal(strFromU8(output['bibliography-2.bib']), sources[1].contents);
  assert.match(qmd, /bibliography-1.bib/);
  assert.match(qmd, /bibliography-2.bib/);
});

test('scoped marker identities survive and the established Lua filter is included', async () => {
  const preparation = await prepare(docx(p('Chapter', 1) + `<w:p>${cite('[@Exact-Key]')}</w:p>`
    + refs('chapter', 'heading1') + p('Section', 2) + `<w:p>${cite('@Other:Key')}</w:p>` + refs('section', 'heading2')));
  const output = readArchive(await finish(preparation, placeholderMarkdown(preparation)));
  const qmd = strFromU8(output['document.qmd']);
  assert.match(qmd, /data-bibword-id="chapter" data-bibword-scope="heading1"/);
  assert.match(qmd, /data-bibword-id="section" data-bibword-scope="heading2"/);
  assert.equal(strFromU8(output['bibdocqmd-multiple-bibliographies.lua']), assets.filter);
});

test('custom CSL is recovered from customXML and verified rather than guessed', async () => {
  const xml = assets.styles.apa;
  const { id, hash } = await customStyleIdentity(xml);
  const part = `<bibwordStyle xmlns="urn:bibword:csl-style:v1" id="${id}" hash="${hash}"><csl encoding="base64">${Buffer.from(xml).toString('base64')}</csl></bibwordStyle>`;
  const input = docx(`<w:p>${cite('[@Exact-Key]')}</w:p>` + refs('custom', 'document', id),
    { 'customXML/item9.xml': strToU8(part) });
  const preparation = await prepare(input);
  assert.equal(strFromU8(readArchive(preparation)[`bibdocqmd-${id}.csl`]), xml);
  const corrupt = readArchive(input);
  corrupt['customXML/item9.xml'] = strToU8(part.replace(hash, '0'.repeat(64)));
  await assert.rejects(prepare(zipSync(corrupt)), /integrity/);
});

test('Word Markdown and images in a ZIP retain their relative paths and exact bytes', async () => {
  const preparation = await prepare(docx(`<w:p>${cite('[@Exact-Key]')}</w:p>`));
  const image = new Uint8Array([1, 2, 3]);
  const exported = zipSync({ 'report/body.md': strToU8(placeholderMarkdown(preparation) + '\n\n![Figure](media/image.png)'),
    'report/media/image.png': image });
  const output = readArchive(await finish(preparation, exported, { markdownZip: true }));
  assert.deepEqual(output['report/media/image.png'], image);
  assert.ok(output['report/document.qmd']);
  assert.ok(output['report/bibliography-1.bib']);
});

test('missing, duplicate, wrong-document, and code-formatted placeholders block export', async () => {
  const preparation = await prepare(docx(`<w:p>${cite('[@Exact-Key]')}</w:p>` + refs()));
  const markdown = placeholderMarkdown(preparation);
  const token = planOf(preparation).replacements[0].token;
  for (const value of [markdown.replace(token, ''), markdown + '\n' + token,
    markdown + '\nBIBDOCQMD' + '0'.repeat(32) + 'C000001END',
    markdown.replace(token, '`' + token + '`'), '```\n' + markdown + '\n```']) {
    await assert.rejects(finish(preparation, value), /placeholder|different preparation|inside code/);
  }
});

test('corrupted sources, duplicate keys, unresolved citations, scopes and unsupported locations are explicit errors', async () => {
  await assert.rejects(prepare(docx(`<w:p>${cite('[@Missing]')}</w:p>`)), /Missing/);
  await assert.rejects(prepare(docx(`<w:p>${cite('[@Exact-Key]')}</w:p>`), { sources: [sources[0], sources[0]] }), /Duplicate/);
  await assert.rejects(prepare(docx(`<w:p>${cite('[@Exact-Key]')}</w:p>` + refs('bad', 'heading1'))), /Heading 1/);
  await assert.rejects(prepare(docx(`<w:p>${cite('[@Exact-Key]')}</w:p>` + refs() + refs())), /copied/);
  await assert.rejects(prepare(docx(`<w:p>${cite('[@Exact-Key]')}</w:p>`,
    { 'word/footnotes.xml': strToU8(`<w:footnotes xmlns:w="${W}">${cite('[@Exact-Key]')}</w:footnotes>`) })), /outside the main body/);
  const preparation = await prepare(docx(`<w:p>${cite('[@Exact-Key]')}</w:p>`));
  const corrupt = readArchive(preparation);
  corrupt['bibliography-1.bib'] = strToU8(sources[0].contents + ' changed');
  await assert.rejects(finish(zipSync(corrupt), placeholderMarkdown(preparation)), /integrity/);
  const corruptPlan = readArchive(preparation);
  corruptPlan[PLAN] = strToU8(strFromU8(corruptPlan[PLAN]).replace('My document', 'Altered title'));
  await assert.rejects(finish(zipSync(corruptPlan), placeholderMarkdown(preparation)), /plan failed its integrity/);
});

test('archive traversal and XML entity declarations are rejected', async () => {
  assert.throws(() => readArchive(zipSync({ '../escape': strToU8('x') })), /Unsafe/);
  const files = readArchive(fixture);
  files['word/document.xml'] = strToU8('<!DOCTYPE doc [<!ENTITY x "x">]><doc/>');
  await assert.rejects(prepare(zipSync(files)), /entity declarations/);
});

// Pandoc is a stand-in for the native Word Markdown writer, not a real-host
// test. It checks the complete prepared-DOCX -> Markdown -> QMD -> Cite AST path.
const candidates = [['pandoc', []], ['quarto', ['pandoc']]];
const pandoc = candidates.find(([command, prefix]) => spawnSync(command, [...prefix, '--version']).status === 0);
test('fixture integration: DOCX through Pandoc Markdown to QMD retains Cite nodes and renders references',
  { skip: !pandoc && 'Pandoc/Quarto not installed' }, async () => {
    const directory = await mkdtemp(join(tmpdir(), 'bibdocqmd-markdown-'));
    const run = (args) => {
      const result = spawnSync(pandoc[0], [...pandoc[1], ...args], { cwd: directory, encoding: 'utf8' });
      assert.equal(result.status, 0, result.stderr);
      return result.stdout;
    };
    try {
      const preparation = await prepare(fixture, { sources: fixtureSources });
      await writeFile(join(directory, 'prepared.docx'), readArchive(preparation)['word-markdown-copy.docx']);
      const markdown = run(['prepared.docx', '-t', 'gfm']);
      const output = readArchive(await finish(preparation, markdown));
      for (const [name, bytes] of Object.entries(output)) {
        await mkdir(dirname(join(directory, name)), { recursive: true });
        await writeFile(join(directory, name), bytes);
      }
      const ast = JSON.parse(run(['document.qmd', '-f', 'markdown', '-t', 'json']));
      const modes = [];
      function walk(node) {
        if (node && typeof node === 'object') {
          if (node.t === 'Cite') modes.push(node.c[0][0].citationMode.t);
          Object.values(node).forEach(walk);
        }
      }
      walk(ast);
      assert.equal(modes.length, 5);
      assert.ok(modes.includes('AuthorInText'));
      assert.ok(modes.includes('SuppressAuthor'));
      const html = run(['document.qmd', '--citeproc', '-t', 'html']);
      assert.match(html, /id="refs"/);
      assert.doesNotMatch(html, /BIBDOCQMD/);
    } finally { await rm(directory, { recursive: true, force: true }); }
  });

test('fixture integration: scoped QMD renders each bibliography once with the original Lua filter',
  { skip: !pandoc && 'Pandoc/Quarto not installed' }, async () => {
    const directory = await mkdtemp(join(tmpdir(), 'bibdocqmd-scoped-markdown-'));
    const run = (args) => {
      const result = spawnSync(pandoc[0], [...pandoc[1], ...args], { cwd: directory, encoding: 'utf8' });
      assert.equal(result.status, 0, result.stderr);
      return result.stdout;
    };
    try {
      const preparation = await prepare(docx(p('First chapter', 1) + `<w:p>${cite('[@Exact-Key]')}</w:p>`
        + refs('first', 'heading1') + p('Second chapter', 1) + `<w:p>${cite('[@Other:Key]')}</w:p>`
        + refs('second', 'heading1')));
      await writeFile(join(directory, 'prepared.docx'), readArchive(preparation)['word-markdown-copy.docx']);
      const output = readArchive(await finish(preparation, run(['prepared.docx', '-t', 'gfm'])));
      for (const [name, bytes] of Object.entries(output)) await writeFile(join(directory, name), bytes);
      const html = run(['document.qmd', '--lua-filter=bibdocqmd-multiple-bibliographies.lua', '-t', 'html']);
      assert.equal((html.match(/id="bibword-refs-first"/g) || []).length, 1);
      assert.equal((html.match(/id="bibword-refs-second"/g) || []).length, 1);
      assert.doesNotMatch(html, /STALE GENERATED|BIBDOCQMD/);
      assert.match(html, /id="bibword-refs-first-ref-Exact-Key"/);
      assert.match(html, /id="bibword-refs-second-ref-Other:Key"/);
    } finally { await rm(directory, { recursive: true, force: true }); }
  });
