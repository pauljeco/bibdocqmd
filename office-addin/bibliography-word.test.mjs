import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import vm from 'node:vm';
import { createLibrary } from './csl-engine.mjs';

const require = createRequire(import.meta.url);
const citation = require('./citation.js');
const source = await readFile(new URL('taskpane.js', import.meta.url), 'utf8');
const library = createLibrary(`
  @report{first, author={Alpha, Ada}, year={2025}, title={First title}}
  @report{last, author={Zulu, Zoe}, year={2025}, title={Last title}}
`);

// Simulated Office.js batches, not a Microsoft Word compatibility test. Keep
// insertion and styling in the real update handler; stub scope discovery/UI.
async function createHarness(paragraphApi) {
  const pending = [];
  const elements = new Map();
  const plans = new Map();
  const markers = ['document', 'heading1', 'heading2', 'heading3'].map((scope, index) => {
    const id = `marker-${index}`;
    const tag = citation.bibliographyTag(id, 'apa', scope);
    const metadata = citation.parseBibliographyTag(tag);
    plans.set(id, { keys: ['first', 'last'], metadata, title: 'References' });
    let items = [];
    const paragraphs = {
      load() { pending.push(() => {}); },
      get items() { return items; },
    };
    return {
      tag, title: 'bibdocqmd bibliography position', text: 'References', paragraphs,
      insertHtml(html, location) {
        assert.equal(location, 'Replace');
        this.html = html;
        items = []; // Imported paragraphs only become available after sync.
        pending.push(() => {
          // Derive paragraph boundaries from the actual generated HTML. The
          // first entry must be a sibling of the heading and later entries.
          assert.doesNotMatch(html, /<\/?div\b/i);
          const imported = [...html.matchAll(/<(h[1-9]|p)\b[^>]*>[\s\S]*?<\/\1>/g)];
          assert.equal(imported.length, 1 + plans.get(id).keys.length);
          items = imported.map((_, index) => {
            let style = 'Heading1';
            return {
              leftIndent: index === 1 ? 72 : 36, firstLineIndent: -36,
              rightIndent: 12, alignment: 'Centered',
              get styleBuiltIn() { return style; },
              set styleBuiltIn(value) {
                style = value;
                // Simulate Word replacing HTML-derived paragraph indents
                // when a native paragraph style is applied.
                pending.push(() => {
                  this.leftIndent = 0;
                  this.firstLineIndent = 0;
                });
              },
            };
          });
        });
      },
      insertParagraph() { assert.fail('An existing exit paragraph should be reused.'); },
    };
  });
  const context = {
    document: {
      contentControls: { items: markers, load() {} },
      body: { getOoxml: () => ({ value: '<fixture/>' }) },
    },
    async sync() { pending.splice(0).forEach((operation) => operation()); },
  };
  const sandbox = {
    BibWordCitation: citation,
    BibWordCsl: { CSL_STYLES: [{ id: 'apa' }] },
    BibWordGithubLibrary: {},
    document: {
      getElementById(id) {
        if (!elements.has(id)) elements.set(id, {
          value: id === 'bibliography-indent' ? '1.27' : 'apa',
          selectedOptions: [{ textContent: 'APA' }], textContent: '',
        });
        return elements.get(id);
      },
    },
    Office: { context: { requirements: { isSetSupported: () => paragraphApi } } },
    Word: { run: async (callback) => callback(context), InsertLocation: { replace: 'Replace' } },
    fixturePlans: plans,
    fixtureLibrary: library,
  };
  vm.runInNewContext(source.slice(0, source.indexOf("  $('bib-file').addEventListener")) + `
    officeReady = true;
    cslLibrary = fixtureLibrary;
    bibliographyPlansFromWordParagraphs = async () => ({
      plans: fixturePlans,
      markerHasFollowingParagraph: new Map([...fixturePlans.keys()].map(id => [id, true])),
    });
    bibliographyPlansFromOoxml = () => fixturePlans;
    renderDocumentBibliography = () => {};
    globalThis.updateBibliography = formatBibliographyInWord;
    globalThis.wordBibliographyHtml = wordBibliographyHtml;
  })();`, sandbox);
  return { sandbox, markers, plans, elements };
}

for (const paragraphApi of [true, false]) {
  test(`simulated bibliography update resets headings and restores hanging indents (${paragraphApi ? 'paragraph' : 'OOXML'} scope reader)`, async () => {
    const { sandbox, markers, plans, elements } = await createHarness(paragraphApi);
    const tags = markers.map(({ tag }) => tag);
    for (const [centimetres, keys] of [
      [1.27, ['first', 'last']], [0.75, ['first', 'last']],
      [0, ['first', 'last']], [2.54, ['last']], [5, ['first', 'last']],
    ]) {
      sandbox.document.getElementById('bibliography-indent').value = String(centimetres);
      const points = centimetres * 72 / 2.54;
      plans.forEach((plan) => { plan.keys = keys; });
      await sandbox.updateBibliography();
      assert.notEqual(elements.get('status').className, 'error', elements.get('status').textContent);
      markers.forEach((marker, index) => {
        assert.deepEqual(marker.paragraphs.items.map(({ styleBuiltIn }) => styleBuiltIn),
          [`Heading${index + 1}`, ...keys.map(() => 'Normal')]);
        assert.deepEqual(marker.paragraphs.items.map(({ leftIndent, firstLineIndent }) =>
          [leftIndent, firstLineIndent]), [[0, 0], ...keys.map(() => [points, points === 0 ? 0 : -points])]);
        marker.paragraphs.items.forEach((paragraph, paragraphIndex) => {
          assert.equal(paragraph.rightIndent, 0);
          if (paragraphIndex > 0) {
            assert.equal(paragraph.alignment, 'Left');
            assert.equal(paragraph.leftIndent + paragraph.firstLineIndent, 0,
              'Every reference first line must start at the same margin.');
          }
        });
        assert.equal(marker.tag, tags[index]);
        assert.match(marker.html, /<i>Last title<\/i>/);
        assert.doesNotMatch(marker.html, /margin-left|text-indent:-/);
        assert.equal((marker.html.match(/<p class="csl-entry"/g) || []).length, keys.length);
      });
    }
  });
}

test('invalid hanging indents are rejected before changing Word content', async () => {
  const { sandbox, markers, elements } = await createHarness(true);
  for (const value of ['', ' ', '-1', '5.01', 'Infinity', 'not-a-number']) {
    sandbox.document.getElementById('bibliography-indent').value = value;
    await sandbox.updateBibliography();
    assert.equal(elements.get('status').className, 'error');
    assert.match(elements.get('status').textContent, /between 0 and 5 cm/);
    markers.forEach((marker) => assert.equal(marker.html, undefined));
  }
});

test('Word bibliography markup preserves nested CSL inline content without block offsets', async () => {
  const { sandbox } = await createHarness(true);
  const html = sandbox.wordBibliographyHtml('References & sources',
    '<div class="csl-bib-body"><div class="csl-entry">'
    + '<div class="csl-left-margin">1.</div><div class="csl-right-inline">'
    + '<i>Title</i> <a href="https://example.org/">Link</a></div></div>'
    + '<div class="csl-entry">Second &amp; third</div></div>', 2);
  assert.match(html, /^<h2>References &amp; sources<\/h2>\s*<p/);
  assert.doesNotMatch(html, /<\/?div\b/i);
  assert.match(html, /<span class="csl-left-margin">1\.<\/span> <span/);
  assert.match(html, /<i>Title<\/i> <a href="https:\/\/example.org\/">Link<\/a>/);
  assert.match(html, /<\/p><p[^>]*>Second &amp; third<\/p>$/);
});
