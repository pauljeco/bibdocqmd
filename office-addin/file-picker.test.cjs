'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const { join } = require('node:path');
const { runInNewContext } = require('node:vm');
const { bindFilePickers, createFileReceiver } = require('./file-picker.js');

const fileIds = ['bib-file', 'csl-file', 'export-docx', 'export-preparation', 'export-markdown'];
function fixture() {
  const nodes = new Map();
  function node(id) {
    if (!nodes.has(id)) nodes.set(id, {
      id, hidden: true, textContent: '', events: {},
      addEventListener(type, listener) { this.events[type] = listener; },
    });
    return nodes.get(id);
  }
  const document = {
    baseURI: 'https://localhost:3000/taskpane.html',
    getElementById: node,
    querySelectorAll: () => fileIds.map(node),
  };
  return { document, node, click(id) {
    const event = { defaultPrevented: false, preventDefault() { this.defaultPrevented = true; } };
    node(id).events.click(event);
    return event;
  } };
}

test('all five file inputs open their picker synchronously and suppress a duplicate native dialog', () => {
  const { document, node, click } = fixture();
  const calls = [];
  fileIds.forEach((id) => { node(id).showPicker = () => calls.push(id); });
  bindFilePickers(document);
  fileIds.forEach((id) => {
    const event = click(id);
    assert.equal(calls.at(-1), id, 'Opening must happen in the click, without awaiting Office.');
    assert.equal(event.defaultPrevented, true);
  });
});

test('desktop Word routes selection through an Office dialog and returns an intact FileList', async () => {
  const { document, node, click } = fixture();
  const handlers = {};
  let request;
  let closed = 0;
  const office = {
    context: { host: 'Word', platform: 'PC', ui: {
      displayDialogAsync(url, options, callback) {
        request = { url: new URL(url), options };
        callback({ status: 'succeeded', value: {
          addEventHandler: (type, handler) => { handlers[type] = handler; },
          close: () => { closed += 1; },
        } });
      },
    } },
    AsyncResultStatus: { Succeeded: 'succeeded' },
    EventType: { DialogMessageReceived: 'message', DialogEventReceived: 'event' },
  };
  const previous = globalThis.DataTransfer;
  globalThis.DataTransfer = class {
    constructor() { this.files = []; this.items = { add: (file) => this.files.push(file) }; }
  };
  try {
    node('export-docx').showPicker = () => assert.fail('Desktop must not use the pane native chooser.');
    let changed = 0;
    node('export-docx').dispatchEvent = (event) => { assert.equal(event.type, 'change'); changed += 1; };
    bindFilePickers(document, { office });
    assert.equal(click('export-docx').defaultPrevented, true);
    assert.equal(request.url.pathname, '/file-picker-dialog.html');
    assert.equal(request.options.displayInIframe, false);
    const token = request.url.searchParams.get('token');
    const send = (message) => handlers.message({ origin: request.url.origin,
      message: JSON.stringify({ version: 1, token, ...message }) });
    send({ type: 'start', files: [{ name: 'sample.docx', size: 4, type: 'application/octet-stream', lastModified: 123 }] });
    send({ type: 'chunk', file: 0, index: 0, data: btoa('test') });
    send({ type: 'finish' });
    assert.equal(closed, 1);
    assert.equal(changed, 1);
    assert.equal(node('export-docx').files[0].name, 'sample.docx');
    assert.equal(await node('export-docx').files[0].text(), 'test');
    send({ type: 'finish' });
    assert.equal(changed, 1, 'Messages after completion must not repeat file loading.');
  } finally {
    if (previous === undefined) delete globalThis.DataTransfer;
    else globalThis.DataTransfer = previous;
  }
});

test('dialog transfer validation rejects wrong tokens, out-of-order and incomplete data', () => {
  const start = { version: 1, token: 'test', type: 'start', files: [{ name: 'sample.bib', size: 4 }] };
  assert.throws(() => createFileReceiver('test', false)({ ...start, token: 'wrong' }), /Invalid file chooser/);
  const receive = createFileReceiver('test', false);
  receive(start);
  assert.throws(() => receive({ version: 1, token: 'test', type: 'chunk', file: 0, index: 1, data: btoa('test') }), /out-of-order/);
  assert.throws(() => receive({ version: 1, token: 'test', type: 'finish' }), /Incomplete/);
  assert.throws(() => createFileReceiver('test', false)({ ...start, files: [start.files[0], start.files[0]] }), /Invalid file selection/);
  assert.throws(() => createFileReceiver('test', true)({ ...start, files: [{ name: 'large.bib', size: 25_000_001 }] }), /25 MB/);
});

test('older webviews retain the native input default', () => {
  const { document, click } = fixture();
  bindFilePickers(document);
  assert.equal(click('bib-file').defaultPrevented, false);
});

test('a blocked picker reports the failure and keeps native fallback without recursive clicks', () => {
  const { document, node, click } = fixture();
  node('export-docx').showPicker = () => { throw Object.assign(new Error('blocked'), { name: 'NotAllowedError' }); };
  node('export-docx').click = () => assert.fail('Fallback must not recursively call input.click().');
  bindFilePickers(document);
  assert.equal(click('export-docx').defaultPrevented, false);
  const feedback = node('export-docx-picker-status');
  assert.equal(feedback.hidden, false);
  assert.match(feedback.textContent, /NotAllowedError/);
  assert.match(feedback.textContent, /Open local tools in browser/);
  node('export-docx').events.change();
  assert.equal(feedback.hidden, true);
  assert.equal(feedback.textContent, '');
});

test('disabled controls do not open a picker and cancellation clears stale errors', () => {
  const { document, node, click } = fixture();
  node('csl-file').disabled = true;
  node('csl-file').showPicker = () => assert.fail('Disabled control');
  bindFilePickers(document);
  assert.equal(click('csl-file').defaultPrevented, false);
  node('csl-file-picker-status').hidden = false;
  node('csl-file').events.cancel();
  assert.equal(node('csl-file-picker-status').hidden, true);
});

test('every file input has local feedback and the browser page omits the Office runtime', () => {
  const html = readFileSync(join(__dirname, 'taskpane.html'), 'utf8');
  const browserHtml = readFileSync(join(__dirname, 'local-tools.html'), 'utf8');
  fileIds.forEach((id) => assert.ok(html.includes(`id="${id}-picker-status"`)));
  assert.match(html, /local-tools\.html\?browser=1#quarto-export-heading/);
  assert.doesNotMatch(browserHtml, /appsforoffice\.microsoft\.com/);
  assert.match(browserHtml, /data-local-tools="true"/);
});

test('local browser mode starts without waiting for Office and browser links use the Office API when present', () => {
  const { document, node } = fixture();
  document.documentElement = { dataset: { localTools: 'true' } };
  document.addEventListener = () => {};
  document.createElement = () => ({ textContent: '' });
  fileIds.forEach((id) => { node(id).showPicker = () => {}; });
  const originalSelect = document.querySelectorAll;
  document.querySelectorAll = (selector) => selector === '[data-open-local-tools]'
    ? [node('browser-link')] : originalSelect(selector);
  const originalGet = document.getElementById;
  document.getElementById = (id) => {
    const item = originalGet(id);
    if (item.value === undefined) item.value = '';
    item.replaceChildren = () => {};
    item.appendChild = () => {};
    return item;
  };
  const opened = [];
  const sandbox = {
    document,
    Office: { onReady() { assert.fail('Browser mode must not wait for Office readiness.'); },
      context: { ui: { openBrowserWindow: (url) => opened.push(url) } } },
    BibWordCitation: require('./citation.js'), BibWordCsl: { CSL_STYLES: [] },
    BibWordGithubLibrary: {}, BibWordFilePicker: { bindFilePickers },
  };
  runInNewContext(readFileSync(join(__dirname, 'taskpane.js'), 'utf8'), sandbox);
  assert.equal(node('browser-mode-note').hidden, false);
  assert.match(node('status').textContent, /Local browser tools ready/);
  let prevented = false;
  node('browser-link').events.click({ currentTarget: { href: 'https://example.org/local-tools.html' },
    preventDefault() { prevented = true; } });
  assert.deepEqual(opened, ['https://example.org/local-tools.html']);
  assert.equal(prevented, true);
  sandbox.Office.context.ui.openBrowserWindow = () => { throw new Error('unavailable'); };
  prevented = false;
  node('browser-link').events.click({ currentTarget: { href: 'https://example.org/local-tools.html' },
    preventDefault() { prevented = true; } });
  assert.equal(prevented, false, 'The ordinary link must survive an Office API failure.');
  assert.match(node('status').textContent, /Open or copy the local tools link manually/);
});
