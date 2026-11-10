'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const { join } = require('node:path');
const { runInNewContext } = require('node:vm');
const { createFileReceiver } = require('./file-picker.js');

const script = readFileSync(join(__dirname, 'file-picker-dialog.js'), 'utf8');
const token = '12345678-1234-4567-890a-1234567890ab';
function dialog(inputId = 'export-docx') {
  const nodes = new Map();
  const messages = [];
  const get = (id) => {
    if (!nodes.has(id)) nodes.set(id, { disabled: true, events: {}, addEventListener(type, handler) { this.events[type] = handler; } });
    return nodes.get(id);
  };
  const context = {
    URLSearchParams, Uint8Array, btoa,
    location: { search: `?input=${inputId}&token=${token}` },
    document: { getElementById: get },
    Office: { onReady: (callback) => callback(), context: { ui: { messageParent: (message) => messages.push(JSON.parse(message)) } } },
    setTimeout: (callback) => callback(),
  };
  runInNewContext(script, context);
  return { get, messages };
}

test('the actual dialog script initializes and enables the correct native file input', () => {
  for (const [input, accept, multiple] of [['bib-file', '.bib,.bibtex,text/plain', true],
    ['csl-file', '.csl,application/xml,text/xml', false], ['export-docx', '.docx', false],
    ['export-preparation', '.zip', false], ['export-markdown', '.md,.zip', false]]) {
    const { get } = dialog(input);
    assert.equal(get('dialog-files').disabled, false);
    assert.equal(get('dialog-cancel').disabled, false);
    assert.equal(get('dialog-files').accept, accept);
    assert.equal(get('dialog-files').multiple, multiple);
    assert.match(get('dialog-status').textContent, /Choose files/);
  }
});

test('the dialog sends bounded chunks that reproduce the file bytes exactly', async () => {
  const { get, messages } = dialog();
  const bytes = Uint8Array.from({ length: 120_000 }, (_, index) => index % 256);
  get('dialog-files').files = [new File([bytes], 'sample.docx')];
  await get('dialog-files').events.change();
  assert.equal(messages[0].type, 'start');
  assert.equal(messages.at(-1).type, 'finish');
  const chunks = messages.filter((message) => message.type === 'chunk');
  assert.equal(chunks.length, 3);
  chunks.forEach((message) => assert.ok(message.data.length <= 65536));
  const receive = createFileReceiver(token, false);
  let result;
  messages.forEach((message) => { result = receive(message); });
  assert.deepEqual(new Uint8Array(await result.files[0].arrayBuffer()), bytes);
});

test('dialog cancellation sends no file content', () => {
  const { get, messages } = dialog();
  get('dialog-cancel').events.click();
  assert.equal(messages.length, 1);
  assert.equal(messages[0].type, 'cancel');
});
