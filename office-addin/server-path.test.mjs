import test from 'node:test';
import assert from 'node:assert/strict';
import { requestPath } from './server-path.mjs';

const rootWithTrailingSlash = '/tmp/bibword-office-addin/';

test('maps valid task-pane paths when the server root has a trailing slash', () => {
  assert.equal(
    requestPath(rootWithTrailingSlash, '/taskpane.html'),
    '/tmp/bibword-office-addin/taskpane.html',
  );
  assert.equal(
    requestPath(rootWithTrailingSlash, '/'),
    '/tmp/bibword-office-addin/taskpane.html',
  );
});

test('rejects encoded traversal and malformed URLs', () => {
  assert.equal(requestPath(rootWithTrailingSlash, '/%2e%2e%2fsecret'), null);
  assert.equal(requestPath(rootWithTrailingSlash, '/%E0%A4%A'), null);
});
