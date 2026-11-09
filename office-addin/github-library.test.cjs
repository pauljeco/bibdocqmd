const test = require('node:test');
const assert = require('node:assert/strict');
const { githubBibliographyUrl, fetchGithubBibliography } = require('./github-library.js');

test('accepts public GitHub bibliography file pages and raw links', () => {
  const raw = 'https://raw.githubusercontent.com/team/papers/main/data/references.bib';
  assert.equal(githubBibliographyUrl('https://github.com/team/papers/blob/main/data/references.bib'), raw);
  assert.equal(githubBibliographyUrl(raw), raw);
  assert.equal(githubBibliographyUrl('https://github.com/team/papers/blob/v1/refs.BIBTEX'),
    'https://raw.githubusercontent.com/team/papers/v1/refs.BIBTEX');
});

test('rejects credentials, other hosts, and non-bibliography links', () => {
  for (const url of [
    'http://github.com/team/papers/blob/main/refs.bib',
    'https://user:token@github.com/team/papers/blob/main/refs.bib',
    'https://github.com.evil.test/team/papers/blob/main/refs.bib',
    'https://github.com/team/papers/blob/main/refs.txt',
    'https://github.com/team/papers/blob/main/refs.bib?token=secret',
    'https://github.com/team/papers/tree/main/refs.bib',
  ]) assert.throws(() => githubBibliographyUrl(url));
});

test('fetches GitHub text as a source and reports HTTP failures', async () => {
  const link = 'https://github.com/team/papers/blob/main/refs.bib';
  const source = await fetchGithubBibliography(link, async (url, options) => {
    assert.equal(url, 'https://raw.githubusercontent.com/team/papers/main/refs.bib');
    assert.deepEqual(options, { cache: 'no-store' });
    return new Response('@article{one, title={One}}');
  });
  assert.deepEqual(source, {
    name: 'https://raw.githubusercontent.com/team/papers/main/refs.bib',
    contents: '@article{one, title={One}}',
  });
  await assert.rejects(() => fetchGithubBibliography(link, async () => new Response('', { status: 404 })),
    /HTTP 404/);
});

test('rejects a streamed file over 20 MB', async () => {
  const chunk = new Uint8Array(1024 * 1024);
  const response = new Response(new ReadableStream({
    start(controller) {
      for (let index = 0; index < 21; index += 1) controller.enqueue(chunk);
      controller.close();
    },
  }));
  await assert.rejects(() => fetchGithubBibliography(
    'https://raw.githubusercontent.com/team/papers/main/refs.bib', async () => response,
  ), /20 MB limit/);
});
