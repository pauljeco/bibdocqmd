import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { listSharedBibliographies, readSharedBibliography } from './shared-library.mjs';

test('an unconfigured shared bibliography folder reports no files', async () => {
  assert.deepEqual(await listSharedBibliographies(''), {
    configured: false, folder: '', files: [],
  });
});

test('lists and reads only bibliography files in the configured directory', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'bibword-shared-'));
  await writeFile(join(directory, 'team.bib'), '@article{team, title={Team}}');
  await writeFile(join(directory, 'extra.BIBTEX'), '@book{extra, title={Extra}}');
  await writeFile(join(directory, 'notes.txt'), 'not a bibliography');

  const listing = await listSharedBibliographies(directory);
  assert.equal(listing.configured, true);
  assert.deepEqual(listing.files.map(({ name }) => name), ['extra.BIBTEX', 'team.bib']);
  assert.equal(
    (await readSharedBibliography(directory, 'team.bib')).contents,
    '@article{team, title={Team}}',
  );
});

test('rejects traversal and ignores bibliography symlinks outside the folder', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'bibword-shared-'));
  const outside = join(tmpdir(), `bibword-outside-${Date.now()}.bib`);
  await writeFile(outside, '@article{outside, title={Outside}}');
  await symlink(outside, join(directory, 'outside.bib'));

  assert.deepEqual((await listSharedBibliographies(directory)).files, []);
  await assert.rejects(() => readSharedBibliography(directory, '../secret.bib'));
  await assert.rejects(() => readSharedBibliography(directory, 'outside.bib'));
  await assert.rejects(() => readSharedBibliography(directory, 'notes.txt'));
});
