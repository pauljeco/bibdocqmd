import { build } from 'esbuild';
import { fileURLToPath } from 'node:url';
import { readFile, writeFile } from 'node:fs/promises';

const browserFetchShim = fileURLToPath(new URL('browser-fetch-shim.mjs', import.meta.url));

// The same local file tools can run outside Office's embedded webview. Keep
// their markup in sync, and omit Office.js entirely from this browser page.
const paneHtml = await readFile(new URL('taskpane.html', import.meta.url), 'utf8');
await writeFile(new URL('local-tools.html', import.meta.url), paneHtml
  .replace('<html lang="en">', '<html lang="en" data-local-tools="true">')
  .replace(/\s*<script src="https:\/\/appsforoffice\.microsoft\.com\/lib\/1\/hosted\/office\.js"><\/script>/, '')
  .replace('<title>bibdocqmd prototype</title>', '<title>bibdocqmd local file tools</title>'));

await build({
  entryPoints: [fileURLToPath(new URL('csl-engine.mjs', import.meta.url))],
  outfile: fileURLToPath(new URL('dist/csl-engine.js', import.meta.url)),
  bundle: true,
  format: 'iife',
  globalName: 'BibWordCsl',
  platform: 'browser',
  target: ['es2020'],
  legalComments: 'eof',
  alias: {
    'node-fetch': browserFetchShim,
    'sync-fetch': browserFetchShim,
  },
});

await build({
  entryPoints: [fileURLToPath(new URL('markdown-export-browser.mjs', import.meta.url))],
  outfile: fileURLToPath(new URL('dist/markdown-export.js', import.meta.url)),
  bundle: true,
  format: 'iife',
  globalName: 'BibWordMarkdown',
  platform: 'browser',
  target: ['es2020'],
  loader: { '.csl': 'text', '.lua': 'text' },
  legalComments: 'eof',
  banner: { js: '/*! fflate 0.8.3\n' + await readFile(new URL('node_modules/fflate/LICENSE', import.meta.url), 'utf8') + '\n*/' },
  alias: {
    'node-fetch': browserFetchShim,
    'sync-fetch': browserFetchShim,
  },
});
