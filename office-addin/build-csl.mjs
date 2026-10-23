import { build } from 'esbuild';
import { fileURLToPath } from 'node:url';

const browserFetchShim = fileURLToPath(new URL('browser-fetch-shim.mjs', import.meta.url));

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
