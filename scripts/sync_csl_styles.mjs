import { createRequire } from 'node:module';
import { mkdir, writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const styles = require('../office-addin/node_modules/@citation-js/plugin-csl/lib/styles.json');
const destination = fileURLToPath(new URL('../src/bibdocqmd/styles/', import.meta.url));

await mkdir(destination, { recursive: true });
await Promise.all([
  writeFile(new URL('../src/bibdocqmd/styles/apa.csl', import.meta.url), styles.apa, 'utf8'),
  writeFile(new URL('../src/bibdocqmd/styles/harvard1.csl', import.meta.url), styles.harvard1, 'utf8'),
]);
