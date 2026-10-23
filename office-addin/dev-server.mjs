import { createServer } from 'node:https';
import { readFile, stat } from 'node:fs/promises';
import { extname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { getHttpsServerOptions } from 'office-addin-dev-certs';
import { requestPath } from './server-path.mjs';
import './build-csl.mjs';

const root = fileURLToPath(new URL('.', import.meta.url));
const port = Number.parseInt(process.env.BIBWORD_PORT || '3000', 10);
if (!Number.isInteger(port) || port < 1 || port > 65535) {
  throw new Error('BIBWORD_PORT must be an integer between 1 and 65535.');
}
const mimeTypes = new Map([
  ['.css', 'text/css; charset=utf-8'],
  ['.html', 'text/html; charset=utf-8'],
  ['.js', 'text/javascript; charset=utf-8'],
  ['.png', 'image/png'],
  ['.xml', 'application/xml; charset=utf-8'],
]);

const options = await getHttpsServerOptions();
const server = createServer(options, async (request, response) => {
  try {
    const path = requestPath(root, request.url || '/', port);
    if (!path || !(await stat(path)).isFile()) throw new Error('not found');
    const body = await readFile(path);
    response.writeHead(200, {
      'Content-Type': mimeTypes.get(extname(path)) || 'application/octet-stream',
      'Cache-Control': 'no-store',
      'X-Content-Type-Options': 'nosniff',
    });
    response.end(body);
  } catch {
    response.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
    response.end('Not found');
  }
});

server.listen(port, '127.0.0.1', () => {
  console.log(`BibWord development server: https://localhost:${port}/taskpane.html`);
});
