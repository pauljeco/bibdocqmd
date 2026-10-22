import { resolve, sep } from 'node:path';

export function requestPath(rootDirectory, requestUrl, port = 3000) {
  try {
    const root = resolve(rootDirectory);
    const pathname = decodeURIComponent(
      new URL(requestUrl, `https://localhost:${port}`).pathname,
    );
    const requested = pathname === '/' ? 'taskpane.html' : pathname.replace(/^\/+/, '');
    const candidate = resolve(root, requested);
    if (candidate !== root && !candidate.startsWith(root + sep)) return null;
    return candidate;
  } catch {
    return null;
  }
}
