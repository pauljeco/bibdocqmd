import { readdir, readFile, realpath, stat } from 'node:fs/promises';
import { basename, extname, resolve, sep } from 'node:path';

const bibliographyExtensions = new Set(['.bib', '.bibtex']);
const maximumBibliographyBytes = 20 * 1024 * 1024;

function isBibliographyName(name) {
  return bibliographyExtensions.has(extname(name).toLowerCase());
}

function isInside(directory, candidate) {
  return candidate !== directory && candidate.startsWith(directory + sep);
}

async function resolvedDirectory(directory) {
  if (!directory || !String(directory).trim()) return null;
  const resolved = await realpath(resolve(String(directory).trim()));
  const details = await stat(resolved);
  if (!details.isDirectory()) throw new Error('The configured shared bibliography location is not a directory.');
  return resolved;
}

async function resolvedBibliography(directory, name) {
  const root = await resolvedDirectory(directory);
  if (!root) throw new Error('No shared bibliography folder is configured.');
  if (!name || basename(name) !== name || !isBibliographyName(name)) {
    throw new Error('Choose a valid bibliography filename.');
  }
  const candidate = await realpath(resolve(root, name));
  if (!isInside(root, candidate)) throw new Error('The bibliography is outside the configured folder.');
  const details = await stat(candidate);
  if (!details.isFile()) throw new Error('The selected bibliography is not a file.');
  if (details.size > maximumBibliographyBytes) {
    throw new Error('The selected bibliography exceeds the 20 MB safety limit.');
  }
  return { candidate, details };
}

export async function listSharedBibliographies(directory) {
  if (!directory || !String(directory).trim()) {
    return { configured: false, folder: '', files: [] };
  }
  const root = await resolvedDirectory(directory);
  const entries = await readdir(root, { withFileTypes: true });
  const files = [];
  for (const entry of entries) {
    if (!isBibliographyName(entry.name)) continue;
    try {
      const { details } = await resolvedBibliography(root, entry.name);
      files.push({
        name: entry.name,
        size: details.size,
        modified: details.mtime.toISOString(),
      });
    } catch {
      // Ignore broken links and links escaping the explicitly configured folder.
    }
  }
  files.sort((left, right) => left.name.localeCompare(right.name));
  return { configured: true, folder: basename(root), files };
}

export async function readSharedBibliography(directory, name) {
  const { candidate, details } = await resolvedBibliography(directory, name);
  return {
    name,
    contents: await readFile(candidate, 'utf8'),
    size: details.size,
    modified: details.mtime.toISOString(),
  };
}
