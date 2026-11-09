/* Public GitHub bibliography links, shared by the task pane and tests. */
(function exposeGithubLibrary(root, factory) {
  const helpers = factory();
  if (typeof module === 'object' && module.exports) module.exports = helpers;
  else root.BibWordGithubLibrary = helpers;
}(typeof globalThis !== 'undefined' ? globalThis : this, () => {
  'use strict';

  const maximumBytes = 20 * 1024 * 1024;

  function githubBibliographyUrl(value) {
    let url;
    try {
      url = new URL(String(value).trim());
    } catch {
      throw new Error('Enter a full HTTPS GitHub .bib or .bibtex file URL.');
    }
    if (url.protocol !== 'https:' || url.username || url.password || url.port || url.search || url.hash) {
      throw new Error('Use a plain HTTPS GitHub file URL without credentials, query, or fragment.');
    }
    const segments = url.pathname.split('/').filter(Boolean);
    if (url.hostname === 'github.com') {
      if (segments.length < 5 || segments[2] !== 'blob') {
        throw new Error('Use a GitHub file link ending in .bib or .bibtex.');
      }
      segments.splice(2, 1);
    } else if (url.hostname !== 'raw.githubusercontent.com' || segments.length < 4) {
      throw new Error('Only public github.com or raw.githubusercontent.com file links are supported.');
    }
    if (segments.some((segment) => !segment || segment === '.' || segment === '..'
      || /%2f|%5c/i.test(segment))) {
      throw new Error('The GitHub file path is invalid.');
    }
    if (!/\.(bib|bibtex)$/i.test(decodeURIComponent(segments.at(-1)))) {
      throw new Error('The GitHub link must point to a .bib or .bibtex file.');
    }
    return `https://raw.githubusercontent.com/${segments.join('/')}`;
  }

  async function fetchGithubBibliography(value, fetcher = fetch) {
    const url = githubBibliographyUrl(value);
    let response;
    try {
      response = await fetcher(url, { cache: 'no-store' });
    } catch {
      throw new Error('Could not reach this public GitHub file. Check the link and network access.');
    }
    if (!response.ok) throw new Error(`GitHub returned HTTP ${response.status} for this file.`);
    const declaredSize = Number(response.headers?.get('content-length'));
    if (declaredSize > maximumBytes) throw new Error('The GitHub bibliography exceeds the 20 MB limit.');
    if (!response.body?.getReader) {
      throw new Error('This browser cannot safely read the GitHub file stream.');
    }
    const reader = response.body.getReader();
    const decoder = new TextDecoder('utf-8', { fatal: true });
    let contents = '';
    let bytes = 0;
    try {
      while (true) {
        const { done, value: chunk } = await reader.read();
        if (done) break;
        bytes += chunk.byteLength;
        if (bytes > maximumBytes) throw new Error('The GitHub bibliography exceeds the 20 MB limit.');
        contents += decoder.decode(chunk, { stream: true });
      }
      contents += decoder.decode();
    } catch (error) {
      await reader.cancel().catch(() => {});
      throw error;
    }
    return { name: url, contents };
  }

  return { githubBibliographyUrl, fetchGithubBibliography };
}));
