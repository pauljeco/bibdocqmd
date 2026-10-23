export const Headers = globalThis.Headers;

function unsupportedSynchronousFetch() {
  throw new Error('Network bibliography loading is disabled; choose a local .bib file.');
}

unsupportedSynchronousFetch.Headers = Headers;

export default unsupportedSynchronousFetch;
