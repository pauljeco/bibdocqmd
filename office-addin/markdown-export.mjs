// Word performs the document-to-Markdown conversion. This module only replaces
// known bibdocqmd controls in a disposable copy and restores their exact syntax.
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import citation from './citation.js';
import { createLibraryFromSources, sha256Hex, validateCslStyle } from './csl-engine.mjs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CSL = 'urn:bibword:csl-style:v1';
const TOKEN = /BIBDOCQMD[A-F0-9]{32}[CB]\d{6}END/g;
const PLAN = 'bibdocqmd-export-plan.json';
const MAX_FILE = 25_000_000;
const MAX_EXPANDED = 100_000_000;
const { parseCitationControl, parseBibliographyTag, citationKeysForBibliographyScope } = citation;

function safePath(name) {
  if (!name || /[\\:\x00-\x1f]/.test(name) || name.startsWith('/')
      || name.split('/').some((part) => part === '..' || part === '.')) {
    throw new Error(`Unsafe archive filename: ${name}`);
  }
  return name;
}

export function readArchive(bytes) {
  if (bytes.length > MAX_FILE) throw new Error('Choose a file smaller than 25 MB.');
  let total = 0;
  const seen = new Set();
  return unzipSync(bytes, { filter(file) {
    safePath(file.name);
    const key = file.name.toLowerCase();
    if (seen.has(key)) throw new Error(`Duplicate archive filename: ${file.name}`);
    seen.add(key);
    total += file.originalSize;
    if (total > MAX_EXPANDED) throw new Error('The expanded archive exceeds 100 MB.');
    return true;
  } });
}

function packageFiles(files) {
  if (Object.values(files).reduce((sum, bytes) => sum + bytes.length, 0) > MAX_EXPANDED) {
    throw new Error('The export package exceeds 100 MB uncompressed.');
  }
  const bytes = zipSync(files);
  if (bytes.length > MAX_FILE) throw new Error('The export package exceeds 25 MB; use smaller source or media files.');
  return bytes;
}

function parseXml(text) {
  if (/<!DOCTYPE\b|<!ENTITY\b/i.test(text)) throw new Error('XML entity declarations are unsupported.');
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error('Invalid document XML.');
  return doc;
}

const descendants = (node, name) => node ? Array.from(node.getElementsByTagNameNS(W, name)) : [];
const first = (node, name) => descendants(node, name)[0];
const children = (node) => Array.from(node.childNodes).filter((child) => child.nodeType === 1);
const direct = (node, name) => children(node).find((child) => child.namespaceURI === W && child.localName === name);
const value = (node) => node?.getAttributeNS(W, 'val') || '';
const text = (node) => descendants(node, 't').map((part) => part.textContent).join('');

function summary(node) {
  const properties = direct(node, 'sdtPr');
  return { tag: value(first(properties, 'tag')), title: value(first(properties, 'alias')),
    text: text(direct(node, 'sdtContent')) };
}

function managed(node) {
  const properties = direct(node, 'sdtPr');
  return value(first(properties, 'tag')).startsWith('bibword:');
}

function inspectControls(doc) {
  const body = first(doc, 'body');
  if (!body) throw new Error('The DOCX has no main document body.');
  for (const name of ['ins', 'del', 'moveFrom', 'moveTo']) {
    if (descendants(body, name).length) throw new Error('Accept or reject tracked changes before preparing export.');
  }
  const paragraphs = [];
  const controls = [];
  function walk(node, paragraph = null, inTable = false) {
    if (node.namespaceURI !== W) return;
    if (node.localName === 'sdt' && managed(node)) {
      const info = summary(node);
      const content = direct(node, 'sdtContent');
      if (!content) throw new Error('A managed control has no content.');
      if (descendants(content, 'sdt').some(managed)) throw new Error('Nested managed controls are unsupported.');
      if (info.tag.startsWith('bibword:cite:v1:')) {
        if (!paragraph || descendants(content, 'p').length) {
          throw new Error('Managed citations must be inline in a body paragraph.');
        }
        const parsed = parseCitationControl(info, controls.filter((item) => item.kind === 'citation').length + 1);
        paragraph.citationKeys.push(...parsed.parsed.keys);
        controls.push({ node, kind: 'citation', tag: info.tag, source: parsed.source });
      } else if (info.tag.startsWith('bibword:refs:')) {
        if (inTable) throw new Error('Move bibliography markers outside tables before exporting.');
        const metadata = parseBibliographyTag(info.tag);
        if (!/^[a-zA-Z0-9-]+$/.test(metadata.id)) throw new Error('Invalid bibliography identity.');
        if (paragraph) {
          const siblings = children(node.parentNode).filter((child) => child !== node && child.localName !== 'pPr');
          if (siblings.some((sibling) => text(sibling).trim() || descendants(sibling, 'drawing').length
              || descendants(sibling, 'object').length || descendants(sibling, 'pict').length)) {
            throw new Error('A bibliography marker must occupy its own paragraph.');
          }
        }
        const titleParagraph = first(content, 'p');
        const title = (titleParagraph ? text(titleParagraph) : info.text).split(/[\r\n]/)[0].trim() || 'References';
        const record = paragraph || { headingLevel: 0, citationKeys: [] };
        record.headingLevel = 0;
        if (!paragraph) paragraphs.push(record);
        record.marker = { metadata, tag: info.tag, text: title };
        controls.push({ node, kind: 'bibliography', tag: info.tag, ...metadata, title, record });
      } else {
        throw new Error(`Unsupported managed control: ${info.tag}`);
      }
      return;
    }
    if (node.localName === 'p') {
      const heading = /^Heading([1-9])$/i.exec(value(first(first(node, 'pPr') || node, 'pStyle')).replace(/\s/g, ''));
      paragraph = { headingLevel: heading ? Number(heading[1]) : 0, citationKeys: [] };
      paragraphs.push(paragraph);
    }
    children(node).forEach((child) => walk(child, paragraph, inTable || node.localName === 'tbl'));
  }
  walk(body);
  if (!controls.some((item) => item.kind === 'citation')) throw new Error('No managed citations found in this DOCX.');
  const ids = new Set();
  controls.filter((item) => item.kind === 'bibliography').forEach((item) => {
    if (ids.has(item.id)) throw new Error(`Bibliography marker ${item.id} was copied; insert a new marker.`);
    ids.add(item.id);
    const keys = citationKeysForBibliographyScope(paragraphs, paragraphs.indexOf(item.record), item.scope);
    if (!keys.length) throw new Error(`Bibliography ${item.title} has no citations in its scope.`);
  });
  return controls;
}

function replaceControl(doc, node, token, bibliography) {
  const run = doc.createElementNS(W, 'w:r');
  const literal = doc.createElementNS(W, 'w:t');
  literal.textContent = token;
  run.appendChild(literal);
  if (bibliography) {
    const paragraph = doc.createElementNS(W, 'w:p');
    const properties = doc.createElementNS(W, 'w:pPr');
    const style = doc.createElementNS(W, 'w:pStyle');
    style.setAttributeNS(W, 'w:val', 'Normal');
    properties.appendChild(style);
    paragraph.appendChild(properties);
    paragraph.appendChild(run);
    if (node.parentNode.localName === 'p') node = node.parentNode;
    node.parentNode.replaceChild(paragraph, node);
  } else node.parentNode.replaceChild(run, node);
}

async function styleXml(files, id, assets) {
  if (assets.styles[id]) return assets.styles[id];
  if (!/^custom-[0-9a-f]{12}$/.test(id)) throw new Error(`Unsupported citation style: ${id}`);
  let found;
  for (const [name, bytes] of Object.entries(files)) {
    if (!/^customxml\/.*\.xml$/i.test(name)) continue;
    const doc = parseXml(strFromU8(bytes));
    const root = doc.documentElement;
    if (root.namespaceURI !== CSL || root.localName !== 'bibwordStyle' || root.getAttribute('id') !== id) continue;
    const encoded = root.getElementsByTagNameNS(CSL, 'csl')[0];
    if (!encoded || encoded.getAttribute('encoding') !== 'base64') throw new Error('Invalid embedded CSL encoding.');
    const xml = new TextDecoder().decode(Uint8Array.from(atob(encoded.textContent.replace(/\s/g, '')), (char) => char.charCodeAt(0)));
    const hash = await sha256Hex(xml);
    if (root.getAttribute('hash') !== hash || id !== `custom-${hash.slice(0, 12)}`) {
      throw new Error('Embedded CSL failed its integrity check.');
    }
    if (found && found !== xml) throw new Error('Conflicting embedded CSL styles.');
    found = xml;
  }
  if (!found) throw new Error(`Embedded CSL style is missing: ${id}`);
  return found;
}

const INSTRUCTIONS = `Open word-markdown-copy.docx in Word for the web and export it as Markdown.
This is a disposable export copy: edit your original Word document, then prepare a fresh copy.
Return to bibdocqmd, choose this preparation ZIP and Word's exported .md (or a ZIP containing
the .md and its accompanying images), and choose Download QMD package.
All processing is local. Native Word Markdown fidelity must be checked in your Word version.
`;

export async function prepareMarkdownExport(docxBytes, { sources, assets, title = 'My document', fallbackStyle = 'apa' }) {
  const files = readArchive(docxBytes);
  if (!files['word/document.xml'] || !files['[Content_Types].xml']) throw new Error('Choose a valid .docx file.');
  for (const [name, bytes] of Object.entries(files)) {
    if (/^word\/.*\.xml$/i.test(name) && name !== 'word/document.xml'
        && /bibword:(?:cite|refs):/.test(strFromU8(bytes))) {
      throw new Error(`Managed controls outside the main body are unsupported (${name}).`);
    }
  }
  const doc = parseXml(strFromU8(files['word/document.xml']));
  if (text(doc).match(TOKEN)) throw new Error('This file already contains export placeholders; choose your original DOCX.');
  const controls = inspectControls(doc);
  const library = createLibraryFromSources(sources);
  const keys = controls.filter((item) => item.kind === 'citation')
    .flatMap((item) => parseCitationControl({ text: item.source }).parsed.keys);
  library.assertKeys(keys);
  const markers = controls.filter((item) => item.kind === 'bibliography');
  const styles = [...new Set(markers.map((item) => item.style).filter(Boolean))];
  if (styles.length > 1) throw new Error('Bibliography markers use different styles; update them consistently first.');
  const style = styles[0] || fallbackStyle;
  const xml = await styleXml(files, style, assets);
  validateCslStyle(xml);
  const nonce = globalThis.crypto.randomUUID().replace(/-/g, '').toUpperCase();
  const replacements = controls.map(({ node, record, version, style: ignoredStyle, ...item }, index) => {
    const token = `BIBDOCQMD${nonce}${item.kind === 'citation' ? 'C' : 'B'}${String(index + 1).padStart(6, '0')}END`;
    replaceControl(doc, node, token, item.kind === 'bibliography');
    return { ...item, token };
  });
  files['word/document.xml'] = strToU8(new XMLSerializer().serializeToString(doc));
  const resources = {};
  const bibliography = [];
  for (const [index, source] of sources.entries()) {
    const file = `bibliography-${index + 1}.bib`;
    resources[file] = strToU8(source.contents);
    bibliography.push({ file, originalName: source.name, hash: await sha256Hex(source.contents) });
  }
  const cslFile = `bibdocqmd-${style}.csl`;
  resources[cslFile] = strToU8(xml);
  const plan = { version: 1, kind: 'bibdocqmd-native-markdown', title: String(title), nonce,
    replacements, bibliography, cslFile, cslHash: await sha256Hex(xml),
    scoped: markers.length > 1 || markers.some((item) => item.scope !== 'document') };
  const serializedPlan = JSON.stringify(plan, null, 2);
  return packageFiles({ ...resources, [PLAN]: strToU8(serializedPlan),
    'bibdocqmd-export-plan.sha256': strToU8(await sha256Hex(serializedPlan)),
    'word-markdown-copy.docx': packageFiles(files), 'READ-ME.txt': strToU8(INSTRUCTIONS) });
}

function bibliographyMarkdown(item, scoped) {
  const level = item.scope === 'document' ? 1 : Number(item.scope.slice(-1)) + 1;
  const title = item.title.replace(/\s+/g, ' ').replace(/[\\`*_{}\[\]<>#!|]/g, '\\$&');
  const attributes = scoped
    ? `#bibword-refs-${item.id} .bibword-bibliography data-bibword-id="${item.id}" data-bibword-scope="${item.scope}"`
    : '#refs';
  return `::: {${attributes}}\n${'#'.repeat(level)} ${title} {.unnumbered}\n:::`;
}

function count(textValue, needle) { return textValue.split(needle).length - 1; }

export async function finishMarkdownExport(preparationBytes, exportedBytes, { assets, markdownZip = false }) {
  const preparation = readArchive(preparationBytes);
  if (!preparation[PLAN]) throw new Error('Choose the preparation ZIP created by bibdocqmd.');
  const serializedPlan = strFromU8(preparation[PLAN]);
  if (!preparation['bibdocqmd-export-plan.sha256']
      || await sha256Hex(serializedPlan) !== strFromU8(preparation['bibdocqmd-export-plan.sha256'])) {
    throw new Error('The export plan failed its integrity check; choose the original preparation ZIP.');
  }
  const plan = JSON.parse(serializedPlan);
  if (plan.version !== 1 || plan.kind !== 'bibdocqmd-native-markdown'
      || !/^[A-F0-9]{32}$/.test(plan.nonce) || !Array.isArray(plan.replacements)
      || !plan.replacements.length || !Array.isArray(plan.bibliography) || !plan.bibliography.length
      || typeof plan.title !== 'string') {
    throw new Error('Unsupported or invalid export plan.');
  }
  if (exportedBytes.length > MAX_FILE) throw new Error('The Markdown export exceeds 25 MB.');
  const output = markdownZip ? readArchive(exportedBytes) : {};
  const mdFiles = Object.keys(output).filter((name) => /\.md$/i.test(name));
  if (markdownZip && mdFiles.length !== 1) throw new Error('Choose a ZIP containing exactly one Markdown file and its images.');
  const mdPath = markdownZip ? mdFiles[0] : 'document.md';
  const prefix = mdPath.slice(0, mdPath.lastIndexOf('/') + 1);
  let markdown = strFromU8(markdownZip ? output[mdPath] : exportedBytes).replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n');
  delete output[mdPath];
  if (/^---\n[\s\S]*?\n(?:---|\.\.\.)\s*(?:\n|$)/.test(markdown)) {
    throw new Error('The Markdown already has YAML metadata; remove that header before importing.');
  }
  const expected = new Set();
  for (const item of plan.replacements) {
    if (typeof item.token !== 'string' || !new RegExp(`^BIBDOCQMD${plan.nonce}[CB]\\d{6}END$`).test(item.token)
        || expected.has(item.token)) throw new Error('Invalid or duplicate export placeholder.');
    expected.add(item.token);
    if (count(markdown, item.token) !== 1) {
      throw new Error(`Missing or duplicated ${item.kind} placeholder. Export the prepared copy without editing it.`);
    }
    if (item.kind === 'citation') parseCitationControl({ text: item.source });
    else if (item.kind === 'bibliography') {
      const metadata = parseBibliographyTag(item.tag);
      if (!metadata || metadata.id !== item.id || metadata.scope !== item.scope
          || !/^[a-zA-Z0-9-]+$/.test(item.id) || typeof item.title !== 'string') {
        throw new Error('Invalid bibliography export metadata.');
      }
      const line = markdown.split('\n').find((part) => part.includes(item.token));
      if (line.trim() !== item.token || /^ {4}|^\t/.test(line)) {
        throw new Error('A bibliography placeholder must be alone in an ordinary paragraph.');
      }
    } else throw new Error('Unsupported export placeholder kind.');
  }
  for (const token of markdown.match(TOKEN) || []) {
    if (!expected.has(token)) throw new Error('The Markdown belongs to a different preparation ZIP.');
  }
  // Refuse inactive placeholders in code: replacing them would silently turn
  // citations into literal text rather than Pandoc Cite nodes.
  let fence = null;
  for (const line of markdown.split('\n')) {
    const delimiter = /^ {0,3}(`{3,}|~{3,})/.exec(line);
    if (delimiter) {
      if (!fence) fence = delimiter[1];
      else if (delimiter[1][0] === fence[0] && delimiter[1].length >= fence.length) fence = null;
    }
    if ((line.match(TOKEN) || []).length && (fence || /^ {4}|^\t/.test(line) || /`[^`]*BIBDOCQMD/.test(line))) {
      throw new Error('A citation placeholder was exported inside code. Remove code formatting in the original and prepare it again.');
    }
  }
  const sources = [];
  const addResource = (name, bytes) => {
    safePath(name);
    const target = prefix + name;
    if (Object.keys(output).some((path) => path.toLowerCase() === target.toLowerCase())) {
      throw new Error(`Exported media conflicts with a project file: ${target}`);
    }
    output[target] = bytes;
  };
  for (const source of plan.bibliography) {
    safePath(source.file);
    const bytes = preparation[source.file];
    if (!bytes || await sha256Hex(strFromU8(bytes)) !== source.hash) throw new Error('Bibliography file failed its integrity check.');
    sources.push({ name: source.file, contents: strFromU8(bytes) });
    addResource(source.file, bytes);
  }
  const library = createLibraryFromSources(sources);
  for (const item of plan.replacements.filter((item) => item.kind === 'citation')) {
    library.assertKeys(parseCitationControl({ text: item.source }).parsed.keys);
  }
  safePath(plan.cslFile);
  const csl = preparation[plan.cslFile];
  if (!csl || await sha256Hex(strFromU8(csl)) !== plan.cslHash) throw new Error('CSL file failed its integrity check.');
  validateCslStyle(strFromU8(csl));
  addResource(plan.cslFile, csl);
  const markers = plan.replacements.filter((item) => item.kind === 'bibliography');
  const scoped = markers.length > 1 || markers.some((item) => item.scope !== 'document');
  const header = ['---', `title: ${JSON.stringify(plan.title)}`, 'bibliography:',
    ...plan.bibliography.map((item) => `  - ${JSON.stringify(item.file)}`),
    `csl: ${JSON.stringify(plan.cslFile)}`];
  if (scoped) {
    header.push('citeproc: false', 'filters:', '  - bibdocqmd-multiple-bibliographies.lua');
    addResource('bibdocqmd-multiple-bibliographies.lua', strToU8(assets.filter));
  }
  header.push('---', '');
  // One pass prevents replacement text from being interpreted as more tokens.
  const replacements = new Map(plan.replacements.map((item) => [item.token,
    item.kind === 'citation' ? item.source : bibliographyMarkdown(item, scoped)]));
  markdown = markdown.replace(TOKEN, (token) => replacements.get(token));
  addResource('document.qmd', strToU8(header.join('\n') + '\n' + markdown));
  addResource('READ-ME.txt', strToU8('Render document.qmd with Quarto. Keep the bibliography, CSL, filter (if present), and media together.\n'
    + 'Word performed the Markdown conversion: check tables, notes, equations, captions, and cross-references against the original.\n'
    + 'Only media included in the selected Markdown ZIP is packaged; retain any separately exported files alongside document.qmd.\n'));
  return packageFiles(output);
}
