"""DOCX -> Quarto/Pandoc Markdown converter for tagged, visible citations.

Prototype scope: body-level inline w:sdt content controls tagged
``bibword:cite:v1:...``. Key view stores visible Pandoc/Quarto syntax;
formatted view stores the same canonical syntax in the control alias.

Pandoc handles non-citation DOCX structure. The converter substitutes unique
sentinels for citation controls in a temporary DOCX, reads the Pandoc AST,
restores real Cite nodes, then writes Markdown. No author-year text matching.
"""
from __future__ import annotations

import argparse
import base64
import copy
import hashlib
import io
import json
from pathlib import Path
import re
import subprocess
import tempfile
import zipfile
from lxml import etree

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
W = '{' + NS['w'] + '}'
CITATION_TAG_PREFIX = 'bibword:cite:v1:'
CITATION_SOURCE_ALIAS_PREFIX = 'BibWord citation source v1: '
BIBLIOGRAPHY_TAG_PREFIX = 'bibword:refs:'
BIBLIOGRAPHY_TAG_V1_PREFIX = 'bibword:refs:v1:'
BIBLIOGRAPHY_TAG_V2_PREFIX = 'bibword:refs:v2:'
BIBLIOGRAPHY_TAG_V3_PREFIX = 'bibword:refs:v3:'
BIBLIOGRAPHY_SCOPES = {'document', 'heading1', 'heading2', 'heading3'}
CUSTOM_CSL_NAMESPACE = 'urn:bibword:csl-style:v1'
CSL_NAMESPACE = 'http://purl.org/net/xbiblio/csl'
SUPPORTED_STYLES = {
    'apa': 'apa.csl',
    'harvard1': 'harvard1.csl',
}
CUSTOM_STYLE_PATTERN = re.compile(r'^custom-[0-9a-f]{12}$')
TOKEN_PATTERN = re.compile(r'BIBWORDCIT\d{6}TOKEN')
BIB_KEY_PATTERN = re.compile(r'^\s*@\w+\s*[({]\s*([^,\s]+)\s*,', re.M)


class CitationError(ValueError):
    """Citation is invalid or could not be mapped to Pandoc syntax."""


def bibliography_metadata_from_tag(tag: str) -> dict[str, str | int | None]:
    """Return marker identity, style, and semantic scope across tag versions."""
    if tag.startswith(BIBLIOGRAPHY_TAG_V1_PREFIX):
        marker_id = tag.removeprefix(BIBLIOGRAPHY_TAG_V1_PREFIX)
        if not marker_id:
            raise CitationError('Bibliography marker has no identity.')
        return {'version': 1, 'id': marker_id, 'style': None, 'scope': 'document'}
    if tag.startswith(BIBLIOGRAPHY_TAG_V3_PREFIX):
        parts = tag.removeprefix(BIBLIOGRAPHY_TAG_V3_PREFIX).split(':')
        if len(parts) != 3 or not all(parts):
            raise CitationError(
                'Bibliography v3 marker must contain an identity, CSL style, and scope.')
        marker_id, style, scope = parts
        if scope not in BIBLIOGRAPHY_SCOPES:
            raise CitationError(f'Unsupported bibliography scope: {scope}')
        version = 3
    elif tag.startswith(BIBLIOGRAPHY_TAG_V2_PREFIX):
        metadata = tag.removeprefix(BIBLIOGRAPHY_TAG_V2_PREFIX)
        marker_id, separator, style = metadata.rpartition(':')
        if not separator or not marker_id or not style:
            raise CitationError('Bibliography v2 marker must contain an identity and CSL style.')
        scope = 'document'
        version = 2
    else:
        raise CitationError(f'Unsupported bibliography marker: {tag}')
    if style not in SUPPORTED_STYLES and not CUSTOM_STYLE_PATTERN.fullmatch(style):
        raise CitationError(f'Unsupported bibliography style: {style}')
    return {'version': version, 'id': marker_id, 'style': style, 'scope': scope}


def bibliography_style_from_tag(tag: str) -> str | None:
    """Return persisted style metadata while accepting legacy markers."""
    return bibliography_metadata_from_tag(tag)['style']


def csl_style_path(style: str) -> Path:
    filename = SUPPORTED_STYLES.get(style)
    if not filename:
        raise CitationError(f'Unsupported bibliography style: {style}')
    path = Path(__file__).with_name('styles') / filename
    if not path.is_file():
        raise CitationError(f'Bundled CSL style is missing: {path}')
    return path


def validate_csl_style(content: bytes) -> None:
    """Reject malformed and dependent styles that cannot render independently."""
    if len(content) > 2_000_000:
        raise CitationError('Embedded CSL style exceeds the 2 MB prototype limit.')
    if re.search(br'<!DOCTYPE\b|<!ENTITY\b', content, re.I):
        raise CitationError('CSL styles containing document type or entity declarations are unsupported.')
    try:
        parser = etree.XMLParser(resolve_entities=False, no_network=True)
        root = etree.fromstring(content, parser=parser)
    except etree.XMLSyntaxError as exc:
        raise CitationError(f'Embedded CSL style is invalid XML: {exc}') from exc
    qname = etree.QName(root)
    if qname.localname != 'style' or qname.namespace != CSL_NAMESPACE:
        raise CitationError('Embedded custom style is not a CSL style document.')
    if root.find(f'{{{CSL_NAMESPACE}}}info/{{{CSL_NAMESPACE}}}title') is None:
        raise CitationError('Embedded CSL style has no title metadata.')
    if root.find(f'{{{CSL_NAMESPACE}}}citation') is None \
            or root.find(f'{{{CSL_NAMESPACE}}}bibliography') is None:
        raise CitationError('Embedded CSL style must define citation and bibliography layouts.')
    dependent = root.xpath('.//csl:link[@rel="independent-parent"]',
                           namespaces={'csl': CSL_NAMESPACE})
    if dependent:
        raise CitationError('Dependent CSL styles are not supported; use the independent parent style.')


def embedded_csl_styles(docx: Path) -> dict[str, bytes]:
    """Read and integrity-check BibWord custom CSL parts stored in a DOCX."""
    styles: dict[str, bytes] = {}
    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    with zipfile.ZipFile(docx) as archive:
        # Word for the web writes this package directory as ``customXML``;
        # accept casing variants because OPC ZIP member names are case-sensitive.
        names = [name for name in archive.namelist()
                 if name.casefold().startswith('customxml/') and name.casefold().endswith('.xml')]
        for name in names:
            try:
                root = etree.fromstring(archive.read(name), parser=parser)
            except etree.XMLSyntaxError:
                continue
            qname = etree.QName(root)
            if qname.namespace != CUSTOM_CSL_NAMESPACE or qname.localname != 'bibwordStyle':
                continue
            style_id = root.get('id', '')
            expected_hash = root.get('hash', '')
            if not CUSTOM_STYLE_PATTERN.fullmatch(style_id) or not re.fullmatch(r'[0-9a-f]{64}', expected_hash):
                raise CitationError(f'Embedded BibWord CSL part has invalid metadata: {name}')
            csl = root.find(f'{{{CUSTOM_CSL_NAMESPACE}}}csl')
            if csl is None or csl.get('encoding') != 'base64' or not csl.text:
                raise CitationError(f'Embedded BibWord CSL part has no encoded style: {name}')
            try:
                content = base64.b64decode(''.join(csl.text.split()), validate=True)
            except (ValueError, TypeError) as exc:
                raise CitationError(f'Embedded BibWord CSL part has invalid base64 data: {name}') from exc
            actual_hash = hashlib.sha256(content).hexdigest()
            if actual_hash != expected_hash or style_id != f'custom-{actual_hash[:12]}':
                raise CitationError(f'Embedded CSL style {style_id} failed its integrity check.')
            validate_csl_style(content)
            if style_id in styles and styles[style_id] != content:
                raise CitationError(f'Conflicting embedded CSL styles use identifier {style_id}.')
            styles[style_id] = content
    return styles


def materialize_csl_style(style: str, destination: Path,
                          custom_styles: dict[str, bytes] | None = None) -> Path:
    """Write the selected bundled or embedded style without silent overwrite."""
    if style in SUPPORTED_STYLES:
        source = csl_style_path(style)
        content = source.read_bytes()
    elif CUSTOM_STYLE_PATTERN.fullmatch(style):
        content = (custom_styles or {}).get(style)
        if content is None:
            raise CitationError(f'Embedded custom CSL style is missing: {style}')
    else:
        raise CitationError(f'Unsupported bibliography style: {style}')
    target = destination / f'bibword-{style}.csl'
    if target.exists() and target.read_bytes() != content:
        raise CitationError(f'Refusing to overwrite a different CSL file: {target}')
    target.write_bytes(content)
    return target


def materialize_multiple_bibliographies_filter(destination: Path) -> Path:
    """Write BibWord's bundled scoped-bibliography Lua filter safely."""
    source = Path(__file__).with_name('filters') / 'multiple_bibliographies.lua'
    if not source.is_file():
        raise CitationError(f'Bundled multiple-bibliography filter is missing: {source}')
    content = source.read_bytes()
    target = destination / 'bibword-multiple-bibliographies.lua'
    if target.exists() and target.read_bytes() != content:
        raise CitationError(f'Refusing to overwrite a different Lua filter: {target}')
    target.write_bytes(content)
    return target


def run_pandoc(*args: str, stdin: str | None = None) -> str:
    command = ['pandoc', *args]
    try:
        completed = subprocess.run(command, input=stdin, capture_output=True,
                                   text=True, check=True)
    except FileNotFoundError as exc:
        raise RuntimeError('Pandoc 3.x is required: https://pandoc.org/installing.html') from exc
    except subprocess.CalledProcessError as exc:
        raise RuntimeError(f"Pandoc failed: {' '.join(command)}\n{exc.stderr}") from exc
    return completed.stdout


def pandoc_default_reference_docx() -> bytes:
    """Return the installed Pandoc version's standard Word reference document."""
    command = ['pandoc', '--print-default-data-file', 'reference.docx']
    try:
        completed = subprocess.run(command, capture_output=True, check=True)
    except FileNotFoundError as exc:
        raise RuntimeError('Pandoc 3.x is required: https://pandoc.org/installing.html') from exc
    except subprocess.CalledProcessError as exc:
        detail = exc.stderr.decode('utf-8', errors='replace')
        raise RuntimeError(f"Pandoc failed: {' '.join(command)}\n{detail}") from exc
    return completed.stdout


def hanging_bibliography_reference_docx() -> bytes:
    """Create a Pandoc reference DOCX with BibWord's bibliography indent."""
    source = io.BytesIO(pandoc_default_reference_docx())
    destination = io.BytesIO()
    bibliography_found = False
    with zipfile.ZipFile(source) as original, zipfile.ZipFile(destination, 'w') as updated:
        for info in original.infolist():
            content = original.read(info.filename)
            if info.filename == 'word/styles.xml':
                root = etree.fromstring(content)
                matches = root.xpath(
                    './/w:style[@w:styleId="Bibliography"]', namespaces=NS)
                if len(matches) != 1:
                    raise RuntimeError(
                        'Pandoc reference.docx does not contain one Bibliography style.')
                style = matches[0]
                paragraph_properties = style.find(W + 'pPr')
                if paragraph_properties is None:
                    paragraph_properties = etree.Element(W + 'pPr')
                    run_properties = style.find(W + 'rPr')
                    if run_properties is None:
                        style.append(paragraph_properties)
                    else:
                        style.insert(style.index(run_properties), paragraph_properties)
                indentation = paragraph_properties.find(W + 'ind')
                if indentation is None:
                    indentation = etree.SubElement(paragraph_properties, W + 'ind')
                indentation.set(W + 'left', '720')
                indentation.set(W + 'hanging', '720')
                indentation.attrib.pop(W + 'firstLine', None)
                content = etree.tostring(
                    root, encoding='UTF-8', xml_declaration=True, standalone=True)
                bibliography_found = True
            updated.writestr(info, content)
    if not bibliography_found:
        raise RuntimeError('Pandoc reference.docx has no word/styles.xml part.')
    return destination.getvalue()


def materialize_reference_docx(destination: Path) -> Path:
    """Write a content-addressed Word reference template without overwriting."""
    content = hanging_bibliography_reference_docx()
    digest = hashlib.sha256(content).hexdigest()[:12]
    target = destination / f'bibword-reference-{digest}.docx'
    if target.exists() and target.read_bytes() != content:
        raise CitationError(f'Refusing to overwrite a different reference DOCX: {target}')
    target.write_bytes(content)
    return target


def citation_ast(citation_text: str) -> dict:
    """Use Pandoc's *actual* parser rather than implementing our own grammar."""
    document = json.loads(run_pandoc('-f', 'markdown+citations', '-t', 'json',
                                      stdin=citation_text))
    blocks = document['blocks']
    if len(blocks) != 1 or blocks[0]['t'] not in ('Para', 'Plain'):
        raise CitationError(f'Not one inline citation: {citation_text!r}')
    inlines = blocks[0]['c']
    if len(inlines) != 1 or inlines[0]['t'] != 'Cite':
        raise CitationError(f'Not valid Quarto citation syntax: {citation_text!r}')
    return inlines[0]


def citation_control_error(position: int, citation_text: str, error: CitationError) -> CitationError:
    """Add enough Word-control context for a writer to repair bad syntax."""
    compact = ' '.join(citation_text.split())
    preview = compact if len(compact) <= 100 else compact[:97] + '...'
    hint = ''
    if compact.startswith('[') and ']' in compact and not compact.endswith(']'):
        hint = ' Keep sentence punctuation outside the blue citation control.'
    return CitationError(
        f'Citation {position} {preview!r} is invalid: {error}{hint}')


def extract_tagged_controls(docx: Path) -> tuple[bytes, dict[str, dict]]:
    """Replace BibWord controls with unique tokens in a temporary document part."""
    with zipfile.ZipFile(docx) as archive:
        original = archive.read('word/document.xml')
    root = etree.fromstring(original)
    mapping: dict[str, dict] = {}
    bibliography_ids: set[str] = set()
    citation_position = 0
    for control in list(root.xpath('.//w:sdt', namespaces=NS)):
        tag = control.find(f'{W}sdtPr/{W}tag')
        value = tag.get(W + 'val', '') if tag is not None else ''
        is_citation = value.startswith(CITATION_TAG_PREFIX)
        is_bibliography = value.startswith(BIBLIOGRAPHY_TAG_PREFIX)
        if not is_citation and not is_bibliography:
            continue
        content = control.find(f'{W}sdtContent')
        if content is None:
            raise CitationError(f'BibWord control has no content: {value}')
        if is_citation:
            citation_position += 1
            if content.xpath('.//w:p', namespaces=NS):
                raise CitationError(f'Citation control must be inline: {value}')
            texts = content.xpath('.//w:t/text()', namespaces=NS)
            visible_text = ''.join(texts).strip()
            alias = control.find(f'{W}sdtPr/{W}alias')
            alias_value = alias.get(W + 'val', '') if alias is not None else ''
            citation_text = (alias_value.removeprefix(CITATION_SOURCE_ALIAS_PREFIX).strip()
                             if alias_value.startswith(CITATION_SOURCE_ALIAS_PREFIX)
                             else visible_text)
            try:
                ast = citation_ast(citation_text)
            except CitationError as exc:
                raise citation_control_error(citation_position, citation_text, exc) from exc
        else:
            marker_paragraphs = content.xpath('./w:p', namespaces=NS)
            if not marker_paragraphs:
                raise CitationError('Bibliography marker must occupy its own paragraph.')
            metadata = bibliography_metadata_from_tag(value)
            marker_id = str(metadata['id'])
            if marker_id in bibliography_ids:
                raise CitationError(
                    f'Bibliography marker {marker_id} was copied; insert a new marker instead.')
            bibliography_ids.add(marker_id)
            # A Word-formatted bibliography may occupy later paragraphs inside
            # the marker. Only the first paragraph is the authoritative title.
            marker_title = ''.join(marker_paragraphs[0].xpath('.//w:t/text()',
                                   namespaces=NS)).strip()
            ast = {
                't': 'BibWordBibliography',
                'id': marker_id,
                'title': marker_title or 'References',
                'style': metadata['style'],
                'scope': metadata['scope'],
            }
        token = f'BIBWORDCIT{len(mapping) + 1:06d}TOKEN'
        mapping[token] = ast
        parent = control.getparent()
        if parent is None:
            raise CitationError('BibWord control has no parent')
        if is_bibliography:
            paragraph = etree.Element(W + 'p')
            run = etree.SubElement(paragraph, W + 'r')
            text = etree.SubElement(run, W + 't')
            text.text = token
            index = parent.index(control)
            parent.remove(control)
            parent.insert(index, paragraph)
            continue
        run = etree.Element(W + 'r')
        text = etree.SubElement(run, W + 't')
        text.text = token
        index = parent.index(control)
        parent.remove(control)
        parent.insert(index, run)
    changed_xml = etree.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)
    return changed_xml, mapping


def temporary_docx(original_path: Path, document_xml: bytes, destination: Path) -> None:
    with zipfile.ZipFile(original_path, 'r') as original, \
         zipfile.ZipFile(destination, 'w') as updated:
        for info in original.infolist():
            content = document_xml if info.filename == 'word/document.xml' else original.read(info.filename)
            updated.writestr(info, content)


def inject_citations(ast: dict, mapping: dict[str, dict]) -> tuple[dict, set[str]]:
    """Restore Cite nodes and the explicit Pandoc bibliography-position Div."""
    seen: set[str] = set()
    markers = [value for value in mapping.values()
               if value.get('t') == 'BibWordBibliography']
    scoped = len(markers) > 1 or any(marker.get('scope') != 'document' for marker in markers)

    def transform(node):
        if isinstance(node, dict):
            if node.get('t') in ('Para', 'Plain'):
                inlines = node.get('c', [])
                if (len(inlines) == 1 and isinstance(inlines[0], dict)
                        and inlines[0].get('t') == 'Str'):
                    token = inlines[0].get('c')
                    replacement = mapping.get(token)
                    if replacement and replacement.get('t') == 'BibWordBibliography':
                        seen.add(token)
                        title_words = replacement['title'].split()
                        title_inlines = []
                        for index, word in enumerate(title_words):
                            if index:
                                title_inlines.append({'t': 'Space'})
                            title_inlines.append({'t': 'Str', 'c': word})
                        scope = replacement.get('scope', 'document')
                        heading_level = (1 if scope == 'document'
                                         else min(int(scope.removeprefix('heading')) + 1, 6))
                        heading = {
                            't': 'Header',
                            'c': [heading_level, ['', ['unnumbered'], []], title_inlines],
                        }
                        if not scoped:
                            return {'t': 'Div', 'c': [['refs', [], []], [heading]]}
                        attributes = [
                            ['data-bibword-id', replacement['id']],
                            ['data-bibword-scope', scope],
                        ]
                        return {
                            't': 'Div',
                            'c': [[f"bibword-refs-{replacement['id']}",
                                   ['bibword-bibliography'], attributes], [heading]],
                        }
            return {key: transform(value) for key, value in node.items()}
        if not isinstance(node, list):
            return node
        result = []
        for part in node:
            if not (isinstance(part, dict) and part.get('t') == 'Str'):
                result.append(transform(part))
                continue
            value = part['c']
            matches = list(TOKEN_PATTERN.finditer(value))
            if not matches:
                result.append(part)
                continue
            cursor = 0
            for match in matches:
                if match.start() > cursor:
                    result.append({'t': 'Str', 'c': value[cursor:match.start()]})
                token = match.group()
                if token not in mapping:
                    raise CitationError(f'Unrecognised sentinel: {token}')
                if mapping[token].get('t') != 'Cite':
                    raise CitationError('Bibliography marker must occupy its own paragraph.')
                seen.add(token)
                result.append(copy.deepcopy(mapping[token]))
                cursor = match.end()
            if cursor < len(value):
                result.append({'t': 'Str', 'c': value[cursor:]})
        return result

    output = transform(ast)
    missing = set(mapping) - seen
    if missing:
        raise CitationError('Pandoc could not locate citation placeholders: ' + ', '.join(sorted(missing)))
    return output, seen


def validate_bibliography_scopes(ast: dict) -> None:
    """Reject scoped markers outside their heading or without any citations."""
    blocks = ast.get('blocks', [])

    def marker_metadata(block: dict) -> tuple[str, str] | None:
        if block.get('t') != 'Div':
            return None
        identifier, classes, attributes = block['c'][0]
        if 'bibword-bibliography' not in classes:
            return None
        values = dict(attributes)
        return identifier, values.get('data-bibword-scope', 'document')

    def contains_citation(value) -> bool:
        if isinstance(value, dict):
            return value.get('t') == 'Cite' or any(
                contains_citation(child) for child in value.values())
        if isinstance(value, list):
            return any(contains_citation(child) for child in value)
        return False

    for marker_index, block in enumerate(blocks):
        marker = marker_metadata(block)
        if not marker:
            continue
        identifier, scope = marker
        if scope == 'document':
            first, last = 0, len(blocks)
        else:
            if scope not in BIBLIOGRAPHY_SCOPES:
                raise CitationError(f'Unsupported bibliography scope: {scope}')
            level = int(scope.removeprefix('heading'))
            first = None
            for index in range(marker_index - 1, -1, -1):
                candidate = blocks[index]
                if candidate.get('t') == 'Header' and candidate['c'][0] <= level:
                    if candidate['c'][0] != level:
                        raise CitationError(
                            f'Bibliography marker {identifier} is not inside a Heading {level} section.')
                    first = index
                    break
            if first is None:
                raise CitationError(
                    f'Bibliography marker {identifier} has no preceding Heading {level}.')
            last = len(blocks)
            for index in range(first + 1, len(blocks)):
                candidate = blocks[index]
                if candidate.get('t') == 'Header' and candidate['c'][0] <= level:
                    last = index
                    break
        if not any(contains_citation(candidate) for candidate in blocks[first:last]):
            label = 'the document' if scope == 'document' else f'its {scope} section'
            raise CitationError(
                f'Bibliography marker {identifier} has no managed citations in {label}.')


def bibliography_keys(bib_file: Path) -> set[str]:
    """Basic validation only; not a complete BibLaTeX parser."""
    return set(BIB_KEY_PATTERN.findall(bib_file.read_text(encoding='utf-8-sig')))


def combined_bibliography_keys(bib_files: list[Path]) -> set[str]:
    """Combine keys while rejecting ambiguous duplicates across all sources."""
    source_by_key: dict[str, Path] = {}
    for bib_file in bib_files:
        keys = BIB_KEY_PATTERN.findall(bib_file.read_text(encoding='utf-8-sig'))
        file_keys: set[str] = set()
        for key in keys:
            if key in file_keys:
                raise CitationError(
                    f'Duplicate citation key "{key}" occurs more than once in {bib_file}.')
            file_keys.add(key)
            if key in source_by_key:
                raise CitationError(
                    f'Duplicate citation key "{key}" occurs in '
                    f'{source_by_key[key]} and {bib_file}. Rename one key before export.')
            source_by_key[key] = bib_file
    return set(source_by_key)


def resolve_bibliography_paths(bib_paths: str | Path | list[str | Path]) -> list[Path]:
    inputs = [bib_paths] if isinstance(bib_paths, (str, Path)) else list(bib_paths)
    if not inputs:
        raise CitationError('At least one bibliography is required.')
    resolved = [Path(path).resolve() for path in inputs]
    for path in resolved:
        if not path.exists():
            raise FileNotFoundError(f'Bibliography not found: {path}')
    return resolved


def export_docx(docx_path: str | Path, qmd_path: str | Path,
                bib_path: str | Path | list[str | Path],
                title: str = 'Word to Quarto prototype') -> int:
    docx_path = Path(docx_path).resolve()
    qmd_path = Path(qmd_path).resolve()
    bib_paths = resolve_bibliography_paths(bib_path)

    xml, mapping = extract_tagged_controls(docx_path)
    if not mapping:
        raise CitationError('No BibWord content controls found in this DOCX.')

    known_keys = combined_bibliography_keys(bib_paths)
    citations = [value for value in mapping.values() if value.get('t') == 'Cite']
    used_keys = {c['citationId'] for value in citations for c in value['c'][0]}
    unknown = used_keys - known_keys
    if unknown:
        raise CitationError('Citation keys missing from bibliography: ' + ', '.join(sorted(unknown)))

    with tempfile.TemporaryDirectory(prefix='bibword-') as temp:
        temp_path = Path(temp)
        clean_docx = temp_path / 'converted.docx'
        temporary_docx(docx_path, xml, clean_docx)
        doc_ast = json.loads(run_pandoc(str(clean_docx), '-f', 'docx', '-t', 'json'))
        patched_ast, _ = inject_citations(doc_ast, mapping)
        validate_bibliography_scopes(patched_ast)
        rendered = run_pandoc('-f', 'json', '-t', 'markdown+citations+pipe_tables',
                              '--wrap=none', stdin=json.dumps(patched_ast))

    qmd_path.parent.mkdir(parents=True, exist_ok=True)
    # A relative bibliography path keeps the output portable with its .bib file.
    from os.path import relpath
    relative_bibs = [relpath(path, qmd_path.parent).replace('\\', '/') for path in bib_paths]
    bibliography_value: str | list[str] = (
        relative_bibs[0] if len(relative_bibs) == 1 else relative_bibs)
    frontmatter = ('---\n'
                   f'title: {json.dumps(title, ensure_ascii=False)}\n'
                   f'bibliography: {json.dumps(bibliography_value, ensure_ascii=False)}\n')
    styles = {value.get('style') for value in mapping.values()
              if value.get('t') == 'BibWordBibliography' and value.get('style')}
    if styles:
        if len(styles) > 1:
            raise CitationError(
                'Bibliography markers use different CSL styles; choose one document style.')
        style = styles.pop()
        custom_styles = embedded_csl_styles(docx_path) if CUSTOM_STYLE_PATTERN.fullmatch(style) else None
        relative_csl = relpath(materialize_csl_style(style, qmd_path.parent, custom_styles),
                               qmd_path.parent).replace('\\', '/')
        frontmatter += f'csl: {json.dumps(relative_csl, ensure_ascii=False)}\n'
    markers = [value for value in mapping.values()
               if value.get('t') == 'BibWordBibliography']
    scoped = len(markers) > 1 or any(marker.get('scope') != 'document' for marker in markers)
    if scoped:
        relative_filter = relpath(
            materialize_multiple_bibliographies_filter(qmd_path.parent),
            qmd_path.parent,
        ).replace('\\', '/')
        frontmatter += ('citeproc: false\n'
                        'filters:\n'
                        f'  - {json.dumps(relative_filter, ensure_ascii=False)}\n')
    relative_reference = relpath(materialize_reference_docx(qmd_path.parent),
                                 qmd_path.parent).replace('\\', '/')
    frontmatter += ('format:\n'
                    '  html: default\n'
                    '  docx:\n'
                    f'    reference-doc: {json.dumps(relative_reference)}\n'
                    '---\n\n')
    qmd_path.write_text(frontmatter + rendered, encoding='utf-8')
    return len(citations)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description='Export tagged Word Quarto-style citations to .qmd')
    parser.add_argument('docx', type=Path)
    parser.add_argument(
        '--bib', type=Path, action='append', required=True,
        help='BibTeX/BibLaTeX source; repeat --bib to combine several files.',
    )
    parser.add_argument('--out', type=Path, required=True)
    parser.add_argument('--title', default='Word to Quarto prototype')
    args = parser.parse_args(argv)
    try:
        count = export_docx(args.docx, args.out, args.bib, args.title)
    except (CitationError, FileNotFoundError, RuntimeError) as exc:
        parser.exit(2, f'Error: {exc}\n')
    print(f'Exported {count} managed citation controls to {args.out}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
