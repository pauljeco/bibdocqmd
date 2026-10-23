"""DOCX -> Quarto/Pandoc Markdown converter for tagged, visible citations.

Prototype scope: body-level inline w:sdt content controls, tagged
bibword:cite:v1:..., containing *visible* Pandoc/Quarto citation syntax.

Pandoc handles non-citation DOCX structure. The converter substitutes unique
sentinels for citation controls in a temporary DOCX, reads the Pandoc AST,
restores real Cite nodes, then writes Markdown. No author-year text matching.
"""
from __future__ import annotations

import argparse
import copy
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
BIBLIOGRAPHY_TAG_PREFIX = 'bibword:refs:v1:'
TOKEN_PATTERN = re.compile(r'BIBWORDCIT\d{6}TOKEN')
BIB_KEY_PATTERN = re.compile(r'^\s*@\w+\s*[({]\s*([^,\s]+)\s*,', re.M)


class CitationError(ValueError):
    """Citation is invalid or could not be mapped to Pandoc syntax."""


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


def extract_tagged_controls(docx: Path) -> tuple[bytes, dict[str, dict]]:
    """Replace BibWord controls with unique tokens in a temporary document part."""
    with zipfile.ZipFile(docx) as archive:
        original = archive.read('word/document.xml')
    root = etree.fromstring(original)
    mapping: dict[str, dict] = {}
    bibliography_found = False
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
            if content.xpath('.//w:p', namespaces=NS):
                raise CitationError(f'Citation control must be inline: {value}')
            texts = content.xpath('.//w:t/text()', namespaces=NS)
            ast = citation_ast(''.join(texts).strip())
        else:
            if bibliography_found:
                raise CitationError('Only one BibWord bibliography marker is allowed.')
            if not content.xpath('./w:p', namespaces=NS):
                raise CitationError('Bibliography marker must occupy its own paragraph.')
            bibliography_found = True
            marker_title = ''.join(content.xpath('.//w:t/text()', namespaces=NS)).strip()
            ast = {'t': 'BibWordBibliography', 'title': marker_title or 'References'}
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
                        heading = {'t': 'Header', 'c': [1, ['', ['unnumbered'], []], title_inlines]}
                        return {'t': 'Div', 'c': [['refs', [], []], [heading]]}
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


def bibliography_keys(bib_file: Path) -> set[str]:
    """Basic validation only; not a complete BibLaTeX parser."""
    return set(BIB_KEY_PATTERN.findall(bib_file.read_text(encoding='utf-8-sig')))


def export_docx(docx_path: str | Path, qmd_path: str | Path,
                bib_path: str | Path, title: str = 'Word to Quarto prototype') -> int:
    docx_path = Path(docx_path).resolve()
    qmd_path = Path(qmd_path).resolve()
    bib_path = Path(bib_path).resolve()
    if not bib_path.exists():
        raise FileNotFoundError(f'Bibliography not found: {bib_path}')

    xml, mapping = extract_tagged_controls(docx_path)
    if not mapping:
        raise CitationError('No BibWord content controls found in this DOCX.')

    known_keys = bibliography_keys(bib_path)
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
        rendered = run_pandoc('-f', 'json', '-t', 'markdown+citations+pipe_tables',
                              '--wrap=none', stdin=json.dumps(patched_ast))

    qmd_path.parent.mkdir(parents=True, exist_ok=True)
    # A relative bibliography path keeps the output portable with its .bib file.
    from os.path import relpath
    relative_bib = relpath(bib_path, qmd_path.parent).replace('\\', '/')
    frontmatter = ('---\n'
                   f'title: {json.dumps(title, ensure_ascii=False)}\n'
                   f'bibliography: {json.dumps(relative_bib, ensure_ascii=False)}\n'
                   'format:\n  html: default\n  docx: default\n'
                   '---\n\n')
    qmd_path.write_text(frontmatter + rendered, encoding='utf-8')
    return len(citations)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description='Export tagged Word Quarto-style citations to .qmd')
    parser.add_argument('docx', type=Path)
    parser.add_argument('--bib', type=Path, required=True)
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
