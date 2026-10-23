"""Regression tests for the DOCX-to-Quarto proof of concept.

Execute from repository root: python -m pytest -q
"""
from pathlib import Path
import base64
import hashlib
import json
import subprocess
import sys
import zipfile
import pytest
from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))
from bibword.export import (
    CitationError, bibliography_style_from_tag, citation_ast, export_docx,
    embedded_csl_styles, extract_tagged_controls, materialize_csl_style,
    validate_csl_style,
)

DOCX = ROOT / 'examples' / 'bibword_demo.docx'
BIB = ROOT / 'examples' / 'references.bib'


@pytest.mark.parametrize('syntax,mode,keys', [
    ('@smith2020', 'AuthorInText', ['smith2020']),
    ('[@smith2020]', 'NormalCitation', ['smith2020']),
    ('[@smith2020; @jones2022]', 'NormalCitation', ['smith2020', 'jones2022']),
    ('[-@jones2022]', 'SuppressAuthor', ['jones2022']),
    ('[@jones2022, p. 45]', 'NormalCitation', ['jones2022']),
    ('@smith2020 [p. 45]', 'AuthorInText', ['smith2020']),
])
def test_pandoc_parses_citation_types(syntax, mode, keys):
    ast = citation_ast(syntax)
    assert ast['t'] == 'Cite'
    assert [c['citationId'] for c in ast['c'][0]] == keys
    assert ast['c'][0][0]['citationMode']['t'] == mode


@pytest.mark.parametrize('invalid', ['plain text', 'Smith (2020)', '[foo]', '@not a key'])
def test_rejects_invalid_syntax(invalid):
    with pytest.raises(CitationError):
        citation_ast(invalid)


def test_fixture_contains_five_real_tagged_controls():
    xml, mapping = extract_tagged_controls(DOCX)
    assert len(mapping) == 5
    assert b'BIBWORDCIT000005TOKEN' in xml
    assert b'bibword:cite:v1:' not in xml  # all managed controls temporarily removed


def test_docx_to_quarto_citations_and_structure(tmp_path):
    out = tmp_path / 'report.qmd'
    count = export_docx(DOCX, out, BIB, 'Example document')
    assert count == 5
    text = out.read_text(encoding='utf-8')
    assert '@smith2020 describes' in text
    assert '[@smith2020]' in text
    assert '[@smith2020; @jones2022]' in text
    assert '[@jones2022, p.' in text
    assert '[-@jones2022]' in text
    assert '# Example citations' in text
    assert '# Document structure' in text
    assert 'Word feature' in text
    assert 'First tested conversion' in text
    assert 'BIBWORDCIT' not in text

    result = subprocess.run(['pandoc', '-f', 'markdown+citations', '-t', 'json', str(out)],
                            capture_output=True, text=True, check=True)
    ast = json.loads(result.stdout)
    modes = []
    identifiers = []
    def walk(x):
        if isinstance(x, dict):
            if x.get('t') == 'Cite':
                modes.extend(c['citationMode']['t'] for c in x['c'][0])
                identifiers.extend(c['citationId'] for c in x['c'][0])
            for v in x.values():
                walk(v)
        elif isinstance(x, list):
            for v in x:
                walk(v)
    walk(ast)
    assert modes.count('AuthorInText') == 1
    assert modes.count('SuppressAuthor') == 1
    assert modes.count('NormalCitation') == 4  # 1 parenthetical + 2 grouped + 1 locator
    assert set(identifiers) == {'smith2020', 'jones2022'}


def test_missing_bibliography_key_is_detected(tmp_path):
    bib = tmp_path / 'bad.bib'
    bib.write_text('@article{smith2020, title={Test}}\n', encoding='utf-8')
    with pytest.raises(CitationError, match='jones2022'):
        export_docx(DOCX, tmp_path / 'bad.qmd', bib)


def test_bibliography_marker_controls_reference_position(tmp_path):
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    document = Document()
    paragraph = document.add_paragraph('Statement ')
    citation = OxmlElement('w:sdt')
    citation_properties = OxmlElement('w:sdtPr')
    citation_tag = OxmlElement('w:tag')
    citation_tag.set(qn('w:val'), 'bibword:cite:v1:test')
    citation_properties.append(citation_tag)
    citation_content = OxmlElement('w:sdtContent')
    citation_run = OxmlElement('w:r')
    citation_text = OxmlElement('w:t')
    citation_text.text = '[@smith2020]'
    citation_run.append(citation_text)
    citation_content.append(citation_run)
    citation.extend([citation_properties, citation_content])
    paragraph._p.append(citation)

    marker_paragraph = document.add_paragraph('References')
    marker = OxmlElement('w:sdt')
    marker_properties = OxmlElement('w:sdtPr')
    marker_tag = OxmlElement('w:tag')
    marker_tag.set(qn('w:val'), 'bibword:refs:v2:test:apa')
    marker_properties.append(marker_tag)
    marker_content = OxmlElement('w:sdtContent')
    marker_parent = marker_paragraph._p.getparent()
    marker_index = marker_parent.index(marker_paragraph._p)
    marker_parent.remove(marker_paragraph._p)
    marker_content.append(marker_paragraph._p)
    generated_paragraph = OxmlElement('w:p')
    generated_run = OxmlElement('w:r')
    generated_text = OxmlElement('w:t')
    generated_text.text = 'Generated reference that must not be exported'
    generated_run.append(generated_text)
    generated_paragraph.append(generated_run)
    marker_content.append(generated_paragraph)
    marker.extend([marker_properties, marker_content])
    marker_parent.insert(marker_index, marker)
    document.add_heading('Content after references', level=1)

    source = tmp_path / 'marker.docx'
    output = tmp_path / 'marker.qmd'
    document.save(source)
    count = export_docx(source, output, BIB)
    text = output.read_text(encoding='utf-8')

    assert count == 1
    assert '::: {#refs}' in text
    assert '# References {.unnumbered}' in text
    assert text.index('::: {#refs}') < text.index('# Content after references')
    assert 'Generated reference that must not be exported' not in text
    assert 'csl:' in text
    assert 'bibword-apa.csl' in text
    assert (tmp_path / 'bibword-apa.csl').is_file()


def test_formatted_and_copied_controls_export_from_canonical_alias(tmp_path):
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    document = Document()
    for visible, syntax in [
        ('(Smith, 2020)', '[@smith2020]'),
        ('Jones (2022)', '@jones2022'),
    ]:
        paragraph = document.add_paragraph('Statement ')
        citation = OxmlElement('w:sdt')
        properties = OxmlElement('w:sdtPr')
        alias = OxmlElement('w:alias')
        alias.set(qn('w:val'), f'BibWord citation source v1: {syntax}')
        tag = OxmlElement('w:tag')
        # Copy/paste may duplicate the BibWord tag. Export is based on each
        # physical control and must not collapse or reject either citation.
        tag.set(qn('w:val'), 'bibword:cite:v1:copied-tag')
        properties.extend([alias, tag])
        content = OxmlElement('w:sdtContent')
        run = OxmlElement('w:r')
        text = OxmlElement('w:t')
        text.text = visible
        run.append(text)
        content.append(run)
        citation.extend([properties, content])
        paragraph._p.append(citation)

    source = tmp_path / 'formatted-copy.docx'
    output = tmp_path / 'formatted-copy.qmd'
    document.save(source)

    xml, mapping = extract_tagged_controls(source)
    assert len(mapping) == 2
    assert b'(Smith, 2020)' not in xml
    assert export_docx(source, output, BIB) == 2
    rendered = output.read_text(encoding='utf-8')
    assert '[@smith2020]' in rendered
    assert '@jones2022' in rendered
    assert '(Smith, 2020)' not in rendered


def test_bibliography_marker_style_metadata_is_versioned_and_validated():
    assert bibliography_style_from_tag('bibword:refs:v1:legacy') is None
    assert bibliography_style_from_tag('bibword:refs:v2:marker-id:apa') == 'apa'
    assert bibliography_style_from_tag('bibword:refs:v2:marker-id:harvard1') == 'harvard1'
    assert bibliography_style_from_tag('bibword:refs:v2:marker-id:custom-0123456789ab') \
        == 'custom-0123456789ab'
    with pytest.raises(CitationError, match='Unsupported bibliography style'):
        bibliography_style_from_tag('bibword:refs:v2:marker-id:unknown')
    with pytest.raises(CitationError, match='identity and CSL style'):
        bibliography_style_from_tag('bibword:refs:v2:broken')
    with pytest.raises(CitationError, match='Unsupported bibliography marker'):
        bibliography_style_from_tag('bibword:refs:v3:future')


def custom_style_docx(source: Path, destination: Path, style: bytes,
                      stored_hash: str | None = None) -> str:
    namespace = 'urn:bibword:csl-style:v1'
    word_namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    word = f'{{{word_namespace}}}'
    digest = hashlib.sha256(style).hexdigest()
    style_id = f'custom-{digest[:12]}'

    store = etree.Element(f'{{{namespace}}}bibwordStyle', nsmap={None: namespace},
                          id=style_id, hash=stored_hash or digest)
    label = etree.SubElement(store, f'{{{namespace}}}label', encoding='base64')
    label.text = base64.b64encode(b'custom-test.csl').decode('ascii')
    data = etree.SubElement(store, f'{{{namespace}}}csl', encoding='base64')
    data.text = base64.b64encode(style).decode('ascii')

    with zipfile.ZipFile(source) as archive:
        document = etree.fromstring(archive.read('word/document.xml'))
        body = document.find(f'{word}body')
        marker = etree.Element(f'{word}sdt')
        properties = etree.SubElement(marker, f'{word}sdtPr')
        tag = etree.SubElement(properties, f'{word}tag')
        tag.set(f'{word}val', f'bibword:refs:v2:custom-marker:{style_id}')
        content = etree.SubElement(marker, f'{word}sdtContent')
        paragraph = etree.SubElement(content, f'{word}p')
        run = etree.SubElement(paragraph, f'{word}r')
        text = etree.SubElement(run, f'{word}t')
        text.text = 'References'
        section = body.find(f'{word}sectPr')
        body.insert(body.index(section) if section is not None else len(body), marker)
        changed = etree.tostring(document, encoding='UTF-8', xml_declaration=True,
                                 standalone=True)
        with zipfile.ZipFile(destination, 'w') as output:
            for info in archive.infolist():
                output.writestr(info, changed if info.filename == 'word/document.xml'
                                else archive.read(info.filename))
            output.writestr('customXML/itemBibWord.xml', etree.tostring(store,
                              encoding='UTF-8', xml_declaration=True, standalone=True))
    return style_id


def test_embedded_custom_csl_style_round_trips_to_export(tmp_path):
    style = (ROOT / 'src/bibword/styles/apa.csl').read_bytes()
    source = tmp_path / 'custom-style.docx'
    style_id = custom_style_docx(DOCX, source, style)

    assert embedded_csl_styles(source) == {style_id: style}
    validate_csl_style(style)
    output = tmp_path / 'custom-style.qmd'
    assert export_docx(source, output, BIB) == 5
    text = output.read_text(encoding='utf-8')
    assert f'csl: "bibword-{style_id}.csl"' in text
    assert (tmp_path / f'bibword-{style_id}.csl').read_bytes() == style


def test_embedded_custom_csl_style_integrity_failure_is_explicit(tmp_path):
    style = (ROOT / 'src/bibword/styles/apa.csl').read_bytes()
    source = tmp_path / 'tampered-style.docx'
    custom_style_docx(DOCX, source, style, stored_hash='0' * 64)
    with pytest.raises(CitationError, match='integrity check'):
        embedded_csl_styles(source)


def test_bundled_style_copy_does_not_overwrite_different_content(tmp_path):
    target = materialize_csl_style('apa', tmp_path)
    assert target.name == 'bibword-apa.csl'
    assert '<title>APA Style 7th edition</title>' in target.read_text(encoding='utf-8')
    target.write_text('user-owned style', encoding='utf-8')
    with pytest.raises(CitationError, match='Refusing to overwrite'):
        materialize_csl_style('apa', tmp_path)


def test_invalid_docx_without_managed_citations(tmp_path):
    from docx import Document
    plain = tmp_path / 'plain.docx'
    Document().save(plain)
    with pytest.raises(CitationError, match='No BibWord'):
        export_docx(plain, tmp_path / 'plain.qmd', BIB)
