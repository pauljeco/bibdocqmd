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
from bibdocqmd.export import (
    CitationError, bibliography_metadata_from_tag, bibliography_style_from_tag,
    citation_ast, export_docx,
    embedded_csl_styles, extract_tagged_controls, materialize_csl_style,
    materialize_multiple_bibliographies_filter, materialize_reference_docx,
    validate_csl_style,
)
from bibword.export import export_docx as legacy_export_docx

DOCX = ROOT / 'examples' / 'bibword_demo.docx'
BIB = ROOT / 'examples' / 'references.bib'


def test_legacy_python_import_uses_current_exporter():
    assert legacy_export_docx is export_docx


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


def test_bibliography_metadata_supports_scopes_and_legacy_tags():
    assert bibliography_metadata_from_tag('bibword:refs:v3:chapter-one:apa:heading1') == {
        'version': 3,
        'id': 'chapter-one',
        'style': 'apa',
        'scope': 'heading1',
    }
    assert bibliography_metadata_from_tag('bibword:refs:v2:old:harvard1')['scope'] == 'document'
    assert bibliography_metadata_from_tag('bibword:refs:v1:old')['scope'] == 'document'
    with pytest.raises(CitationError, match='Unsupported bibliography scope'):
        bibliography_metadata_from_tag('bibword:refs:v3:test:apa:page')


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
    assert 'reference-doc: "bibdocqmd-reference-' in text
    reference_docs = list(tmp_path.glob('bibdocqmd-reference-*.docx'))
    assert len(reference_docs) == 1

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


def test_multiple_bibliography_files_are_combined_in_quarto_metadata(tmp_path):
    articles = tmp_path / 'articles.bib'
    books = tmp_path / 'books.bib'
    articles.write_text(
        '@article{smith2020, author={Smith, Alex}, title={Article}, year={2020}}\n',
        encoding='utf-8',
    )
    books.write_text(
        '@book{jones2022, author={Jones, Morgan}, title={Book}, year={2022}}\n',
        encoding='utf-8',
    )
    output = tmp_path / 'combined.qmd'

    assert export_docx(DOCX, output, [articles, books]) == 5
    text = output.read_text(encoding='utf-8')
    assert 'bibliography: ["articles.bib", "books.bib"]' in text
    assert '@smith2020' in text
    assert '@jones2022' in text
    rendered = subprocess.run(
        ['pandoc', str(output), '--citeproc', '--to', 'html'],
        cwd=tmp_path, capture_output=True, text=True, check=True,
    ).stdout
    assert 'Article' in rendered
    assert 'Book' in rendered


def test_duplicate_keys_across_bibliography_files_are_rejected(tmp_path):
    first = tmp_path / 'first.bib'
    second = tmp_path / 'second.bib'
    first.write_text('@article{shared, title={First}}\n', encoding='utf-8')
    second.write_text('@article{shared, title={Second}}\n', encoding='utf-8')

    with pytest.raises(CitationError, match=r'Duplicate citation key "shared".*first\.bib.*second\.bib'):
        export_docx(DOCX, tmp_path / 'duplicate.qmd', [first, second])


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
    assert 'bibdocqmd-apa.csl' in text
    assert (tmp_path / 'bibdocqmd-apa.csl').is_file()
    assert 'reference-doc: "bibdocqmd-reference-' in text


def test_multiple_heading_scoped_bibliographies_export_and_render(tmp_path):
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    document = Document()

    def add_citation(key):
        paragraph = document.add_paragraph('Evidence ')
        citation = OxmlElement('w:sdt')
        properties = OxmlElement('w:sdtPr')
        tag = OxmlElement('w:tag')
        tag.set(qn('w:val'), f'bibword:cite:v1:{key}')
        properties.append(tag)
        content = OxmlElement('w:sdtContent')
        run = OxmlElement('w:r')
        text = OxmlElement('w:t')
        text.text = f'[@{key}]'
        run.append(text)
        content.append(run)
        citation.extend([properties, content])
        paragraph._p.append(citation)

    def add_marker(marker_id):
        paragraph = document.add_heading('References', level=2)
        marker = OxmlElement('w:sdt')
        properties = OxmlElement('w:sdtPr')
        tag = OxmlElement('w:tag')
        tag.set(qn('w:val'), f'bibword:refs:v3:{marker_id}:apa:heading1')
        properties.append(tag)
        content = OxmlElement('w:sdtContent')
        parent = paragraph._p.getparent()
        index = parent.index(paragraph._p)
        parent.remove(paragraph._p)
        content.append(paragraph._p)
        marker.extend([properties, content])
        parent.insert(index, marker)

    document.add_heading('Chapter one', level=1)
    add_citation('smith2020')
    add_marker('chapter-one')
    document.add_heading('Chapter two', level=1)
    add_citation('jones2022')
    add_citation('smith2020')
    add_marker('chapter-two')

    source = tmp_path / 'scoped.docx'
    output = tmp_path / 'scoped.qmd'
    document.save(source)
    assert export_docx(source, output, BIB) == 3

    qmd = output.read_text(encoding='utf-8')
    assert 'bibword-refs-chapter-one' in qmd
    assert 'bibword-refs-chapter-two' in qmd
    assert 'data-bibword-scope="heading1"' in qmd
    assert 'citeproc: false' in qmd
    assert 'bibdocqmd-multiple-bibliographies.lua' in qmd
    filter_path = tmp_path / 'bibdocqmd-multiple-bibliographies.lua'
    assert filter_path.is_file()

    rendered = subprocess.run([
        'pandoc', str(output), '--lua-filter', str(filter_path), '--to', 'html',
    ], cwd=tmp_path, capture_output=True, text=True, check=True).stdout
    normalized = rendered.casefold()
    first_chapter = normalized.index('chapter one')
    second_chapter = normalized.index('chapter two')
    first_reference = normalized.index('illustrative citation workflow')
    second_reference = normalized.index('demonstration of structured citations')
    assert first_chapter < first_reference < second_chapter < second_reference
    assert normalized.count('illustrative citation workflow') == 2
    assert normalized.count('demonstration of structured citations') == 1
    assert 'id="bibword-refs-chapter-one-ref-smith2020"' in rendered
    assert 'id="bibword-refs-chapter-two-ref-smith2020"' in rendered
    assert 'href="#ref-' not in rendered


def test_scoped_filter_is_not_silently_overwritten(tmp_path):
    target = materialize_multiple_bibliographies_filter(tmp_path)
    target.write_text('user-owned filter', encoding='utf-8')
    with pytest.raises(CitationError, match='Refusing to overwrite'):
        materialize_multiple_bibliographies_filter(tmp_path)


def test_generated_reference_docx_has_hanging_bibliography_style(tmp_path):
    target = materialize_reference_docx(tmp_path)
    assert target.name.startswith('bibdocqmd-reference-')
    assert target.suffix == '.docx'

    with zipfile.ZipFile(target) as archive:
        styles = etree.fromstring(archive.read('word/styles.xml'))
    namespace = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    indentation = styles.xpath(
        './/w:style[@w:styleId="Bibliography"]/w:pPr/w:ind', namespaces=namespace)
    assert len(indentation) == 1
    word = '{' + namespace['w'] + '}'
    assert indentation[0].get(word + 'left') == '720'
    assert indentation[0].get(word + 'hanging') == '720'
    assert indentation[0].get(word + 'firstLine') is None
    assert materialize_reference_docx(tmp_path) == target


def test_generated_reference_docx_is_not_silently_overwritten(tmp_path):
    target = materialize_reference_docx(tmp_path)
    target.write_bytes(b'user-owned document')
    with pytest.raises(CitationError, match='Refusing to overwrite'):
        materialize_reference_docx(tmp_path)


def test_formatted_and_copied_controls_export_from_canonical_alias(tmp_path):
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    document = Document()
    for visible, syntax, alias_prefix in [
        ('(Smith, 2020)', '[@smith2020]', 'BibWord citation source v1: '),
        ('Jones (2022)', '@jones2022', 'bibdocqmd citation source v1: '),
    ]:
        paragraph = document.add_paragraph('Statement ')
        citation = OxmlElement('w:sdt')
        properties = OxmlElement('w:sdtPr')
        alias = OxmlElement('w:alias')
        alias.set(qn('w:val'), f'{alias_prefix}{syntax}')
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


def test_invalid_control_error_identifies_position_text_and_punctuation(tmp_path):
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    document = Document()
    for index, visible in enumerate(['[@smith2020]', '[@bennett1998-spatial]. '], 1):
        paragraph = document.add_paragraph()
        citation = OxmlElement('w:sdt')
        properties = OxmlElement('w:sdtPr')
        tag = OxmlElement('w:tag')
        tag.set(qn('w:val'), f'bibword:cite:v1:test-{index}')
        properties.append(tag)
        content = OxmlElement('w:sdtContent')
        run = OxmlElement('w:r')
        text = OxmlElement('w:t')
        text.text = visible
        run.append(text)
        content.append(run)
        citation.extend([properties, content])
        paragraph._p.append(citation)

    source = tmp_path / 'invalid-control.docx'
    document.save(source)
    with pytest.raises(CitationError) as caught:
        extract_tagged_controls(source)
    message = str(caught.value)
    assert "Citation 2 '[@bennett1998-spatial].' is invalid" in message
    assert 'Keep sentence punctuation outside the blue citation control.' in message


def test_bibliography_marker_style_metadata_is_versioned_and_validated():
    assert bibliography_style_from_tag('bibword:refs:v1:legacy') is None
    assert bibliography_style_from_tag('bibword:refs:v2:marker-id:apa') == 'apa'
    assert bibliography_style_from_tag('bibword:refs:v2:marker-id:harvard1') == 'harvard1'
    assert bibliography_style_from_tag('bibword:refs:v2:marker-id:custom-0123456789ab') \
        == 'custom-0123456789ab'
    assert bibliography_style_from_tag('bibword:refs:v3:marker-id:apa:heading2') == 'apa'
    with pytest.raises(CitationError, match='Unsupported bibliography style'):
        bibliography_style_from_tag('bibword:refs:v2:marker-id:unknown')
    with pytest.raises(CitationError, match='identity and CSL style'):
        bibliography_style_from_tag('bibword:refs:v2:broken')
    with pytest.raises(CitationError, match='identity, CSL style, and scope'):
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
    style = (ROOT / 'src/bibdocqmd/styles/apa.csl').read_bytes()
    source = tmp_path / 'custom-style.docx'
    style_id = custom_style_docx(DOCX, source, style)

    assert embedded_csl_styles(source) == {style_id: style}
    validate_csl_style(style)
    output = tmp_path / 'custom-style.qmd'
    assert export_docx(source, output, BIB) == 5
    text = output.read_text(encoding='utf-8')
    assert f'csl: "bibdocqmd-{style_id}.csl"' in text
    assert (tmp_path / f'bibdocqmd-{style_id}.csl').read_bytes() == style


def test_embedded_custom_csl_style_integrity_failure_is_explicit(tmp_path):
    style = (ROOT / 'src/bibdocqmd/styles/apa.csl').read_bytes()
    source = tmp_path / 'tampered-style.docx'
    custom_style_docx(DOCX, source, style, stored_hash='0' * 64)
    with pytest.raises(CitationError, match='integrity check'):
        embedded_csl_styles(source)


def test_bundled_style_copy_does_not_overwrite_different_content(tmp_path):
    target = materialize_csl_style('apa', tmp_path)
    assert target.name == 'bibdocqmd-apa.csl'
    assert '<title>APA Style 7th edition</title>' in target.read_text(encoding='utf-8')
    target.write_text('user-owned style', encoding='utf-8')
    with pytest.raises(CitationError, match='Refusing to overwrite'):
        materialize_csl_style('apa', tmp_path)


def test_invalid_docx_without_managed_citations(tmp_path):
    from docx import Document
    plain = tmp_path / 'plain.docx'
    Document().save(plain)
    with pytest.raises(CitationError, match='No bibdocqmd'):
        export_docx(plain, tmp_path / 'plain.qmd', BIB)
