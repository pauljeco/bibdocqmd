"""Regression tests for the DOCX-to-Quarto proof of concept.

Execute from repository root: python -m pytest -q
"""
from pathlib import Path
import json
import subprocess
import sys
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))
from bibword.export import CitationError, citation_ast, export_docx, extract_tagged_controls

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


def test_invalid_docx_without_managed_citations(tmp_path):
    from docx import Document
    plain = tmp_path / 'plain.docx'
    Document().save(plain)
    with pytest.raises(CitationError, match='No BibWord'):
        export_docx(plain, tmp_path / 'plain.qmd', BIB)
