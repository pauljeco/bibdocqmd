"""Produce a sample Word DOCX containing genuine inline Word content controls."""
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

BASE = Path(__file__).resolve().parent


def insert_citation(paragraph, syntax: str, tag_suffix: str) -> None:
    sdt = OxmlElement('w:sdt')
    props = OxmlElement('w:sdtPr')
    alias = OxmlElement('w:alias')
    alias.set(qn('w:val'), 'BibWord citation')
    tag = OxmlElement('w:tag')
    tag.set(qn('w:val'), 'bibword:cite:v1:' + tag_suffix)
    props.extend([alias, tag])
    content = OxmlElement('w:sdtContent')
    run = OxmlElement('w:r')
    style = OxmlElement('w:rPr')
    fonts = OxmlElement('w:rFonts')
    fonts.set(qn('w:ascii'), 'Consolas')
    fonts.set(qn('w:hAnsi'), 'Consolas')
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '165789')
    style.extend([fonts, color])
    text = OxmlElement('w:t')
    text.text = syntax
    run.extend([style, text])
    content.append(run)
    sdt.extend([props, content])
    paragraph._p.append(sdt)


def build_sample() -> None:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(.78)
    section.bottom_margin = Inches(.78)
    styles = doc.styles
    styles['Normal'].font.name = 'Aptos'
    styles['Normal'].font.size = Pt(10.5)
    doc.add_heading('BibWord: Quarto syntax in Microsoft Word', 0)
    doc.add_paragraph('Development fixture • version 0.1 | illustrative references only', style='Subtitle')
    doc.add_paragraph('This DOCX demonstrates citations displayed exactly as Quarto/Pandoc markup. '
                      'Blue monospaced snippets are managed Word content controls, not ordinary pasted text.')
    doc.add_heading('Example citations', 1)

    p = doc.add_paragraph('Narrative citation: ')
    insert_citation(p, '@smith2020', 'narrative-001')
    p.add_run(' describes an illustrative workflow.')

    p = doc.add_paragraph('Parenthetical citation: an illustrative statement ')
    insert_citation(p, '[@smith2020]', 'parenthetical-002')
    p.add_run('.')

    p = doc.add_paragraph('Grouped citation: this example draws on two sample records ')
    insert_citation(p, '[@smith2020; @jones2022]', 'group-003')
    p.add_run('.')

    p = doc.add_paragraph('With a page locator: ')
    insert_citation(p, '[@jones2022, p. 45]', 'locator-004')
    p.add_run('.')

    p = doc.add_paragraph('Suppress author: Jones is mentioned separately, then ')
    insert_citation(p, '[-@jones2022]', 'suppress-005')
    p.add_run(' can be used to show only the year after processing.')

    doc.add_heading('Document structure', 1)
    doc.add_paragraph('Normal paragraphs, headings and tables are converted using Pandoc; citation content controls are restored as native Pandoc Cite nodes.')
    tab = doc.add_table(rows=1, cols=2)
    tab.style = 'Light Shading Accent 1'
    tab.rows[0].cells[0].text = 'Word feature'
    tab.rows[0].cells[1].text = 'Quarto result'
    for a, b in [('Heading 1', '# Heading'), ('Managed citation', 'Native Pandoc citation'),
                 ('Table', 'Markdown table')]:
        cells = tab.add_row().cells
        cells[0].text = a
        cells[1].text = b
    doc.add_paragraph('First conceptual deliverable', style='List Bullet')
    doc.add_paragraph('First tested conversion', style='List Bullet')
    doc.save(BASE / 'bibword_demo.docx')


if __name__ == '__main__':
    build_sample()
