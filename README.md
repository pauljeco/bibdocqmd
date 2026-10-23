# BibWord | Quarto-native Word citations (technical spike v0.1)

A **working first development step** toward a BibTeX-native citation manager for Microsoft Word on Windows and **Word for the web**. The design preserves the visible Quarto/Pandoc markup (`@key`, `[@key]`) directly in Word.

> **Status:** DOCX → Quarto conversion and the task pane's citation-building logic are tested locally. Microsoft's validator accepts the manifest schema and configuration. On 23 October 2026, the pane loaded in Word for the web on Ubuntu/Firefox and six inserted citation controls, including a two-key group, retained their tags in the downloaded DOCX. All six controls then exported through Quarto with the real BibLaTeX library and rendered successfully to HTML and DOCX. Save/reopen, collaboration, and Word desktop remain unverified. This is not yet a finished add-in or a fork of JabRef.

## Deliverables

- `examples/bibword_demo.docx`: sample Word document with **five genuine Word inline content controls**, each showing Quarto citation notation (fictional example citations).
- `examples/references.bib`: the matching demonstration bibliography (fictional sources, not scholarly references).
- `examples/bibword_demo.qmd`: exported Quarto document containing *real* Pandoc citations with their original keys, modes, and locators.
- `src/bibword/export.py`: runnable DOCX-to-Quarto converter, retaining Word document structure through Pandoc and replacing citation control markers with native Pandoc `Cite` nodes.
- `office-addin/`: experimental minimal Office.js task pane. Choose a `.bib` file (local, in browser), enter a key, select narrative/parenthetical/suppress-author mode, and insert tagged citation syntax.
- `tests/test_export.py`: regression tests for citation modes, preservation and invalid references.

## Why the visible syntax is the stored syntax

An in-Word citation *looks* like `[@smith2020]`, not `(Smith, 2020)`. This serves three purposes:

1. Writers learn the same citation notation used in Quarto and Pandoc.
2. Citation identity remains visible even if the Word add-in is unavailable.
3. The Word document can round-trip into semantic Pandoc `Cite` nodes without text-recognition heuristics.

The Word content control has a unique tag beginning `bibword:cite:v1:`. It is tagged so the exporter knows which substrings are managed citations, and does not replace ordinary text that just happens to look like `[@key]`.

Quarto's canonical citation grammar is parsed by **Pandoc itself**, rather than by a custom regular expression.

## Developer quick start on Ubuntu

Requirements: Python 3.10+, Pandoc 3.x, Node.js 20+, npm, and Git.

```bash
make setup
make check
.venv/bin/bibword-export examples/bibword_demo.docx \
  --bib examples/references.bib \
  --out examples/bibword_demo.qmd \
  --title "BibWord demonstration"
```

`make setup` creates a local Python virtual environment and uses the committed npm lockfile. See [`docs/development.md`](docs/development.md) for independent setup commands and the Word sideloading checklist.

The `.qmd` and `.bib` must be distributed together. Pandoc/Quarto citations can be rendered to a finished bibliography using:

```bash
cd examples
pandoc bibword_demo.qmd --from markdown+citations --citeproc --to html \
  --output example_rendered.html
```

When using Pandoc directly, run from `examples/` (or set `--resource-path`) so it resolves `references.bib` correctly.

## Try the experimental Word add-in (not yet verified in Word)

The minimal task pane is deliberately dependency-light and does **not** use the JabRef GraphQL backend.

1. Install Node.js 20 or newer, then run `make setup-addin` from the repository root.
2. In `office-addin/`, run `npm run cert` and `npm run cert:verify` to generate and trust local HTTPS credentials.
3. Run `npm run dev` (serves `https://localhost:3000/`). Certificate files are read from the development certificate store and are not copied into the repository.
4. Open a document in Word for the web. If your tenant allows developer sideloading, upload `office-addin/manifest.xml` through the Word Add-ins UI; see [Microsoft's sideloading guide](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).
5. In the task pane, load `examples/references.bib`, enter `smith2020`, and insert `[@smith2020]` or `@smith2020`.
6. Save, reopen, and download the document as `.docx`; inspect the retained content-control tag and run `bibword_export.py` against the downloaded file.

**Limitations:** browsers require a trusted HTTPS certificate and some Microsoft 365 tenants disallow sideloading. One Word-for-the-web insertion/download path has been verified, but save/reopen, copy/paste, coauthoring, and broad host compatibility remain pending. The manifest passes Microsoft's automated validator, but that does not prove runtime compatibility. The `.bib` scanner only extracts basic keys and is **not** a full BibLaTeX parser. The add-in does not yet insert bibliographies or provide a formatted CSL preview.

## Architecture and next steps

See `docs/technical_audit.md` for the upstream repository assessment, `docs/development.md` for reproducible setup, and `docs/word-host-test-matrix.md` for real-host evidence.

```text
BibTeX file → Office.js task pane → inline tagged Quarto citation in Word
                                              │
                                    saved Word DOCX
                                              │
                                 custom citation-aware converter
                                              │
                              Pandoc AST with native Cite nodes
                                              │
                                         report.qmd
                                              │
                                         references.bib
```

Next: test save/reopen persistence in **Word for the web**, validate the Office.js insertion operation in **Windows Word**, then evaluate which parts of [JabRef-Word-Addin](https://github.com/JabRef/JabRef-Word-Addin) can be reused under its MIT licence. The repo uses React 17, citeproc, Apollo/GraphQL, an external API setting, and Office.js. That backend is not needed for this first proof of concept. Build a robust BibLaTeX parser and key-validation workflow before shipping.

## Known limitations and test boundaries

- Only **inline content controls in the DOCX main document body** are supported (not headers, footnotes, text boxes, bibliography fields, cross-references or tracked-change contexts).
- Copy/paste, concurrent editing, and Word Online persistence of controls have not been verified on actual Microsoft products.
- Word users can change the text within a control; invalid citation syntax then produces an explicit error on export.
- Untagged Quarto-looking text is left untouched rather than guessed to be a citation.
- Raw `.bib` entries are not rewritten, so BibLaTeX-specific fields are preserved in the source file.
- The exporter is a prototype CLI requiring Python and Pandoc on the export computer, **not** on colleagues' Word computers.
- Figures, advanced layout and DOCX formatting are delegated to Pandoc, which cannot always recreate Word layouts faithfully.
- The Office add-in has no server-side data storage. A locally chosen `.bib` file is only read in the browser, not uploaded.

## Licence and attribution

The independently written prototype code is provided under the MIT licence (see `LICENSE`).


The included source was independently written as a proof of concept; the JabRef repository has **not been cloned or vendored** into this archive. When integrating JabRef upstream code later, retain its MIT licence and copyright notices.
