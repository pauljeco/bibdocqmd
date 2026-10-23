# BibWord | Quarto-native Word citations (technical spike v0.1)

A **working first development step** toward a BibTeX-native citation manager for Microsoft Word on Windows and **Word for the web**. The design preserves the visible Quarto/Pandoc markup (`@key`, `[@key]`) directly in Word.

> **Status:** DOCX → Quarto conversion and the task pane's citation-building logic are tested locally. Microsoft's validator accepts the manifest schema and configuration. On 23 October 2026, the pane loaded in Word for the web on Ubuntu/Firefox; citation editing, save/reopen persistence, a three-key group, the bibliography marker, Citation.js/citeproc-js previews, style switching, and styled-marker persistence worked in the real host. A downloaded document retained its selected Harvard style, and the exporter produced matching Harvard citations and references at the explicit marker. Copy/paste, collaboration, and Word desktop remain unverified. This is not yet a finished add-in or a fork of JabRef.

## Deliverables

- `examples/bibword_demo.docx`: sample Word document with **five genuine Word inline content controls**, each showing Quarto citation notation (fictional example citations).
- `examples/references.bib`: the matching demonstration bibliography (fictional sources, not scholarly references).
- `examples/bibword_demo.qmd`: exported Quarto document containing *real* Pandoc citations with their original keys, modes, and locators.
- `src/bibword/export.py`: runnable DOCX-to-Quarto converter, retaining Word document structure through Pandoc and replacing citation control markers with native Pandoc `Cite` nodes.
- `office-addin/`: experimental Office.js task pane. Choose a `.bib` file locally, insert or update tagged citation syntax, inspect APA or Harvard CSL citation/reference previews, preview all document references, and insert an explicit bibliography-position marker.
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

The `.qmd` and `.bib` must be distributed together. When a document has a styled `v2` bibliography marker, the exporter also writes a `bibword-*.csl` file beside the `.qmd`; distribute that file with them. Pandoc/Quarto citations can be rendered to a finished bibliography using:

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
5. In the task pane, load `examples/references.bib`, enter `smith2020`, choose a preview style, inspect the CSL-formatted preview, and insert `[@smith2020]` or `@smith2020`.
6. To edit, put the Word cursor inside the blue citation, choose **Load selected citation**, change its fields, then choose **Update selected**.
7. Choose **Refresh document preview** to format all unique managed references in the pane.
8. Put the cursor where the references must appear and choose **Insert bibliography marker**. Use one marker per document. Changing the citation style after inserting the marker updates its stored export style.
9. Save, reopen, and download the document as `.docx`; inspect the retained content-control tags and run `bibword_export.py` against the downloaded file.

**Limitations:** browsers require a trusted HTTPS certificate and some Microsoft 365 tenants disallow sideloading. Word-for-the-web insertion, selected-citation editing, bibliography-marker insertion, save/reopen persistence, download, and export have been verified in one Ubuntu/Firefox/Microsoft 365 path. Copy/paste, coauthoring, Word desktop, and broad host compatibility remain pending. The manifest passes Microsoft's automated validator, but that does not prove runtime compatibility. Citation.js provides local BibLaTeX parsing and citeproc-js provides CSL previews. APA 7th edition and Cite Them Right Harvard are currently the two bundled styles. Word stores visible keys plus a visible, versioned bibliography-position marker containing the selected style identifier.

## Architecture and next steps

See `docs/technical_audit.md` for the upstream repository assessment, `docs/development.md` for reproducible setup, and `docs/word-host-test-matrix.md` for real-host evidence.

```text
BibTeX file → Office.js task pane → tagged Quarto citations in Word
                      │                         + bibliography marker
                      └→ local preview                  │
                                             saved Word DOCX
                                                    │
                                       citation-aware converter
                                                    │
                              Pandoc Cite nodes + explicit #refs Div
                                                    │
                                               report.qmd
                                                    │
                                            Pandoc/CSL output
```

Next: validate the workflow in **Windows Word**, then add user-supplied CSL files and evaluate which parts of [JabRef-Word-Addin](https://github.com/JabRef/JabRef-Word-Addin) can be reused under its MIT licence without adopting its external backend.

## Known limitations and test boundaries

- Citation controls must be **inline controls in the DOCX main document body**. The one supported block control is the bibliography-position marker; headers, footnotes, text boxes, cross-references, and tracked-change contexts remain unsupported.
- Copy/paste and concurrent editing have not been verified on actual Microsoft products; Word-for-the-web persistence has been tested only in the recorded Ubuntu/Firefox path.
- Word users can change the text within a control; invalid citation syntax then produces an explicit error on export.
- Untagged Quarto-looking text is left untouched rather than guessed to be a citation.
- Raw `.bib` entries are not rewritten, so BibLaTeX-specific fields are preserved in the source file.
- The exporter is a prototype CLI requiring Python and Pandoc on the export computer, **not** on colleagues' Word computers.
- Figures, advanced layout and DOCX formatting are delegated to Pandoc, which cannot always recreate Word layouts faithfully.
- The Office add-in has no server-side data storage. A locally chosen `.bib` file is only read in the browser, not uploaded.
- New bibliography markers use `bibword:refs:v2:<identity>:<style>`. Legacy `v1` markers remain exportable but have no stored style and therefore use Pandoc's default.
- The in-pane preview and exporter use the same bundled APA or Harvard CSL source. Exported Quarto, HTML, and DOCX output remains the final formatting authority.

## Licence and attribution

The independently written prototype code is provided under the MIT licence (see `LICENSE`). Third-party packages and licences are recorded in [`THIRD_PARTY_NOTICES.md`](THIRD_PARTY_NOTICES.md).


The included source was independently written as a proof of concept; the JabRef repository has **not been cloned or vendored** into this archive. When integrating JabRef upstream code later, retain its MIT licence and copyright notices.
