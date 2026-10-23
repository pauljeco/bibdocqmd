# Technical audit / initial design decision — 21 October 2026

## Foundation verification — 22 October 2026

- The supplied archive is internally consistent and contains a small Office.js task pane plus a Python/Pandoc exporter; there is no React or TypeScript implementation yet.
- The existing Python suite initially produced 13 passes and one environment failure because `python-docx` was documented but not declared. Project metadata now declares runtime and development dependencies explicitly.
- Microsoft `office-addin-manifest` 3.0.1 confirmed that the XML follows the current add-in schema. Its broader acceptance check initially rejected the prototype's `0.1.0.0` manifest version and missing support URL; both were corrected. The manifest now declares the WordApi 1.1 requirement used by range insertion and rich-text content controls.
- Pure citation construction and basic key scanning were separated from the Office host so they can be exercised by Node's built-in test runner.
- The local HTTPS server now reads credentials directly from `office-addin-dev-certs`; developers no longer copy private keys into the source tree.
- Real Word-for-the-web tests on Ubuntu/Firefox succeeded on 23 October 2026: eight citation controls and a bibliography marker survived save/reopen and DOCX download; bundled and custom CSL styles persisted; and a generated nine-entry bibliography retained italics and hanging indents. Copy/paste, collaboration, and Word desktop remain pending; automated tests and manifest validation are not treated as substitutes.
- The first formatted-citation toggle host test converted one citation and then returned `ApiNotFound`. The downloaded package showed that replacement and alias persistence had succeeded before Word rejected `Range.font.reset()`. The unsupported call was removed; the exporter successfully recovered all eight citations from the resulting mixed formatted/key-view document.
- The current `office-addin-dev-certs` 3.0.1 dependency brings a high-severity `node-forge` advisory through `mkcert`, with no npm-published fix reported on this date. It is a development-only certificate tool and is not shipped in the task pane, but it should be replaced or upgraded once a fixed upstream version is available.

## User-facing requirement

Word desktop (Windows) and Word for the web (Ubuntu/Linux browser) display **Quarto/Pandoc syntax** by default. `@key`, `[@key]`, grouped citations, `[-@key]`, and locators export to `.qmd` with actual Pandoc `Cite` nodes. BibTeX/BibLaTeX libraries remain the source of bibliographic records. CSL formatting is available for previews, the generated bibliography, and an optional document-wide citation reading view. In reading view, the canonical syntax is stored in the content-control alias and remains the export authority.

## Upstream project assessment: JabRef-Word-Addin

Source: <https://github.com/JabRef/JabRef-Word-Addin>

- MIT licence (reported in GitHub repository metadata and package.json).
- React + Office.js integration documented; `yarn start:web` / sideloading workflow in README.
- package.json at review: React ^17.0.2, citeproc ^2.4.63, Apollo client ^3.7.13, GraphQL ^16.6.0; external API endpoint required in README.
- Implication: reuse selectively, avoid inheriting backend coupling, and update/replace dated dependencies after an in-depth code audit.
- **Not cloned or executed here** because this runtime's shell cannot resolve github.com. The web-accessible README/package.json/manifest were inspected instead.
- Do not claim Word web support has been validated merely because the upstream README includes web sideloading instructions.

## Office.js APIs to verify on real hosts

- `Word.Range.insertText(...)`, returning a `Word.Range`.
- `Word.Range.insertContentControl()` around the inserted range, returning a `Word.ContentControl`.
- `Word.ContentControl.tag` / `title` persistence into DOCX (`w:sdtPr/w:tag`).
- Confirm that Word for the web produces *inline* controls, and that this exporter reads the saved Word XML in a native DOCX (not a LibreOffice re-save).

API references: <https://learn.microsoft.com/en-us/javascript/api/word/word.range> and <https://learn.microsoft.com/en-us/javascript/api/word/word.contentcontrol>

## Conversion prototype

- Uses Pandoc's DOCX reader for headings, paragraphs, lists, tables.
- Reads the original Word package directly, extracts tagged content controls, parses visible citation syntax with Pandoc (no custom grammar), and stores Pandoc `Cite` nodes.
- A temporary copy replaces controls with neutral sentinels.
- Inserts the native Cite nodes into Pandoc's document AST, then writes Quarto-compatible Markdown with YAML and `.bib` link.
- Sample citations render successfully using Pandoc's `--citeproc` with sample references.

## Product risks / priorities

1. **Word web validation**: content-control save/persist semantics and tenant sideloading rights.
2. **Reliable `.bib` support**: current task pane extracts basic keys, but cannot parse multiline/nested BibLaTeX records fully. Plan a real parser and library state management.
3. **Collaborative editing**: content-control merges, copy/paste, deleted citations, and duplicate ID handling require real-host tests.
4. **DOI-based metadata editing**: out of scope for first spike; optional future differentiation.
5. **Better Quarto exporter**: decide how to handle citations typed without a managed content control; avoid false positives inside prose/code.
6. **Advanced structure**: figures, footnotes, equations, cross-references and tracked changes require test documents.
7. **Cross-platform**: only `Office.js` task pane is envisaged; do not use Windows-only .NET/VSTO as core architecture.

## Explicit next engineering task

Test the citation-display toggle and copied controls in Word for the web, including save/reopen and DOCX export while formatted display is active. Then add a generated reference DOCX so Quarto's DOCX bibliography receives the same hanging-indent treatment as the bibliography generated directly in Word, and validate the complete workflow in a current Windows Word desktop build.
