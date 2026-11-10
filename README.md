# bibdocqmd | Quarto-native Word citations (technical spike v0.1)

A **working first development step** toward a BibTeX-native citation manager for Microsoft Word on Windows and **Word for the web**. The design preserves Quarto/Pandoc markup (`@key`, `[@key]`) directly in each Word citation control while allowing an optional formatted reading view.

> **Status:** DOCX → Quarto conversion and the task pane's citation-building logic are tested locally. Microsoft's validator accepts the manifest schema and configuration. On 23 October 2026, the pane loaded in Word for the web on Ubuntu; citation editing, save/reopen persistence, grouped citations, multiple scoped bibliographies, Citation.js/citeproc-js previews, bundled-style switching, custom CSL import, in-document bibliography generation, the document-wide formatted/editable citation toggle, formatted-view saving, citation copy/paste in both display modes, and safe cursor exit after a final bibliography all worked in the real host. Multiple bibliography files now pass automated combined-query and exporter tests but still need a real Word-host test. Coauthoring and Word desktop remain unverified. This is not yet a finished add-in or a fork of JabRef.

## Deliverables

- `examples/bibword_demo.docx`: sample Word document with **five genuine Word inline content controls**, each showing Quarto citation notation (fictional example citations).
- `examples/references.bib`: the matching demonstration bibliography (fictional sources, not scholarly references).
- `examples/bibword_demo.qmd`: exported Quarto document containing *real* Pandoc citations with their original keys, modes, and locators.
- `src/bibdocqmd/export.py`: runnable DOCX-to-Quarto converter, retaining Word document structure through Pandoc and replacing citation control markers with native Pandoc `Cite` nodes.
- `office-addin/`: experimental Office.js task pane. Load one or more `.bib` files locally or from public GitHub links, query their combined citation-key index, insert or update tagged citation syntax, preview references with bundled or imported CSL, generate formatted Word bibliographies, and control their positions with explicit markers.
- `tests/test_export.py`: regression tests for citation modes, preservation and invalid references.

## Why the visible syntax is the stored syntax

An in-Word citation *looks* like `[@smith2020]`, not `(Smith, 2020)`. This serves three purposes:

1. Writers learn the same citation notation used in Quarto and Pandoc.
2. Citation identity remains visible even if the Word add-in is unavailable.
3. The Word document can round-trip into semantic Pandoc `Cite` nodes without text-recognition heuristics.

The Word content control has a unique tag beginning `bibword:cite:v1:`. It is tagged so the exporter knows which substrings are managed citations, and does not replace ordinary text that just happens to look like `[@key]`. The `bibword` tag prefix remains in place so documents created before the bibdocqmd rename still open and export correctly.

The installable Python package is `bibdocqmd`; use `import bibdocqmd` or the `bibdocqmd-export` command. The former `bibword` import and `bibword-export` command remain available for existing scripts.

Quarto's canonical citation grammar is parsed by **Pandoc itself**, rather than by a custom regular expression.

## Developer quick start on Ubuntu

Requirements: Python 3.10+, Pandoc 3.x, Node.js 20+, npm, and Git.

```bash
make setup
make check
.venv/bin/bibdocqmd-export examples/bibword_demo.docx \
  --bib examples/references.bib \
  --out examples/bibword_demo.qmd \
  --title "bibdocqmd demonstration"
```

`make setup` creates a local Python virtual environment and uses the committed npm lockfile. See [`docs/development.md`](docs/development.md) for independent setup commands and the Word sideloading checklist.

Repeat `--bib path/to/library.bib` to export against several bibliography files. bibdocqmd validates citation keys across their union and rejects duplicates rather than silently choosing one record. The `.qmd`, all selected `.bib` files, and generated `bibdocqmd-reference-*.docx` must be distributed together. The reference DOCX is based on Pandoc's installed default and gives Quarto Word bibliographies a 0.5-inch hanging indent. A styled bibliography marker also causes the exporter to write a verified `bibdocqmd-*.csl` file. Multiple or heading-scoped markers additionally produce `bibdocqmd-multiple-bibliographies.lua`; keep that beside the QMD. Custom styles are embedded in the source DOCX, so the original external `.csl` file is not required at export time.

```bash
cd examples
pandoc bibword_demo.qmd --from markdown+citations --citeproc --to html \
  --output example_rendered.html
```

When using Pandoc directly, run from `examples/` (or set `--resource-path`) so it resolves `references.bib` correctly.
For multiple or heading-scoped bibliographies, render with Quarto so the QMD's bundled filter entry is applied:

```bash
quarto render exported.qmd --to html
quarto render exported.qmd --to docx
```

## Try the experimental Word add-in

For colleagues using **Word for the web**, use the hosted pilot's install page
and sample files; no local development setup is needed. The colleague onboarding
page is `office-addin/pilot.html`, deployed as the site's `index.html`. See
[the colleague pilot checklist](docs/colleague-pilot.md) before sharing a release.
The commands below are for local development.

The minimal task pane is deliberately dependency-light and does **not** use the JabRef GraphQL backend.

1. Install Node.js 20 or newer, then run `make setup-addin` from the repository root.
2. In `office-addin/`, run `npm run cert` and `npm run cert:verify` to generate and trust local HTTPS credentials.
3. Run `npm run dev` (serves `https://localhost:3000/`). Certificate files are read from the development certificate store and are not copied into the repository.
4. Open a document in Word for the web. If your tenant allows developer sideloading, upload `office-addin/manifest.xml` through the Word Add-ins UI; see [Microsoft's sideloading guide](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).
5. In the task pane, add one or more `.bib` files. Repeat the file selection to add libraries from other directories; **Clear libraries** resets the combined set. To load a public GitHub bibliography, paste its `.bib` or `.bibtex` file link and choose **Load or refresh GitHub file**. For a locally synced or rclone-mounted SharePoint directory, start the server with `BIBDOCQMD_LIBRARY_DIR="/absolute/path/to/folder" npm run dev`, then use **Browse synced SharePoint** and **Load or refresh selected**. Enter a key from the combined index, choose a bundled style or import an independent `.csl` file, inspect the CSL-formatted preview, and insert the citation.
6. To edit, put the Word cursor inside the blue citation, choose **Load selected citation**, change its fields, then choose **Update selected**.
7. Choose **Show formatted citations** for a reading view or **Show editable keys** to restore the canonical Quarto syntax. bibdocqmd remembers source names and hashes in the document. After reopening Word, choose **Reload saved GitHub files** for remembered GitHub links, and reselect local files.
8. Choose **Refresh document preview** to format all unique managed references in the pane.
9. Put the cursor where references must appear, choose **Entire document** or the current Heading 1–3 scope, and choose **Insert marker**. Add one scoped marker to each chapter or section that needs its own list. Scoped markers require a preceding Word heading of the chosen level.
10. Choose **Update in Word** to rebuild every marker from the citations in its scope. A source cited in two chapters appears in both chapter bibliographies. Repeat this after citation, library, heading, or style changes.
11. Save, reopen, and download the document as `.docx`; inspect the retained content-control tags and run `bibdocqmd-export` against the downloaded file.

**Limitations:** browsers require a trusted HTTPS certificate and some Microsoft 365 tenants disallow sideloading. Word-for-the-web insertion, selected-citation editing, bibliography-marker insertion, bibliography generation, formatted/editable citation display, save/reopen persistence, citation copy/paste in both display modes, custom-style embedding, download, and export have been verified in one Ubuntu/Firefox/Microsoft 365 path. Coauthoring, Word desktop, and broad host compatibility remain pending real-host validation. The manifest passes Microsoft's automated validator, but that does not prove runtime compatibility. Citation.js provides local BibLaTeX parsing and citeproc-js provides CSL previews. APA 7th edition and Cite Them Right Harvard are bundled; independent CSL files can also be imported. Word stores recoverable citation syntax plus a visible, versioned bibliography-position marker containing the selected style identifier.

## Architecture and next steps

The pane also has an **experimental / in development** local **Export to Quarto with Word Markdown** workflow:
prepare a disposable DOCX copy, export that copy with Word for the web's native
Markdown feature (where available), and import the exported Markdown to download
a QMD package with bibliography sources, CSL, scoped filter, and supplied media.
No Python or conversion server is required for this route. Automated integration
tests currently use Pandoc as a stand-in; native Word Markdown export still needs
a real-host test. See [the Markdown export guide](docs/markdown-export.md).

Export development is parked at the current prototype. The next priority is the
colleague pilot in Word for the web, using the hosted add-in, short install guide,
and fictional sample files. A one-button exporter is not part of this phase.

See `docs/technical_audit.md` for the upstream repository assessment, `docs/development.md` for reproducible setup, and `docs/word-host-test-matrix.md` for real-host evidence.

```text
BibTeX file → Office.js task pane → tagged Quarto citations in Word
                      │                         + bibliography marker
                      └→ local preview                  │
                                             saved Word DOCX
                                                    │
                                       citation-aware converter
                                                    │
                           Pandoc Cite nodes + scoped reference Divs
                                                    │
                                               report.qmd
                                                    │
                                            Pandoc/CSL output
```

Next: verify combined querying across multiple `.bib` files in Word for the web, including a cross-file grouped citation and duplicate-key error, then run the complete workflow in **Windows Word**.

## Known limitations and test boundaries

- Citation controls must be **inline controls in the DOCX main document body**. The one supported block control is the bibliography-position marker; headers, footnotes, text boxes, cross-references, and tracked-change contexts remain unsupported.
- Citation copy/paste in formatted and key views has been verified in the recorded Word-for-the-web Ubuntu/Firefox path. Concurrent editing remains unverified.
- A sentence mark pasted into a blue control becomes part of the managed citation and makes that control invalid. Batch operations identify the citation number and text; insertion and selected-citation updates now move the caret outside the control automatically.
- Word users can change the text within a control; invalid citation syntax then produces an explicit error on export.
- Untagged Quarto-looking text is left untouched rather than guessed to be a citation.
- Raw `.bib` entries are not rewritten, so BibLaTeX-specific fields are preserved in the source file.
- The exporter is a prototype CLI requiring Python and Pandoc on the export computer, **not** on colleagues' Word computers.
- Figures, advanced layout and DOCX formatting are delegated to Pandoc, which cannot always recreate Word layouts faithfully.
- The Office add-in has no server-side data storage. A locally chosen `.bib` file is only read in the browser, not uploaded. Public GitHub `.bib` links are fetched directly in the browser; private repository authentication is not supported. The optional synced-folder selector reads `.bib` files through the local development server from the one directory named by `BIBDOCQMD_LIBRARY_DIR`; it neither writes to that directory nor connects directly to SharePoint.
- New bibliography markers use `bibword:refs:v3:<identity>:<style>:<scope>`, where scope is `document`, `heading1`, `heading2`, or `heading3`. Legacy `v1` and `v2` markers remain exportable as whole-document markers.
- All markers currently share one selected CSL style and use the union of all selected `.bib` sources. Bibliography scope controls which cited keys appear, independently of the source file containing each key.
- Imported independent CSL styles use a `custom-<hash>` style identifier and a namespaced custom XML part inside the DOCX. The exporter verifies the full SHA-256 hash before using the embedded data. This path requires WordApi 1.4; bundled styles continue to work on the prototype's WordApi 1.1 baseline.
- The generated Word bibliography lives inside the bibliography marker. The exporter intentionally discards those display paragraphs and regenerates the final bibliography from managed citation keys, preventing duplicate or stale references in Quarto output.
- Formatted citations retain their canonical Quarto syntax in the content-control alias. Switching back to editable keys restores that source, and the exporter reads it directly when the formatted display is active. Direct edits to formatted display text do not change citation semantics; use **Load selected citation** or return to key view before editing.
- The in-pane preview and exporter use the same bundled or embedded CSL source. Exported Quarto, HTML, and DOCX output remains the final formatting authority.
- The bibliography generated directly in Word and Quarto's exported DOCX both use hanging indents. The exporter emits a content-addressed reference DOCX based on the locally installed Pandoc version and links it from the QMD front matter.

## Licence and attribution

The independently written prototype code is provided under the MIT licence (see `LICENSE`). Third-party packages and licences are recorded in [`THIRD_PARTY_NOTICES.md`](THIRD_PARTY_NOTICES.md).


The included source was independently written as a proof of concept; the JabRef repository has **not been cloned or vendored** into this archive. When integrating JabRef upstream code later, retain its MIT licence and copyright notices.
