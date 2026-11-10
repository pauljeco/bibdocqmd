# Export through Word's native Markdown writer

**Status: Experimental / in development. Development parked at the current
prototype, at the user's request (2026-11-10).** The manual workflow below remains
available behind a collapsed experimental section. The proposed one-button export
and browser-based Pandoc conversion are not implemented. The current priority is
making the hosted citation and bibliography pilot easy for colleagues to use in
Word for the web.

The task pane's **Export to Quarto with Word Markdown** section prepares a
disposable DOCX copy, then turns Word's exported Markdown into a QMD project.
It runs locally in the browser, including the hosted pilot. It needs no Python,
Pandoc, local server, or document-upload service on the exporting computer.
Rendering the resulting project still requires Quarto on the rendering computer.

Microsoft announced `.md`/`.docx` conversion in **Word for the web** in its
[October 2026 update](https://techcommunity.microsoft.com/blog/microsoft-copilot-blog/what%E2%80%99s-new-in-microsoft-copilot--october-2026/4568528).
This workflow requires that Word's Markdown export is available to your account;
it does not add that feature to desktop Word or call an undocumented export API.

If Word does not open file or download dialogs, use **Open local tools in
browser** beside the file controls. This opens the generated `local-tools.html`
page, which runs the same local library/export tools without loading Office.js
or waiting for Word initialization. Select the files again there; selections are
not transferred between the browser and Word pane. Word document editing remains
in the Word pane. `npm run build` generates this page, and the hosted pilot
workflow deploys it alongside the pane.

In Windows desktop Word, clicking a file input now opens a separate Word file
selection window. Choose the file(s) there; their bytes are returned directly to
the original field in the pane, and the window closes. This route applies to
bibliography, CSL, DOCX, preparation-ZIP, and Markdown inputs. Select at most
25 MB in one dialog transfer. Close the selection window before reloading the
task pane; reloading an owning pane can leave its existing dialog disconnected.

## Steps

1. Save the original document as `.docx`. Load its matching `.bib` files in the
   pane. The saved DOCX, rather than unsaved changes in the open document, is used.
2. Under **Export to Quarto with Word Markdown**, choose the saved DOCX and click
   **Download preparation ZIP**.
3. Extract that ZIP. Open **word-markdown-copy.docx** in Word for the web and use
   Word's Markdown export. Citation and bibliography placeholders are intentional:
   do not edit this disposable copy. Keep editing the original document and
   prepare a fresh export copy when it changes.
4. In the pane, choose the original preparation ZIP and Word's exported `.md`.
   If Word exports accompanying images, select a ZIP containing the `.md` and
   those files with their original relative paths. Exactly one `.md` is allowed.
5. Click **Download QMD package** and extract it. Open **document.qmd**, keeping
   the included `.bib`, `.csl`, optional Lua filter, and media beside it. If the
   Word export ZIP has a containing directory, the QMD is placed in that directory.
6. Render with Quarto. Compare headings, emphasis, lists, tables, images, notes,
   equations, captions, and cross-references with the original Word document.
   Only media included in the selected ZIP is packaged; retain any separately
   exported files alongside the project.

The source bibliographies are copied without rewriting their contents, using
`bibliography-1.bib`, `bibliography-2.bib`, etc. Their original filenames and hashes
are listed in the preparation plan. The saved document's bibliography marker
selects its CSL style; when no marker records a style, the pane selection is used.

## Citation and scope preservation

Preparation only changes known managed content controls in `word/document.xml`
inside a new DOCX package. All other package members, including styles, media,
relationships, and notes, retain their original bytes. General document conversion
is delegated to Word; this is not a new general DOCX-to-Markdown parser.

Each physical citation receives a unique plain-text placeholder, even when copied
citations share the same managed tag. The temporary export plan retains the
canonical citation syntax and control tag. Formatted citations use their stored
source alias, never reconstructed author-year text. Restoring whole syntax strings
preserves key case, citation order, modes, and supported modifiers. The existing
v0.1 citation parser validates sources; unsupported syntax is rejected explicitly.

Generated reference lists are replaced by standalone placeholders in the copy,
then restored as Quarto bibliography markers. Multiple or heading-scoped markers
use the existing `multiple_bibliographies.lua` filter; Quarto regenerates their
lists from the bibliography sources. The selected bundled or integrity-checked
embedded custom CSL is included.

Finalization requires every placeholder exactly once and rejects missing,
duplicated, foreign, or code-formatted placeholders. It checks the plan, source,
and CSL hashes, and validates citation keys again. It does not globally unescape
Markdown or guess citation identities. ZIPs are limited to 25 MB compressed and
100 MB expanded; unsafe and duplicate paths are rejected.

Managed citations outside the main body, block/nested managed citations,
bibliography markers in tables or mixed with other paragraph content, invalid
heading scopes, and unresolved keys stop preparation. Accept or reject tracked
changes before using this export path. Word's handling of other advanced document
features remains a host-dependent conversion limitation.

## Validation boundary

Automated tests cover local pane downloads, known-control replacement, source
and package-member preservation, citation modes/groups/locators, copied-control
tags, multiple sources, custom CSL, scoped markers, media paths, and explicit
failure cases. Pandoc is used as a **stand-in** for Word's Markdown writer in
integration tests; the resulting QMD is parsed into Cite nodes and the scoped Lua
filter is exercised during rendering. This does **not** validate Microsoft's
native Markdown writer.

No connected Word session was available during implementation. A real-host test
must record the exact Word version/platform, browser, source save/reopen path,
prepared DOCX and native Markdown inspection, and rendered QMD results in
[`word-host-test-matrix.md`](word-host-test-matrix.md).

The direct Python/Pandoc `bibdocqmd-export` command remains available when native
Markdown export is unavailable or when its fidelity is insufficient.
