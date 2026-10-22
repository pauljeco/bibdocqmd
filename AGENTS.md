# BibWord development instructions

These instructions apply to the entire repository.

## Product invariants

- BibTeX/BibLaTeX files are the authoritative bibliography sources. Do not create a proprietary citation database.
- Visible Quarto/Pandoc syntax is the default and canonical citation representation in Word. Preserve original citation keys exactly.
- Managed citations must retain enough structured identity to represent Pandoc citation groups, modes, locators, prefixes, suffixes, and a schema version. Never reconstruct identity from formatted author-year text.
- Support Word for the web and current Word desktop through Office.js. Do not introduce VSTO, COM, or another Windows-only core.
- Keep the basic add-in local-first and usable without JabRef, Zotero, Python, R, LaTeX, cloud authentication, or privileged Microsoft Graph permissions. Python and Pandoc are acceptable only for the interim developer exporter.
- Multiple bibliography files and document-, chapter-, and section-scoped bibliographies are core future requirements. Semantic scopes come from Word heading structure, not page-layout section breaks.
- Word-to-Quarto export must preserve citation keys, modes, order, locators, modifiers, and bibliography scope. Prefer Pandoc AST transforms over a new DOCX parser.

## Engineering rules

- Extend the v0.1 architecture incrementally until the upstream/competitor assessment justifies a larger rewrite.
- Prefer small, maintained, permissively licensed dependencies. Record attribution and do not silently import upstream code.
- Treat citation corruption or silent metadata loss as release blockers. Surface invalid and unresolved citations explicitly.
- Keep browser-facing code free of secrets and backend assumptions. Bibliography files selected in the task pane must remain local unless a later feature clearly tells the user otherwise.
- Add regression tests for parsing, insertion logic, metadata persistence, and export defects. Clearly distinguish simulated/fixture tests from tests performed in Microsoft Word.
- Never claim Word-host compatibility from unit tests alone. Record the Word version, platform, browser, save/reopen path, and downloaded-DOCX inspection for every real-host result.
- Preserve semantic Word structure during export where practical: headings, paragraphs, emphasis, lists, tables, links, media, captions, notes, equations, and cross-references.

## Repository checks

- Python/exporter: `make setup-python && make test-python`
- Office pane: from `office-addin/`, run `npm ci && npm run check`
- Full check after setup: `make check`
- Local task pane: from `office-addin/`, run `npm run cert`, then `npm run dev`

Do not commit private keys, local certificates, virtual environments, `node_modules`, downloaded user documents, or real bibliography data.
