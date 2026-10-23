# Word host validation matrix

Real-host results are deliberately narrow. Do not generalize a successful test beyond the operations recorded here.

| Date | Host | OS/browser | Add-in loads | Citation inserted | Save/reopen retains tag | Downloaded DOCX retains tag | Export succeeds | Notes |
|---|---|---|---|---|---|---|---|---|
| 2026-10-23 | Word for the web | Ubuntu 24.04 / Firefox (version not recorded) | Yes | Yes, eight citation controls plus one bibliography marker | Yes; all eight pre-existing control identities survived, including an edited grouped citation and the marker | Yes, all nine tags | Yes; Quarto DOCX rendered against `bspi.bib` | **Update selected** preserved the grouped control tag while adding a third key; a new narrative citation also persisted. Parenthetical, narrative, suppressed-author, locator, and three-key grouped-citation semantics parsed correctly. The marker placed a titled eight-entry reference section before following document text. The later Citation.js/citeproc-js preview was confirmed working in the pane; whole-document refresh and styled `v2` marker persistence still require a host test. |
| Pending | Word desktop | Windows 11 | Not tested | Not tested | Not tested | Not tested | Not tested | Test a current Microsoft 365 build |

For each run, record the exact Word build, browser version, tenant/sideload method, document save path, and whether `word/document.xml` contains the expected `bibword:cite:v1:` tag.
