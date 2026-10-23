# Word host validation matrix

Real-host results are deliberately narrow. Do not generalize a successful test beyond the operations recorded here.

| Date | Host | OS/browser | Add-in loads | Citation inserted | Save/reopen retains tag | Downloaded DOCX retains tag | Export succeeds | Notes |
|---|---|---|---|---|---|---|---|---|
| 2026-10-23 | Word for the web | Ubuntu 24.04 / Firefox (version not recorded) | Yes | Yes, six controls | Not separately verified | Yes, all six tags | Yes; Quarto HTML and DOCX rendered against `bspi.bib` | Parenthetical, narrative, suppressed-author, locator, and two-key grouped-citation semantics parsed correctly from downloaded DOCX; grouped key order is preserved and all five unique keys resolve |
| Pending | Word desktop | Windows 11 | Not tested | Not tested | Not tested | Not tested | Not tested | Test a current Microsoft 365 build |

For each run, record the exact Word build, browser version, tenant/sideload method, document save path, and whether `word/document.xml` contains the expected `bibword:cite:v1:` tag.
