# Word host validation matrix

No real Microsoft Word host has been validated in this repository yet. Do not convert blank cells into claims of support.

| Date | Host | OS/browser | Add-in loads | Citation inserted | Save/reopen retains tag | Downloaded DOCX retains tag | Export succeeds | Notes |
|---|---|---|---|---|---|---|---|---|
| Pending | Word for the web | Ubuntu 24.04 / Firefox | Not tested | Not tested | Not tested | Not tested | Not tested | Requires tenant sideloading access |
| Pending | Word desktop | Windows 11 | Not tested | Not tested | Not tested | Not tested | Not tested | Test a current Microsoft 365 build |

For each run, record the exact Word build, browser version, tenant/sideload method, document save path, and whether `word/document.xml` contains the expected `bibword:cite:v1:` tag.
