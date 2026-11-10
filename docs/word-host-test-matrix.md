# Word host validation matrix

Real-host results are deliberately narrow. Do not generalize a successful test beyond the operations recorded here.

| Date | Host | OS/browser | Add-in loads | Citation inserted | Save/reopen retains tag | Downloaded DOCX retains tag | Export succeeds | Notes |
|---|---|---|---|---|---|---|---|---|
| 2026-10-23 | Word for the web | Ubuntu 24.04 / Firefox (version not recorded) | Yes | Yes, eight citation controls plus one bibliography marker | Yes; all eight pre-existing control identities survived, including an edited grouped citation and the marker | Yes, all nine tags; the marker was upgraded to `bibword:refs:v2:mvkrm25o-vbwn7b:harvard1` | Yes; Quarto DOCX rendered against `bspi.bib` with the bundled Harvard CSL | **Update selected** preserved the grouped control tag while adding a third key; a new narrative citation also persisted. Parenthetical, narrative, suppressed-author, locator, and three-key grouped-citation semantics parsed correctly. Citation.js/citeproc-js previews and live APA/Harvard switching worked in the pane. The downloaded DOCX retained the selected Harvard style; export rendered a titled eight-entry Harvard reference section at the marker before following document text. Whole-document preview refresh was not separately recorded. |
| 2026-10-23 | Word for the web | Ubuntu 24.04 / Firefox (version not recorded) | Yes | Reused eight controls; the final grouped citation gained a second key | Yes; custom APA reopened as the selected default | Yes; marker retained `custom-b3b87c2d572d` and the 66 KB style survived in `customXML/item.xml` | Yes; Quarto rendered nine APA references from the emitted `bibword-custom-b3b87c2d572d.csl` | This validates the WordApi 1.4 custom XML path through import, marker update, close/reopen restoration, download, integrity verification, export, and DOCX rendering. The real package exposed a case-sensitive `customXML/` path difference from the fixture; exporter matching and the regression fixture were corrected. |
| 2026-10-23 | Word for the web | Ubuntu 24.04 / Firefox (version not recorded) | Yes | Reused eight controls containing nine unique keys | Yes; the custom-APA bibliography survived in the marker after reopen | Yes; the marker contains one heading and nine reference paragraphs with italics and 0.5-inch hanging indents | Yes; generated display paragraphs were discarded and Quarto rendered each of the nine references once | **Update in Word** generated the bibliography at the explicit marker. The downloaded `test 8.docx` retained the bibliography, custom CSL, eight citation controls, and following document text. Two-page visual inspection found no clipping or duplication; Quarto rendered the citations and reference section successfully. |
| 2026-10-23 | Word for the web | Ubuntu 24.04 / Firefox (version not recorded) | Yes | **Show formatted citations** converted only the first of eight controls before `ApiNotFound` | Not reached | `test 9.docx` retained one formatted control with canonical alias plus seven unchanged key-view controls | Yes; the mixed-view document exported all eight controls and nine unique keys correctly | Package evidence isolated the failure to the queued `Range.font.reset()` call after the first replacement. Word had already persisted `(Abarca et al., 2022)` and its `[@abarca2022spo-a]` alias. The unsupported reset call was removed and a compatibility regression test added. |
| 2026-10-23 | Word for the web | Ubuntu 24.04 / Firefox (version not recorded) | Yes | User confirmed **Show formatted citations** and **Show editable keys** now work across the document | Yes for the returned key view; all eight original control identities remained | Yes; `test 10.docx` contains all eight citation controls, the bibliography marker, and the embedded custom APA style | Yes; Quarto rendered all eight controls and nine unique references once, with following text preserved | This validates the formatted → editable round trip after removing `Font.reset`. The downloaded file was saved in key view, so formatted aliases were correctly removed. Two-page source and export renders were clean. Copy/paste was not separately recorded. |
| 2026-10-23 | Word for the web | Ubuntu 24.04 / Firefox (version not recorded) | Yes | Eleven formatted controls after copying and editing citations | Saved in formatted view; explicit close/reopen was not separately reported | Yes; every citation retained a canonical source alias. Two pasted controls reused the source BibWord tag but received distinct Word `w:id` values | Yes; all 11 physical controls exported as 14 citation items over 10 unique keys, and Quarto generated 10 hanging-indent references | `test 11.docx` validates formatted-view save and copy/paste tolerance. The original and pasted Jasiewicz controls share a BibWord tag but preserve different parenthetical and narrative aliases, showing that the edited occurrence remained independent. The saved in-Word bibliography still had nine entries because it was not refreshed after Bennett was added; export correctly regenerated all ten. Two-page source and export renders were clean. |
| 2026-10-23 | Word for the web | Ubuntu 24.04 / Firefox (version not recorded) | Yes | Thirteen controls after key-view copies and independent grouped-citation edits | Saved and later formatted successfully; explicit close/reopen was not separately reported | Yes; copied controls reused source BibWord tags but received distinct Word `w:id` values and retained independent canonical aliases | Yes; `test 13.docx` exported 13 controls as 19 citation items over 12 unique keys, and Quarto generated all 12 references | `test 12.docx` first exposed an unrelated malformed Bennett control containing `[@bennett1998-spatial]. ` with punctuation inside the control; that one control blocked the all-or-nothing display and bibliography batches. After the orphan was deleted, `test 13.docx` switched to formatted display and exported cleanly. Diagnostics now identify the citation number and text and advise keeping punctuation outside the blue control. Two-page source and export renders were clean. |
| 2026-10-23 | Word for the web | Ubuntu 24.04 / browser version not recorded | Yes | Reused 14 managed citation controls | Multiple-bibliography update succeeded; explicit close/reopen was not separately reported | Yes; `test 15.docx` retains 14 citation controls and two distinct v3 bibliography markers | Not tested in this run | The paragraph-based scope reader fixed the prior live-host marker lookup failure. The document contains a 14-entry document bibliography and a 2-entry Heading 1 bibliography using the embedded custom APA style. Three-page visual inspection found both lists correctly positioned and formatted. Package inspection also showed that the final bibliography was the last body element, with no ordinary paragraph after its content control; this motivated automatic creation of one normal exit paragraph without duplicating it on later updates. |
| 2026-10-23 | Word for the web | Ubuntu 24.04 / browser version not recorded | Yes | Fifteen managed citation controls | User confirmed the normal-paragraph exit after a bibliography; explicit close/reopen was not separately reported | Yes; `test 16.docx` retains 15 citation controls, three v3 bibliography markers, and one ordinary paragraph immediately after every marker | Yes; all 15 controls exported with two `--bib` sources and Quarto rendered the three scoped lists to DOCX | This validates the cursor-exit repair in the real host and downloaded package. The document contains one document bibliography and two Heading 1 bibliographies, including a final one-entry list followed by an ordinary paragraph. The second export bibliography contained no cited keys, so cross-source citation lookup still requires a Word-host test. Both three-page source and exported-DOCX renders were clean. |
| Pending | Word desktop | Windows 11 | Not tested | Not tested | Not tested | Not tested | Not tested | Test a current Microsoft 365 build |

For each run, record the exact Word build, browser version, tenant/sideload method, document save path, and whether `word/document.xml` contains the expected `bibword:cite:v1:` tag.

## Bibliography paragraph style regression (2026-11-10)

A user testing Word on Windows reported a two-entry bibliography whose final
entry had the blue, enlarged appearance of the References heading. The exact
Word build and desktop/web host were not recorded; the screenshot alone does
not establish the stored paragraph style. No downloaded DOCX was inspected.

Bibliography updates now explicitly set the first paragraph inside each marker
to the scope's heading style and all subsequent paragraphs to Normal after HTML
import. Simulated Office.js regression tests cover heading inheritance, all four
scopes, multiple markers, and repeated updates with a changed entry count. A
follow-up user test reported that hanging indents were lost after the style
change. Updates now explicitly apply a 36-point left indent and a -36-point
first-line indent to entries after setting Normal. A further screenshot showed
the first entry shifted right relative to the following entry. The pane now
offers a hanging indent in cm (default 1.27, range 0–5, with 0 disabling it),
applied to every marker by Update in Word. Bibliography HTML is flattened into
sibling entry paragraphs, and native styles finish synchronizing before indents
are applied. Every entry receives the same left and negative first-line indents,
zero right indent, and left alignment. The title receives zero indents.

Simulated tests use actual Citation.js HTML and check flat paragraph boundaries,
inline italics/links, equal first-line positions, custom and zero indents, invalid
input, all four scopes, and repeated updates. They also simulate style application
clearing imported indents during synchronization. These tests do not establish
real-host compatibility or rendered formatting fidelity.

User-reported host result (2026-11-10): after the configurable-indent and
first-entry alignment changes, the user confirmed that tests ran successfully in
Microsoft Word, in the Windows testing context above. This records user
confirmation of the bibliography changes; individual test cases and indent values
were not specified. The exact Word version/build, Windows version, desktop/web
host, browser/version (if applicable), and sideload method were not recorded.
The user subsequently confirmed that closing and reopening the file, refreshing,
and editing entries all succeeded in Microsoft Word. This is a user-reported
real-host result; the save location/reopen path and the specific refresh/edit
operations were not recorded. Downloaded-DOCX inspection and export were not
reported for this run.

Remaining detailed host validation: update a two-entry bibliography, confirm only its title
appears in Word's Navigation pane, and verify title italics and hanging indents
remain. Inspect the downloaded DOCX paragraph styles and citation/marker
identities after the confirmed close/reopen, refresh, and editing workflow.
Try 0, 0.75, and 1.27 cm, confirming that
the first line of every reference starts at the same left margin and every
wrapped line uses the selected indent. Record the host/build and save path.

## Native Word Markdown export (implementation validation)

The new two-stage pane workflow prepares a separate DOCX with protected citation
and bibliography placeholders, then finalizes Word-exported Markdown as a QMD
package. No connected Codex Word session was available during implementation.
There is no real-host result for the native Markdown writer yet.

Automated fixture results: the five-citation example DOCX was prepared, converted
to GFM by Pandoc as a stand-in, finalized as QMD, parsed back into five Cite nodes
including narrative and suppressed-author modes, and rendered with references.
A two-chapter fixture rendered one scoped bibliography per marker using the
existing Lua filter. Simulated pane tests verified local preparation/finalization
downloads without calling Word mutation APIs. See `docs/markdown-export.md` for
the workflow and validation boundaries.

Pending real-host result: record Word build, platform, browser, source save/reopen
path, prepared-DOCX inspection, native `.md` inspection, and final Quarto render.
Check grouped citations with locators, formatted aliases, copied citations,
multiple bibliography sources, custom CSL, chapter scopes, and supplied media.

## File chooser investigation (2026-11-10)

The user reported that all five Choose file controls stopped opening a dialog,
including the existing bibliography and CSL controls. The exact Word build and
desktop/web host were not supplied. No connected Word session was available for
inspection; the underlying Word-host cause is unconfirmed.

File inputs now synchronously request `HTMLInputElement.showPicker()` during the
click when available. Success cancels the duplicate native default; unsupported
or blocked calls retain that default, and exceptions display feedback beside the
affected input. Shared click handling leaves file inputs alone. A local-tools
browser link provides a separate page without Office.js; files must be selected
again there, and no browser-to-Word file transfer is claimed.

Automated tests cover all five inputs, synchronous activation, older-webview
fallback, blocked-picker feedback, cancellation, disabled inputs, browser-mode
startup, and the Office browser-opening API/fallback. A headless Chrome
157.0.8103.39 smoke check on Windows observed exactly one intercepted
`Page.fileChooserOpened` event for each input: multiple selection for `.bib`,
single selection for CSL/DOCX/preparation ZIP/Markdown. No runtime exceptions were
reported. This checks browser picker activation, not a visible OS dialog or
Microsoft Word. Word-host retesting remains pending.

### Desktop Word follow-up (2026-11-10)

The user confirmed that the embedded picker still failed after reloading and
showed a separate file-selection window stuck at "Connecting to Word". Inspection
of the live desktop pane confirmed the localhost assets and the absence of a
JavaScript exception during a pane chooser request. A syntax error in the newly
added dialog script prevented it from initializing; this was corrected and the
actual script is now executed by a regression test, in addition to syntax checks.

Windows desktop now routes file-input clicks through a same-origin Office dialog
with an unmodified native input. Files return to the originating pane field as
bounded, validated binary chunks; origin, transfer token, sequence, declared
sizes, and completion are checked before a FileList is assigned and one change
event is dispatched. Native browser and non-PC paths retain their previous
selection behavior. The task pane and dialog assets are deployed together.

Real-host environment: Word desktop **16.0.20527.20150**, Windows **25H2 build
26200.9711**, embedded Edge WebView2 **157.0.4341.42**, localhost HTTPS development
manifest. The Word debugging connection on port 9229 was used for inspection.

Observed results:

- The corrected Office dialog initialized with enabled file and Cancel controls.
- A native Windows **Open** dialog was visibly present after requesting file
  selection from that Office dialog (verified through its Windows dialog handle).
- `examples/bibword_demo.docx` (37,540 bytes, fictional demonstration content) was
  assigned to the dialog input through CDP, transferred through Office's message
  bridge, and read back from the originating task-pane FileList. SHA-256 matched
  the original fixture. This verifies the actual Office transfer route but uses
  automated file assignment rather than automating the OS filename selection.
- The export input's previous FileList was restored after the fixture test, and
  the temporary dialog/chooser windows were cancelled. The Word document was not
  edited by these picker tests.

Save/reopen path and downloaded user-DOCX inspection were not exercised in this
picker-only run. Native Word Markdown export was not exercised. The earlier
browser-only chooser test must not be taken as validation of the embedded pane's
native picker, which remained unsuccessful in the user's desktop run.
