# Ubuntu development workflow

## Prerequisites

- Ubuntu 24.04
- Git
- Python 3.10 or newer
- Pandoc 3.x
- Node.js 20 or newer with npm
- Firefox for Word for the web testing
- A Microsoft 365 tenant that permits add-in sideloading

Quarto is useful for end-to-end rendering but is not required by the unit tests.

## Set up a clean checkout

```bash
make setup
make check
```

This creates `.venv`, installs the exporter and its test dependencies, installs the locked npm dependencies with `npm ci`, runs Python and JavaScript tests, and validates the Office manifest.

The add-in test and development commands build `office-addin/dist/csl-engine.js` locally from the pinned Citation.js components. The generated bundle is intentionally ignored by Git.

To set up the parts independently:

```bash
make setup-python
make setup-addin
make test-python
make test-addin
make validate-manifest
```

## Run the task pane over HTTPS

Office add-ins require HTTPS, including local development. Generate and trust Microsoft's development certificate, then start the server:

```bash
cd office-addin
npm run cert
npm run cert:verify
npm run dev
```

Open `https://localhost:3000/taskpane.html` in Firefox first and confirm that it loads without a certificate warning. The server reads the certificate from the location managed by `office-addin-dev-certs`; no certificate or private-key files need to be copied into the repository.

The certificate tool may request `sudo` on Linux to add its local development CA to the operating-system trust store. Firefox may need to be restarted after certificate installation.

## Sideload in Word for the web

1. Keep the HTTPS development server running.
2. Open a Word document in Word for the web.
3. Use **Add-ins** > **More Add-ins** > **My Add-ins** > **Upload My Add-in** and select `office-addin/manifest.xml`. The exact labels can vary by tenant.
4. Open the BibWord task pane, select `examples/references.bib`, enter `smith2020`, confirm the APA CSL preview, and insert a citation.
5. Put the cursor inside the inserted citation, load it into the task pane, change its key or locator, and update it.
6. Choose **Show formatted citations**, then **Show editable keys**. Confirm that all citations change together and the exact original Quarto syntax returns. Format them again, select one formatted citation, load it, update it, and confirm only that occurrence changes.
7. Copy and paste one formatted citation. Load the pasted occurrence, update its locator or key, and confirm the original occurrence is unchanged. Copy a key-view citation and repeat the same check.
8. Refresh the document bibliography preview, change between APA and Harvard, and confirm that all unique managed keys are represented. Optionally import an independent `.csl` file and verify its previews.
9. Put the cursor on a separate line where the references should appear and insert the bibliography marker. Change the style once more after insertion, then choose **Update in Word**. Confirm that the marker now contains the heading and all formatted references.
10. Save and reopen the document. Reload the `.bib` file, confirm that the saved style and generated bibliography remain, and toggle formatted citations back to editable keys. Then download the document as DOCX.
11. Inspect `word/document.xml` for tags beginning `bibword:cite:v1:` and a marker tag matching `bibword:refs:v2:<identity>:<style>`. Formatted controls must have a `w:alias` beginning `BibWord citation source v1:`. For a custom style, also confirm that the DOCX contains a `customXML/item*.xml` part in the `urn:bibword:csl-style:v1` namespace.
12. Export it with the command below. Confirm that the `.qmd` contains the matching `csl` path, the generated `bibword-reference-*.docx` path, and a `#refs` Div at the marker position. Confirm that formatted display text does not leak into the QMD, generated Word references are not duplicated, rendered references occur before any text following the marker, and Quarto DOCX references use a 0.5-inch hanging indent.

If a document-wide operation reports `Citation N "..." is invalid`, inspect that numbered blue control. Sentence punctuation must sit outside the control; the diagnostic adds an explicit hint when it detects text after a closing citation bracket. Display switching remains all-or-nothing so a failed batch does not leave the document in a mixed state.

```bash
.venv/bin/bibword-export downloaded.docx \
  --bib examples/references.bib \
  --out exported.qmd
```

Tenant policy can disable **Upload My Add-in**. That is an external deployment-policy blocker, not something the add-in can bypass.

Custom CSL persistence uses `Word.CustomXmlPart`, which requires WordApi 1.4. The pane rejects malformed styles, dependent styles that need a separately fetched parent, styles without both citation and bibliography layouts, and files larger than 2 MB. The imported file remains local and is embedded as base64 in a namespaced document part. Its short marker identifier and full stored SHA-256 digest are verified by the exporter.

## What automated checks do and do not prove

`make check` verifies the citation-syntax helpers, DOCX fixture conversion, Pandoc AST restoration, and Microsoft's manifest validation. It does not prove that a real Word host loads the pane or preserves content controls through save, reopen, copy/paste, coauthoring, and download. Record those manual results separately in `docs/word-host-test-matrix.md`.
