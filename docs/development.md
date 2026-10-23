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
4. Open the BibWord task pane, select `examples/references.bib`, enter `smith2020`, and insert a citation.
5. Put the cursor inside the inserted citation, load it into the task pane, change its key or locator, and update it.
6. Put the cursor on a separate line where the references should appear and insert the bibliography marker.
7. Save and reopen the document. Download it as DOCX.
8. Inspect `word/document.xml` for tags beginning `bibword:cite:v1:` and `bibword:refs:v1:`.
9. Export it with the command below. Confirm that the `.qmd` contains a `#refs` Div at the marker position and that rendered references occur before any text following the marker.

```bash
.venv/bin/bibword-export downloaded.docx \
  --bib examples/references.bib \
  --out exported.qmd
```

Tenant policy can disable **Upload My Add-in**. That is an external deployment-policy blocker, not something the add-in can bypass.

## What automated checks do and do not prove

`make check` verifies the citation-syntax helpers, DOCX fixture conversion, Pandoc AST restoration, and Microsoft's manifest validation. It does not prove that a real Word host loads the pane or preserves content controls through save, reopen, copy/paste, coauthoring, and download. Record those manual results separately in `docs/word-host-test-matrix.md`.
