# Hosted pilot

The hosted Word add-in is deployed as static assets at:

- `https://pauljeco.github.io/BibWord/`
- manifest: `https://pauljeco.github.io/BibWord/manifest.xml`
- task pane: `https://pauljeco.github.io/BibWord/taskpane.html`

The production manifest is `office-addin/manifest.pages.xml`. It deliberately keeps the same add-in identity as the localhost manifest so a test document continues to recognize BibWord citations. Its version is incremented independently when the deployed manifest changes.

## One-time GitHub configuration

In the repository, open **Settings > Pages** and set **Build and deployment > Source** to **GitHub Actions**. Then run **Deploy hosted BibWord pilot** from the Actions tab, or push a change covered by the workflow's path filters.

The workflow installs the locked npm dependencies, builds `dist/csl-engine.js`, assembles only the required static assets, and deploys them through GitHub Pages. It does not publish development certificates, `node_modules`, the Python exporter, example documents, or bibliography files.

## Pilot installation

1. Download the hosted `manifest.xml`.
2. In Word for the web, use **Add-ins > More Add-ins > My Add-ins > Upload My Add-in** and select that file.
3. Open **BibWord hosted pilot** and run the Word-host smoke test in `docs/development.md`.

The manifest adds an **Open BibWord** button to the **Home** ribbon tab. If an older
pilot manifest is already installed, remove it and upload the current hosted manifest
again so Word refreshes the ribbon definition.

The hosted task pane reads user-selected `.bib`, `.bibtex`, and CSL files inside the browser. It does not upload them. Native SharePoint selection and the optional localhost synced-folder experiment are not part of this hosted pilot.
