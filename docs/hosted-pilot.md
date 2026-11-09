# Hosted pilot

The hosted Word add-in is deployed as static assets at:

- `https://pauljeco.github.io/bibdocqmd/`
- manifest: `https://pauljeco.github.io/bibdocqmd/manifest.xml`
- task pane: `https://pauljeco.github.io/bibdocqmd/taskpane.html`

These URLs become active after the GitHub repository is renamed to `bibdocqmd` and the Pages workflow has deployed. GitHub does not redirect the old `/BibWord/` Pages URL.

The production manifest is `office-addin/manifest.pages.xml`. It deliberately keeps the same add-in identity as the localhost manifest so a test document continues to recognize bibdocqmd citations. Its version is incremented independently when the deployed manifest changes.

## Move the pilot to the new repository name

1. In the existing `pauljeco/BibWord` repository, open **Settings > General**, change **Repository name** to `bibdocqmd`, and choose **Rename**.
2. In this checkout, run `git remote set-url origin https://github.com/pauljeco/bibdocqmd.git`.
3. Review and commit the intended local changes, including `office-addin/github-library.js`, the hosted manifest, and the Pages workflow, then push `main`.
4. In **Settings > Pages**, set **Build and deployment > Source** to **GitHub Actions**. The push triggers **Deploy hosted bibdocqmd pilot**; it can also be run manually from **Actions**.
5. Open the deployed site and its `manifest.xml` and `taskpane.html` URLs above. In Word, remove the old sideloaded pilot manifest and upload the newly downloaded one.

The workflow installs the locked npm dependencies, builds `dist/csl-engine.js`, assembles only the required static assets, and deploys them through GitHub Pages. It does not publish development certificates, `node_modules`, the Python exporter, example documents, or bibliography files.

## Pilot installation

1. Download the hosted `manifest.xml`.
2. In Word for the web, use **Add-ins > More Add-ins > My Add-ins > Upload My Add-in** and select that file.
3. Open **bibdocqmd hosted pilot** and run the Word-host smoke test in `docs/development.md`.

The manifest adds an **Open bibdocqmd** button to the **Home** ribbon tab. If an older
pilot manifest is already installed, remove it and upload the current hosted manifest
again so Word refreshes the ribbon definition.

The hosted task pane reads user-selected `.bib`, `.bibtex`, and CSL files inside the browser. It does not upload them. Native SharePoint selection and the optional localhost synced-folder experiment are not part of this hosted pilot.
