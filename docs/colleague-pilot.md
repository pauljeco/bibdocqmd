# Colleague pilot: Word for the web

Current focus: make citation insertion, editing, bibliography generation, and
save/reopen straightforward for a small group of colleagues. QMD export is
**experimental / in development** and parked at its current manual prototype.

## What colleagues receive

One link to the hosted pilot's landing page. The intended site is
`https://pauljeco.github.io/bibdocqmd/`; confirm the deployed release before sharing
it. The page provides the hosted manifest, installation steps, a fictional sample
DOCX and bibliography, and a short first-use guide. They need Word for the web and
a browser. They do not need Node.js, Python, Pandoc, a local server, or development
certificates to use the citation/bibliography pilot.

Manual add-in upload is a pilot/testing installation method, subject to the
organization's Microsoft 365 settings. The documented web route is Home → Add-ins
→ More Settings → Upload My Add-in; labels can vary. Browser sideloading is local
to that browser, so changing browsers or clearing its cache can require reinstalling.
See [Microsoft's web sideloading guide](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).

## Prepare the hosted release

1. Build and run the add-in checks. Review existing platform-specific failures
   separately; do not describe a failing full check as passing.
2. Review the pending repository changes and publish the intended pilot revision
   using the GitHub Pages workflow. The workflow builds the bundles and deploys
   `pilot.html` as `index.html`, `manifest.pages.xml` as `manifest.xml`, and the
   fictional examples as `demo.docx` and `demo.bib`.
3. Verify the live landing page, XML download, task pane assets, and both sample
   downloads. Check that the manifest uses the hosted HTTPS URLs.
4. Use a fresh browser profile/account to install through Word for the web. This
   checks a colleague's path without depending on the developer's local server or
   existing add-in cache.
5. Record the browser/version, Word web environment, install path, save/reopen
   path, and downloaded-DOCX inspection in `word-host-test-matrix.md`.

Publishing and fresh-account installation still need to be performed for the
current changes; preparing this guide does not establish that the live site has
been updated or that a new colleague can install successfully.

## Small pilot

Start with two or three colleagues using separate document copies. Have each
person install, load the sample bibliography, switch citation display, edit one
citation, generate a bibliography, and save/reopen. Then repeat with their own
`.bib` source. Check one citation group spanning two bibliography files and the
duplicate-key error. Verify downloaded DOCX identities where practical.

Local bibliography files need to be reselected after reopening. Their original
contents remain the authoritative source. Simultaneous coauthoring is a separate,
unverified workflow; test it before recommending shared-document editing.
