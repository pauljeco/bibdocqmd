# Third party notices

BibWord's Office task pane uses the following packages to parse BibTeX and BibLaTeX and produce local CSL previews.

- Citation.js core, BibTeX plugin, and CSL plugin, version 0.8.2. Copyright Lars Willighagen and contributors. Licensed under the MIT License. Source: <https://github.com/citation-js/citation-js>
- citeproc-js, version 2.4.63. Copyright Frank Bennett and contributors. Used through the Citation.js CSL plugin. Dual-licensed under CPAL 1.0 or AGPL; BibWord uses it under CPAL 1.0. Source: <https://github.com/Juris-M/citeproc-js>

The task pane displays attribution to Citation.js and citeproc-js beside formatted previews. Dependency versions and integrity hashes are recorded in `office-addin/package-lock.json`.

The bundled APA 7th edition and Cite Them Right 12th edition Harvard CSL style files are sourced from the Citation.js 0.8.2 distribution and ultimately from the Citation Style Language styles project. Each file retains its authorship, source links, and Creative Commons Attribution-ShareAlike 3.0 notice in its `<info>` metadata. Source: <https://github.com/citation-style-language/styles>
