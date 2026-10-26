import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';

const manifestCases = [
  {
    file: new URL('./manifest.xml', import.meta.url),
    baseUrl: 'https://localhost:3000',
  },
  {
    file: new URL('./manifest.pages.xml', import.meta.url),
    baseUrl: 'https://pauljeco.github.io/BibWord',
  },
];

for (const manifestCase of manifestCases) {
  test(`${manifestCase.file.pathname.split('/').at(-1)} defines the BibWord ribbon command`, async () => {
    const manifest = await readFile(manifestCase.file, 'utf8');

    assert.match(manifest, /<VersionOverrides[\s\S]*xsi:type="VersionOverridesV1_0">/);
    assert.match(manifest, /<ExtensionPoint xsi:type="PrimaryCommandSurface">/);
    assert.match(manifest, /<OfficeTab id="TabHome">/);
    assert.match(manifest, /<Control xsi:type="Button" id="BibWord\.OpenTaskpaneButton">/);
    assert.match(manifest, /<Action xsi:type="ShowTaskpane">/);
    assert.ok(manifest.includes(`<SourceLocation resid="BibWord.Taskpane.Url"/>`));
    assert.ok(manifest.includes(`DefaultValue="${manifestCase.baseUrl}/taskpane.html"`));
    assert.ok(manifest.includes(`DefaultValue="${manifestCase.baseUrl}/commands.html"`));

    for (const size of [16, 32, 80]) {
      assert.ok(manifest.includes(`<bt:Image size="${size}" resid="BibWord.Icon${size}"/>`));
      assert.ok(manifest.includes(`DefaultValue="${manifestCase.baseUrl}/icon-${size}.png"`));
    }
  });
}
