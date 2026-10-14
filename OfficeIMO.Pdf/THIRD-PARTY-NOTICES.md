# Third-party notices

## Adobe predefined character maps

OfficeIMO.Pdf embeds horizontal UCS-2 code-to-CID mappings for Adobe-Japan1,
Adobe-GB1, Adobe-CNS1 and Adobe-Korea1, and their CID-to-Unicode mappings.
These data files are copyright Adobe and distributed under BSD-3-Clause.
The complete copyright and license notices remain in each embedded resource.

Sources are [Adobe CMap Resources](https://github.com/adobe-type-tools/cmap-resources/tree/f5cf3bca7fdfeaceb77aa82847e974f2306c20b4)
and [Adobe Mapping Resources for PDF](https://github.com/adobe-type-tools/mapping-resources-pdf/tree/2dd5e53fb74a01718b9dfd448a0d1cce6fff2aa5).
`licenses/AdobeCMaps/source.json` in the NuGet package records each source URL,
revision, byte length and SHA-256. The original upstream licenses are included
as `licenses/AdobeCMaps/LICENSE.md` and `licenses/AdobeCMaps/LICENSE.txt`.

Only mapping data is bundled. No Adobe executable, font program or network
service is required.
