# Third-party notices

## Adobe predefined character maps

OfficeIMO.Pdf embeds mapping data derived from the predefined CMaps (code to
CID) of Adobe-GB1, Adobe-CNS1, Adobe-Japan1, Adobe-Korea1 and Adobe-KR, and from
the CID-to-Unicode mappings of those character collections. The data is copyright
1990-2023 Adobe and distributed under BSD-3-Clause; each embedded resource carries
the copyright notice.

Sources are [Adobe CMap Resources](https://github.com/adobe-type-tools/cmap-resources/tree/f5cf3bca7fdfeaceb77aa82847e974f2306c20b4)
and [Adobe Mapping Resources for PDF](https://github.com/adobe-type-tools/mapping-resources-pdf/tree/2dd5e53fb74a01718b9dfd448a0d1cce6fff2aa5).
`Build/Generate-PdfPredefinedCMaps.py` converts them to a compact compressed form.
`licenses/AdobeCMaps/source.json` in the NuGet package records the upstream
revisions, the SHA-256 of every input file and of each generated resource. The
original upstream licenses are included as `licenses/AdobeCMaps/LICENSE.md` and
`licenses/AdobeCMaps/LICENSE.txt`.

Only mapping data is bundled. No Adobe executable, font program or network
service is required.
