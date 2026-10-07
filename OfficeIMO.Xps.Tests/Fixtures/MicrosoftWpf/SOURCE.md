# Microsoft WPF XPS fixtures

Unmodified documents from the .NET Foundation's WPF test repository, pinned by
source URL and SHA-256 in `manifest.json`. The upstream MIT license is retained
in `LICENSE`. These files are test inputs, not shipped package assets.

`word.xps` records Microsoft XPS Document Converter 0.3.7600.16385 in its page
markup and includes an obfuscated font, print tickets and a thumbnail.
`Test_Document.xps` is the single-page input for the upstream document-structure
test; it does not itself contain a document-structure part. `PrintingDrt.xps`
contains two outlined red squares. All three use Microsoft XPS and atomic ZIP
entries. They do not qualify OpenXPS producers or interleaved production.

At 96 DPI, all three pages load, render, save and reopen. GhostXPS 10.08.0
comparison gives mean absolute RGB differences of 3.38/255, 0.34/255 and 0.68/255
respectively; glyph and edge rasterization differ. GhostXPS renders original and
OfficeIMO-saved packages pixel-identically. The contract tests retain
native page markup, text, fonts, print tickets and thumbnails through page edits.
