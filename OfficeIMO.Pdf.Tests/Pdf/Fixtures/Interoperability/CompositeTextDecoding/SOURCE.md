# Composite text decoding fixture

`unijis-without-to-unicode.pdf` is a synthetic, two-page document produced with ReportLab 4.4.9 using `UnicodeCIDFont("HeiseiMin-W3")`, then saved with pypdf 6.10.0. Each page contains `Before private account 123 after page`, set through `UniJIS-UCS2-H`. It contains no ToUnicode stream and no embedded font program.

The document is licensed under MIT. ReportLab and pypdf are fixture-generation tools under their BSD licenses; neither is an OfficeIMO runtime dependency. The byte length and SHA-256 are pinned in `source.json`.

This fixture exercises extraction and reviewed redaction through the bundled Adobe character maps when no ToUnicode is supplied. It checks the two-page retained text and source immutability. The separate `PredefinedCjk` corpus exercises CJK characters, mapped CID widths and a mismatched encoding/character-collection refusal. These selected files do not qualify every font, viewer or predefined encoding.
