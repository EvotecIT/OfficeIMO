# Composite text decoding fixture

`unijis-without-to-unicode.pdf` is a synthetic, two-page document produced with ReportLab 4.4.9 using `UnicodeCIDFont("HeiseiMin-W3")`, then saved with pypdf 6.10.0. Each page contains `Before private account 123 after page`, set through `UniJIS-UCS2-H`. It contains no ToUnicode stream and no embedded font program.

The document is licensed under MIT. ReportLab and pypdf are fixture-generation tools under their BSD licenses; neither is an OfficeIMO runtime dependency. The byte length and SHA-256 are pinned in `source.json`.

This fixture protects an explicit refusal: extraction and redaction search cannot treat composite character codes as WinAnsi bytes. It does not qualify predefined CJK font rendering or add Adobe CMap data to OfficeIMO.
