# Independent redaction readback

These opt-in checks consume exact source/result PDF pairs. PDFKit and Poppler are independent validation tools; neither becomes an OfficeIMO runtime prerequisite.

Build `OfficeIMO.Pdf.Benchmarks` as described in the [redaction runtime lane](../Benchmarks/README.md#pdf-redaction-workflow), then export two-page samples at each quarter-turn rotation:

```powershell
./Build/PdfViewerVerification/Export-PdfRedactionViewerFixtures.ps1 -OutputRoot ./Ignore/PdfViewerVerification
```

For an independently produced PDF, supply `-InputPath` and `-Pattern`. Export fails if search is blocked, no text matches, output verification fails, pages change or selected text survives. The source and output SHA-256 fingerprints and reviewed rectangles are saved beside each pair.

On macOS, verify a pair with Apple's independent PDFKit reader:

```sh
swift Build/PdfViewerVerification/Verify-PdfRedactionPdfKit.swift \
  Ignore/PdfViewerVerification/rotation-90/source.pdf \
  Ignore/PdfViewerVerification/rotation-90/redacted.pdf \
  Ignore/PdfViewerVerification/rotation-90/pdfkit \
  'private account [0-9]{3}' Before 'after page'
```

PDFKit must find the removal criterion in the source and no matches in the output. Every specified retained marker must occur in the source. Its occurrence count must match on each page, and each occurrence must keep its selection bounds within 0.05 PDF points. Page count, MediaBox and rotation must remain unchanged. The check writes per-occurrence geometry and counts to a JSON report, plus source/result page PNGs for visual inspection.

Use Poppler's `pdftotext -bbox-layout`, `pdfinfo` and `pdftoppm -png` on the same pair for a second reader and renderer. Inspect the PNGs and compare changes against the reviewed rectangles; successful extraction alone does not establish visual preservation. Keep source hashes, tool versions, reports and representative images with the run, and remove superseded output.

These checks qualify the selected files and readers. They do not prove Acrobat behavior, accessibility, every imported font or structure, or native Studio usability.

## Independent-producer regression corpus

The [redaction corpus](../../OfficeIMO.Pdf.Tests/Pdf/Fixtures/Interoperability/Redaction/corpus-manifest.json) records producer versions, font licenses, feature coverage and source SHA-256 fingerprints. Its ReportLab and PyMuPDF inputs exercise imported TrueType and CFF fonts, inherited text state, numeric font descriptors, bookmarks and ink annotations. The outlined-letter input requires a reviewed area; text search cannot select letters represented by vector paths.

`PdfRedactionImportedCorpusTests` applies the public redaction API to these inputs, verifies the saved output, preserves neighboring text and navigation/ink, and checks source immutability. Use the opt-in pair export and independent reader commands above when changing the engine. For fonts with non-breaking spaces, a criterion such as `private\s+account\s+123` matches the extracted whitespace explicitly. A substring whose conservative glyph envelope intersects unselected text remains blocked.

The corpus covers these selected font and document structures. Predefined CJK character maps, other producers, Acrobat and native platform accessibility require their own evidence.
