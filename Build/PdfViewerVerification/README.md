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

PDFKit must find the removal criterion in the source and no matches in the output. Every specified retained marker must occur in the source and remain on each page where it occurred, with unchanged selection bounds within 0.05 PDF points. Page count, MediaBox and rotation must remain unchanged. The check writes a JSON report and source/result page PNGs for visual inspection.

Use Poppler's `pdftotext -bbox-layout`, `pdfinfo` and `pdftoppm -png` on the same pair for a second reader and renderer. Inspect the PNGs and compare changes against the reviewed rectangles; successful extraction alone does not establish visual preservation. Keep source hashes, tool versions, reports and representative images with the run, and remove superseded output.

These checks qualify the selected files and readers. They do not prove Acrobat behavior, accessibility, every imported font or structure, or native Studio usability.
