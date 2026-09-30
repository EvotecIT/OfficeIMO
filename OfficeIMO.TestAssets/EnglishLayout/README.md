# English layout qualification

These two controlled documents cover aligned prose columns with a centered closing line, and a borderless four-column ledger with negative amounts and wider row spacing. ReportLab produces the native PDFs; Poppler rasterizes them; ReportLab wraps those images as scans. OfficeIMO does not generate the inputs or expected reading order and table cells. The centered line supplies horizontal evidence for a page-wide closing band. Vertical separation alone does not establish page-wide intent; an isolated column-aligned paragraph retains its column membership.

`manifest.json` records producer versions, input hashes, recorded Tesseract TSV hashes, and acceptance limits declared before measurement. The TSV supplies repeatable provider geometry for the normal regression suite. The opt-in runner separately invokes the installed Tesseract engine and saves searchable PDFs for readback.

```sh
dotnet run --project Build/PdfQualityCorpus -c Release -f net8.0 -- layout OfficeIMO.TestAssets/EnglishLayout artifacts/english-layout
```

Tesseract needs the `eng` model. The runner uses 300 DPI, page segmentation mode 3, zero minimum confidence, and `ReconstructLayout = true`. Each document must have zero normalized character and word error, complete labelled reading order, and exactly the expected table structure in native, OCR, and searchable-PDF readback modes. Unexpected tables also fail qualification. Exit code 1 means a declared acceptance limit failed; exit code 2 means execution or input validation failed. Cases in other manifests without acceptance limits remain measurements and report `Qualified = null`.

The report binds measurements to the fixture manifest and PDF assembly hashes and records the provider version and available configured trained-data hashes. These are two controlled regressions from a second producer family, not a general document-accuracy estimate. Real upstream scans remain in the [scan quality corpus](../ScanQuality/README.md); mixed-script and rotated cases remain in the [Pango/Cairo corpus](../MultilingualLayout/README.md).

Run `generate.py` with ReportLab, Poppler's `pdftoppm`, and Tesseract available to regenerate the fixtures and recorded TSV. The committed versions are in the manifest. Review geometry, labels, and changed hashes together after regeneration. The text, labels, generator, and generated fixtures use the repository license. The generator is contributor tooling; none of its dependencies enters an OfficeIMO runtime package.
