# Independent-producer PDF corpus run — 2026-11-10

This run widens the PDF interoperability corpus with 60 hash-pinned public files from producers OfficeIMO does not control. It records where `OfficeIMO.Pdf` fails on them. It is discovery evidence for the [corpus roadmap item](../ROADMAP.md#pdf-engine-and-conversion-fidelity), not a compatibility percentage: most new files are small conformance fixtures, and several sources share a producer.

## Corpus

[`pdf-corpus-extended.json`](../../OfficeIMO.Pdf.Benchmarks.Comparisons/Corpus/pdf-corpus-extended.json) adds 60 entries to the 16 in [`pdf-corpus.json`](../../OfficeIMO.Pdf.Benchmarks.Comparisons/Corpus/pdf-corpus.json). Each entry pins a commit or publisher URL, its SHA-256, and its license. The 51 files up to 2 MB are now committed under [`OfficeIMO.TestAssets/Pdf/IndependentCorpus`](../../OfficeIMO.TestAssets/Pdf/IndependentCorpus/PROVENANCE.md), and `PdfIndependentCorpusTests` reads them in ordinary test runs; the nine larger files remain download-only.

| Source | Entries | License |
| --- | ---: | --- |
| [qpdf test suite](https://github.com/qpdf/qpdf/tree/4eba95899886e851cc41d76886483b347612f2a8/qpdf/qtest/qpdf) | 22 | Apache-2.0 |
| [PDF/UA reference suite](https://pdfa.org/resource/pdfua-reference-suite/) (via veraPDF regression tests) | 8 | CC-BY-4.0 |
| [veraPDF corpus](https://github.com/veraPDF/veraPDF-corpus/tree/bb75f4f0073d9350dfd058c0162a367e6fadf25e) | 7 | CC-BY-4.0 |
| [PDF 2.0 examples](https://github.com/pdf-association/pdf20examples/tree/c20f2c17bfcc4baab7cfe62e70fae64caf14d5fa) | 7 | CC-BY-SA-4.0 |
| ACL Anthology papers | 4 | CC-BY-4.0 |
| NIST, NASA, IRS, National Archives | 6 | United States government work |
| Japanese Digital Agency | 3 | Public Data License 1.0 |
| RFC Editor | 2 | IETF Trust Legal Provisions |
| LibreOffice documentation | 1 | CC-BY-4.0 or GPL-3.0-or-later |

The producers include Adobe InDesign (5, three remediated by AbleDocs), pdfTeX (5), LibreOffice Writer (4), Microsoft Word through Acrobat PDFMaker (3, one Japanese), Ghostscript 5.x (2), three scanners (Internet Archive LuraDocument with JBIG2 and JPX, Pdf.Capture with CCITT, and Aspose with JBIG2), Antenna House, cairo/xml2rfc, JUST PDF, and Acrobat Distiller. Feature labels cover xref and object streams (13), hybrid xref, linearization (13), incremental updates (9), RC4-40, RC4-128, AES-128 and AES-256 encryption (6, three with a known user password), AcroForms (12), tagged and PDF/UA files (16), PDF 2.0 (8), Japanese, Chinese, Arabic, Devanagari and Cyrillic text, Type3, CIDFontType0C and CIDFontType2 fonts, JBIG2, JPX and CCITT images, transparency and blend modes, optional content, six damaged or repairable files, and four documents of 132–297 pages.

The new optional `password` field opens encrypted entries. Those entries skip the select, split and merge step, because the runner's manipulation helpers do not take credentials.

Not covered: no Chrome/Skia-produced file with a clear license turned up. The [HTML head-to-head lane](../../OfficeIMO.Pdf.Benchmarks.Comparisons/README.md#html-head-to-head-corpus) already reads Chromium output. Hebrew and vertical CJK are also missing. I left out pdf.js test PDFs because most of them are bug-report attachments without a per-file license.

## Labels

`expectedPages` is set only where Poppler `pdfinfo` and MuPDF agree. Ten entries have `expectedText`, and each string appears in Poppler, MuPDF, and PdfPig or iText output. Arabic strings are not labelled because Poppler and MuPDF disagree on Arabic word order. The recall threshold is 0.90 for born-digital files and 0.75 for scans with an OCR layer.

## Method

All runs used macOS 27 on Apple silicon with .NET 10 and the heavy-run lock. The source was `923a2f612e` plus the runner changes listed at the end of this report, which the corpus command needs for the `password` field and offset headers.

1. The repository `corpus` command, over both manifests:

   ```sh
   dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons -c Release -f net10.0 -- corpus --download \
     --output <scratch>/corpus --additional-manifest OfficeIMO.Pdf.Benchmarks.Comparisons/Corpus/pdf-corpus-extended.json
   ```

2. A three-reader probe, run in a separate process for each file. It used OfficeIMO's structured `Read()` with artifact text, OfficeIMO's per-page `ExtractText()`, PdfPig 0.1.15 `ContentOrderTextExtractor`, and iText 9.7.0 `LocationTextExtractionStrategy`, with the iText Bouncy Castle adapter for AES. Each reader ran twice and the warm pass is reported. Token recall uses the corpus tokenizer (letters and digits, three or more characters, duplicate-aware, per page). The reference is the better of Poppler `pdftotext` and MuPDF, so neither one's word joining counts against a reader.
3. Rendering of the first three pages with OfficeIMO's managed renderer to PNG at 72 DPI. Each page is compared with Poppler `pdftoppm` at the same resolution. The score is unmatched ink: ink pixels with no counterpart within 2 px, divided by all ink.
4. `AssessMutations()` for every file, counting `Blocked` plans and their blocker codes.

Raw text, PNGs and JSON stay in the scratch folder.

## Results

`corpus` command: the 16 base entries gave 13 passed and 3 blocked. The 60 new entries gave 34 passed, 11 blocked, 14 failed and 1 skipped (a password entry).

| Reader | Files read | Mean recall (62 files all four read) | Median warm time vs PdfPig | Total warm time | Total allocation |
| --- | ---: | ---: | ---: | ---: | ---: |
| OfficeIMO structured `Read()` | 64/70 | 0.969 | 6.0× | 67.2 s | 101.5 GB |
| OfficeIMO `ExtractText()` | 67/70 | 0.970 | 1.11× | 6.4 s | 6.3 GB |
| PdfPig 0.1.15 | 69/70 | 0.984 | 1.0× | 5.0 s | 4.0 GB |
| iText 9.7.0 | 69/70 | 0.983 | 1.42× | 9.1 s | 14.4 GB |

Times and allocations are summed over the 29 files where every reader succeeded and PdfPig took at least 5 ms. One file, `verapdf-cmap-embedded` (recall 0), accounts for most of the OfficeIMO recall gap. On the other files OfficeIMO's recall matches the reference readers, and its precision is 0.999.

Rendering: 148 of the first 150 pages were attempted and 133 produced PNGs. One LibreOffice page did not finish, and the AES-256 file failed to load. 125 of the 133 produced pages are within 5% unmatched ink of Poppler. The most frequent diagnostic is `render.resource.font-substitution`, on 53 pages.

| Operation | Blocked of 70 | Leading blockers (files) |
| --- | ---: | --- |
| `ModifyPageTree` | 39 | `TaggedContent` 22, `XrefStreamPreservation` 17, `Forms` 16, `ObjectStreamPreservation` 13, `Outlines` 12 |
| `Redact` | 21 | `XrefStreamPreservation` 10, `TaggedContent` 8, `ObjectStreamPreservation` 7, `Outlines` 6, `Forms` 3, `Read.ContentStreamDecodeFailure` 1 |
| `MergeDocuments` | 20 | `Outlines` 12, `XrefStreamPreservation` and `ObjectStreamPreservation` 7, `Encryption` 3, `PageLabels` 2 |
| `ExtractPages` | 10 | `Encryption` 3, `Outlines` 3, xref/object-stream preservation 3, `OpenActions` 1, `IncompleteObjectGraph` 1, `Read.MissingHeader` 1, `Read.NoPages` 1 |
| `UpdateMetadata` | 6 | `AppendOnlyRequired` 5, `Encryption` 3 |

Every blocked `ModifyPageTree`, `Redact`, `MergeDocuments` and `ExtractPages` plan also reports `AppendOnly.NotImplemented.<operation>`. Append-only metadata updates are implemented, so the six `UpdateMetadata` blocks come from append-only preconditions (encryption, signatures, usage rights, an incomplete object graph) or from a read blocker.

## Ranked failures by root cause

1. **Clip-intersection work exhausts limits on vector-heavy LibreOffice pages** (`libreoffice-custom-shape-tutorial`, LibreOffice 7.3, 132 pages). Text extraction aborts with `PdfReadLimitException` once `ClippingIntersectionWork` passes 1,000,000. Rendering page 1 alone did not finish in 150 s; Poppler renders it in under a second. `Redact` is blocked by `Read.ContentStreamDecodeFailure`. PdfPig and iText read the file at recall 0.999.
2. **The cost of the structured read grows with document size.** Structured `Read()` is a median 6× PdfPig. `verapdf-large-pdfa` takes 18.9 s and allocates 61.8 GB, against 48 ms and 65 MB for `ExtractText()`. `pdfua-ref-magazine-danish` takes 12.3 s and 3.2 GB, `nist-sp800-53-rev5` 6.9 s, and `acl-2023-findings-cjk` 25× PdfPig. On the 70-page CCITT scan `naca-report-1135-scan`, structured reading fails with `UnderstandingWork` 10,000,001 > 10,000,000, while `ExtractText()` reads it at recall 0.995.
3. **Digits disappear from extracted text** in `acl-2021-mt5-multilingual` (pdfTeX, NimbusRomNo9L Type1 subsets with `/Differences`). "mT5" comes out as "mT" and "(2020)" as "()", 499 tokens in all. They are missing at the span level in every Type1 font of the file, though rendering draws them correctly. The file also has a Type3 DejaVuSans font whose `/Differences` name the same digit glyphs, but that link is unconfirmed. Other pdfTeX papers are unaffected.
4. **CID-keyed text mapping.**
   - `verapdf-cmap-embedded` (CIDFontType0C with an embedded CMap) extracts no text, and the render drops glyphs ("ello rld").
   - `verapdf-composite-font` (Adobe-Korea1-2, no `ToUnicode`) raises `PdfUnsupportedTextMappingException` from both text extraction and rendering. Rendering should not depend on a Unicode mapping.
5. **Encryption.**
   - `qpdf-enc-r6-aes256-user-view` (revision 6, AES-256, password `view`) is rejected with "AES-256 Standard security dictionary entries have invalid lengths". Poppler, MuPDF, PdfPig and iText open it.
   - With the default permission policy, the two RC4 files whose permissions deny copying refuse both text and rendering. That is intended behavior, but rendering is stricter than in viewers.
6. **Image decoding in the managed renderer.**
   - Opaque 8-bit RGB/ICC JPX images fail with "could not be decoded or projected: JPXDecode" in three producers: the veraPDF Test Builder, InDesign CS6 (`pdfua-ref-brochure`), and the LuraDocument scan of NIST SP 800-12 (pages 2–3).
   - On a mixed-raster page (JPEG background, JPEG foreground, JBIG2 mask; `verapdf-type0-tounicode`) the composite comes out as noise, with 0.90 unmatched ink.
   - Ghostscript 5.50 inline images at 600 dpi (`qpdf-inline-images`) exceed the managed raster limits.
7. **Nested clip outside drawing bounds.** Page 1 of `pdfua-ref-magazine-danish` (InDesign CC 13) throws `ArgumentOutOfRangeException: Nested drawing clip must fit inside the drawing bounds`.
8. **Manipulation failures not reported as blocks.**
   - Splitting a LibreOffice form and merging its pages back fails with `AcroForm field name collision` in three files.
   - The optional-content merge refusal (`nist-cswp-29`) and the tagged-merge attachment refusal (`irs-pub17-2023`) surface as `NotSupportedException` rather than `PdfMutationBlockedException`, so callers cannot tell them apart from crashes.
9. **Inconsistent blocks.** `pdf20-offset-start` reads correctly but is blocked from `ExtractPages` by `Read.MissingHeader`; ISO 32000 allows bytes before the header. Outlines alone block merge in 12 files, the largest single refusal driver.
10. **Smaller visual gaps.**
    - Rotated widget borders (`/MK` rotation) are missing in `qpdf-rotated-annotations`.
    - Links with no `/Border` entry get no default border in `w3c-atag-1`, where Poppler draws one.
    - The Type3 page of `verapdf-type3-fonts` renders lighter, with an ellipsis glyph in place of "m²".

Other readers fail too. PdfPig cannot open `acl-2021-mt5-multilingual` (Type3 font without `/FirstChar`), and iText rejects the deep page tree in `qpdf-deep-pages`. For Arabic, recall depends on which reference is used: OfficeIMO scores 0.88 against PdfPig in the corpus command and 0.95 against the better of Poppler and MuPDF.

## Runner changes made for this run

- Optional `password` entries.
- Accept a `%PDF-` header anywhere in the first 1024 bytes.
- Compare select/split/merge output with the same extractor that produced the source baseline. Before this change, when PdfPig could not read a source, iText source text was compared with PdfPig output text.
