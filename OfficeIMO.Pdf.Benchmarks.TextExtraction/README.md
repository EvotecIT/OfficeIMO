# OfficeIMO PDF text extraction over public corpora

This opt-in runner measures whole-document plain-text extraction over directories of real PDFs, such as the veraPDF, pdf.js and SafeDocs test corpora, and scores reading quality on trial protocols from ClinicalTrials.gov. It is outside `OfficeIMO.slnx` and is never packed. PDFium (through Docnet.Core), PdfPig, Tabula and PdfOxide are measurement references only.

## What it measures

For every file, each engine opens the PDF from disk and returns the text of every page. The timed region covers exactly that, from open to the final page's text:

| Engine | Call |
| --- | --- |
| `officeimo` | `PdfDocument.Load(path).ExtractText()` |
| `officeimo-read` | `PdfDocument.Load(path).Read().Text`, the structured semantic read. It does much more work than plain extraction and is reported separately. |
| `pdfoxide` | `PdfDocument.Open(path)`, then `ExtractText(page)` for every page. `ExtractAllText()` silently drops the text of a failing page, so the per-page call keeps failures visible. |
| `pdfium` | `GetDocReader(path)`, then `GetPageReader(page).GetText()` for every page. |
| `pdfpig` | `PdfDocument.Open(path)`, then `Page.Text` for every page. |
| `pdfpig-tabula` | PdfPig `ContentOrderTextExtractor.GetText(page)` for every page, then Tabula's lattice extractor (`SpreadsheetExtractionAlgorithm`) on the same page, with each table row appended as tab-separated cells. This is the usual open-source .NET pairing for structured reading. |

Every engine runs in its own child process, single-threaded, with a 60-second limit per file. Engines are interleaved per file, so load changes on the machine affect all of them equally. There are two modes:

- `cold`: a new process for each file and engine, with no warm-up. .NET engines therefore also pay for JIT compilation on every file. Native engines pay only for loading their library.
- `warm`: one long-lived process per engine receives the files in sequence, the way a corpus script loops in one interpreter. Nothing is measured twice. A crash or timeout replaces the process and is recorded against that file.

Outcomes are `ok`, `error` (the engine threw), `timeout` and `crash` (the process died). The pass rate counts `ok` only. An engine that returns empty text without an error still passes, so read the parity results alongside it.

## Engines do different work

The engines extract text with different defaults:

- **OfficeIMO `ExtractText()`** builds lines from glyph geometry without reordering columns. It honours `/ActualText`. It leaves out `/Artifact` content, hidden optional content and form widgets unless `PdfLoadOptions` asks for them. It keeps invisible OCR text (render mode 3). It leaves out text runs whose font codes have no Unicode mapping and page content streams it cannot decode, and returns the rest. A page content stream with a CCITTFax or JBIG2 filter that fails to decode, and a read limit, still count as `error`. Encrypted files that need a password and permission-restricted files raise typed exceptions instead of returning partial text.
- **PdfOxide `ExtractText(page)`** sorts text into reading order (XY-cut on multi-column pages, structure-tree order for tagged PDFs) and inlines detected tables. By default it includes `/Artifact` content, FreeText and Stamp annotation text and form-field values. It ignores optional-content visibility. It returns empty text instead of an error for documents it cannot decrypt or decode.
- **PDFium** returns page text in content order without filtering `/Artifact` content.
- **PdfPig** returns `Page.Text` in content-stream order without inserting spaces between separately positioned words.

## Run

```powershell
dotnet build OfficeIMO.Pdf.Benchmarks.TextExtraction -c Release
$runner = 'OfficeIMO.Pdf.Benchmarks.TextExtraction/bin/Release/net10.0/OfficeIMO.Pdf.Benchmarks.TextExtraction.dll'
dotnet $runner run --mode warm --write-text --output <out> --suite verapdf=<veraPDF corpus> --suite pdfjs=<pdf.js test/pdfs>
dotnet $runner run --mode cold --engines officeimo,pdfoxide,pdfium,pdfpig --shard 1/4 --output <out> --suite verapdf=<veraPDF corpus>
dotnet $runner summary --output <out>
dotnet $runner parity --output <out>
```

- `run` writes `results-<mode>-<suites>-<engines>-shard<i>of<n>.csv`, with one row per file and engine, so runs with different `--engines` selections keep separate files. `--write-text` also stores each extracted text under `<out>/texts/<engine>/<suite>/`. Use `--shard i/n` to split a long cold run into shorter sessions.
- `summary` writes `summary.md` and `metrics.csv`, with the file count, pass rate, mean, p50, p99 and maximum milliseconds over successful files, plus errors, timeouts and crashes per suite and engine.
- `parity` compares the stored texts in pairs (by default each engine against PDFium, and OfficeIMO against PdfOxide) and writes `parity.csv`:
  - **Normalisation:** texts are NFKC-normalised, control characters are removed and whitespace collapses.
  - **`similarity`:** `2 × LCS / (|a| + |b|)` over the characters with all whitespace removed. Spacing does not count, but missing, extra or reordered text does. Pairs too large for the diff budget fall back to a word-multiset Dice ratio, recorded in `method`.
  - **`lax_category`:** reproduces PdfOxide's own published equivalence check, which accepts any output whose character set covers the reference or whose length is within 0.8–1.2×. Read it next to `similarity`, which is stricter.

Results depend on the machine and its load; record both with any published numbers.

## ClinicalTrials.gov protocols

Sponsors upload trial protocols to ClinicalTrials.gov next to the structured registry record. The record is the scoring reference: its official title, NCT number, lead sponsor, primary outcome measures, arm labels, intervention names and eligibility criteria should all appear in the protocol. The corpus scores how much of that each engine's text keeps intact and in order, and whether it sits under the right section heading. It isn't exact ground truth: sponsors word the record and the protocol differently, and older protocols predate registration and lack the NCT number. A low score for every engine on the same field usually means that, not lost text, so compare engines rather than reading a single score as an extraction failure.

`ClinicalTrials/manifest.json` lists the sample: 200 trials spread across the registry's sponsor classes (industry 50, other 50, NIH 25, other federal 25, other government 20, network 20, individual 10). Within each class the trials with a protocol PDF of at most 25 MB are ranked by SHA-256 of the seed and NCT number, and the lowest ranks are kept, which is a reproducible uniform sample. The manifest pins each protocol's file name and SHA-256. The PDFs and registry records are downloaded to a local folder and never committed.

```powershell
$runner = 'OfficeIMO.Pdf.Benchmarks.TextExtraction/bin/Release/net10.0/OfficeIMO.Pdf.Benchmarks.TextExtraction.dll'
$manifest = 'OfficeIMO.Pdf.Benchmarks.TextExtraction/ClinicalTrials/manifest.json'
dotnet $runner ctgov-fetch --manifest $manifest --root <corpus>
dotnet $runner run --mode warm --write-text --engines officeimo-read,pdfpig-tabula --output <out> --suite ctgov=<corpus>/pdfs
dotnet $runner ctgov-score --manifest $manifest --root <corpus> --output <out>
```

- `ctgov-fetch` downloads `<corpus>/pdfs/<NCT>.pdf` and the current record to `<corpus>/records/<NCT>.json`, keeps files that already match the pinned hash and leaves out any protocol whose bytes changed. It reports records updated since the sample was drawn, because a changed record changes the expected fields.
- `ctgov-sample` draws a new sample and `ctgov-fetch --record-hashes` pins it. Only use them to replace the manifest.
- `ctgov-score` writes `ctgov-fields.csv` (one row per field and engine) and `ctgov-summary.md`, broken down by sponsor class and by field, with time per page and the protocols with the lowest intact share.

Scoring:

- **Matching:** text and fields are NFKC-normalised and lower-cased, hyphens are removed (so a word hyphenated across a line break joins up) and everything that isn't a letter or digit separates words.
- **Intact:** the field's words appear contiguously. **In order:** they all appear in order within three times the field's length. **Words:** at least 90% of them appear somewhere, which points at order rather than missing text. Eligibility criteria become one field per criterion of five or more words.
- **Heading:** some line of twelve words or fewer, not ending like a sentence, names the field's section (outcomes or endpoints, arms or design, interventions or study drug, inclusion or exclusion). Table-of-contents lines count, so this is generous.
- **In section:** some occurrence of the field follows, within 2,000 words, a line that opens a section with the field's section name, with no line opening another section in between. A line opens a section when, after its numbering (`5.2.1`, `A.`, `IV.`, `(b)`) is removed, it starts with a section name or with a qualifier and a section name (`Study drug`, `Primary endpoints`), where a joined heading (`Inclusion and exclusion criteria`, `Inclusion / Exclusion`, `Inclusion and/or Exclusion`) opens both sections, and it reads as a heading: eight words or fewer before an early colon, or eight words or fewer in all that either don't end like a sentence or are in title case (`Primary Outcomes.`). A field matched inside the heading itself doesn't count as following it. Dot-leader table-of-contents lines don't count. Protocol headings range from `5.2 Inclusion Criteria` to `A. Inclusion criteria (all must be met): ...`, so this reads a line's opening words rather than its layout. For the NCT number, title and sponsor, it means the field appears in the front matter (the first 5% of the text, at least 2,000 words).
- **Time per page:** the run's time to open and read each protocol, divided by PdfPig's page count.

Every engine is scored on its own text output by the same rules, so the score measures text order and completeness, not each engine's own heading detection.
