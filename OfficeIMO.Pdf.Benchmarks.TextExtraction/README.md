# OfficeIMO PDF text extraction over public corpora

This opt-in runner measures whole-document plain-text extraction over directories of real PDFs, such as the veraPDF, pdf.js and SafeDocs test corpora. It is outside `OfficeIMO.slnx` and is never packed. PDFium (through Docnet.Core), PdfPig and PdfOxide are measurement references only.

## What it measures

For every file, each engine opens the PDF from disk and returns the text of every page. The timed region covers exactly that, from open to the final page's text:

| Engine | Call |
| --- | --- |
| `officeimo` | `PdfDocument.Load(path).ExtractText()` |
| `officeimo-read` | `PdfDocument.Load(path).Read().Text`, the structured semantic read. It does much more work than plain extraction and is reported separately. |
| `pdfoxide` | `PdfDocument.Open(path)`, then `ExtractText(page)` for every page. `ExtractAllText()` silently drops the text of a failing page, so the per-page call keeps failures visible. |
| `pdfium` | `GetDocReader(path)`, then `GetPageReader(page).GetText()` for every page. |
| `pdfpig` | `PdfDocument.Open(path)`, then `Page.Text` for every page. |

Every engine runs in its own child process, single-threaded, with a 60-second limit per file. Engines are interleaved per file, so load changes on the machine affect all of them equally. There are two modes:

- `cold`: a new process for each file and engine, with no warm-up. .NET engines therefore also pay for JIT compilation on every file. Native engines pay only for loading their library.
- `warm`: one long-lived process per engine receives the files in sequence, the way a corpus script loops in one interpreter. Nothing is measured twice. A crash or timeout replaces the process and is recorded against that file.

Outcomes are `ok`, `error` (the engine threw), `timeout` and `crash` (the process died). The pass rate counts `ok` only. An engine that returns empty text without an error still passes, so read the parity results alongside it.

## Engines do different work

The engines extract text with different defaults:

- **OfficeIMO `ExtractText()`** builds lines from glyph geometry without reordering columns. It honours `/ActualText`. It leaves out `/Artifact` content, hidden optional content and form widgets unless `PdfLoadOptions` asks for them. It keeps invisible OCR text (render mode 3). Encrypted files that need a password, permission-restricted files and unmappable composite fonts raise typed exceptions instead of returning partial text.
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
