# OfficeIMO.Reader.Ocr - OCR enrichment for Reader

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Ocr)](https://www.nuget.org/packages/OfficeIMO.Reader.Ocr)

`OfficeIMO.Reader.Ocr` applies any `OfficeIMO.Ocr.IOcrEngine` to image candidates emitted by modular OfficeIMO readers. It adds recognized content to `OfficeDocumentReadResult` while preserving native text, source locations, assets, diagnostics, nested document results, and provider evidence. `ApplyOcrAsync` processes document- and page-owned candidates in one result. `ApplyOcrTreeAsync` also visits nested results with one shared execution budget.

## Install

Install the integration, the required Reader format adapter, and one provider. For example, Word plus Tesseract:

```powershell
dotnet add package OfficeIMO.Reader.Ocr
dotnet add package OfficeIMO.Reader.Word
dotnet add package OfficeIMO.Ocr.Tesseract
```

`OfficeIMO.Reader.All` does not include OCR providers or the OCR integration.

## Recognize embedded document images

```csharp
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddWordHandler()
    .Build();

OfficeDocumentReadResult document = await reader.ReadDocumentAsync("scanned-notes.docx");
var engine = TesseractOcrEngine.CreateDefault();

OfficeDocumentOcrExecutionResult result = await document.ApplyOcrAsync(
    engine,
    new OfficeDocumentOcrExecutionOptions {
        Language = "eng+pol",
        MaxCandidates = 50,
        MaxDegreeOfParallelism = 2,
        CandidateTimeout = TimeSpan.FromMinutes(1)
    });

Console.WriteLine(result.Document.Markdown);
Console.WriteLine($"Recognized: {result.Report.RecognizedCandidateCount}");
```

Reader adapters can emit OCR candidates for raster images in Word, Excel, PowerPoint, OneNote, EPUB, email, PDF, standalone image files, and future formats. Candidate metadata stays in `OfficeIMO.Reader.Core`; no OCR runs during ordinary parsing.

This package owns candidate-to-asset validation, payload and hash checks, per-candidate timeout configuration, deterministic scheduling, aggregate result/span/diagnostic limits, diagnostic mapping, and normalized-document enrichment. Shared engine serialization and timeout supervision come from `OfficeIMO.Ocr.OcrEngineRunner`, so a non-concurrent instance is protected even when Reader and PDF use it simultaneously. `OfficeDocumentOcrExecutionResult.Recognitions` retains each bounded neutral `OcrResult`, including detailed geometry and provider provenance.

Use the same engine instance with `OfficeIMO.Pdf.Ocr` when PDF page rendering and searchable output are required. Use `DelegateOcrEngine` or another `IOcrEngine` implementation for hosted, native, or application-specific providers.

## Recognize attachments and archive entries

Read the container with its required format handlers, then invoke the tree operation:

```csharp
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Email;
using OfficeIMO.Reader.Image;
using OfficeIMO.Reader.Zip;

var reader = new OfficeDocumentReaderBuilder()
    .AddEmailHandler()
    .AddZipHandler()
    .AddImageHandler()
    .Build();
var document = await reader.ReadDocumentAsync("message.eml");
var result = await document.ApplyOcrTreeAsync(TesseractOcrEngine.CreateDefault(),
    new OfficeDocumentOcrExecutionOptions {
        MaxDocuments = 50,
        MaxNestedDepth = 4,
        MaxCandidates = 100,
        MaxTotalInputBytes = 64L * 1024 * 1024,
        MaxTotalRecognizedCharacters = 2 * 1024 * 1024,
        TotalTimeout = TimeSpan.FromMinutes(3)
    });
Console.WriteLine(result.Document.Markdown);
foreach (var recognition in result.Recognitions)
    Console.WriteLine($"{recognition.DocumentPath}: {recognition.CandidateId}");
```

Add other format handlers when those attachments are supported. Tree execution preserves rich child results and adds newly recognized text to each containing result's blocks, chunks and Markdown without repeating its native text. `DocumentId` identifies the structural node; `DocumentPath` retains its source or virtual container path. Candidate and asset identifiers remain local to that node.

When native text is available only as chunks, enrichment captures it in `chunk` blocks before adding OCR blocks. The original chunks and their structured evidence remain available.

Candidate selection, materialized input bytes, newly accepted text, spans and span characters share their limits across all visited documents. The total engine-call deadline also spans the tree; local normalization and projection can finish after that deadline. Concurrent raw responses are retained only within the in-flight window and normalized in source order. Limit diagnostics and unresolved candidates remain visible. Document/depth limits preserve skipped subtrees and report them through `SkippedDocumentCount`. These bounds do not replace Reader's decoding limits or a provider's own process-memory limits.

`OfficeDocumentOcrProcessor` remains a per-document Reader processor. Use `ApplyOcrTreeAsync` after reading a container when the whole operation needs a shared budget and refreshed parent text. Reader JSON excludes binary payloads: materialize asset bytes again before executing OCR on a transported result.

## Targets and dependency footprint

- Targets: `netstandard2.0`, `net8.0`, `net10.0` (`net472` is also included on Windows builds).
- OfficeIMO dependencies: `OfficeIMO.Ocr` and `OfficeIMO.Reader.Core`.
- Not dependencies: PDF, Tesseract, process execution, cloud SDKs, native runtimes, or other Reader format packages.
- License: MIT.

See the [Reader Core README](../OfficeIMO.Reader.Core/README.md) for reader construction and result contracts.

## Recognition evidence

OCR-enriched `OfficeDocumentBlock.Recognition` carries an immutable `OfficeDocumentRecognitionEvidence` with the provider, model, language and available confidence/review measurements. JSON round trips and nested OCR projection preserve it. Direct `ApplyOcrResults` callers may supply detailed `Recognition`; omitted assessments remain unknown. Engine execution records word-confidence checks and preserves adaptive comparison outcomes. Truncation or normalization loss requires review.

Consumers must not treat confidence or matching OCR variants as approval. OfficeIMO.AI carries this evidence into its snapshot identity, citations and review reports.
