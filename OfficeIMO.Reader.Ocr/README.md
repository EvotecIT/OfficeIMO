# OfficeIMO.Reader.Ocr - OCR enrichment for Reader

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Ocr)](https://www.nuget.org/packages/OfficeIMO.Reader.Ocr)

`OfficeIMO.Reader.Ocr` applies any `OfficeIMO.Ocr.IOcrEngine` to image candidates emitted by modular OfficeIMO readers. It adds recognized content to `OfficeDocumentReadResult` while preserving native text, source locations, assets, diagnostics, and provider evidence.

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

## Targets and dependency footprint

- Targets: `netstandard2.0`, `net8.0`, `net10.0` (`net472` is also included on Windows builds).
- OfficeIMO dependencies: `OfficeIMO.Ocr` and `OfficeIMO.Reader.Core`.
- Not dependencies: PDF, Tesseract, process execution, cloud SDKs, native runtimes, or other Reader format packages.
- License: MIT.

See the [Reader Core README](../OfficeIMO.Reader.Core/README.md) for reader construction and result contracts.
