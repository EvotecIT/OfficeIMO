# OfficeIMO.Reader.Ocr.Tesseract - Tesseract OCR provider

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Ocr.Tesseract)](https://www.nuget.org/packages/OfficeIMO.Reader.Ocr.Tesseract)

`OfficeIMO.Reader.Ocr.Tesseract` is an optional `IOfficeOcrEngine` backed by an installed Tesseract command-line executable. It does not bundle native binaries or trained language data.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Ocr.Tesseract
```

Install Tesseract separately for the host operating system, then verify the executable and required languages:

```text
tesseract --version
tesseract --list-langs
```

Tesseract 5 is the current stable major line. Its command contract supports image input, language expressions, and TSV output; see the [official Tesseract manual](https://github.com/tesseract-ocr/tesseract/blob/main/doc/tesseract.1.asc).

## Recognize Reader candidates

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Ocr.Tesseract;

OfficeDocumentReadResult source = DocumentReader.ReadDocument("scanned.pdf");
var engine = new TesseractOcrEngine(new TesseractOcrEngineOptions {
    ExecutablePath = "tesseract",
    Language = "eng+pol",
    PageSegmentationMode = 3,
    Timeout = TimeSpan.FromMinutes(1)
});

OfficeDocumentOcrExecutionResult execution = await source.ApplyOcrAsync(
    engine,
    new OfficeDocumentOcrExecutionOptions {
        MaxCandidates = 25,
        MaxDegreeOfParallelism = 2,
        MaxTotalInputBytes = 64L * 1024L * 1024L
    });

foreach (OfficeDocumentOcrRecognition recognition in execution.Recognitions) {
    foreach (OfficeOcrTextSpan span in recognition.Result.Spans) {
        Console.WriteLine($"{span.Level}: {span.Text} ({span.Confidence:P0})");
    }
}
```

The provider parses Tesseract TSV into line and word spans with pixel bounding boxes and normalized confidence. Tesseract TSV does not expose character boxes, so `SupportsCharacterSpans` is false. A process or delegate engine can still return character spans through the shared core contract.

`GetVersionAsync()` and `GetLanguagesAsync()` provide explicit installation discovery. Missing executables, trained data, unsupported input formats, and nonzero process exits surface as engine failures; `ApplyOcrAsync` converts them to structured diagnostics when `ContinueOnError` is enabled.

## Targets and licenses

- Targets: `netstandard2.0`, `net8.0`, `net10.0` (`net472` is also included on Windows builds).
- OfficeIMO provider license: MIT.
- Tesseract is an external dependency distributed under its own Apache 2.0 license.
