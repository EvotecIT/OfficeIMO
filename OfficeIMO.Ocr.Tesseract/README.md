# OfficeIMO.Ocr.Tesseract - Tesseract OCR provider

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Ocr.Tesseract)](https://www.nuget.org/packages/OfficeIMO.Ocr.Tesseract)

`OfficeIMO.Ocr.Tesseract` is an optional `IOcrEngine` backed by an installed Tesseract command-line executable. It does not bundle native binaries or trained language data.

## Install

```powershell
dotnet add package OfficeIMO.Ocr.Tesseract
```

Install Tesseract separately for the host operating system. `TesseractOcrEngine.CreateDefault()` discovers an explicit path, `OFFICEIMO_TESSERACT_PATH`, `TESSERACT_PATH`, the process `PATH`, and common platform locations:

```csharp
TesseractOcrEngine engine = TesseractOcrEngine.CreateDefault();
```

You can verify the executable and required languages directly:

```text
tesseract --version
tesseract --list-langs
```

Tesseract 5 is the current stable major line. Its command contract supports image input, language expressions, and TSV output; see the [official Tesseract manual](https://github.com/tesseract-ocr/tesseract/blob/main/doc/tesseract.1.asc).

## Recognize an image

```csharp
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Ocr;

TesseractOcrSession session = await TesseractOcr.CreateSessionAsync(new TesseractOcrSessionOptions {
    Languages = TesseractOcrLanguage.English | TesseractOcrLanguage.Polish,
    Engine = new TesseractOcrEngineOptions {
        PageSegmentationMode = 3,
        Timeout = TimeSpan.FromMinutes(1)
    }
});

OcrResult result = await session.RecognizeAsync(
    File.ReadAllBytes("scan.png"),
    "image/png",
    "scan.png");

foreach (OcrTextSpan span in result.Spans) {
    Console.WriteLine($"{span.Level}: {span.Text} ({span.Confidence:P0})");
}
```

`session.Engine` is the same neutral `IOcrEngine` accepted by `OfficeIMO.Reader.Ocr`, `OfficeIMO.Pdf.Ocr`, and future format integrations. The Tesseract package itself does not reference those packages.

The provider parses Tesseract TSV into line and word spans with pixel bounding boxes and normalized confidence. Tesseract TSV does not expose character boxes, so `SupportsCharacterSpans` is false. A process or delegate engine can still return character spans through the shared core contract.

`GetVersionAsync()` and `GetLanguagesAsync()` provide explicit installation evidence. `TesseractLanguageData.EnsureAsync("eng+pol")` can provision any model in the 28-language catalog, plus orientation data, into a versioned user cache. Downloads come from one immutable official `tessdata_fast` commit and must match package-pinned lengths and SHA-256 digests. Missing executables, unavailable trained data, unsupported input formats, and nonzero process exits surface as engine failures. A consuming format integration decides whether to stop or convert those failures into document diagnostics.

Per-request payload and output files use owner-only Unix directories and permissions. Temporary files are deleted by default; enable `KeepTemporaryFiles` only for controlled diagnostics.

## Targets and licenses

- Targets: `netstandard2.0`, `net8.0`, `net10.0` (`net472` is also included on Windows builds).
- OfficeIMO provider license: MIT.
- Tesseract is an external dependency distributed under its own Apache 2.0 license.

## Dependency footprint

- **External:** An installed Tesseract CLI. Language data is not bundled; callers may use their system data or explicitly invoke the checksum-pinned provisioner.
- **OfficeIMO:** `OfficeIMO.Ocr` owns the neutral contracts. `OfficeIMO.Ocr.Process` supplies the bounded cross-platform process runner used by this provider.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
