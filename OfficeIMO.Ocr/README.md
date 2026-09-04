# OfficeIMO.Ocr - shared OCR contracts

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Ocr)](https://www.nuget.org/packages/OfficeIMO.Ocr)

`OfficeIMO.Ocr` is the dependency-light contract between image-producing applications, document-format integrations, and OCR providers. It contains no document parser, network client, native runtime, or process runner.

## Install

```powershell
dotnet add package OfficeIMO.Ocr
```

## Use a custom or hosted recognizer

Implement `IOcrEngine`, or adapt an existing SDK with `DelegateOcrEngine`:

```csharp
using OfficeIMO.Ocr;

IOcrEngine engine = new DelegateOcrEngine(
    "custom-provider",
    async (request, cancellationToken) => {
        HostedRecognition response = await client.RecognizeAsync(
            request.Payload,
            request.MediaType,
            request.Language,
            cancellationToken);

        return new OcrResult {
            Text = response.Text,
            Confidence = response.Confidence,
            Provider = "custom-provider",
            Spans = response.Words.Select((word, index) => new OcrTextSpan {
                Sequence = index,
                Level = OcrTextSpanLevel.Word,
                Text = word.Text,
                Confidence = word.Confidence,
                Region = new OcrRegion {
                    X = word.Left,
                    Y = word.Top,
                    Width = word.Width,
                    Height = word.Height
                },
                CoordinateUnit = OcrCoordinateUnit.Pixels
            }).ToArray()
        };
    },
    new OcrEngineCapabilities {
        SupportedMediaTypes = new[] { "image/png", "image/jpeg" },
        SupportsWordSpans = true,
        SupportsConfidence = true,
        SupportsConcurrentRequests = true
    });
```

Each request carries one validated raster payload plus optional source, candidate, page, region, language, and provider metadata. Results can return plain text, normalized confidence, provider/model provenance, diagnostics, and line, word, or character spans. Span geometry is explicit: pixels, PDF-style points, or normalized `0..1` coordinates. A provider should preserve its logical sequence and hierarchy instead of deriving structure from language-specific words.

The host or format integration owns input validation, returned-output limits, retry policy, and how recognized evidence is merged into a document. `OcrEngineRunner.RecognizeAsync` applies a total timeout and serializes every caller that shares an engine whose `SupportsConcurrentRequests` capability is `false`. For a multi-candidate document operation, call `OcrEngineRunner.CreateExecution(engine)` once and reuse the returned `OcrEngineExecution`; it captures one validated identity and capability snapshot so provider properties cannot change provenance or concurrency behavior between candidates. If a timed-out provider ignores cancellation, or its cancellation callback is still running, the runner keeps that engine's gate until all provider-owned work settles. Reader, PDF, and future format integrations use this shared runner rather than creating incompatible concurrency rules.

An engine should still honor cancellation and accurately advertise whether the same instance accepts concurrent calls. Applications invoking `IOcrEngine.RecognizeAsync` directly opt out of the shared runner policy.
Engine identifiers are stable, non-empty provenance values and are limited to 256 untrimmed characters.

## Discover optional providers

Hosts that offer selectable OCR can register provider factories in an explicit `OcrEngineCatalog`. The catalog performs no ambient assembly scanning and the core package still carries no provider runtime:

```csharp
var catalog = new OcrEngineCatalog()
    .Register(new MyOcrEngineProvider());

foreach (OcrEngineDescriptor provider in catalog.Discover()) {
    Console.WriteLine($"{provider.Id}: {provider.DisplayName}");
}

IOcrEngine engine = catalog.Create(
    "my-provider",
    new Dictionary<string, string> { ["model"] = "document" });
```

Provider identifiers are case-insensitive and unique. Registration snapshots identity and capabilities, while engine creation snapshots at most 128 scalar options with a 64 KiB aggregate character limit. Secrets remain host-owned: pass an environment-variable or secret-store reference understood by the provider instead of serializing a credential into a recipe or evidence file.

## Integrations and providers

- `OfficeIMO.Reader.Ocr` recognizes image candidates from Word, Excel, PowerPoint, OneNote, EPUB, email, PDF, and other Reader adapters.
- `OfficeIMO.Pdf.Ocr` renders PDF pages, filters OCR/native overlap, reconstructs the logical document, and can add a searchable text layer.
- `OfficeIMO.Ocr.Process` adapts a caller-configured executable through a bounded versioned protocol.
- `OfficeIMO.Ocr.Tesseract` supplies an optional engine for an installed Tesseract CLI.

All integrations accept the same `IOcrEngine`; providers do not reference Reader, PDF, or another document format.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0` (`net472` is also included on Windows builds).
- Dependencies: none.
- License: MIT.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
