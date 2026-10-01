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

## Retry weak recognition evidence

`AdaptiveOcrEngine` wraps one through four caller-configured engines. It runs the first variant, then tries later variants only when the selected evidence fails the configured checks. All variants receive isolated copies of the same raster and coordinate frame. Segmentation and language can vary; image cleanup and coordinate transforms remain with the scanning or format owner.

```csharp
var adaptive = new AdaptiveOcrEngine(
    "document-ocr",
    new[] {
        new OcrRecognitionAttempt("baseline", baselineEngine),
        new OcrRecognitionAttempt("alternate", alternateEngine)
    },
    new OcrReviewPolicy(minimumWordConfidence: 0.8,
        maximumUncertainWordFraction: 0.1),
    timeout: TimeSpan.FromSeconds(45));

AdaptiveOcrResult recognition = await adaptive.RecognizeWithReviewAsync(request);
Console.WriteLine(recognition.Result.Text);
Console.WriteLine($"Review recommended: {recognition.ReviewRecommended}");
```

`baselineEngine`, `alternateEngine`, and `request` are your configured providers and raster request. See the [Tesseract example](../OfficeIMO.Ocr.Tesseract/README.md#bounded-segmentation-retries) for a concrete setup. Pass `adaptive` anywhere an `IOcrEngine` is accepted, including Reader and PDF OCR. Its normal result carries a content-free `adaptive-ocr-review-recommended` warning or `adaptive-ocr-thresholds-met` information diagnostic. `RecognizeWithReviewAsync` additionally returns attempt outcomes and selected word-level evidence.

The default checks require at least one word span, confidence of at least 0.8 for at least 90% of words, and no warning/error diagnostics or omitted spans. Unknown or invalid confidence counts as uncertain. A retry can replace the baseline only if it retains at least 90% of the baseline word count, has usable evidence, and improves uncertainty or repairs missing evidence. This count check cannot prove that the same facts survived. Overall confidence never ranks results. Text disagreement, after Unicode and whitespace normalization, and failed or timed-out retries recommend review even when the selected confidence checks pass.

Limits default to one minute across attempts, 25 MiB of input, 100,000 retained spans, and one million characters including span text per attempt. Caller cancellation propagates; a failed retry retains the baseline. Orientation detection delegates to the baseline with the same input, result, and timeout bounds and copying rules. Engines remain caller-owned, and shared runner gates retain their normal lifetime rules.

These thresholds are starting settings, not calibrated correctness probabilities or human approval. The [native quality corpus](../OfficeIMO.TestAssets/OcrQuality/README.md) measures confidence false passes against independent gold text. In particular, high-confidence output can still omit text or misorder columns.

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

## Execution outcomes

`OcrEngineRunner` rejects null results and an `Error` diagnostic whose `IsRecoverable` is false before Reader or PDF integrations consume recognized text. These failures use `OcrEngineExecutionException.Kind`. Provider exceptions become content-free `ProviderFailure` errors without an inner exception; inspect private provider logs for details. Caller cancellation and the shared execution timeout retain their existing contracts.

Execution captures independent result collections and geometry while the deadline remains active. Use the `OcrEngineExecution.RecognizeAsync` overload with `OcrResultCaptureLimits` to retain bounded span, diagnostic, and attribute prefixes. `OmittedSpanCount`, `OmittedDiagnosticCount`, and `OmittedAttributeCount` report discarded items; terminal error diagnostics still reject recognition outside the retained prefix. Reader passes its configured retention limits, and PDF preserves count-limit rejection. Runner-owned copying stops at cancellation or timeout between provider collection accesses.

Recoverable diagnostics remain recognition evidence. Provider diagnostic messages are provider-authored content; providers must keep credentials out of them. A recoverable result does not imply that every recognized word is correct. Confidence must be finite and within zero through one. Reader removes invalid confidence values, and PDF excludes words with invalid confidence.

## Integrations and providers

- `OfficeIMO.Reader.Ocr` recognizes image candidates from Word, Excel, PowerPoint, OneNote, EPUB, email, PDF, and other Reader adapters.
- `OfficeIMO.Pdf.Ocr` renders PDF pages, filters OCR/native overlap, reconstructs the logical document, and can add a searchable text layer.
- `OfficeIMO.Workflows` combines bounded OCR geometry with decoded raster pixels for concealed-text assessment and explicitly selected opaque-region redaction.
- `OfficeIMO.Ocr.Process` adapts a caller-configured executable through a bounded versioned protocol.
- `OfficeIMO.Ocr.Tesseract` supplies an optional engine for an installed Tesseract CLI.

All integrations accept the same `IOcrEngine`; providers do not reference Reader, PDF, or another document format.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0` (`net472` is also included on Windows builds).
- Dependencies: none.
- License: MIT.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
