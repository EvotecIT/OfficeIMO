# OfficeIMO.Reader.Ocr - easy local OCR

`OfficeIMO.Reader.Ocr` is the one-call facade for local image OCR and searchable PDFs. It uses the engine-neutral Reader contracts, the optional Tesseract CLI provider, and the OfficeIMO.Pdf 3.3 rendering and mutation pipeline.

## Runtime prerequisite

Install Tesseract once with the operating-system package manager. The facade discovers common installations and `PATH` automatically:

```text
Windows: winget install --id UB-Mannheim.TesseractOCR --exact
macOS:   brew install tesseract
Ubuntu:  apt-get install tesseract-ocr
```

You can also set `OFFICEIMO_TESSERACT_PATH` or pass `Tesseract.ExecutablePath`. OfficeIMO does not silently install a system executable or bundle a community native build.

## Searchable PDF in one call

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Reader.Ocr;

PdfSearchableOcrResult result = await OfficeOcr.MakePdfSearchableAsync(
    "scanned.pdf",
    "scanned-searchable.pdf");

Console.WriteLine($"Added {result.AddedWordCount} OCR words.");
```

English is the default. Language choices are typed and discoverable, so callers do not need to know Tesseract language codes. The curated catalog contains 28 typed languages across Latin, Cyrillic, Arabic, Hebrew, Devanagari, Chinese, Japanese, Korean, Greek, and Vietnamese text. Requested language data is downloaded on demand from an immutable official `tessdata_fast` commit; orientation data is included automatically when the selected page segmentation mode needs it. Every file is checked against a package-pinned size and SHA-256 digest before it enters the versioned user cache.

```csharp
var options = new OfficeOcrOptions();
options.Languages = OfficeOcrLanguage.English | OfficeOcrLanguage.Polish | OfficeOcrLanguage.German;

await OfficeOcr.MakePdfSearchableAsync("scan.pdf", "searchable.pdf", options);
```

Use `OfficeOcrLanguages.Supported` to populate a UI without maintaining another language list. Callers with their own trained-data files can set `CustomLanguageExpression`. That advanced escape hatch accepts raw Tesseract identifiers; ordinary use should stay on the typed property. The earlier `Tesseract.Language` configuration route remains supported for compatibility. Combining multiple non-default language routes is rejected instead of silently choosing one.

## Reuse a session

Create a session once for repeated images or PDFs. Session creation verifies the executable version and available languages.

```csharp
OfficeOcrSession session = await OfficeOcr.CreateSessionAsync(options);
OfficeOcrEngineResult image = await session.RecognizeImageAsync(pngBytes, "image/png");
PdfSearchableOcrResult pdf = await session.MakePdfSearchableAsync(PdfDocument.Load("scan.pdf"));
```

For other OCR engines, use `IOfficeOcrEngine`, `OfficeOcrEnginePdfProvider`, or `IPdfOcrProvider` directly. This keeps cloud SDKs, native inference runtimes, and experimental providers out of the default package graph.

## Boundaries

- Tesseract runs out of process with bounded input, output, time, and descendant-process containment.
- Searchable PDF output is a full rewrite through OfficeIMO.Pdf mutation policy; signatures that forbid rewriting fail closed.
- OCR words overlapping native PDF text are excluded by default.
- The package is MIT. Tesseract and official trained data use Apache 2.0 and remain external runtime assets.
