---
title: "Charts, OCR and Provenance"
description: "Place ChartForgeX charts and Mermaid diagrams in documents, recognize text with pluggable OCR engines, and verify C2PA Content Credentials."
order: 57
meta.seo_title: "Charts, OCR and C2PA provenance for .NET documents"
---

These optional packages add capabilities that several document formats share. None of them is a dependency of the Word, Excel, PowerPoint, or PDF packages: you install them only in applications that need charts, text recognition, or Content Credentials.

| Capability | Package(s) | What it adds | External requirement |
|---|---|---|---|
| Charts and diagrams | `OfficeIMO.ChartForgeX` | ChartForgeX visuals in Word, Excel, PowerPoint, PDF, and editable Visio | ChartForgeX NuGet package |
| Mermaid in Markdown | `OfficeIMO.ChartForgeX.Markdown` | Static Mermaid rendering for Markdown, Word, PDF, HTML, and Markup output | ChartForgeX Mermaid NuGet package |
| OCR contracts | `OfficeIMO.Ocr` | Engine-neutral request, result, geometry, and review contracts | None |
| External OCR process | `OfficeIMO.Ocr.Process` | Runs a caller-configured executable through a versioned JSON protocol | Your OCR executable |
| Tesseract OCR | `OfficeIMO.Ocr.Tesseract` | `IOcrEngine` backed by an installed Tesseract command-line executable | Tesseract, installed separately |
| C2PA provenance | `OfficeIMO.Provenance.C2pa` | Content Credential verification and signing through `c2patool` | `c2patool`, supplied by the host |

## Charts and diagrams with ChartForgeX

`OfficeIMO.ChartForgeX` places any ChartForgeX `VisualArtifact` in a Word paragraph, an Excel range, a PowerPoint slide, or PDF content. Each conversion returns an `OfficeVisualConversionResult` with SVG or PNG placement bytes, an `OfficeDrawing` scene for PDF, sizes in points, accessible text, and a fidelity report. Topology, flow, and sequence artifacts can also be projected into native, editable Visio diagrams with shapes, connectors, and Shape Data.

```shell
dotnet add package OfficeIMO.ChartForgeX
```

```csharp
using ChartForgeX.Core;
using ChartForgeX.Primitives;
using ChartForgeX.Rendering;
using ChartForgeX.VisualArtifacts;
using OfficeIMO.ChartForgeX;
using OfficeIMO.Word;

var chart = Chart.Create().AddLine("Revenue", new[] {
    new ChartPoint(1, 42), new ChartPoint(2, 58), new ChartPoint(3, 73)
});
chart.ConfigureAccessibility(value => value.WithTextAlternative("Quarterly sales", "Revenue: 42, 58 and 73.", "en"));

var context = OfficeVisualDocumentStyle.Default.CreateContext(
    frame: new VisualFrame("Quarterly sales", showLegend: false));
VisualArtifact artifact = chart.Prepare(context).ToArtifact("sales-quarter");

using WordDocument document = WordDocument.Create("sales.docx");
document.AddParagraph().AddVisualArtifact(artifact);
document.Save();
```

The same artifact can be passed to `sheet.AddVisualArtifact("B4:I18", artifact)` and `slide.AddVisualArtifact(...)`. Word fits an unsized visual to the paragraph's available content width.

Limits: the default `RasterizeWhenNeeded` policy falls back to PNG when SVG import is incomplete, which is reported as an approximation and loses vector editability; choose `PreserveVector` or `RequireVector` to change that. Cropped `Cover` fitting is not supported. Native Visio output does not draw static watermarks and recomputes the layout of flow and sequence diagrams; both are reported. See the [OfficeIMO.ChartForgeX README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.ChartForgeX/README.md).

## Mermaid diagrams in Markdown

`OfficeIMO.ChartForgeX.Markdown` replaces Mermaid code fences with rendered images, using ChartForgeX for parsing, layout, and SVG or PNG output. Diagrams are embedded as SVG by default; PDF draws them as vector content, and Word stores the SVG with a PNG fallback picture. The adapter also configures the static HTML renderer and plugs into the OfficeIMO Markup parser for Word and PowerPoint export.

```shell
dotnet add package OfficeIMO.ChartForgeX.Markdown
```

```csharp
using OfficeIMO.ChartForgeX.Markdown;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Word.Markdown;

MarkdownDocument document = MermaidMarkdownAdapter.Materialize(
    MarkdownDocument.Parse(File.ReadAllText("report.md")),
    diagnostic => Console.WriteLine(diagnostic.Message));

using var word = document.ToWord(new MarkdownWordExportOptions { FitImagesToContextWidth = true });
word.Save("report.docx");
document.SaveAsPdf("report.pdf");
```

Limits: supported Mermaid grammar and its static approximations are defined by ChartForgeX. Invalid or unsupported diagrams keep their source fence and report a diagnostic. Mermaid's browser runtime is not used. See the [OfficeIMO.ChartForgeX.Markdown README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.ChartForgeX.Markdown/README.md).

## OCR

OCR in OfficeIMO is built around one engine-neutral contract. `OfficeIMO.Ocr` defines `IOcrEngine`, the request and result models, span geometry in pixels, points, or normalized coordinates, and a shared runner that applies timeouts and concurrency rules. It contains no document parser, network client, native runtime, or process runner. `AdaptiveOcrEngine` can retry weak recognition with up to four engine configurations and reports when a human review is recommended.

The same engine plugs into `OfficeIMO.Reader.Ocr` for images found in Word, Excel, PowerPoint, OneNote, EPUB, email, and PDF content, into `OfficeIMO.Pdf.Ocr` for searchable PDFs, and into the DjVu PDF adapter. See the [Reader](/docs/reader/) and [PDF](/docs/pdf/) guides for those integrations.

```shell
dotnet add package OfficeIMO.Ocr
dotnet add package OfficeIMO.Ocr.Process
dotnet add package OfficeIMO.Ocr.Tesseract
```

### Tesseract

`OfficeIMO.Ocr.Tesseract` is an optional `IOcrEngine` backed by an installed Tesseract command-line executable. It does not bundle native binaries or trained language data; install Tesseract separately for the host operating system. Discovery checks an explicit path, `OFFICEIMO_TESSERACT_PATH`, `TESSERACT_PATH`, the process `PATH`, and common platform locations. The provider returns line and word spans with pixel bounding boxes and normalized confidence, and can detect page orientation.

```csharp
using OfficeIMO.Ocr;
using OfficeIMO.Ocr.Tesseract;

TesseractOcrSession session = await TesseractOcr.CreateSessionAsync(new TesseractOcrSessionOptions {
    Languages = TesseractOcrLanguage.English | TesseractOcrLanguage.Polish,
    ProvisionMissingLanguageData = false,
    Engine = new TesseractOcrEngineOptions {
        PageSegmentationMode = 3,
        Timeout = TimeSpan.FromMinutes(1)
    }
});

OcrResult result = await session.RecognizeAsync(File.ReadAllBytes("scan.png"), "image/png", "scan.png");
foreach (OcrTextSpan span in result.Spans)
    Console.WriteLine($"{span.Level}: {span.Text} ({span.Confidence:P0})");
```

`TesseractOcr.CreateSessionAsync` downloads missing language data by default from one pinned official `tessdata_fast` commit, checked against package-pinned lengths and SHA-256 digests. Set `ProvisionMissingLanguageData = false`, as above, to use only the language data already installed. `session.Engine` is the `IOcrEngine` to pass to Reader and PDF integrations.

### Your own OCR executable

`OfficeIMO.Ocr.Process` runs an executable you configure, without a command shell, and exchanges a versioned JSON request and response through an isolated per-call directory. Input size, output size, and runtime are bounded, and the process tree is contained so it can be terminated on timeout.

```csharp
using OfficeIMO.Ocr;
using OfficeIMO.Ocr.Process;

IOcrEngine engine = new ProcessOcrEngine(new ProcessOcrEngineOptions {
    FileName = "/opt/my-ocr/recognize",
    Arguments = new[] { "--request", "{request}" },
    Id = "my-ocr",
    Timeout = TimeSpan.FromMinutes(1)
});
```

Hosted or in-process recognizers can implement `IOcrEngine` directly or wrap an existing SDK with `DelegateOcrEngine`.

Limits: confidence values and review thresholds are starting settings, not calibrated accuracy; high-confidence output can still omit text or misorder columns. Tesseract TSV has no character boxes, so the Tesseract provider returns line and word spans only. Executable paths, arguments, and provider options are trusted host configuration and must not come from document content. See the READMEs for [OfficeIMO.Ocr](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Ocr/README.md), [OfficeIMO.Ocr.Process](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Ocr.Process/README.md), and [OfficeIMO.Ocr.Tesseract](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Ocr.Tesseract/README.md).

## C2PA Content Credentials

`OfficeIMO.Provenance.C2pa` connects OfficeIMO's provider-neutral provenance contracts to the official `c2patool` command-line application for cryptographic C2PA verification and signing. Structural inspection and provenance policy stay in `OfficeIMO.Core` and the format packages; ordinary format packages do not depend on this adapter.

```shell
dotnet add package OfficeIMO.Provenance.C2pa
```

OfficeIMO does not download, bundle, or auto-discover `c2patool`. Install a supported build from the [official c2pa-rs releases](https://github.com/contentauth/c2pa-rs/releases) and pass an absolute path or a command name on your application's controlled `PATH`. On Unix hosts the adapter also requires `setsid`; without it, verification and signing fail closed.

```csharp
using OfficeIMO.Provenance;
using OfficeIMO.Provenance.C2pa;

var verifier = new C2paToolProvenanceVerifier("/opt/c2pa/c2patool");
OfficeProvenanceVerificationResult result = verifier.Verify("image.jpg", new OfficeProvenanceVerificationOptions {
    TrustAnchorsPath = "/etc/my-app/c2pa-trust-anchors.pem",
    IncludeRawReport = false
});

Console.WriteLine(result.Status);
foreach (string finding in result.Findings)
    Console.WriteLine(finding);
```

`Valid` means a manifest was found, verified, and produced no findings. `Untrusted`, `Invalid`, `NotPresent`, `ProviderUnavailable`, `Indeterminate`, and `Error` are reported separately. Remote manifest and OCSP fetching are disabled by default. `C2paToolProvenanceSigner` signs through `c2patool` 0.27.0 or newer with a signer process backed by your HSM, KMS, or signing service, and installs the signed asset atomically.

Limits: the adapter does not emit watermark actions, because they require soft-binding assertions that OfficeIMO does not create. Built-in test credentials are for development only and are marked as such in results. See the [OfficeIMO.Provenance.C2pa README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Provenance.C2pa/README.md).
