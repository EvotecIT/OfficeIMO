# OfficeIMO.Ocr.Process - external OCR process provider

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Ocr.Process)](https://www.nuget.org/packages/OfficeIMO.Ocr.Process)

`OfficeIMO.Ocr.Process` adapts a caller-configured executable to the shared `OfficeIMO.Ocr.IOcrEngine` contract through a versioned JSON file protocol. It works with images from any source; it does not depend on Reader, PDF, or another document-format package. The executable runs directly, without an inserted command shell.

## Install

```powershell
dotnet add package OfficeIMO.Ocr.Process
```

## Configure an engine

```csharp
using OfficeIMO.Ocr;
using OfficeIMO.Ocr.Process;

var engine = new ProcessOcrEngine(new ProcessOcrEngineOptions {
    FileName = "/opt/my-ocr/recognize",
    Arguments = new[] { "--request", "{request}" },
    Id = "my-ocr",
    Timeout = TimeSpan.FromMinutes(1),
    MaxOutputBytes = 8L * 1024L * 1024L
});

OcrResult result = await engine.RecognizeAsync(new OcrRequest {
    Payload = File.ReadAllBytes("scan.png"),
    MediaType = "image/png",
    FileName = "scan.png",
    CandidateKind = "image",
    Language = "eng"
});

Console.WriteLine(result.Text);
```

Each call receives an isolated request directory. The provider writes the raster payload and a camel-case request JSON file with schema id `officeimo.ocr.process-request`, version `2`. The request's `outputPath` identifies where the executable must write a response envelope with schema id `officeimo.ocr.process-response`, version `2`, and an `OcrResult` in its `result` property. Use `ProcessOcrProtocol.SerializeResult(...)` when the external bridge is implemented in .NET.

Available argument placeholders are `{request}`, `{input}`, `{output}`, `{language}`, `{candidateId}`, `{sourceId}`, and `{pageNumber}`. They are substituted as individual process arguments, not shell text.

## Operational boundaries

- Input bytes, process stdout and stderr, response JSON size, and runtime are bounded by `ProcessOcrEngineOptions`.
- A format integration can add its own document-level limits. `OfficeIMO.Reader.Ocr`, for example, bounds candidate count, aggregate input bytes, concurrency, recognized text, and span counts.
- The runner contains descendants in a kill-on-close Windows Job Object, a `setsid` process group on Linux/Unix, or a POSIX session launched through the system Perl on macOS. It fails closed when the host cannot provide one of those containment boundaries.
- Executable paths, arguments, environment variables, and provider options are trusted host configuration. Do not build them directly from document content.
- Payload bytes are stored in owner-only per-request directories/files on Unix and are deleted by default. Set `KeepTemporaryFiles` only for controlled diagnostics.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0` (`net472` is also included on Windows builds).
- License: MIT.

## Dependency footprint

- **External:** A caller-configured executable. `System.Text.Json` is used for the versioned protocol on legacy targets.
- **OfficeIMO:** `OfficeIMO.Ocr` owns the engine-neutral request, result, geometry, capability, and diagnostic contracts.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
