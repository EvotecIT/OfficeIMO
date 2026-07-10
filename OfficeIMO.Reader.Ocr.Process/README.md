# OfficeIMO.Reader.Ocr.Process - external OCR process provider

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Ocr.Process)](https://www.nuget.org/packages/OfficeIMO.Reader.Ocr.Process)

`OfficeIMO.Reader.Ocr.Process` connects `OfficeIMO.Reader` to a caller-configured executable through a versioned JSON file protocol. The executable runs directly—no command shell is inserted by the provider.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Ocr.Process
```

## Configure an engine

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Ocr.Process;

var engine = new ProcessOfficeOcrEngine(new ProcessOfficeOcrEngineOptions {
    FileName = "/opt/my-ocr/recognize",
    Arguments = new[] { "--request", "{request}" },
    Id = "my-ocr",
    Timeout = TimeSpan.FromMinutes(1),
    MaxOutputBytes = 8L * 1024L * 1024L
});

OfficeDocumentReadResult source = DocumentReader.ReadDocument("scan.docx");
OfficeDocumentOcrExecutionResult execution = await source.ApplyOcrAsync(engine);

Console.WriteLine(execution.Document.Markdown);
```

Each call receives an isolated request directory. The provider writes the candidate payload and a camel-case request JSON file with schema id `officeimo.reader.ocr.process-request`, version `1`. The request's `outputPath` identifies where the executable must write a response envelope with schema id `officeimo.reader.ocr.process-response`, version `1`, and an `OfficeOcrEngineResult` in its `result` property. Use `ProcessOfficeOcrProtocol.SerializeResult(...)` when the external bridge is implemented in .NET.

Available argument placeholders are `{request}`, `{input}`, `{output}`, `{language}`, `{candidateId}`, and `{assetId}`. They are substituted as individual process arguments, not shell text.

## Operational boundaries

- Candidate count, input bytes, concurrency, engine duration, recognized text, and span counts remain bounded by `OfficeDocumentOcrExecutionOptions`.
- Process stdout and stderr, response JSON size, and runtime are bounded separately by `ProcessOfficeOcrEngineOptions`.
- Executable paths, arguments, environment variables, and provider options are trusted host configuration. Do not build them directly from document content.
- Payload bytes are temporary local files and are deleted by default. Set `KeepTemporaryFiles` only for controlled diagnostics.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0` (`net472` is also included on Windows builds).
- License: MIT.
