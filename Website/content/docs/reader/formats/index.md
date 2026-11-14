---
title: Check Supported Formats
description: List the formats a configured OfficeIMO.Reader instance can read, detect a file's kind, and register every adapter at once, with C#, PSWriteOffice, and the officeimo CLI.
order: 60
packages:
  - OfficeIMO.Reader.Core
cmdlets:
  - Get-OfficeDocumentCapability
limits: "A registered handler is not a claim of complete format preservation; each capability lists its format qualifications."
search_terms:
  - supported formats
  - file types
  - reader formats
commands:
  - officeimo reader capabilities
verified: 2026-11-14
---

A reader only handles the formats you registered, so the supported set is a property of your configuration. The full list of adapter packages and extensions is in the [supported formats table](/docs/reader/#supported-formats); this page shows how to check it at run time.

## List what a reader can read

### C#

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.Word;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddWordHandler()
    .AddPdfHandler()
    .Build();

foreach (ReaderHandlerCapability capability in reader.GetCapabilities()) {
    Console.WriteLine($"{capability.DisplayName}: {string.Join(", ", capability.Extensions)}");
}

File.WriteAllText("reader-capabilities.json", reader.GetCapabilityManifestJson(indented: true));
```

### PowerShell

```powershell
Get-OfficeDocumentCapability | Format-Table DisplayName, Extensions
Get-OfficeDocumentCapability -Manifest
```

### CLI

```sh
officeimo reader capabilities
officeimo reader capabilities --format json
```

## Detect a file's kind

Check a file before reading it, for example to route uploads:

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddPdfHandler().Build();
Console.WriteLine(reader.DetectKind("upload.pdf"));
```

## Register every adapter

`OfficeIMO.Reader.All` registers all adapters in one call. Use it for hosts that accept any document, and selective adapters for smaller deployments.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddAllOfficeIMOHandlers(new ReaderAllOptions())
    .Build();
```
