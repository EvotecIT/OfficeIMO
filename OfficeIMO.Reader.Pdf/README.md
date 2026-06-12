# OfficeIMO.Reader.Pdf

Thin PDF adapter for `OfficeIMO.Reader`.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;

DocumentReaderPdfRegistrationExtensions.RegisterPdfHandler();

IReadOnlyList<ReaderChunk> chunks = DocumentReader
    .Read("invoice.pdf")
    .ToList();
```

For service hosts that load modular reader adapters by assembly, bootstrap the
PDF adapter and export the merged capability manifest in one step:

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;

ReaderHostBootstrapResult result = DocumentReader.BootstrapHostFromAssemblies(
    new[] { typeof(DocumentReaderPdfRegistrationExtensions).Assembly },
    new ReaderHostBootstrapOptions {
        ReplaceExistingHandlers = true,
        IncludeBuiltInCapabilities = true,
        IncludeCustomCapabilities = true
    });

string manifestJson = result.ManifestJson;
```

The adapter uses `OfficeIMO.Pdf`'s logical read model and emits page-aware chunks with `ReaderLocation.Page`, Markdown text, detected tables with column profiles and diagnostics, source/security/form/open-action/active-content counters in `ReaderChunk.Diagnostics`, table confidence aggregates, table/image geometry coverage, selected form-widget appearance coverage, passive action summaries in `ReaderChunk.Actions`, image placeholders plus image visual geometry in `ReaderChunk.Visuals`, link annotations, and typed form fields with widget geometry, current appearance state names, and normal appearance state names when available. Action summaries identify document-open, catalog, page, and annotation sources plus scope, trigger, path, and action type without exposing JavaScript bodies or other executable payloads.

```csharp
using OfficeIMO.Reader.Pdf;

ReaderPdfProfileContract contract = ReaderPdfProfileContracts.OfficeIMO;

Console.WriteLine(contract.Id);
Console.WriteLine(contract.OutputContract);
```

`ReaderPdfProfileContracts.OfficeIMO` exposes the stable handler identifier,
pipeline, chunk metadata contract, safety behavior, and unsupported scope for
hosts that need a capability manifest or user-facing adapter description.
