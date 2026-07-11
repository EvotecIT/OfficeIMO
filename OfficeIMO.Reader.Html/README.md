# OfficeIMO.Reader.Html - HTML reader adapter

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Html)](https://www.nuget.org/packages/OfficeIMO.Reader.Html)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Html?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Html)

`OfficeIMO.Reader.Html` registers a modular HTML ingestion adapter for `OfficeIMO.Reader`.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Html
```

## Register

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Html;

DocumentReaderHtmlRegistrationExtensions.RegisterHtmlHandler();
```

For untrusted or size-sensitive HTML:

```csharp
DocumentReaderHtmlRegistrationExtensions.RegisterHtmlHandler(
    htmlOptions: ReaderHtmlOptions.CreateUntrustedHtmlProfile(maxInputCharacters: 100_000),
    replaceExisting: true);
```

## Examples

### Convert HTML to Markdown chunks

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Html;

DocumentReaderHtmlRegistrationExtensions.RegisterHtmlHandler(
    htmlOptions: ReaderHtmlOptions.CreatePortableProfile(),
    replaceExisting: true);

foreach (var chunk in DocumentReader.Read("page.html", new ReaderOptions {
    MarkdownChunkByHeadings = true,
    MaxChars = 5_000
})) {
    Console.WriteLine($"{chunk.Id}: {chunk.Location.HeadingPath}");
    Console.WriteLine(chunk.Markdown ?? chunk.Text);
}
```

### Extract tables from HTML

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Html;

DocumentReaderHtmlRegistrationExtensions.RegisterHtmlHandler();

IReadOnlyList<ReaderTable> tables = DocumentReader.ReadTables("report.html",
    new ReaderOptions {
        MaxTableRows = 250
    });

foreach (var table in tables) {
    Console.WriteLine($"{table.Rows.Count} row(s), {table.ColumnProfiles.Count} column profile(s)");
}
```

### Read the rich document result

```csharp
OfficeDocumentReadResult document =
    DocumentReaderHtmlExtensions.ReadHtmlDocument("application.html");

foreach (OfficeDocumentFormField field in document.Forms) {
    Console.WriteLine($"{field.Name}: {field.Value}");
}

foreach (OfficeDocumentLink link in document.Links) {
    Console.WriteLine(link.Uri ?? link.DestinationName);
}
```

After registration, `DocumentReader.ReadDocument("application.html")` dispatches to this native rich mapping as well.

## What it emits

- HTML converted to Markdown through `OfficeIMO.Markdown.Html`.
- Markdown-shaped `ReaderChunk` output.
- Table extraction with `ReaderTable.ColumnProfiles`.
- Heading-aware chunk metadata when `ReaderOptions.MarkdownChunkByHeadings` is enabled.
- HTML-to-Markdown profile, transform, converter, and visual round-trip option pass-through.
- A schema-v5 rich result containing semantic blocks, figures, tables, links, form controls, media visuals, metadata, and bounded data-URI image assets.

## Boundaries

- Reader adapter registration belongs here.
- HTML to Markdown conversion belongs in `OfficeIMO.Markdown.Html`.
- Shared extraction contracts belong in `OfficeIMO.Reader`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
