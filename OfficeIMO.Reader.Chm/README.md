# OfficeIMO.Reader.Chm

Register bounded CHM ingestion in the modular Reader. Each HTML topic passes through the existing rich HTML reader, retaining paragraphs, tables, links, assets and diagnostics. Citations identify the archive and topic; source hashes cover the original CHM bytes.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Chm;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddChmHandler(new ReaderChmOptions {
        ReadOptions = new OfficeIMO.Chm.ChmReadOptions {
            MaxInputBytes = 64L * 1024 * 1024
        }
    })
    .Build();
OfficeDocumentReadResult document = reader.ReadDocument("manual.chm");
Console.WriteLine(document.Markdown);
```

`ReaderChmOptions` holds archive limits, topic-selection/aggregate projection limits, and rich HTML reader options. Registrations snapshot the supplied settings. `.chm` extensions and the `ITSF` signature participate in Reader detection. `OfficeIMO.Reader.All` includes this handler.

Results use `ReaderInputKind.Chm` (`27`) and document transport schema version 11. Per-topic identifiers are prefixed to keep rich content unique across the book. Locations such as `manual.chm!/guide/install.html` identify HTML topics rather than fixed pages. The archive is inert, and resources remain inside the CHM. See [the archive contract](../OfficeIMO.Chm/README.md) and [supported profiles](../OfficeIMO.Chm/SUPPORT.md).
