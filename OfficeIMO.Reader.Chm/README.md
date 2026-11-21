# OfficeIMO.Reader.Chm

Register bounded CHM ingestion in the modular Reader. Each HTML topic passes through the existing rich HTML reader, retaining paragraphs, tables, links, assets and diagnostics. Citations identify the archive and topic; source hashes cover the original CHM bytes.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Chm;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddChmHandler(new ReaderChmOptions {
        ReadOptions = new OfficeIMO.Chm.ChmLoadOptions {
            MaxInputBytes = 64L * 1024 * 1024
        }
    })
    .Build();
OfficeDocumentReadResult document = reader.ReadDocument("manual.chm");
Console.WriteLine(document.Markdown);
```

`ReaderChmOptions` holds archive limits, topic-selection/aggregate projection limits, and rich HTML reader options. Registrations snapshot the supplied settings. `.chm` extensions and the `ITSF` signature participate in Reader detection. `OfficeIMO.Reader.All` includes this handler.

## Options

`ReaderChmOptions` has three parts. Each is cloned when the handler is registered.

| Property | Default | Purpose |
|---|---|---|
| `ReadOptions` | `new ChmLoadOptions()` | Archive limits. `MaxInputBytes` also becomes the handler's input limit. |
| `ConversionOptions` | `new ChmConversionOptions()` | Topic selection and aggregate projection limits. |
| `HtmlOptions` | `ReaderHtmlOptions.CreateOfficeIMOProfile()` | Rich HTML reader settings applied to every topic. |

| `ChmLoadOptions` | Default |
|---|---|
| `MaxInputBytes` | 128 MiB |
| `MaxExpandedBytes` | 256 MiB |
| `MaxEntries` | 100,000 |
| `MaxEntryBytes` | 32 MiB |
| `MaxPathLength` | 4,096 |
| `MaxNavigationItems` / `MaxNavigationReferences` | 100,000 each |
| `MaxNavigationCharacters` | 16 Mi characters |
| `MaxNavigationDepth` | 128 |
| `MaxHtmlNodes` / `MaxHtmlDepth` (contents and index files) | 200,000 / 256 |
| `TextEncoding` | `null`: the help-book code page, falling back to Windows-1252 with a `CHM_ENCODING_FALLBACK` warning |

| `ChmConversionOptions` used by Reader | Default | Effect |
|---|---|---|
| `TopicPaths` | `null` (all HTML topics) | Reads only the listed topics; an unknown or non-HTML path throws `ArgumentException`. |
| `MaxTopics` | 20,000 | Selections above the limit fail with `CONVERSION_LIMIT`. |
| `MaxTotalHtmlCharacters` | 64 Mi characters | Aggregate topic HTML read across the book. |
| `MaxHtmlNodes` | 1,000,000 | Aggregate HTML nodes parsed across all topics. |

Exceeding a limit fails the read instead of returning a silently partial book.

## Output

- **Chunks:** the rich HTML reader's chunks for each topic, re-labelled with `Kind = ReaderInputKind.Chm`, an id prefix `chm-topic-00001-`, `chm-topic-00002-`, ... in topic order, `Location.SourceBlockKind = "topic"`, and `Location.SourceBlockIndex` set to the zero-based topic index.
- **Heading paths** start with the topic title: `Installation / Requirements`.
- **Locations:** chunk paths use `manual.chm!/guide/install.html` style virtual paths.
- **Document:** `Source.Title` is the book title. Blocks, tables, links, forms, visuals and assets from all topics are merged into one result.
- **Diagnostics:** topic-level HTML diagnostics plus recoverable archive warnings from the CHM container, each located at `manual.chm!/<entry>`.
- **Capabilities:** `officeimo.reader.chm.rich-v5`, `officeimo.chm.archive`, `officeimo.chm.contents` and `officeimo.chm.index`, plus those reported by the HTML reader.

Results use `ReaderInputKind.Chm` (`27`) and document transport schema version 11 or later. Per-topic identifiers are prefixed to keep rich content unique across the book. Locations such as `manual.chm!/guide/install.html` identify HTML topics rather than fixed pages. The archive is inert, and resources remain inside the CHM. See [the archive contract](../OfficeIMO.Chm/README.md) and [supported profiles](../OfficeIMO.Chm/SUPPORT.md).
