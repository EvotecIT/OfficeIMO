# OfficeIMO.Chm

Read compiled HTML Help (`.chm`) in process on Windows, macOS, and Linux. `ChmDocument` exposes HTML topics, exact entry bytes, embedded resources, hierarchical contents, and the keyword index. Loading does not launch the Windows help viewer, execute scripts, or fetch external resources.

## Read a book

```csharp
using OfficeIMO.Chm;

ChmDocument book = ChmDocument.Load("manual.chm", new ChmReadOptions {
    MaxInputBytes = 64L * 1024 * 1024,
    MaxExpandedBytes = 128L * 1024 * 1024
});
Console.WriteLine(book.Title);
foreach (ChmTopic topic in book.Topics) {
    Console.WriteLine($"{topic.Title}: {topic.Path}");
    string html = topic.ReadHtml();
}
foreach (ChmNavigationItem item in book.TableOfContents) {
    Console.WriteLine($"{item.Name}: {item.Children.Count} children");
}
byte[] image = book.FindEntry("/images/logo.png")!.GetBytes();
```

`Load` accepts paths, bytes, and streams; `LoadAsync` accepts paths and streams. Seekable streams are read from the beginning, restored to their original position, and left open. Forward-only streams are consumed from their current position and left open. The result owns an independent snapshot. `GetBytes()` returns a copy; `OpenRead()` exposes a read-only stream.

Topics follow contents order, then unlisted HTML topics in ordinal path order. Contents and index items retain children, multiple links, and See Also labels. `FindEntry(reference, sourcePath)` resolves relative paths, case differences, URI escapes, queries, and fragments inside the archive. It returns null for missing, unsafe, external, and merged-book references. Raw directory identity remains available as `ChmEntry.Name`.

## Convert through existing engines

| Package | Entry point | Result |
| --- | --- | --- |
| `OfficeIMO.Chm` | `ToHtmlDocumentResult()` | Linked canonical HTML book and fidelity report |
| [OfficeIMO.Chm.Markdown](../OfficeIMO.Chm.Markdown/README.md) | `ToMarkdownResult()` | Linked Markdown book and fidelity report |
| [OfficeIMO.Chm.Epub](../OfficeIMO.Chm.Epub/README.md) | `ToEpubBytesResult()` | Reflowable EPUB with contents and embedded resources |
| [OfficeIMO.Chm.Pdf](../OfficeIMO.Chm.Pdf/README.md) | `ToPdfBytesResult()` | Independently rendered topic pages and combined fidelity evidence |
| [OfficeIMO.Reader.Chm](../OfficeIMO.Reader.Chm/README.md) | `AddChmHandler()` | Reader chunks, tables, links, assets, hashes, and topic citations |

```csharp
ChmConversionResult<OfficeIMO.Html.HtmlConversionDocument> htmlBook =
    book.ToHtmlDocumentResult(new ChmConversionOptions {
        TopicPaths = new[] { "/introduction.html", "/installation.html" },
        MaxOutputBytes = 32L * 1024 * 1024
    });
foreach (var finding in htmlBook.Report.FidelityDiagnostics)
    Console.WriteLine($"{finding.Code}: {finding.Message}");
```

Topic selection retains book order and rejects unknown or non-topic paths. `RequireValue()` rejects conversion failure; `Report.RequireNoLoss()` additionally rejects reported approximations and omissions. Raw archive extraction and converted-document fidelity are separate contracts.

## Bounds and ownership

Default read limits are 128 MiB input, 256 MiB expanded LZX data, 32 MiB per entry, 100,000 entries/navigation items, and bounded HTML/tree depth. Limits reject the operation rather than returning a truncated book. `ChmReadException.Code` identifies malformed or unsupported archives and exhausted limits. Cancellation is cooperative during reads, directory/navigation parsing, decompression, and conversion.

`ChmTopic.ToHtmlDocument()` exposes the existing inert HTML model. `ConfigureRenderOptions(options, topic.Path)` supplies a virtual `chm://archive/` base and an archive-only resource resolver. It replaces an existing resolver; remote images, fonts, stylesheets, and files are never loaded by CHM conversion.

The package references `OfficeIMO.Core` and `OfficeIMO.Html`. LZX decoding is shared with OneNote in Core; HTML parsing and charset services retain their existing provider boundary. There is no CHMLib, Windows API, external executable, or new third-party codec dependency in the product.

See [supported profiles and qualification](SUPPORT.md) for exact format and export boundaries.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 0 | 0 | 0 | 0 | 1 | 0 |
| Read | 1 | 0 | 0 | 0 | 0 | 0 |
| Edit | 0 | 0 | 0 | 0 | 1 | 0 |
| Preserve | 0 | 0 | 0 | 0 | 1 | 0 |
| Inspect | 1 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Chm` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
