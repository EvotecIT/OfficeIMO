# OfficeIMO.Opml

`OfficeIMO.Opml` creates, reads, edits, validates, converts, and writes OPML 1.0 and 2.0 documents without a third-party parser. A declared OPML 1.1 document is read using the 1.0 profile, as specified by OPML 2.0.

```csharp
using OfficeIMO.Opml;

OpmlDocument document = OpmlDocument.Create(OpmlVersion.Opml20);
document.Head.Title = "Subscriptions";

OpmlOutline folder = document.AddOutline("Technology");
OpmlOutline feed = folder.AddChild("OfficeIMO");
feed.Type = "rss";
feed.XmlUrl = "https://example.com/feed.xml";
feed.HtmlUrl = "https://example.com/";

document.Validate();
document.Save("subscriptions.opml");
```

Standard subscription attributes are typed, while `GetAttribute` and `SetAttribute` accept `XName` for namespaced extensions. Unknown attributes, elements, comments, processing instructions, and namespace declarations remain in the backing `XDocument`. An unchanged file or stream is written byte-for-byte; after an edit, preserved XML is serialized around the changed values.

`ToOfficeDocumentModel()` and `FromOfficeDocumentModel()` map nested outlines through `OfficeDocumentModel.Structure`. Conversion results implement `IOfficeConversionReport`, so callers can inspect diagnostics or call `RequireNoLoss()`.

## Safety and profile limits

- XML DTDs and external entity resolution are disabled.
- Default limits are 16 MiB encoded input, 16 million XML characters, 128 levels, 100,000 outlines, and 500,000 attributes. Override them with `OpmlReadOptions`.
- Validation covers the OPML root/head/body contract, declared version, required `text`, RSS `xmlUrl`, and link/include `url`.
- Reader integration is provided by `OfficeIMO.Reader.Opml`.

The implemented vocabulary follows the [OPML 2.0 specification](https://opml.org/spec2.opml). Targets: `netstandard2.0`, `net8.0`, `net10.0`, and `net472` on Windows.

## Dependency footprint

- **External:** None.
- **OfficeIMO:** `OfficeIMO.Core`.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
