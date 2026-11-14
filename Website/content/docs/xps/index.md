---
title: "XPS and OpenXPS"
description: "Read, create, edit and save XPS and OpenXPS documents, extract text and logical structure, and convert fixed pages to SVG, images and tagged PDF in .NET."
order: 57
meta.seo_title: "Read, create and convert XPS and OpenXPS in .NET"
---

`OfficeIMO.Xps` reads, creates, edits and saves Microsoft XPS and ECMA-388 OpenXPS packages without Windows desktop APIs or an external conversion process. Pages convert to SVG and images through the shared OfficeIMO drawing engine. `OfficeIMO.Xps.Pdf` adds vector PDF output with a searchable text layer and tags from the document's native structure.

## Install

```shell
dotnet add package OfficeIMO.Xps
dotnet add package OfficeIMO.Xps.Pdf
```

Both packages target `netstandard2.0`, `net8.0`, `net10.0` and `net472` on Windows.

## Read and convert pages

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Xps;

XpsDocument document = XpsDocument.Load("report.oxps");
XpsPage page = document.Pages[0];

string text = page.ExtractText();
File.WriteAllText("page-1.svg", page.ToSvg().Svg);
File.WriteAllBytes("page-1.png", page.ToImage(OfficeImageExportFormat.Png).Bytes);
```

Coordinates are XPS units, 96 per inch. Pages follow the native document sequence, including multiple fixed documents. `ExtractText()` returns `UnicodeString` runs in markup order; it does not infer paragraphs or rebuild text from glyph IDs.

SVG export outlines embedded TrueType glyphs at their native positions and does not substitute installed fonts. Conversion is strict: a feature outside the rendering profile throws instead of disappearing. Pass `allowPartial: true` to `ToSvg` to get partial output with a list of `Diagnostics`.

## Create a document

```csharp
using OfficeIMO.Xps;

XpsDocument document = XpsDocument.Create(XpsFormat.OpenXps);
string font = document.AddFont(File.ReadAllBytes("licensed-font.ttf"));
XpsPage page = document.AddPage(816, 1056);
page.AddPath("M48,48 H768 V1008 H48 Z", "#FFF2F5FA");
page.AddText("Quarterly report", font, 28, 72, 112);
string image = document.AddResource("Resources/chart.png", File.ReadAllBytes("chart.png"), "image/png");
page.AddImage(image, 72, 160, 400, 240);
document.Save("report.oxps");
```

Use `XpsFormat.Xps` for Microsoft's original dialect; the file extension does not change the dialect. Fonts are obfuscated by default, and you need permission to embed them. Typed creation covers pages, vector paths, embedded fonts, Unicode glyph runs and image placement; it is a bounded fixed-page profile rather than a complete schema object model.

## Edit pages and documents

```csharp
using OfficeIMO.Xps;

XpsDocument document = XpsDocument.Load("report.oxps");
XpsFixedDocument appendix = document.AddDocument();
document.Documents[0].MovePageTo(0, appendix, 0);
appendix.AddPage(816, 1056);
document.MoveDocument(document.Documents.Count - 1, 0);
document.Save("reordered.oxps");
```

Each fixed document supports page insertion, removal, reordering and transfer. Moves keep page parts, resource references and link targets, and DocumentStructure outlines and story fragments stay aligned. For changes beyond the typed API, `GetMarkup()` returns a page's native XML as a detached `XElement`; edit it and call `ReplaceMarkup()`. `ReplaceResource()` swaps a resource's bytes while keeping its URI.

Saving keeps the original dialect, opaque package parts and native page markup. ZIP packaging and XML serialization are rebuilt, so output is not byte-identical to the source.

## Read logical structure

```csharp
using OfficeIMO.Xps;

XpsDocument document = XpsDocument.Load("report.oxps");
XpsLogicalStructure structure = document.ReadLogicalStructure();
foreach (XpsLogicalStory story in structure.Stories)
    foreach (XpsStructureNode block in story.Blocks)
        Console.WriteLine(block.Text);
foreach (string diagnostic in structure.Diagnostics)
    Console.WriteLine(diagnostic);
```

`ReadLogicalStructure()` follows native DocumentStructure and StoryFragments, merging paragraphs, lists and tables that continue across pages. Pages without StoryFragments get no inferred structure. `ToOfficeDocumentModel()` projects the same reading order into the shared document model, and `OfficeIMO.Reader.Xps` adds bounded chunks and source hashes for ingestion.

## Convert to PDF

```csharp
using OfficeIMO;
using OfficeIMO.Xps;

XpsDocument document = XpsDocument.Load("report.xps");
OfficeConversionReport report = document.SaveAsPdf("report.pdf");

document.SaveAsPdf("report-untagged.pdf", new XpsPdfExportOptions {
    PreserveLogicalStructure = false
});
```

The bridge keeps the fixed page canvas as vector output and adds an invisible Unicode text layer for search and copy. Native StoryFragments become paragraph, list, table and figure tags; authored `AutomationProperties` names become figure alternative text. Links, named destinations and DocumentStructure outlines become PDF navigation. Set `PdfOptions` for encryption or other PDF settings, and `FailOnLoss` to reject reported loss.

Unresolved or overlapping structure rejects strict export; `PreserveLogicalStructure = false` keeps the page canvas and markup-order search text without tags. Clipped or transparent source text stays searchable, so clipping is not redaction.

## Limits

- Loading is materialized in memory, bounded by compressed size, expanded size, part count, page count and XML depth. External resources and DTDs are rejected.
- Digitally signed packages can be read but not saved; signatures are not validated, and protected packages are not supported.
- Export does not establish PDF/UA or PDF/X conformance, and print tickets and signatures do not carry into PDF.
- Glyph runs without `UnicodeString` have no recoverable text, and missing figure descriptions are not inferred.
- Some image encodings, color profiles and extended stroke rules are diagnosed rather than rendered; see the support matrix.

## Related

- [XPS support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Xps/SUPPORT.md)
- [XPS API reference](/api/xps/)
- [Reader guide](/docs/reader/)
- [OfficeIMO.Xps product page](/products/xps/)
