# OfficeIMO.Xps.Pdf

Convert XPS/OpenXPS pages to vector PDF through `OfficeIMO.Xps`,
`OfficeIMO.Core`, and `OfficeIMO.Pdf`. No external renderer is used.

```csharp
using OfficeIMO.Xps;

XpsDocument document = XpsDocument.Load("report.xps");
File.WriteAllBytes("report.pdf", document.ToPdf());
```

Page dimensions convert from XPS's 96 units per inch to PDF's 72 points per inch.
The bridge preserves the fixed page canvas instead of reflowing it. Embedded
text remains positioned vector outlines. A separate invisible text layer preserves
native Unicode clusters for search and copy, including ligatures, surrogate pairs,
spaces, right-to-left advances and nested affine transforms. Text inside decorative
visual brushes is excluded.

Search text follows source markup order; DocumentStructure reading order and PDF
accessibility tags are not reconstructed. Selection regions follow cluster advances
and glyph bounds, rather than inferred words. Source text remains searchable even
when native clipping or opacity hides it; clipping is not text redaction. Glyph-only
runs without UnicodeString have no recoverable logical text.

Export rejects known conversion losses. XPS-native metadata, print tickets,
structure tags, signatures, and document navigation are not PDF preservation
contracts. Consult the [XPS support matrix](../OfficeIMO.Xps/SUPPORT.md) and the
PDF engine's own rendering limits before choosing an archival workflow.
