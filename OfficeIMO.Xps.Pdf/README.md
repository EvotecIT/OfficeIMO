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
text becomes positioned vector outlines, so the PDF does not contain searchable
text. Use `XpsPage.ExtractText()` when Unicode extraction is needed.

Export rejects known conversion losses. XPS-native metadata, print tickets,
structure tags, signatures, and document navigation are not PDF preservation
contracts. Consult the [XPS support matrix](../OfficeIMO.Xps/SUPPORT.md) and the
PDF engine's own rendering limits before choosing an archival workflow.
