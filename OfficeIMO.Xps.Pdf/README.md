# OfficeIMO.Xps.Pdf

Convert XPS/OpenXPS pages to vector PDF through `OfficeIMO.Xps`,
`OfficeIMO.Core`, and `OfficeIMO.Pdf`. No external renderer is used.

```csharp
using OfficeIMO.Xps;

XpsDocument document = XpsDocument.Load("report.xps");
File.WriteAllBytes("report.pdf", document.ToPdf());
```

For a caller-owned output stream, use `SavePdf`. It returns the PDF engine's
`PdfSaveResult` and accepts its existing output settings:

```csharp
using OfficeIMO.Pdf;

using var output = File.Create("report.pdf");
PdfSaveResult saved = document.SavePdf(output, new XpsToPdfOptions {
    PdfOptions = new PdfOptions().SetEncryption("document-password")
}, cancellationToken);
```

The stream remains open. A successful save replaces and rewinds seekable output.
Validation runs before writing, but a serialization failure can leave partial
stream contents. Stage and reopen output before publishing it when the destination
must survive a failed operation. [OfficeIMO.Workflows](../OfficeIMO.Workflows/README.md)
provides that publication contract for `.xps` and `.oxps` through `xps-pdf`.

Page dimensions convert from XPS's 96 units per inch to PDF's 72 points per inch.
The bridge preserves the fixed page canvas instead of reflowing it. Embedded
text remains positioned vector outlines. A separate invisible text layer preserves
native Unicode clusters for search and copy, including ligatures, surrogate pairs,
spaces, right-to-left advances and nested affine transforms. Text inside decorative
visual brushes is excluded.

Radial Repeat/Reflect brushes remain vector PDF shadings even when tangent
bounds or high-frequency fields prevent finite stop expansion. Color interpolation,
stop alpha and native outside-cone endpoint paint are retained. Print-condition
conversion configured through `XpsToPdfOptions.PdfOptions` uses the caller-supplied
CMYK ICC profile and bounded vector color samples; alpha remains independent.
This color conversion alone does not establish PDF/X conformance.

Native StoryFragments and DocumentStructure supply paragraph, list, table and figure
tags. Continued blocks share their logical containers across pages; declared story
order can differ from physical page order. List markers and table cell spans retain
their authored roles. Headers and footers remain searchable as artifacts outside the
body structure tree. Page-local fragments use page order when no story addresses
were authored. Unreferenced text remains in markup order in a generic division.
Unstructured documents retain markup-order search text without inferred tags.

Authored `AutomationProperties.Name` and `AutomationProperties.HelpText` on a
figure's named Path or Canvas become its PDF alternative text. Name precedes
HelpText; multiple named targets follow the figure's reference order. Descriptions
remain separate from the painted and searchable Unicode layer.

Selection regions follow cluster advances and glyph bounds, rather than inferred
words. Source text remains searchable even when native clipping or opacity hides
it; clipping is not text redaction. Glyph-only runs without UnicodeString have no
recoverable logical text. Missing figure descriptions are not inferred, and export
does not establish PDF/UA conformance.

Unsupported structure extensions, unresolved references and overlapping semantic
ownership reject strict export. To retain the fixed canvas and markup-order search
text without mapping native semantics, use:

```csharp
File.WriteAllBytes("report.pdf", document.ToPdf(preserveLogicalStructure: false));
```

Safe web/mail links and native page, document and sequence targets become PDF
link annotations. Named elements become destinations with page-space positions;
DocumentStructure outlines retain their titles, order, depth and safe URI actions.
Repeated page references keep local fragment links on their own occurrence, while
absolute page references resolve to the first occurrence. Path target positions use
conservative geometry bounds; link hit areas are rectangles.

Export rejects known conversion losses and unsafe or unresolved outline targets.
Opaque XPS metadata, print tickets and signatures are not PDF preservation contracts. Consult the [XPS support matrix](../OfficeIMO.Xps/SUPPORT.md) and the
PDF engine's own rendering limits before choosing an archival workflow.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Xps.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
