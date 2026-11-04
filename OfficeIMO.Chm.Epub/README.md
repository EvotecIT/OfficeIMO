# OfficeIMO.Chm.Epub

Convert compiled help to a reflowable EPUB through the shared EPUB manuscript importer and writer. Topic contents hierarchy, internal links, supported HTML structures and embedded resources flow into the publication.

```csharp
using OfficeIMO.Chm;

ChmDocument book = ChmDocument.Load("manual.chm");
OfficeConversionReport report = book.SaveAsEpub("manual.epub");
foreach (var diagnostic in report.Diagnostics)
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
```

Use `ToEpub(options, out report)` or `ToEpubAsync` to obtain an editable `EpubPublication` before writing. `ChmEpubExportOptions` groups the settings: `ConversionOptions` (`ChmConversionOptions`) selects topics and bounds aggregate projection/output; `EpubManuscriptOptions` configures the existing importer, and `EpubSaveOptions` bounds package writing. Resource resolution is replaced with the archive-only resolver.

The report combines CHM projection, manuscript import and package-write evidence. Set `FailOnLoss` to reject reported loss before anything is written. Reflow can change topic CSS scope; the compiled keyword index, See Also and Windows help behavior are not recreated. Multi-target contents entries use one EPUB destination and report the reduction. See [CHM support](../OfficeIMO.Chm/SUPPORT.md) and [EPUB authoring](../OfficeIMO.Epub/README.md).

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Chm.Epub` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
