# OfficeIMO 3.0 public API review

This review records the complete public-surface comparison used for the coordinated 3.0 release. It is separate from package validation: package validation checks the assets inside each newly built package, while this review identifies the intentional changes from the audited 2.x source line.

## Compared release states

| Side | Git commit | Package metadata |
|---|---|---|
| Audited 2.x baseline | `955ce7b589512499aa42ba1da654b4a7742817f4` | 81 coordinated projects at `2.0.1` |
| Reviewed 3.0 tree | `584bff01202c7a1da2f8fc51b1d2b9636cc66821` | 81 coordinated projects at `3.0.0` |

The framework inventory is identical on both sides:

| Target | Compared assembly pairs |
|---|---:|
| `netstandard2.0` | 79 |
| `net8.0` | 81 |
| `net10.0` | 81 |
| `net472` | 80 |
| `net8.0-windows` and `net10.0-windows` WPF assets | 2 |
| **Total** | **323** |

Both trees were built in Release for every declared target. The .NET Framework assets were cross-targeted with Microsoft's reference assemblies; the WPF assets were cross-targeted with Windows targeting enabled. Windows CI remains the authoritative runtime check for those assets.

`Meziantou.Framework.PublicApiGenerator.Tool` 2.0.2 generated one deterministic public API stub for each of the 323 assembly/target pairs. Each pair was compared by package identity, with `OfficeIMO.Epub.Html` paired to its 3.0 replacement, `OfficeIMO.Epub.Image`.

## Complete result

Two hundred ninety-seven of the 323 assembly/target pairs are byte-for-byte identical. Seventy-four of the 81 package identities are unchanged on every target. Seven packages changed:

| 2.x identity | 3.0 identity | Changed targets | Added public declarations | Removed public declarations |
|---|---|---|---:|---:|
| `OfficeIMO.Drawing` | `OfficeIMO.Drawing` | `netstandard2.0`, `net472` | 0 | 1 |
| `OfficeIMO.Epub.Html` | `OfficeIMO.Epub.Image` | `netstandard2.0`, `net8.0`, `net10.0`, `net472` | 11 | 11 |
| `OfficeIMO.Excel` | `OfficeIMO.Excel` | `netstandard2.0`, `net8.0`, `net10.0`, `net472` | 1 | 735 |
| `OfficeIMO.Excel.Pdf` | `OfficeIMO.Excel.Pdf` | `netstandard2.0`, `net8.0`, `net10.0`, `net472` | 19 | 16 |
| `OfficeIMO.Pdf` | `OfficeIMO.Pdf` | `netstandard2.0`, `net8.0`, `net10.0`, `net472` | 34 | 0 |
| `OfficeIMO.PowerPoint.Pdf` | `OfficeIMO.PowerPoint.Pdf` | `netstandard2.0`, `net8.0`, `net10.0`, `net472` | 19 | 16 |
| `OfficeIMO.Word` | `OfficeIMO.Word` | `netstandard2.0`, `net8.0`, `net10.0`, `net472` | 14 | 25 |

The counts include declarations beginning with `public`; namespace, brace, and blank lines are excluded. The two WPF-only Windows pairs are unchanged.

## Breaking changes reviewed

### Compatibility shim visibility

The one older-target-only removal is `System.Runtime.CompilerServices.IsExternalInit`:

```csharp
namespace System.Runtime.CompilerServices {
    public static class IsExternalInit { }
}
```

The type was emitted by `OfficeIMO.Drawing` only for `netstandard2.0` and `net472`. It is a compiler compatibility shim rather than an OfficeIMO contract. The 3.0 assemblies still compile record and `init` syntax through an internal shim, but applications must not reference the OfficeIMO-provided type directly.

### PDF table adapters

The Excel and PowerPoint adapters now say “tables” in method and result names. Their reports add `SourceScope` and `HasOmittedPageContent`, so non-table PDF content can no longer be silently described as lossless. The exact old-to-new names are in the [3.0 migration guide](officeimo-3.0-migration.md#pdf-table-imports).

### Legacy XLS

`LegacyXlsLoadResult.Workbook`, `LegacyXlsLoadResult.ImportReport`, and `LegacyXlsLoadResult.CreateAdvancedImportReport()` are removed. `LegacyXlsLoadResult.CreateImportReport()` is added as the supported report entry point.

`LegacyXlsImportReport` retains exactly these 23 public declarations:

`WorksheetCount`, `ChartSheetCount`, `UnsupportedSheetCount`, `CellCount`, `FormulaCellCount`, `CommentCount`, `HyperlinkCount`, `DataValidationCount`, `ConditionalFormattingCount`, `AutoFilterCriteriaCount`, `DefinedNameCount`, `ExternalReferenceCount`, `PivotTableRecordCount`, `ChartRecordCount`, `DrawingRecordCount`, `UnsupportedFeatureCount`, `UnsupportedProjectionGapCount`, `PreservedFeatureRecordCount`, `ErrorCount`, `WarningCount`, `HasImportErrors`, `HasUnsupportedFeatures`, and `ToMarkdown()`.

The other 732 former `LegacyXlsImportReport` telemetry declarations are removed from the public contract. Structured diagnostics, unsupported and preserved feature collections, `HasImportErrors`, `HasUnsupportedFeatures`, `HasConversionLoss`, and the stable summary remain public through the owning result/report APIs. Exhaustive BIFF corpus counters remain internal for OfficeIMO tests and compatibility analysis.

### Word

`FormattingHelper`, `HorizontalAlignmentHelper`, `ImageShapeStyleHelper`, `InlineRunHelper`, `WordHelpers.GetNextSdtId`, and the mutable `WordListLevel._level` field are not 3.0 contracts. Their supported replacements are `WordParagraph.GetFormattedRuns()`, `WordFormattedRun`, owner alignment/image/content-control APIs, and the read-only `WordListLevel.OpenXmlElement`. `WordHelpers` is now static.

### EPUB image export

`OfficeIMO.Epub.Html` and its namespace become `OfficeIMO.Epub.Image`. The adapter exports images through the shared HTML image pipeline; it does not claim EPUB-to-HTML conversion.

## Additive PDF facade

`OfficeIMO.Pdf` has no removals in this comparison. The added surface covers document preflight, merge and resize conveniences, external-signature completion, image export, table-extraction scope, and text preflight. These additions keep the fluent consumer path on the owning PDF document rather than creating another facade package.

No other coordinated package or target-specific asset changed its generated public API between the compared states. The migration guide covers every intentional breaking category above; current generated API reference pages remain the source of truth for the complete 3.0 member inventory.
