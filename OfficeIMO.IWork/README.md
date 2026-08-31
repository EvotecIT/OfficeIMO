# OfficeIMO.IWork - bounded Apple iWork readers for .NET

`OfficeIMO.IWork` reads modern Apple Pages, Numbers, and Keynote packages without running iWork or executing embedded content. It owns ZIP, directory-bundle, nested `Index.zip`, Snappy-framed IWA, protobuf-envelope, package-resource, and unsupported-record preservation. Word, Excel, and PowerPoint remain the owners of editable destination documents.

Keep `OfficeIMO.IWork`, `OfficeIMO.Word`, `OfficeIMO.Excel`, and `OfficeIMO.PowerPoint` on the same coordinated OfficeIMO version.

## Install

Install the bounded reader and the destination owner for the format you need. For example, Numbers-to-Excel uses:

```powershell
dotnet add package OfficeIMO.IWork
dotnet add package OfficeIMO.Excel
```

Use `OfficeIMO.Word` with `OfficeIMO.IWork` for Pages, or `OfficeIMO.PowerPoint` with `OfficeIMO.IWork` for Keynote.

## Reference from a source checkout

For source-based development, reference the bounded reader and the semantic owner you need:

```xml
<ItemGroup>
  <ProjectReference Include="../OfficeIMO.IWork/OfficeIMO.IWork.csproj" />
  <ProjectReference Include="../OfficeIMO.Excel/OfficeIMO.Excel.csproj" />
</ItemGroup>
```

Use `OfficeIMO.Word` for Pages or `OfficeIMO.PowerPoint` for Keynote in place of the Excel owner. Keep all project references on the same checkout so their coordinated API and package contracts stay aligned.

## Read and inspect a source

```csharp
using OfficeIMO.IWork;

IWorkSourceDocument source = IWorkSourceDocument.Open("report.pages");
IWorkPagesProjection pages = source.ReadPages();

Console.WriteLine(source.Kind);                 // Pages
Console.WriteLine(source.ContainerKind);        // ZipPackage, DirectoryBundle, or nested Index.zip
Console.WriteLine(string.Join(", ", source.BuildVersions));
Console.WriteLine(pages.Paragraphs.Count);

IWorkImportReport report = pages.CreateImportReport(
    IWorkProjectionKind.EditableReconstruction);
foreach (IWorkArchiveRecord record in report.UnsupportedRecords) {
    Console.WriteLine($"{record.EntryPath}: {record.MessageType}");
}
```

Path and stream entry points use the same bounded parser. Streams require an explicit `IWorkDocumentKind` because they have no reliable filename:

```csharp
using FileStream stream = File.OpenRead("budget.numbers");
IWorkSourceDocument source = IWorkSourceDocument.Open(
    stream,
    IWorkDocumentKind.Numbers,
    new IWorkReadOptions {
        MaximumPackageBytes = 64 * 1024 * 1024,
        MaximumMaterializedCells = 1_000_000
    });

IWorkNumbersProjection workbook = source.ReadNumbers();
```

## Project into OfficeIMO owners

The public destination APIs live on their semantic owners:

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.IWork;
using OfficeIMO.IWork;

using IWorkNumbersLoadResult result = ExcelDocument.LoadNumbersWithReport(
    "budget.numbers",
    new IWorkReadOptions { ImportMode = IWorkImportMode.Auto });

ExcelDocument workbook = result.Document;
Console.WriteLine(result.ImportReport.ProjectionKind);
Console.WriteLine(result.HasConversionLoss);
workbook.Save("budget.xlsx");
```

The equivalent entry points are `WordDocument.LoadPages*` and `PowerPointPresentation.LoadKeynote*`. The short overload returns the destination document. The `WithReport` overload also returns the bounded source, typed projection, preserved records, diagnostics, producer build history, and the exact projection kind.

This is extended semantic reconstruction rather than plain-text extraction:

- Pages recovers rich paragraphs, source-proven list levels mapped to native Word numbering, page layout, section-specific headers/footers, positioned and sized accessible rich-text boxes, images, tables, and merges for editable Word projection.
- Numbers recovers sparse typed cells, supported formulas with cached values, merges, table metadata, and default sizing for editable Excel projection. Each source table receives its own worksheet so table-local formulas and column sizing remain stable; sheet-level text receives a separate worksheet when present.
- Keynote recovers slide size, order and names, positioned rich text with explicit inline breaks and source-proven list labels and levels, shape/run and presenter-note hyperlinks, notes, images, positioned and rotated tables, and merges for editable PowerPoint projection.

Advanced charts, vector effects, animations, comments/change tracking, masks/crops, and other application-only structures remain available in the preserved source records and are reported as conversion loss rather than silently claimed as editable.

`IWorkReadOptions` bounds decoded text characters, text items and attribute boundaries, cross-record style inheritance, projected sheets/slides/tables/images, repeated encoded destination-image bytes, merged ranges, and source-wide materialized cells in addition to the package/IWA limits.

`Auto` prefers editable semantic reconstruction. `EditableOnly` fails when supported editable structure cannot be recovered. `VisualOnly` selects the package's raster preview without traversing the application-specific semantic graph and reports `VisualFallback`; the corresponding `ReadPages`, `ReadNumbers`, or `ReadKeynote` call returns a diagnostic-only projection and does not claim that preview text or objects are editable. A preview may cover only the first page or a producer-generated composite, and that coverage is exposed on `IWorkPreviewAsset`. Embedded PDF inspection accepts bounded classic cross-reference tables and rejects unvalidated cross-reference streams.

## Preservation and authoring boundary

Every package entry and every decoded IWA payload remains available as defensive bytes on `IWorkSourceDocument`. Import reports conservatively retain every payload that is not losslessly represented, including records whose supported text or values were only partially consumed. The destination DOCX, XLSX, or PPTX contains the supported reconstruction or visual fallback; it is not a lossless iWork package rewrite.

There is deliberately no Pages, Numbers, or Keynote writer. OfficeIMO will not expose iWork save-back until an independently produced corpus demonstrates a stable deterministic round-trip contract across supported producer versions.

See the [iWork support matrix](../Docs/officeimo.iwork-support-matrix.md) for the version corpus, limits, semantic coverage, and known boundaries.

## Target frameworks and dependencies

`OfficeIMO.IWork` targets .NET Standard 2.0, .NET 8, .NET 10, and .NET Framework 4.7.2 on Windows. It depends only on `OfficeIMO.Core`; its IWA, Snappy, protobuf-envelope, and package readers are first-party implementations.
