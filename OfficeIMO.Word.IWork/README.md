# OfficeIMO.Word.IWork

`OfficeIMO.Word.IWork` is the opt-in adapter for importing modern Apple Pages files into editable `OfficeIMO.Word` documents. Installing `OfficeIMO.Word` alone does not add the iWork reader.

```bash
dotnet add package OfficeIMO.Word.IWork
```

```csharp
using OfficeIMO.IWork;
using OfficeIMO.Word.IWork;

IWorkSourceDocument source = IWorkSourceDocument.Open("source.pages");
using PagesToWordResult result = source.ToWordDocumentResult(
    new IWorkConversionOptions { Mode = IWorkConversionMode.Auto });

Console.WriteLine(result.Report.ProjectionKind);
Console.WriteLine(result.HasLoss);
result.Value.Save("converted.docx");
```

`IWorkSourceDocument.Open` reads and bounds the source independently of destination policy. `ToWordDocument` returns the converted document directly; `ToWordDocumentResult` also exposes the typed Pages projection, diagnostics, preserved source records, and exact editable-versus-visual-fallback result. `WordIWorkConverter.ConvertPagesToWord*` provides equivalent path and stream convenience entry points.

The adapter directly depends on `OfficeIMO.Core`, `OfficeIMO.IWork`, and `OfficeIMO.Word`. It does not add iWork support to the default Word package graph.

See the [iWork support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-support-matrix.md) for supported structures and conversion limits.
