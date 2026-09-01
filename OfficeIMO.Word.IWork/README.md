# OfficeIMO.Word.IWork

`OfficeIMO.Word.IWork` is the opt-in adapter for importing modern Apple Pages files into editable `OfficeIMO.Word` documents. Installing `OfficeIMO.Word` alone does not add the iWork reader.

```bash
dotnet add package OfficeIMO.Word.IWork
```

```csharp
using OfficeIMO.IWork;
using OfficeIMO.Word.IWork;

using IWorkPagesLoadResult result = WordIWorkConverter.LoadPagesWithReport(
    "source.pages",
    new IWorkReadOptions { ImportMode = IWorkImportMode.Auto });

Console.WriteLine(result.ImportReport.ProjectionKind);
result.Document.Save("converted.docx");
```

`LoadPages` returns the editable Word document directly. `LoadPagesWithReport` also retains the bounded source model, typed Pages projection, diagnostics, unsupported records, and exact editable-versus-visual-fallback result.

The adapter directly depends on `OfficeIMO.Core`, `OfficeIMO.IWork`, and `OfficeIMO.Word`. It does not add iWork support to the default Word package graph.

See the [iWork support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-support-matrix.md) for supported structures and conversion limits.
