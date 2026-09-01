# OfficeIMO.PowerPoint.IWork

`OfficeIMO.PowerPoint.IWork` is the opt-in adapter for importing modern Apple Keynote files into editable `OfficeIMO.PowerPoint` presentations. Installing `OfficeIMO.PowerPoint` alone does not add the iWork reader.

```bash
dotnet add package OfficeIMO.PowerPoint.IWork
```

```csharp
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint.IWork;

using IWorkKeynoteLoadResult result = PowerPointIWorkConverter.LoadKeynoteWithReport(
    "source.key",
    new IWorkReadOptions { ImportMode = IWorkImportMode.Auto });

Console.WriteLine(result.ImportReport.ProjectionKind);
result.Document.Save("converted.pptx");
```

`LoadKeynote` returns the editable PowerPoint presentation directly. `LoadKeynoteWithReport` also retains the bounded source model, typed Keynote projection, diagnostics, unsupported records, and exact editable-versus-visual-fallback result.

The adapter depends on `OfficeIMO.IWork` and `OfficeIMO.PowerPoint`. It does not add iWork support to the default PowerPoint package graph.

See the [iWork support matrix](../Docs/officeimo.iwork-support-matrix.md) for supported structures and conversion limits.
