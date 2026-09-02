# OfficeIMO.PowerPoint.IWork

`OfficeIMO.PowerPoint.IWork` is the opt-in adapter for importing modern Apple Keynote files into editable `OfficeIMO.PowerPoint` presentations. Installing `OfficeIMO.PowerPoint` alone does not add the iWork reader.

```bash
dotnet add package OfficeIMO.PowerPoint.IWork
```

```csharp
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint.IWork;

IWorkSourceDocument source = IWorkSourceDocument.Open("source.key");
using KeynoteToPowerPointResult result = source.ToPowerPointPresentationResult(
    new IWorkConversionOptions { Mode = IWorkConversionMode.Auto });

Console.WriteLine(result.Report.ProjectionKind);
Console.WriteLine(result.HasLoss);
result.Value.Save("converted.pptx");
```

`IWorkSourceDocument.Open` reads and bounds the source independently of destination policy. `ToPowerPointPresentation` returns the converted presentation directly; `ToPowerPointPresentationResult` also exposes the typed Keynote projection, diagnostics, preserved source records, and exact editable-versus-visual-fallback result. `PowerPointIWorkConverter.ConvertKeynoteToPowerPoint*` provides equivalent path and stream convenience entry points.

The adapter directly depends on `OfficeIMO.Core`, `OfficeIMO.IWork`, and `OfficeIMO.PowerPoint`. It does not add iWork support to the default PowerPoint package graph.

See the [iWork support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-support-matrix.md) for supported structures and conversion limits.
