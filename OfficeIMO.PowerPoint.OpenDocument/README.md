# OfficeIMO.PowerPoint.OpenDocument

Explicit conversion between `OfficeIMO.PowerPoint` presentations and native `OfficeIMO.OpenDocument` presentations.

```csharp
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;

using PowerPointPresentation presentation = PowerPointPresentation.Open("deck.pptx");
var conversion = presentation.ToOpenDocument();
conversion.Document.Save("deck.odp");

foreach (var mapping in conversion.Report.Mappings) {
    Console.WriteLine($"{mapping.Feature}: {mapping.Status} ({mapping.Count})");
}
```

The adapter maps slide size and order, hidden slides, text boxes and basic run formatting, images, tables and merges, basic shapes, solid backgrounds, common transitions, and plain speaker notes. Masters, complex geometry, charts, SmartArt, media, animations, unsupported transition families, and other detected advanced features are called out in the conversion report.
