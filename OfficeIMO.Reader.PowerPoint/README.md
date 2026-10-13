# OfficeIMO.Reader.PowerPoint

PowerPoint presentation support for `OfficeIMO.Reader.Core`.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.PowerPoint;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddPowerPointHandler()
    .Build();

OfficeDocumentReadResult result = reader.ReadDocument("slides.pptx");
```

The package supports PPTX/PPTM plus OfficeIMO.PowerPoint's legacy PPT/POT/PPS import path.
Its direct OfficeIMO dependencies are Reader.Core, PowerPoint and Core.

## Inspect comments and authored visibility

```csharp
OfficeDocumentReadResult inspected = reader.ReadDocument("review.pptx", new ReaderOptions {
    Inspection = new ReaderInspectionOptions()
});
foreach (OfficeDocumentMetadataEntry entry in inspected.Metadata) {
    if (entry.Category == OfficeDocumentInspectionCategories.Comment ||
        entry.Category == OfficeDocumentInspectionCategories.Visibility)
        Console.WriteLine($"Slide {entry.Location?.Slide}: {entry.Name} = {entry.Value}");
}
```

The adapter projects classic and modern comments from the owner model, preserving
available author, parent, status and shape identifiers. Visibility records include
authored hidden slides and hidden shapes inside groups. Original-package structural
observations, embedded payload/VBA descriptors and loaded-content safety evidence
remain separate categories. They do not execute presentation content or validate
signature mathematics.

The [shared inspection policy](../OfficeIMO.Reader.Core/README.md#inspect-review-and-saved-content)
controls requested categories and aggregate record/character limits. Exceeded
inspection or native safety ceilings terminate the operation. Legacy import limits
and unsupported-feature diagnostics still apply.
