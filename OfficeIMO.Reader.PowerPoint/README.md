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

## Project an open presentation

With `OfficeIMO.Reader.PowerPoint` imported, call
`presentation.ReadDocument(sourceName, options, formatOptions, cancellationToken)`
on a caller-owned `PowerPointDocument`. It uses the canonical Reader mapping,
snapshots both option sets and enforces the projection's output budgets. `Full`
is the default. The presentation remains open for subsequent
`InspectEmbeddedPayloads` reads.

The source name supplies citations without reopening a file or inventing its
byte hash, length or timestamp. Input and package admission settings cannot
retroactively constrain the earlier native load; configure them when loading
the presentation. Lean projection still requires a modern XML source model.

## Text and metadata projection

`ReaderOptions.Projection = ReaderProjection.TextAndMetadata` supports PPTX,
PPTM, POTX, POTM, PPSX, PPSM and PPAM. It projects plain slide/table/note text,
native slide provenance, links and requested inspection evidence without image
asset reads, chart JSON visuals or Markdown rendering. Native slide pages retain
their authored membership; no layout engine or raster renderer runs.

Requested package evidence uses bounded inert payload descriptors. Unknown sizes
remain explicit without a scan. VBA part/signature-relationship metadata is
retained; binary module decoding is marked `module-payload-unrequested`.

Binary PPT/POT/PPS/PPA require `Full`. Lean reads refuse the legacy projection
before pictures and embedded storages are reconstructed, including binary input
with a misleading modern filename. Ordinary reads retain their current support.
Native parser limits remain separate from Reader's aggregate result budgets.

`OpenXmlMaxPartUncompressedBytes` and `OpenXmlMaxTotalUncompressedBytes` apply
native declared part/aggregate expansion admission before model loading, even
when `Inspection` is null. Configured ceilings produce terminal
`ReaderResourceLimitException` with the native failure retained; they are not
retried as password failures. These are package admission bounds, not a hard
limit on process memory or every decoder's actual work.


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
