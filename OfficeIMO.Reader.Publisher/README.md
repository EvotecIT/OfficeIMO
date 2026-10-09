# OfficeIMO.Reader.Publisher

`OfficeIMO.Reader.Publisher` reads native `.pub` publications through
[OfficeIMO.Publisher](../OfficeIMO.Publisher/README.md). It retains complete source
stories, native bullet markers, physical page inventory, embedded-image descriptors
and recovery diagnostics in the shared Reader result.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Publisher;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddPublisherHandler(new ReaderPublisherOptions {
        ReadOptions = new OfficeIMO.Publisher.PublisherReadOptions {
            MaximumPages = 256
        },
        IncludeImagePayloads = false
    })
    .Build();

OfficeDocumentReadResult result = reader.ReadDocument("newsletter.pub",
    new ReaderOptions { MaxChars = 4_000, ComputeHashes = true });

foreach (OfficeDocumentBlock block in result.Blocks)
    Console.WriteLine(block.Text);
foreach (OfficeDocumentDiagnostic diagnostic in result.Diagnostics)
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
```

The handler is also part of `OfficeIMO.Reader.All`. Its stable registration ID is
`officeimo.reader.publisher`; configure it through `ReaderAllOptions.Publisher`.
Content detection recognizes the compound-file Contents and Quill stream pair,
including publications with an unfamiliar filename extension. The native codec
then validates the actual supported generation.

## Projection contract

Every source story appears once in native story order, including text that
overflows its printable frames. Paragraph blocks preserve source characters and
carry a story anchor, paragraph index and logical order that survives canonical
traversal, JSON transport and semantic PDF projection. Native bullet labels remain separate
markers; Markdown uses an ordinary list marker and escapes literal source
punctuation. Character references preserve indentation without turning source
paragraphs into Markdown code blocks. Long paragraphs split into bounded chunks without discarding text
or separating a UTF-16 surrogate pair.

`Pages` describes recovered physical pages and their dimensions. Paragraphs and
images have no inferred page assignment. Page placement, frame flow, tables and
typography remain available in the native Publisher model and are omitted from
this semantic projection with `PUB_READER_LAYOUT_OMITTED`. Use
[OfficeIMO.Publisher.Pdf](../OfficeIMO.Publisher.Pdf/README.md) for positioned output.
Native recovery findings retain their codes, loss kinds and source locations in
Reader diagnostics.

Image descriptors contain native image-store IDs, media types and byte counts.
Set `IncludeImagePayloads = true` to retain detached copies of the original bytes.
With `ComputeHashes = true`, retained payloads also receive hashes. External links
are not fetched and active content is not executed.

## Limits and source identity

Registration snapshots `ReaderPublisherOptions` and its native read settings.
`ReaderOptions.MaxInputBytes` can tighten the native input ceiling. Native item
and text ceilings also bound projection work, and shared Reader resource limits
apply to emitted chunks, pages and assets. `MaximumPages` counts both document
and master definitions.

Stream input follows Reader's whole-stream snapshot contract: seekable streams
are read from the beginning, their original position is restored, and the
caller's stream stays open. Non-seekable streams are captured from their current
forward position. Source hashes cover the captured bytes.

Selected Publisher 2002-and-later Contents/Quill publications are qualified.
Earlier generations throw `NotSupportedException`; they do not fall back to text
salvage. Malformed native data and exceeded native limits fail explicitly.
See the [native support contract and fixture evidence](../OfficeIMO.Publisher/SUPPORT.md).
The adapter references only `OfficeIMO.Reader.Core` and `OfficeIMO.Publisher`.
