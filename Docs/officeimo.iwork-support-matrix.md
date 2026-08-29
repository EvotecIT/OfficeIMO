# Apple iWork source-reader support

OfficeIMO reads modern IWA-based Pages, Numbers, and Keynote sources through one bounded `OfficeIMO.IWork` package layer. It then projects supported semantics into the existing Word, Excel, and PowerPoint owners. The source layer is read-only.

## Public ownership

| Source | Typed source projection | Editable destination owner | Entry points |
|---|---|---|---|
| Pages | `IWorkPagesProjection` | `WordDocument` | `LoadPages`, `LoadPagesWithReport` |
| Numbers | `IWorkNumbersProjection` | `ExcelDocument` | `LoadNumbers`, `LoadNumbersWithReport` |
| Keynote | `IWorkKeynoteProjection` | `PowerPointPresentation` | `LoadKeynote`, `LoadKeynoteWithReport` |

`IWorkSourceDocument.Open` is the advanced inspection entry point. It exposes normalized package entries, IWA payload records, producer build history, previews, diagnostics, and application-specific typed projections. The owner `WithReport` APIs retain that same source model beside the generated Office document.

## Package and IWA boundary

| Capability | Current contract |
|---|---|
| Containers | ZIP packages, directory bundles, and packages with nested `Index.zip` |
| Package safety | Configurable package, entry, entry-count, aggregate-uncompressed-size, and path bounds; duplicate, absolute, traversal, empty-segment, and linked directory entries are rejected |
| IWA framing | Raw Snappy chunks used by modern iWork, with declared-size, chunk-size, aggregate-size, copy-offset, truncation, and integer-overflow checks |
| Object envelope | Bounded ArchiveInfo and MessageInfo protobuf parsing with field-count, depth, record-count, record-size, wire-type, varint, and reference validation |
| Preservation | Defensive access to all package entries and all primary or auxiliary IWA payloads; import reports list payloads not consumed by the selected semantic projection |
| Active content | No macros, scripts, external links, embedded executables, or application services are executed |
| Legacy packages | Pre-IWA `index.xml` and `index.apxl` packages are rejected as unsupported rather than guessed |

`IWorkReadOptions` supplies all configurable limits. Defaults cap a package and aggregate expanded entries at 512 MiB, one entry at 128 MiB, one compressed IWA archive at 64 MiB, one decompressed archive at 256 MiB, all decompressed IWA archives at 512 MiB, one Snappy chunk at 64 MiB, one record at 128 MiB, and the source-wide record count at 1,000,000. Numbers adds row, column, and source-wide sparse materialized-cell limits.

## Editable semantic coverage

| Application | Reconstructed today | Preserved or diagnosed rather than reconstructed |
|---|---|---|
| Pages | Body paragraphs, section header/footer text, and text-box text | Exact typography, layout geometry, styles, lists, tables, charts, equations, media, comments, change tracking, fields, and application-only metadata |
| Numbers | Sheets, tables, declared dimensions, sparse typed text/number/Boolean/date/duration cells, cached formula markers, cell decode errors, and text-box text | Pre-BNC cell storage, exact canvas positions, formatting, formula expressions, rich-text runs, merges, filters, names, charts, forms, comments, media, and application-only metadata |
| Keynote | Slide order, skipped-slide state, title/body text, and presenter notes | Masters, layouts, exact geometry, typography, themes, tables, charts, builds, media, transitions, animations, comments, and application-only metadata |

Editable reconstruction means the supported content is represented as normal DOCX, XLSX, or PPTX objects and can be edited and saved through its owner. It does not mean the destination is visually identical or that unsupported iWork records are written into the Office package.

## Visual fallback

`IWorkImportMode.Auto` uses editable reconstruction when supported semantics exist and otherwise uses an embedded raster preview. `EditableOnly` rejects sources without supported editable structure. `VisualOnly` always requests the raster preview.

Every owner report exposes `IWorkProjectionKind.EditableReconstruction` or `IWorkProjectionKind.VisualFallback`. `IWorkPreviewAsset.Coverage` distinguishes a known full-document asset from a first-page or composite preview. Current owner adapters embed PNG or JPEG previews; embedded PDF previews remain available on the source model but are not silently rasterized.

## Corpus evidence

The checked-in interoperability corpus uses unmodified, MIT-licensed fixtures with recorded source revisions:

| Application | Producer/build history exercised |
|---|---|
| Pages | 14.1 and 14.5 |
| Numbers | 11.1, build histories spanning 13.x and 14.x, 14.5, and 15.1 |
| Keynote | 8.1 and 14.5 |

Tests assert path/stream parity, cumulative decompression and materialization bounds, application detection, source-record retention, typed projections including wide Numbers offsets, explicit handling of pre-BNC cell storage, strict preview selection, destination row limits, explicit visual fallback, and save/reopen of the resulting DOCX, XLSX, and PPTX packages. Fixture sources, revisions, expected content, and licenses are recorded in `OfficeIMO.TestAssets/Documents/IWorkCorpus/README.md`.

This corpus proves the current read contract; it does not establish a stable iWork write contract.

## Intentional authoring boundary

OfficeIMO does not create, edit in place, or write `.pages`, `.numbers`, or `.key` files. Authoring remains deferred until a broader independently produced corpus can prove deterministic package reconstruction, unknown-record retention, reopen behavior in multiple iWork versions, and stable semantic and visual round trips. Save the editable projections as DOCX, XLSX, or PPTX instead.
