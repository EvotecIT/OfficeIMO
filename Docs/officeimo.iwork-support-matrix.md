# Apple iWork source-reader support

OfficeIMO reads modern IWA-based Pages, Numbers, and Keynote sources through one bounded `OfficeIMO.IWork` package layer. It then projects supported semantics into the existing Word, Excel, and PowerPoint owners. The source layer is read-only.

The current level is **extended semantic reconstruction**. Normal document content becomes editable destination objects with recovered structure, typography, sizing, and geometry where the corpus proves those fields. It is not a pixel-identical renderer or a claim that every application-only feature can be translated.

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
| Package safety | Configurable package, entry, entry-count, aggregate-uncompressed-size, and path bounds; duplicate, absolute, traversal, empty-segment, and linked directory entries are rejected; directory reads verify the opened regular-file handle remains under the captured physical package root |
| IWA framing | Raw Snappy chunks used by modern iWork, with declared-size, chunk-size, aggregate-size, copy-offset, truncation, and integer-overflow checks |
| Object envelope | Bounded ArchiveInfo and MessageInfo protobuf parsing with field-count, depth, record-count, record-size, wire-type, varint, reference, and unique primary-object validation |
| Preservation | Defensive access to all package entries and all primary or auxiliary IWA payloads; import reports conservatively retain payloads not losslessly represented, including partially consumed records |
| Active content | No macros, scripts, external links, embedded executables, or application services are executed |
| Legacy packages | Pre-IWA `index.xml` and `index.apxl` packages are rejected as unsupported rather than guessed |

`IWorkReadOptions` supplies all configurable limits. Defaults cap a package and aggregate expanded entries at 512 MiB, one entry at 128 MiB, one compressed IWA archive at 64 MiB, one decompressed archive at 256 MiB, all decompressed IWA archives at 512 MiB, one Snappy chunk at 64 MiB, one record at 128 MiB, and the source-wide record count at 1,000,000. Semantic projection also bounds decoded text characters, text items and attribute boundaries, cross-record style inheritance, sparse materialized cells, tables, projected images, merged ranges, table dimensions, Numbers sheets, and Keynote slides. Word and PowerPoint owner preflight additionally bounds the aggregate destination table-cell area before allocating editable table grids.

## Editable semantic coverage

| Application | Reconstructed today | Preserved or diagnosed rather than reconstructed |
|---|---|---|
| Pages | Rich body paragraphs and runs, run hyperlinks, paragraph alignment and native Word list structure, including nested levels recoverable from the source indentation table, page size and margins, section-specific headers/footers, positioned and sized rich-text boxes with accessibility descriptions, embedded PNG/JPEG images, and editable tables with typed cells and merges | Ambiguous list-level metadata, inline objects represented only by unresolved replacement markers, drawable-level hyperlinks, exact floating-object order/wrapping/rotation, vector shapes, charts, equations, advanced table styling, masks/crops, comments, change tracking, fields, and application-only metadata |
| Numbers | Ordered sheets and tables, declared dimensions, sparse typed text/number/Boolean/date/duration cells, complete supported formulas plus typed cached values, merged ranges, header/footer metadata, default row/column sizing, and text-box text. Each table maps to its own worksheet so table-local formulas and column sizing remain stable; sheet-level text maps to a separate worksheet when present. | Pre-BNC cell storage, unsupported formula functions and cross-table references, per-cell rich text and exact styling, filters, names, charts, forms, comments, media, and application-only metadata |
| Keynote | Slide order, size and names, skipped state, positioned rich title/body text, typography, list labels and levels recoverable from source metadata, explicit inline line breaks, shape/run and presenter-note hyperlinks, presenter notes, embedded PNG/JPEG images, and positioned and rotated editable tables with typed cells and merges | Ambiguous list-level metadata, master/layout recreation, exact themes, vector shapes, charts, builds, transitions, animations, masks/crops, comments, and application-only metadata |

Editable reconstruction means the supported content is represented as normal DOCX, XLSX, or PPTX objects and can be edited and saved through its owner. It does not mean the destination is visually identical or that unsupported iWork records are written into the Office package.

## Visual fallback

`IWorkImportMode.Auto` uses editable reconstruction when supported semantics exist and otherwise uses an embedded raster preview. `EditableOnly` rejects sources without supported editable structure. `VisualOnly` always requests the raster preview.

Every owner report exposes `IWorkProjectionKind.EditableReconstruction` or `IWorkProjectionKind.VisualFallback`. `IWorkPreviewAsset.Coverage` distinguishes a known full-document asset from a first-page or composite preview. Current owner adapters embed PNG or JPEG previews; embedded PDF previews remain available on the source model but are not silently rasterized.

## Corpus evidence

The checked-in interoperability corpus uses unmodified fixtures with recorded source revisions and reproduced MIT or 0BSD notices:

| Application | Producer/build history exercised |
|---|---|
| Pages | 14.1 and 14.5, including a 14.4.1 package history with images and tables |
| Numbers | 11.1, build histories spanning 13.x and 14.x, 14.5, and 15.1, including formulas and merged ranges |
| Keynote | 8.1, 14.5, and 15.2.1, including independently maintained image and editable-table fixtures |

Tests assert path/stream parity, cumulative decompression and materialization bounds, application detection, source-record retention, rich text and drawable geometry, bounded style inheritance, section-specific headers/footers, formulas and cached values, merges and tables across all three applications, embedded images, explicit handling of pre-BNC cell storage, strict preview selection, destination limits, explicit visual fallback, and save/reopen of the resulting DOCX, XLSX, and PPTX packages. Fixture sources, revisions, expected content, checksums, and licenses are recorded in `OfficeIMO.TestAssets/Documents/IWorkCorpus/README.md`.

This corpus proves the current read contract; it does not establish a stable iWork write contract.

## Intentional authoring boundary

OfficeIMO does not create, edit in place, or write `.pages`, `.numbers`, or `.key` files. Authoring remains deferred until a broader independently produced corpus can prove deterministic package reconstruction, unknown-record retention, reopen behavior in multiple iWork versions, and stable semantic and visual round trips. Save the editable projections as DOCX, XLSX, or PPTX instead.
