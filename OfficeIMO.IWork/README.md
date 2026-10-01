# OfficeIMO.IWork - bounded Apple iWork readers for .NET

`OfficeIMO.IWork` reads modern Apple Pages, Numbers, and Keynote packages without running iWork or executing embedded content. It owns ZIP, directory-bundle, nested `Index.zip`, Snappy-framed IWA, protobuf-envelope, package-resource, and source-record preservation. Word, Excel, and PowerPoint remain the owners of editable destination documents.

## Reference from a source checkout

For source-based development, reference the bounded reader and the opt-in adapter you need:

```xml
<ItemGroup>
  <ProjectReference Include="../OfficeIMO.IWork/OfficeIMO.IWork.csproj" />
  <ProjectReference Include="../OfficeIMO.Excel.IWork/OfficeIMO.Excel.IWork.csproj" />
</ItemGroup>
```

Use `OfficeIMO.Word.IWork` for Pages or `OfficeIMO.PowerPoint.IWork` for Keynote in place of the Excel adapter. Keep all project references on the same checkout so their coordinated API and package contracts stay aligned.

## Read and inspect a source

```csharp
using OfficeIMO.IWork;

IWorkSourceDocument source = IWorkSourceDocument.Open("report.pages");
IWorkPagesProjection pages = source.ReadPages();

Console.WriteLine(source.Kind);                 // Pages
Console.WriteLine(source.ContainerKind);        // ZipPackage, DirectoryBundle, or nested Index.zip
Console.WriteLine(string.Join(", ", source.BuildVersions));
Console.WriteLine(pages.Paragraphs.Count);

IWorkConversionReport report = pages.CreateConversionReport(
    IWorkProjectionKind.EditableReconstruction);
foreach (IWorkArchiveRecord record in report.PreservedRecords) {
    Console.WriteLine($"{record.EntryPath}: {record.MessageType}");
}
foreach (IWorkSourceUnitCount count in report.SourceUnitCounts) {
    Console.WriteLine($"{count.Kind}: {count.ReconstructedCount} reconstructed, {count.OmittedCount} omitted, {count.UnassessedCount} unassessed");
}
foreach (IWorkSourceReferenceIssue issue in report.SourceReferenceIssues) {
    Console.WriteLine($"{issue.Owner.RecordIdentifier}/{issue.FieldPath}[{issue.ReferenceIndex}]: {issue.Kind}, target {issue.TargetIdentifier}");
}
foreach (IWorkSourceDeclarationIssue issue in report.SourceDeclarationIssues) {
    Console.WriteLine($"{issue.Owner.RecordIdentifier}/{issue.FieldPath}: {issue.Kind}, outer values {issue.DeclaredValueCount}");
}
```

Pages body runs expose qualified drawable attachments through `IWorkTextRun.InlineObject`. The attachment keeps its original zero-based UTF-16 `CharacterOffset`, native attachment identity, and referenced drawable identity. Marker-only runs have empty `Text`; use `InlineObject` to distinguish them from ordinary text. Supported zero-offset image and standalone-table placement is described in the [support matrix](../Docs/officeimo.iwork-support-matrix.md#pages-inline-attachments).

Selected modern cell styles expose `IWorkTableCell.Fill`, including supported solid colors and explicit no-fill overrides. Selected styles also expose `Padding` in points and `VerticalAlignment`. A declared empty padding message resets all four sides to zero. Empty cells carrying supported fills, padding or alignment remain in `Cells` with `Kind == Empty` and count toward the materialized-cell budget. Filter by `Kind` when counting value cells. Missing or unsupported selected fills retain source diagnostics; conditional fills, borders and complete cell styling remain outside this qualified subset.

Shared tables expose explicit `RowHeights` and `ColumnWidths` as read-only maps keyed by one-based positions, measured in points. `GetRowHeight(row)` and `GetColumnWidth(column)` return an explicit size or the table default. Native zero-size entries use the default. `MaximumTableDimensionEntries` bounds dimension headers and bucket references across the source. Unsupported sizing records remain in source evidence and emit a diagnostic.

Positive native hidden/filtered row or column counts, and malformed count declarations, emit `IWORK_TABLE_VISIBILITY_UNASSESSED` and require explicit partial conversion for editable output. Partial output and Reader can include content hidden in the source. Counts do not identify which positions are hidden and can be zero even when native hidden-state records select content.

Selected base/summary hidden-state flags, collapsed-group and other unqualified extent declarations, active filters, and unreadable visibility envelopes emit `IWORK_TABLE_HIDDEN_STATES_UNASSESSED`. Empty extents, false selectors and disabled filter sets retain the existing editable path; disabled rules are not traversed. `MaximumTableDimensionEntries` also bounds hidden-state declarations and filter references. Hidden positions, filtering behavior and collapsed-group reconstruction remain unqualified.

Path and stream entry points use the same bounded parser. Stream and byte-array overloads detect the application kind from bounded package content. Pass an expected `IWorkDocumentKind` when the caller already knows the route and wants a mismatch rejected:

```csharp
using FileStream stream = File.OpenRead("budget.numbers");
IWorkSourceDocument source = IWorkSourceDocument.Open(
    stream,
    new IWorkReadOptions {
        MaximumPackageBytes = 64 * 1024 * 1024,
        MaximumArchiveReferenceCount = 1_000_000,
        MaximumMaterializedCells = 1_000_000,
        MaximumTableCatalogEntries = 100_000,
        MaximumImageMetadataEntries = 16_384,
        MaximumFormulaRenderingOperations = 64L * 1024 * 1024
    });

IWorkNumbersProjection workbook = source.ReadNumbers();
```

The verifying form is `IWorkSourceDocument.Open(stream, IWorkDocumentKind.Numbers, options)`.

`IWorkTable.FillStyles` exposes supported body, header-row, header-column and footer-row fills, plus `BandedBody` for every second body row after the headers. `table.GetFill(row, column)` resolves a one-based position without materializing unstored cells. Selected fills override region defaults, including explicit no-fill; unresolved selected fills stay null. Header and footer rows take precedence over header columns. Banding applies only to body cells. Unsupported or unresolved banding suppresses fill defaults and produces `IWORK_TABLE_FILL_DEFAULTS_UNSUPPORTED`; selected cell fills remain available. The [support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-support-matrix.md#table-fill-defaults-and-banding) defines the native Numbers export evidence and remaining appearance limits.

Table text formatting is separate from typed values. `IWorkTable.TextStyles` exposes applicable region defaults, `IWorkTableCell.ParagraphStyle` exposes a selected text style, and `table.GetParagraphStyle(row, column)` resolves the supported style for a one-based position, including an unstored empty cell. Selected styles override region defaults; an unresolved selected style stays null and produces `IWORK_TABLE_TEXT_STYLE_UNSUPPORTED`. The [support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-support-matrix.md#table-text-defaults-and-selected-styles) defines the qualified fields and limits.

## Cancellation

Path, stream, and byte-array `Open` overloads accept a `CancellationToken` after the read options. The token governs loading and all later semantic projections and conversions from that source. Once it is cancelled, reopen the source with a new token for another operation. Caller-owned streams remain open when loading succeeds or is cancelled.

```csharp
using var cancellation = new CancellationTokenSource();
IWorkSourceDocument source = IWorkSourceDocument.Open(
    "budget.numbers", IWorkDocumentKind.Numbers, options: null,
    cancellationToken: cancellation.Token);
IWorkNumbersProjection workbook = source.ReadNumbers();
```

The destination adapters also accept a token after their read and conversion options. Cancellation is cooperative during loading, projection, and destination construction; saving uses the destination owner's separate save API. These APIs do not establish a fixed cancellation latency or memory budget.

The [bounded NativeAOT contract](../Docs/officeimo.iwork-support-matrix.md#bounded-nativeaot-qualification) covers source projections and shared Reader extraction on macOS arm64 under .NET 8 and 10 using hash-pinned Pages, Numbers and Keynote fixtures. Destination conversion/save, rendering and Apple sandbox/device acceptance have separate qualification requirements.

## Opt in to an Office destination adapter

Install only the adapter for the destination format you need. The Word, Excel, and PowerPoint packages do not depend on iWork.

- [Pages to Word](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Word.IWork/README.md)
- [Numbers to Excel](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Excel.IWork/README.md)
- [Keynote to PowerPoint](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.PowerPoint.IWork/README.md)

Each adapter README owns its conversion API and example. The source reader remains useful on its own for inspection, extraction, and application-owned projection workflows.

This is extended semantic reconstruction rather than plain-text extraction:

- Pages recovers rich paragraphs, source-proven list levels mapped to native Word numbering, page layout, section-specific headers/footers, positioned and sized accessible rich-text boxes, images, rich-text table cells, and merges for editable Word projection.
- Numbers recovers sparse typed cells, rich-text cell runs, supported formulas with cached values, merges, table metadata, and default sizing for editable Excel projection. Each source table receives its own worksheet so table-local formulas and column sizing remain stable; sheet-level text receives a separate worksheet when present. Finite rectangular cross-table references resolve unique native identities among selected tables. Their source `Formula` uses quoted sheet/table labels, such as `=SUM('Sheet'::'Table'::$A$1:$B$2)`; the Excel adapter renders the retained expression with actual worksheet names. Table-local and cross-table whole-axis references, including header-named rows and columns, retain coordinate source expressions. The Excel adapter writes the current body extent and reports `IWORK_NUMBERS_TABLE_BODY_RANGE_APPROXIMATED` because named labels and automatic expansion are not preserved. Missing or ambiguous identities and ambiguous body metadata remain incomplete with valid caches retained.
- `IWorkNumbersSheet.Drawables` retains tables and text shapes in their shared source order. The separate `Tables` and `TextBoxes` lists remain available for type-specific access.
- Keynote recovers slide size, order and names, positioned rich text with explicit inline breaks and source-proven list labels and levels, shape/run and presenter-note hyperlinks, notes, images, rich-text table cells, positioned and rotated tables, and merges for editable PowerPoint projection.

Advanced charts, vector effects, animations, comments/change tracking, masks/crops, and other application-only structures remain available in source records. Reachable unsupported structures produce conversion diagnostics; preserved records alone do not establish a field-level fidelity assessment. Keynote measurements finer than PPTX's integral EMU grid stay editable and emit `IWORK_KEYNOTE_PPTX_PRECISION` when they are quantized to the nearest destination unit.

`IWorkReadOptions` bounds decoded text characters, text items and attribute boundaries, cross-record style inheritance, projected sheets/slides/tables/images, repeated encoded destination-image bytes, merged ranges, source-wide table catalogs, materialized cells, and ArchiveInfo references in addition to the package/IWA byte limits.

All conversion modes use the same bounded semantic source read, so package and projection limits are enforced before the destination representation is chosen. `Auto` prefers editable semantic reconstruction. `EditableOnly` fails when supported editable structure cannot be recovered. `VisualOnly` selects the package's raster preview for the destination and reports `VisualFallback`; it does not erase or bypass the semantic `ReadPages`, `ReadNumbers`, or `ReadKeynote` projection. Set `RequireCompleteVisualCoverage = true` to reject fallback assets without known full-document coverage. `AllowPartialEditableReconstruction = true` retains bounded recoverable objects and exposes `IsPartialEditableReconstruction`; source and destination limits still apply.

A preview may cover only the first page or a producer-generated composite, and that coverage is exposed on `IWorkPreviewAsset`. Embedded PDF inspection accepts bounded classic cross-reference tables and rejects unvalidated cross-reference streams.

## Preservation and authoring boundary

Projected documents, sheets, slides, tables, images, and text expose `SourceIdentity` with the native IWA identifier, message type, entry path, and payload position. Adapter reports retain these identities in `SourceUnits` and summarize them by kind in `SourceUnitCounts`, even when source payload preservation is disabled. The inventory counts identified units selected by the semantic projection: text storages are counted once, auxiliary table models and tiles are excluded, and inactive template records are excluded. Explicitly dropped selected units are reported as omitted. Selected unsupported drawables and Numbers sheet references with unsupported native types retain their identities under `UnsupportedObject` when no recognized kind applies. A preview leaves individual unit coverage unassessed. Unresolved references and content paths that the projection has not assessed remain outside this inventory. A reconstructed unit can still contain omitted, approximated, or unassessed fields.

`SourceReferenceIssues` separately records failed declared reference occurrences in assessed content paths: Pages body, stacking-order, decoded floating-canvas, shape text-storage, section, declared header/footer template and header/footer storage references; Numbers sheets, sheet drawables and shape text-storage references; and Keynote show, slide-tree nodes, slides, slide drawables, selected drawable text-storage and presenter-note references. Each issue retains the owning record, protobuf field path, one-based occurrence and readable target identifier. Missing targets, malformed references and readable siblings in a rejected set are distinct. Repeated occurrences within a source field remain separate; selecting a shared archive several times does not multiply its evidence. These occurrences do not identify distinct omitted objects, and preview output does not resolve them. Selected tables also assess table-info/model, tile, row/column sizing-bucket, string/formula/rich-text catalog, selected rich-text wrapper and wrapper/storage references. Tile paths retain physical entry positions; entries rejected before their target is read remain unassessed. Declared unresolved or malformed string/formula catalog links prevent complete editable reconstruction while valid cell caches remain recoverable through the explicit partial-conversion policy. Unreferenced templates, unused storage fallbacks and unparseable graph containers remain outside this assessment. `MaximumSourceReferenceIssues` defaults to 100,000 and bounds the cumulative occurrences inspected in distinct fields with failures, including readable siblings, before evidence is materialized. Adapter reports classify this evidence as `Unassessed` through `IWORK_SOURCE_REFERENCES_UNRESOLVED`.

`SourceDeclarationIssues` records unreadable, rejected or unsupported declarations separately from failed references and identified objects. Each item retains its owner, protobuf path, failure kind and known outer field-value count. A whole payload uses `$` and has no outer count; nested repeated entries use one-based positions. These counts do not reveal nested reference totals or omitted content. Coverage includes selected document roots, Pages body/shape/section/header/footer records and stacking/floating graph envelopes, Numbers sheet/plain-text record failures, Keynote show/tree/node/slide/drawable/note record failures, table-info/model/store/tile-list and row-header containers, selected string/formula/rich-text catalog envelopes and entry keys, plain-string values and formula-message envelopes, row/column sizing-bucket payloads and header metadata, selected merge-owner/store/pair/formula declarations, selected rich-text wrapper/storage records, decoded rich-text attribute containers and invalid entry positions, unreadable styles or style-parent envelopes actually traversed for text, and invalid values in supported character and paragraph properties, list label/type/indent vectors, style names and selected hyperlink text. Inactive records and unused rich-text value traversal are excluded. Other unsupported fields and remaining unmaterialized cell storage beyond the selected modern-row evidence remain coverage gaps. Shared paths report once per projection. `MaximumSourceDeclarationIssues` defaults to 100,000 and bounds distinct retained paths; evidence survives disabled source-payload preservation and visual fallback, with category `Unassessed` through `IWORK_SOURCE_DECLARATIONS_UNASSESSED`. Configured protobuf field, repeated-value and depth limits remain fatal; text-budget failures propagate rather than being reported as malformed declarations.

Modern row selection retains `InvalidValue` evidence at tile paths `5[n]/2` for a malformed, ambiguous or inconsistent declared cell count, and `5[n]/6` for a non-empty modern buffer with no selected offsets. A readable count is compared with selected physical cell records, including unformatted empty records that need no destination cell. An absent count remains supported. These paths describe outer field values, not missing-cell totals; no cell coordinates or formula presence are inferred from unselected bytes. Recoverable cells remain available through explicit partial conversion, while `IWORK_TABLE_ROW_STORAGE_UNASSESSED` reports `Unassessed` fidelity. Empty buffers with no selected offsets remain valid. Legacy companion buffers are not treated as unselected modern content when modern storage is present. Shared paths use the existing declaration budget and report once per projection.

`IWorkTableCell.UnsupportedFeatures` identifies selected modern conditional-style, applied-rule and comment fields after the complete cell storage has been decoded. These fields remain unassessed and are not reconstructed. `None` does not establish complete cell fidelity. Cells with these features, including otherwise empty cells, remain materialized under `MaximumMaterializedCells`. Their row buffer retains `UnsupportedField` declaration evidence at tile path `5[n]/6`, deduplicated under the existing declaration budget; its outer count is not a cell, rule or comment count. `IWORK_TABLE_CELL_FEATURES_UNASSESSED` reports the number of affected selected cells as `Unassessed` and requires the existing partial-conversion policy to retain editable values. Reader retains the warning. Truncated cells keep decode-failure evidence instead of assessed feature presence. Unused feature catalogs, referenced feature contents and remaining modern selectors stay outside this inventory.

Selected style-value failures use `InvalidValue` with their physical property path, such as `11/3` for font size, `11/7` or `11/26` for an unresolved foreground/background color value, `12/1` for alignment, `1/1` for a style name and `2` for hyperlink text. Counts describe that property’s outer values, including zero for missing required hyperlink text. Ambiguous Boolean values and clear-font/color selectors do not replace qualified inherited values. Rejected list vectors retain qualified parent metadata rather than removing invalid entries and shifting later levels. List evidence uses root fields `11` for label types, `13` for indents and `16` for labels; a packed label-type field counts as one outer value rather than its decoded element count.

Catalog keys are indexed before values are trusted. A malformed value with a readable, distinct key leaves healthy sibling values recoverable. Duplicate keys remain unresolved, including when a duplicate value is malformed; an unreadable entry or key leaves uniqueness unassessed and prevents resolving any catalog value. Formula cells retain valid caches when their expressions cannot be resolved. Catalog evidence uses `3[n]` for an unreadable physical entry, `3[n]/1` for key metadata, `3[n]/3` for invalid plain-string values (`InvalidValue`), and `3[n]/5` for rejected or unreadable formula messages. The catalog-entry budget includes unreadable entries. Rich-text value, wrapper, storage and style traversal stays limited to entries selected by decoded cells.

Sizing-bucket evidence uses `2[n]` for unreadable physical headers, `/1` for invalid or duplicate indexes, `/2` for invalid sizes, and `/3` for unsupported visibility values. `InvalidValue` covers size and visibility failures. An unknown index or unresolved bucket makes overrides on that axis untrusted, including overrides in other selected buckets; the other axis remains recoverable. Invalid values at readable, distinct indexes leave healthy sibling sizes available. Native zero continues to select the table default. Repeated bucket references retire their readable indexes while retaining unrelated overrides, and are reported at model path `4/1/2`; shared physical header paths report once. Explicit partial conversion retains recovered sizes using each destination owner's existing default and geometry rules.

Merged-range evidence uses model paths `47` and `47/2` for rejected or unreadable owner/store messages, `47/2/3[n]` for unreadable physical pairs, and `47/2/3[n]/2` for formula declarations with unsupported bounds or conflicts. Counts describe the outer field at that path. Conflicting rectangles are all retired; disjoint, valid merges remain recoverable through explicit partial conversion. A decoded out-of-bounds rectangle disqualifies only intersecting merges and is never exported as a clipped merge. Exact valid duplicates normalize to one merge; every physical alias retains evidence when that merge conflicts. An unreadable range can conceal any overlap, so cells remain recoverable without applying merges. Configured pair, syntax-node, protobuf and evidence limits remain fatal. Destination guards still protect populated covered cells.

Selected native table tiles retain unreadable payloads as `$`, rejected row sets as `5`, and unreadable or invalid rows as `5[n]` using their physical one-based positions. Tile-list selection failures retain model paths `4/3/1[n]`. Invalid row indices, modern buffer/offset metadata and unsupported legacy storage retain row-level evidence without inventing cell counts. Readable sibling rows survive malformed entries; the table remains incomplete and requires explicit partial reconstruction for editable output. Shared tile paths report once per projection. Rejected envelope/count bounds are checked before nested rows are parsed; configured parser limits remain fatal on content actually parsed.

Pages body, shape, header/footer, Keynote drawable/note and selected table-cell rich text also report failed references in decoded paragraph, list, character-style and hyperlink tables, plus parent references in styles traversed for that text. Each valid-offset decoded attribute entry is assessed, including end-of-text declarations; its path preserves the physical table-entry position. Unused style records are not traversed. Malformed attribute containers and entries with invalid offsets remain unassessed. Numbers shape text is decoded as plain text, so its formatting references remain outside this evidence. Table rich-text catalogs index all declared entries under the catalog budget, preserving structural and duplicate-key checks. Text, styles and wrapper/storage references are resolved when a decoded cell selects a catalog entry; aliases share storage decoding while each cell use consumes its text budget. Unused rich-text entries do not consume text budgets or add reference evidence. Unparseable table containers and unsupported table dependency fields remain unassessed.

`SourceCellIssues` identifies materialized selected cells whose storage or value could not be decoded, with table identity, one-based coordinates and the decoding message. `IWorkTableCell.HasDecodeError` distinguishes these failures from recovered native `#ERROR` markers; it does not assess every field of a successfully decoded cell. The inventory survives visual fallback and disabled record preservation, stays within the source-wide materialized-cell budget and reports `Unassessed` through `IWORK_SOURCE_CELLS_UNDECODED`. Undecoded formula cells also retain their separate formula/cache assessment. Unreadable rows, rejected storage and inactive tables do not establish cell coordinates and add no guessed cell counts. These source assessments do not describe destination cell retention.

Cross-table cell and rectangular-range references remain incomplete until their native table identities can be resolved. Their typed cached values stay available; dropping a table qualifier never produces a complete local expression. A merge qualifier must identify its owning table through the qualified four-word UUID representation; references to another table and ambiguous declarations retain source evidence instead of defining a local merge.

`FormulaCells` identifies source formula declarations by table identity and one-based coordinates. `ExpressionIsAssessed` distinguishes reconstructed or unresolved expressions from undecoded cell contents. Cache status is complete, partial, missing, approximate for a high-precision decimal or generic error marker, or `Unassessed` when the cell cannot be decoded. A supported version-5 header with the formula flag establishes `SourceFormulaIsDeclared`, even if later value decoding fails; the cell keeps its existing error output. Unsupported versions and unreadable headers do not establish formula presence. `FormulaSummary` includes separate unassessed expression and cache counts. These source assessments survive visual fallback and disabled record preservation; they do not establish destination formula retention or cache freshness. The inventory remains bounded by materialized-cell limits and excludes inactive tables.

Every package entry and every decoded IWA payload remains available as defensive bytes on `IWorkSourceDocument`. Import reports expose source payloads through `PreservedRecords` when `PreserveSourceRecords` is enabled. `PreservedRecordCount` counts source payloads regardless of that detail setting. `UnassessedRecordCount` includes consumed and auxiliary records whose field-level fidelity has not been assessed; it is not an omission count. `IWORK_RECORD_FIDELITY_UNASSESSED` uses `OfficeConversionLossKind.Unassessed`, so strict no-loss policies reject it without claiming those records are missing content. The destination DOCX, XLSX, or PPTX contains the supported reconstruction or visual fallback; it is not a lossless iWork package rewrite.

There is deliberately no Pages, Numbers, or Keynote writer. OfficeIMO will not expose iWork save-back until an independently produced corpus demonstrates a stable deterministic round-trip contract across supported producer versions.

See the [iWork support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-support-matrix.md) for the version corpus, limits, semantic coverage, and known boundaries.

`IWorkTableCell.NumberFormat` exposes supported modern number, percentage, currency, and scientific semantics: decimal places (`null` for automatic), digit grouping, and negative-value style. Scientific formats support the default minus style without grouping; decimal places describe the mantissa. Currency metadata also exposes the three-letter uppercase source `CurrencyCode` and `UseAccountingStyle`; the code does not imply a symbol or locale. It does not change `Value`, `DisplayText`, or `CachedDisplayText`. The Excel adapter applies this metadata to numeric XLSX cells; the Word and PowerPoint adapters apply supported invariant display text and red format color while reporting approximation for locale, automatic precision and source appearance. Cells whose display cannot be safely formatted retain raw text and report omission. Unsupported selected numeric formats retain source diagnostics and physical catalog-declaration evidence when a declaration can be identified.

Fraction formats expose `Kind == IWorkNumberFormatKind.Fraction` and `FractionAccuracy`: one-, two-, or three-digit denominators, or fixed halves, quarters, eighths, sixteenths, tenths, and hundredths. `DecimalPlaces` is `null` for fractions and does not mean automatic decimal formatting. The qualified subset uses minus negatives without grouping; other controls remain unsupported. The Excel adapter uses mixed-fraction formats and reports `IWORK_NUMBERS_FRACTION_DISPLAY_APPROXIMATED` for rounding, normalization, and spacing differences.

Finite Decimal128 values remain numeric even when their coefficient exceeds fifteen significant digits. Such cells set `NumericValueIsApproximate`, retain the exact normalized source value in `SourceNumberText`, and emit `IWORK_TABLE_NUMERIC_VALUE_APPROXIMATED` with `Approximation` fidelity. Formula caches receive the same metadata and approximate cache assessment. The retained text counts toward projection text limits. Nonfinite, noncanonical, overflowing, and nonzero underflowing values remain decode failures.

## Target frameworks and dependencies

`OfficeIMO.IWork` targets .NET Standard 2.0, .NET 8, .NET 10, and .NET Framework 4.7.2 on Windows. The source reader depends only on `OfficeIMO.Core`; its IWA, Snappy, protobuf-envelope, and package readers are first-party implementations. Destination projection is opt-in through `OfficeIMO.Word.IWork`, `OfficeIMO.Excel.IWork`, or `OfficeIMO.PowerPoint.IWork`.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 0 | 0 | 0 | 0 | 1 | 0 |
| Read | 1 | 0 | 0 | 0 | 0 | 0 |
| Edit | 0 | 0 | 0 | 0 | 1 | 0 |
| Preserve | 0 | 0 | 0 | 0 | 1 | 0 |
| Inspect | 1 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.IWork` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
