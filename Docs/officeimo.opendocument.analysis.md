# Native OpenDocument Support for OfficeIMO

## Decision

Add one dependency-free package named `OfficeIMO.OpenDocument` that owns the OpenDocument family:

- `OdtDocument` for text documents
- `OdsDocument` for spreadsheets
- `OdpPresentation` for presentations

Keep all three formats in one assembly because ODF gives them the same package, manifest, metadata, style, drawing, value, and XML infrastructure. Split the code by format inside the project, not into three packages that either duplicate the shared engine or depend on a fourth package.

The native package should:

- use only target-framework APIs such as `System.IO.Compression`, `System.Xml`, and `System.Xml.Linq`
- have no `PackageReference` and no `ProjectReference`
- create, inspect, edit, and save ODF without LibreOffice, UNO, Microsoft Office, or a conversion service at runtime
- preserve package entries and XML it does not understand
- expose typed APIs over the retained XML rather than deserialize a file into a lossy set of POCOs
- read ODF 1.2 through 1.4, including extended documents, and write conforming ODF 1.4 by default with an ODF 1.3 compatibility profile

Reader and conversion support should remain in thin packages:

- `OfficeIMO.Reader.OpenDocument`
- `OfficeIMO.Word.OpenDocument`
- `OfficeIMO.Excel.OpenDocument`
- `OfficeIMO.PowerPoint.OpenDocument`

Those adapters can depend on the document packages they connect. The native `OfficeIMO.OpenDocument` package should not.

## Why This Fits OfficeIMO

The repository already contains the pieces of the right pattern:

- `OfficeIMO.Rtf` owns a dependency-free syntax layer, semantic document model, deterministic writer, diagnostics, and preservation-aware editing.
- `OfficeIMO.Word.Rtf` and `OfficeIMO.Rtf.Pdf` keep conversion outside the native RTF engine.
- `OfficeIMO.Pdf` owns a native reader/writer rather than shelling out to another application.
- `OfficeIMO.Zip` owns safe archive traversal rules.
- `OfficeIMO.Epub` shows a small ZIP/XML format owner with a separate `OfficeIMO.Reader.Epub` adapter.
- `OfficeIMO.Reader` has a modular registration contract for formats that should not become dependencies of the reader facade.

Word, Excel, and PowerPoint are useful API references, but not implementation bases for ODF. Their object models wrap the Open XML SDK and OOXML-specific concepts. Putting ODF serialization inside those packages would make ODF a conversion feature, prevent native round-trip editing, and add three implementations of the same ODF package and style rules.

## What “No Dependencies” Means

The production `OfficeIMO.OpenDocument.csproj` should contain neither package nor project references. It can use the internal source linked from `OfficeIMO.Shared` by `Directory.Build.props`, just as other projects do, because that source is compiled into the package rather than shipped as a NuGet dependency.

Archive safety is reusable behavior. The clean implementation is to move the reusable path normalization, entry limits, compression-ratio checks, duplicate-entry checks, and bounded copy helpers behind internal types in `OfficeIMO.Shared/Packaging`. `OfficeIMO.Zip` can delegate to the same source. This keeps one implementation without making OpenDocument depend on `OfficeIMO.Zip` at runtime.

Do not reference `OfficeIMO.Drawing` from the native package. ODF needs format-native lexical types such as `OdfColor`, `OdfLength`, `OdfAngle`, and `OdfPercentage` that can preserve the exact source spelling. Conversion adapters can map those types to `OfficeIMO.Drawing` where required.

Development and CI may use external validators and office applications as test oracles. They are not runtime dependencies.

## Format Constraints That Shape the Design

### Package rules are stricter than an ordinary ZIP

An ODF package is a ZIP containing `META-INF/manifest.xml`. When a `mimetype` entry is present, it must be the first entry, stored without compression, and have no ZIP extra field. Its ASCII content must match the media type on the root manifest entry.

`ZipArchive` is a viable starting point when `mimetype` is created first with `CompressionLevel.NoCompression`. A local .NET 8 probe produced a first entry at offset zero, stored, with no extended local header and a zero-byte extra field. The same byte-level assertion still needs to gate every supported target, but there is no evidence that justifies a custom ZIP writer.

The custom work should be an `OdfPackageWriter` that enforces entry order and manifest policy while using `ZipArchive` for ZIP encoding.

### The three formats share most infrastructure

Packaged documents normally contain:

- `content.xml`
- `styles.xml`
- `meta.xml`
- `settings.xml`
- `META-INF/manifest.xml`
- images, thumbnails, embedded objects, RDF metadata, scripts, and vendor files

ODT, ODS, and ODP differ mainly in the body under `office:text`, `office:spreadsheet`, or `office:presentation`. They share namespaces, styles, data styles, master pages, drawing objects, metadata, links, values, and packaging.

### Styles are a graph, not a flat property bag

ODF has common, automatic, default, and master styles. Styles can inherit through `style:parent-style-name`, and content may refer to styles from both `content.xml` and `styles.xml`.

The engine needs a central `OdfStyleRepository` indexed by style family and name. Effective-style lookup should resolve inheritance without flattening the stored XML. It must detect missing parents and cycles and return diagnostics instead of recursing indefinitely.

### Real documents contain extensions

A basic file saved by LibreOfficeDev 26.8 as ODF 1.4 contained standard namespaces plus `loext`, `calcext`, `officeooo`, `tableooo`, `drawooo`, and Microsoft-interoperability extension namespaces. Dropping foreign elements or attributes would damage ordinary files, not only exotic ones.

The default open/edit/save contract must therefore preserve foreign namespaces and unknown package entries. Strict conformance is a validation or save profile, not the default read behavior.

### ODS is sparse and run-length encoded

LibreOffice represented unused columns in a small sample with `table:number-columns-repeated="16381"`. Materializing repeated rows or cells into individual objects would turn small files into large allocations and create a denial-of-service path.

The spreadsheet model must keep repeated rows, columns, and cells as runs. Accessing or changing one cell should split only the affected run. Enumeration needs explicit used-range or budget semantics; it must not expand an entire theoretical sheet by default.

### Formulas are a separate language

ODF formulas use OpenFormula expressions and namespace-qualified formula values. The storage layer preserves formulas and cached results, exposes typed formula text, and invalidates stale cached values when a formula changes.

The evaluator is a separate, bounded parser/evaluation layer inside the OpenDocument owner. It covers a documented local OpenFormula subset without translating formulas through Excel, fetching external data, or executing active content.

## Recommended Project Layout

```text
OfficeIMO.OpenDocument/
  Core/
    OdfDocument.cs
    OdfDocumentKind.cs
    OdfOpenOptions.cs
    OdfSaveOptions.cs
    OdfVersion.cs
  Packaging/
    OdfPackage.cs
    OdfPackageEntry.cs
    OdfPackageReader.cs
    OdfPackageWriter.cs
    OdfManifest.cs
    OdfBackingStore.cs
  Xml/
    OdfNamespaces.cs
    OdfXmlPart.cs
    OdfXmlReader.cs
    OdfXmlWriter.cs
  Styles/
    OdfStyleRepository.cs
    OdfStyle.cs
    OdfResolvedStyle.cs
    OdfDataStyle.cs
  Drawing/
    OdfFrame.cs
    OdfImage.cs
    OdfRect.cs
    OdfShape.cs
    OdfTransform.cs
  Values/
    OdfColor.cs
    OdfLength.cs
    OdfPercentage.cs
    OdfCellValue.cs
    OdfFormula.cs
  Text/
    OdtDocument.cs
    OdtParagraph.cs
    OdtSpan.cs
    OdtList.cs
    OdtTable.cs
    OdtSection.cs
  Spreadsheet/
    OdsDocument.cs
    OdsSheet.cs
    OdsRowRun.cs
    OdsCellRun.cs
    OdsCell.cs
    OdsRange.cs
  Presentation/
    OdpPresentation.cs
    OdpSlide.cs
    OdpMasterPage.cs
    OdpTextBox.cs
    OdpNotes.cs
  Diagnostics/
    OdfDiagnostic.cs
    OdfFeatureReport.cs
  Validation/
    OdfValidator.cs
    OdfValidationResult.cs
```

Use partial classes only when one public type remains cohesive but has clearly separate responsibilities. Keep package IO, XML policy, styles, and format-specific semantics in dedicated types rather than growing one `OdfDocument` catch-all.

## Internal Representation

### Package store with a dirty overlay

Opening a file should create a package store containing:

- the original package bytes in a memory or temporary-file backing store
- an ordered entry index with exact, case-sensitive names
- lazily loaded XML documents for known XML parts
- a dirty flag per entry
- added and removed entry overlays

`OdfOpenOptions` should choose `Memory`, `TemporaryFile`, or `Auto` backing. `Auto` can keep small packages in memory and spill larger packages to a temporary file. The document owns and disposes this backing store so input streams do not need to stay open.

Saving should create a new ZIP:

1. Write a canonical `mimetype` entry first.
2. Copy untouched entries from the backing store without parsing or normalizing their payloads.
3. Serialize only dirty XML parts.
4. Copy `META-INF/manifest.xml` unchanged when the entry graph and media types are unchanged; otherwise rebuild it from retained manifest metadata plus actual entries.
5. Write to a separate stream or an adjacent temporary file, then replace the target atomically where the platform permits.

Do not update a source archive in place.

### XML-backed typed wrappers

Known XML parts should be loaded with safe reader settings and retained as `XDocument` trees. Public wrappers such as `OdtParagraph`, `OdsCell`, and `OdpSlide` should reference their owning `XElement` nodes.

This gives the API the same useful property as the Open XML wrappers: a typed operation changes a small part of the native tree. Unknown attributes, child elements, namespace declarations, and sibling order remain present unless the caller explicitly replaces their owner.

A detached semantic model that rewrites all XML from POCOs should be avoided. It would make extension preservation, style identity, whitespace semantics, embedded content, and future feature support much harder.

### Preservation contract

The public contract should distinguish three outcomes:

- **Preserved:** unsupported content was retained unchanged.
- **Editable:** content is understood and can be changed through a typed API.
- **Normalized or dropped:** content cannot be preserved across the requested edit; saving requires explicit consent and reports a diagnostic.

For an ordinary open/save with no edits:

- untouched entry payloads should remain byte-identical
- unknown entries should remain present
- unknown XML should remain present
- entry compression bytes, timestamps, and ZIP central-directory layout are not promised to be byte-identical

When one XML part is edited, preservation is semantic within that part rather than byte-for-byte. The save result should report which entries were rewritten, added, removed, or copied unchanged.

## Public API Direction

The API should feel familiar to current OfficeIMO users without pretending ODF and OOXML have identical semantics. The following is a design target, not a frozen signature set:

```csharp
using OfficeIMO.OpenDocument;

using OdtDocument document = OdtDocument.Create();
document.Metadata.Title = "Quarterly report";
document.AddHeading("Summary", level: 1);
OdtParagraph paragraph = document.AddParagraph("Revenue increased by ");
paragraph.AddText("12%.").Bold();
document.Save("report.odt");
```

```csharp
using OdsDocument workbook = OdsDocument.Create();
OdsSheet sheet = workbook.AddSheet("Summary");
sheet.Cell("A1").SetValue("Region");
sheet.Cell("B1").SetValue("Revenue");
sheet.Cell("B2").SetValue(125000m);
sheet.Cell("B3").SetFormula("of:=SUM([.B2])");
workbook.Save("report.ods");
```

```csharp
using OdpPresentation presentation = OdpPresentation.Create();
OdpSlide slide = presentation.AddSlide();
slide.AddTextBox("Quarterly report", OdfRect.FromCentimeters(2, 2, 20, 2));
slide.AddImage("chart.png", OdfRect.FromCentimeters(2, 5, 20, 10));
presentation.Save("report.odp");
```

All three roots should support:

- `Create()`
- `Open(string|Stream, OdfOpenOptions?)`
- `Save(string|Stream, OdfSaveOptions?)`
- `SaveAsync(...)`
- `ToBytes(...)`
- `Validate(...)`
- `InspectFeatures()`
- explicit `Dispose()` ownership for backing storage

Use `Open` consistently across this new package rather than copying the current Word/Excel `Load` and PowerPoint `Open` difference.

## Common Engine Responsibilities

### Packaging

- validate media type against document kind and extension
- require and parse `META-INF/manifest.xml`
- enforce entry and uncompressed-byte budgets
- reject path traversal, absolute paths, NULs, unsafe duplicates, and malformed ZIPs
- retain images, RDF, thumbnails, scripts, embedded documents, and unknown files
- manage part media types and manifest entries when content is added or removed

### XML

- prohibit DTD processing and external entity resolution
- apply document-character and depth budgets
- use namespace URIs, never source prefixes, for semantic lookup
- retain original prefixes where practical when editing existing XML
- centralize namespace and qualified-name constants
- keep ODF whitespace rules in one text reader/writer rather than relying on `XElement.Value`

### Styles and values

- index styles by family and name
- distinguish common, automatic, default, and master styles
- resolve parent chains with cycle diagnostics
- generate deterministic, collision-free automatic style names
- preserve unrecognized properties
- parse lengths, colors, dates, durations, percentages, and angles without unnecessary `double` round trips

### Diagnostics

Diagnostics should have stable IDs and include the package part and XML location when known. They should cover malformed input, unsupported-but-preserved features, lossy edits, version mismatches, extension content, stale formula results, signatures, encryption, and conformance problems.

## Format Slices

### ODT first useful slice

Implement:

- paragraphs, headings, spans, tabs, spaces, and line breaks
- named and automatic text/paragraph styles
- ordered and unordered lists
- tables with row/column spans
- hyperlinks and bookmarks
- sections, page layouts, margins, and page breaks
- headers and footers through master pages
- inline and anchored images
- metadata and custom metadata
- read-only inspection of annotations and tracked changes before editing them

Defer initially:

- tracked-change authoring and merge logic
- bibliography databases and complex indexes
- forms, scripts, macros, mail merge, and embedded databases
- formula editing inside text documents
- encryption and signature creation

### ODS first useful slice

Implement:

- sheets and sheet order
- sparse rows, columns, and cells with repeat-run splitting
- string, number, decimal, boolean, date, time, duration, percentage, and currency values
- formulas as preserved OpenFormula text plus cached values
- cell, row, column, and table styles
- number/date/time formats
- merged cells
- row heights, column widths, hidden rows/columns, and basic print settings
- named ranges, hyperlinks, and basic validation rules
- deterministic used-range enumeration

Defer initially:

- formula evaluation
- pivot tables and data pilots
- external data refresh
- macros and scripts
- complete chart editing
- every conditional-formatting extension

The sparse run model is a foundation requirement, not a later optimization.

### ODP first useful slice

Implement:

- presentation page size
- slides, ordering, naming, and visibility
- master pages and presentation layouts
- text boxes, paragraphs, runs, and lists
- rectangles, ellipses, lines, groups, and common transforms
- images and basic cropping
- tables
- speaker notes
- basic backgrounds and transitions that have direct ODF representations

Defer initially:

- animation timing trees
- audio and video authoring
- 3D scenes
- complete chart editing
- vendor-specific theme round trips beyond preservation
- arbitrary enhanced geometry editing

ODP should reuse the common ODF drawing and style layers rather than copy shape behavior from `OfficeIMO.PowerPoint`.

## Version and Interoperability Policy

Recommended profiles:

- `OdfCompatibilityProfile.Odf14` — default for new documents; standard ODF 1.4 only
- `OdfCompatibilityProfile.Odf13` — output intended for older LibreOffice and Microsoft Office versions
- `OdfCompatibilityProfile.PreserveSource` — retain the opened document version and extensions where possible
- a future explicit `LibreOfficeExtended` profile if OfficeIMO starts creating LibreOffice extension markup

Read ODF 1.2, 1.3, and 1.4 in both conforming and extended forms. Older ODF 1.0/1.1 can be accepted in a tolerant, diagnostic mode later, but should not expand the first implementation.

Do not create vendor-extension markup in the default profile. Preserve existing vendor markup unless a caller requests strict normalization.

Flat XML variants (`.fodt`, `.fods`, `.fodp`) use the same schema model. They are supported through explicit flat-XML open/save APIs, including embedded raster image binary data, while package-only and exotic embedded-object features remain a documented lossy boundary.

## Signatures, Encryption, and Active Content

### Signatures

An unchanged signed document can be copied while retaining signature-related entries. A mutation can invalidate signatures. The default save behavior for a changed signed document should fail with a clear diagnostic rather than silently retain an invalid signature. An explicit save option may remove invalidated signatures.

Signature creation and cryptographic verification can be a later capability. It should not force a production dependency into the first package.

### Encryption

Encrypted entries are described in the ODF manifest and include algorithm compatibility requirements that are larger than the core document work. The first release should detect encryption early and return a specific unsupported-encryption result. It must not produce partially decrypted or partially rewritten packages.

Do not implement custom cryptography casually to satisfy a feature checklist.

### Active content and external resources

The library must never execute scripts, macros, formulas, event listeners, or embedded objects. It must not fetch `xlink` targets while opening a document. Active content and external links can be preserved and reported through feature inspection.

## Validation Strategy

Production validation should be dependency-free and cover the contracts OfficeIMO controls:

- ZIP structure and `mimetype` placement/compression/extra-field rules
- manifest completeness and media-type consistency
- required parts and correct body kind
- XML well-formedness under safe parser settings
- `office:version` and manifest-version consistency
- style references, parent cycles, master-page references, and duplicate names
- ODS repeat counts, merges, formulas, and value/type consistency
- relationship between package entries and referenced images or objects

Tests should add stronger external evidence without turning it into a runtime dependency:

1. Validate generated XML against the official pinned OASIS 1.3 and 1.4 Relax NG schemas in a Linux CI lane.
2. Open and resave generated fixtures with a pinned LibreOffice headless build, then inspect package and semantic results.
3. Keep small LibreOffice-authored and Microsoft Office-authored ODT/ODS/ODP fixtures with source/version notes.
4. Verify untouched-entry hashes and targeted-edit preservation across the corpus.
5. Convert representative outputs to PDF or images with LibreOffice for optional visual baselines.
6. Run focused Windows desktop canaries for current Microsoft Office when that environment is available.
7. Fuzz malformed ZIP paths, duplicate entries, XML entities, deep XML, oversized repeat counts, cyclic styles, invalid formulas, broken manifests, and truncated media.
8. Build the native package for `netstandard2.0`, `net8.0`, `net10.0`, and `net472` on Windows, matching the repository’s format packages.

The OASIS schema gate proves conformance. LibreOffice and Microsoft Office gates prove interoperability. Neither replaces preservation and public-API contract tests.

## Delivery Plan

Keep the implementation in reviewable vertical slices.

### Milestone 0: package kernel

- [x] Add `OfficeIMO.OpenDocument` and `OfficeIMO.OpenDocument.Tests` with no package or project references.
- [x] Extract shared internal archive-safety primitives and make `OfficeIMO.Zip` use them.
- [x] Implement package open/save, manifest handling, safe XML loading, dirty-part tracking, metadata, diagnostics, and validation.
- [x] Create minimal conforming ODT, ODS, and ODP packages that LibreOffice opens without repair.
- [x] Prove first-entry `mimetype` structure on every target framework lane.
- [x] Add unknown-entry and foreign-XML preservation tests.

### Milestone 1: useful ODT

- [x] Implement the ODT slice listed above through XML-backed wrappers.
- [x] Add LibreOffice- and Word-authored preservation fixtures.
- [x] Add a public README and examples only for supported behavior.
- [x] Add `OfficeIMO.Reader.OpenDocument` extraction for ODT.

### Milestone 2: useful ODS

- [x] Implement the sparse row/cell run model before general cell APIs.
- [x] Implement typed values, formulas with cached results, styles, merges, ranges, and sheet operations.
- [x] Add LibreOffice- and Excel-authored preservation fixtures, including extreme repeat counts.
- [x] Extend `OfficeIMO.Reader.OpenDocument` with sheet/table chunks.

### Milestone 3: useful ODP

- [x] Implement slides, masters, text, common drawing shapes, images, tables, and notes.
- [x] Add LibreOffice- and PowerPoint-authored preservation fixtures.
- [x] Extend `OfficeIMO.Reader.OpenDocument` with slide-aligned chunks and notes.

### Milestone 4: conversion adapters

- [x] Add `OfficeIMO.Word.OpenDocument` for explicit ODT/Word conversion.
- [x] Add `OfficeIMO.Excel.OpenDocument` for explicit ODS/Excel conversion.
- [x] Add `OfficeIMO.PowerPoint.OpenDocument` for explicit ODP/PowerPoint conversion.
- [x] Publish feature-mapping reports rather than promise silent full-fidelity conversion.

### Milestone 5: advanced capabilities

- [x] Add bounded formula parsing and evaluation after the storage and preservation contracts are stable.
- [x] Add tracked-change editing, advanced-chart preservation, basic animations, flat XML, encryption detection, and signature preservation as separate capability lines.
- [x] Measure package size and build time before deciding whether to split the dependency-free package.

The M5 measurement on Apple M4/macOS with .NET SDK 10.0.102 produced a 302,725-byte `OfficeIMO.OpenDocument` package with empty NuGet dependency groups. A clean Release build of its three non-Windows targets took 1.18 seconds. This does not justify splitting the native package; the format-specific source folders remain the simpler boundary. BenchmarkDotNet coverage lives in `OfficeIMO.OpenDocument.Benchmarks` for ODT open/enumeration, extreme sparse ODS writing, and bounded range-formula evaluation.

## Repository Integration Points

An implementation needs coordinated updates to:

- `OfficeIMO.sln`
- `Build/project.build.json`
- `.github/codeql/codeql-config.yml`
- `.github/workflows/dotnet-tests.yml`
- package README and top-level package documentation
- `Website/pipeline.json` and the website dependency page when the package is published
- `OfficeIMO.TestAssets` with provenance notes for external fixtures
- `OfficeIMO.Examples` for small supported scenarios
- `OfficeIMO.Reader` capability documentation after the adapter exists

Generated command or API documentation should be updated through its generator rather than edited directly.

## Rejected Shapes

### Three independent native packages

`OfficeIMO.Odt`, `OfficeIMO.Ods`, and `OfficeIMO.Odp` would either copy packaging/styles/drawing code or require a fourth common dependency. Both outcomes conflict with the dependency-free requirement and create more versioning work.

### A common package plus three leaf packages

`OfficeIMO.OpenDocument.Core` plus three public leaf packages is architecturally tidy but makes every format package depend on the core package. It solves an organizational problem the source folders already solve. Revisit this only if one assembly becomes measurably burdensome.

### ODF backends inside Word, Excel, and PowerPoint

This would make native ODF editing depend on the Open XML SDK packages, duplicate ODF infrastructure three times, and encourage a lowest-common-denominator model. Conversion belongs in adapters.

### Generated classes for the complete ODF schema

The schema is broad and heavily reused across document kinds. Generating a class for every element would produce a large low-level API without solving preservation, style resolution, sparse sheets, or fluent authoring. Use a namespace-aware XML tree plus focused wrappers for supported semantics.

### A universal OfficeIMO document model

ODT, ODS, ODP, DOCX, XLSX, and PPTX do not share enough editing semantics to justify a new canonical model as a prerequisite. Direct, thin converters can introduce small mapping plans where useful. Promote a shared intermediate model only after multiple real converters prove the same abstraction.

### A custom ZIP implementation

The BCL already produces the required `mimetype` entry shape. A custom ZIP writer would add binary-format and security risk without a demonstrated need.

## Definition of Done for the First Stable Release

The native package is ready for a stable release when:

1. its NuGet dependency group is empty
2. new ODT, ODS, and ODP files open in current LibreOffice and Microsoft Office without repair prompts
3. supported typed edits survive reopen in OfficeIMO and LibreOffice
4. unsupported package entries and extension XML in the compatibility corpus are preserved by default
5. ODS repeat runs remain bounded under adversarial counts
6. generated standard-profile XML passes the pinned OASIS schema gate
7. save diagnostics identify every rewritten, removed, or lossy part
8. encrypted, signed, active-content, and external-link boundaries are explicit
9. the public feature matrix distinguishes editable, preserved, inspected, and unsupported behavior
10. all target frameworks build, and the focused test suite passes on Windows and Linux/macOS lanes

## Source Basis

- [OpenDocument 1.4, Part 1: Introduction](https://docs.oasis-open.org/office/OpenDocument/v1.4/os/part1-introduction/OpenDocument-v1.4-os-part1-introduction.html)
- [OpenDocument 1.4, Part 2: Packages](https://docs.oasis-open.org/office/OpenDocument/v1.4/os/part2-packages/OpenDocument-v1.4-os-part2-packages.html)
- [OpenDocument 1.4, Part 3: Schema](https://docs.oasis-open.org/office/OpenDocument/v1.4/os/part3-schema/OpenDocument-v1.4-os-part3-schema.html)
- [OpenDocument 1.4, Part 4: OpenFormula](https://docs.oasis-open.org/office/OpenDocument/v1.4/os/part4-formula/OpenDocument-v1.4-os-part4-formula.html)
- [LibreOffice OpenDocument file structure and version support](https://help.libreoffice.org/latest/en-GB/text/shared/00/00000021.html)
- [Microsoft Office implementation information for ODF 1.3](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oodf13/cef24f17-3e5e-4a13-9e16-aa1ebff5e1dc)
- [Microsoft Office application and ODF version support](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offdi/80d0a9bc-6d23-4cfa-84ca-20886a5c94e8)
- [Microsoft’s ODT/Word feature comparison](https://support.microsoft.com/en-us/word/differences-between-the-opendocument-text-odt-format-and-the-word-docx-format)
- [Microsoft’s ODP/PowerPoint feature comparison](https://support.microsoft.com/en-us/powerpoint/use-powerpoint-to-save-or-open-a-presentation-in-the-opendocument-presentation-odp-format)
