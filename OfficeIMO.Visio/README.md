# OfficeIMO.Visio - Visio diagrams for .NET

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Visio)](https://www.nuget.org/packages/OfficeIMO.Visio)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Visio?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Visio)

`OfficeIMO.Visio` creates, edits, inspects, validates, and exports `.vsdx`, `.vstx`, `.vssx`, `.vsdm`, `.vstm`, and `.vssm` packages without COM automation and without Microsoft Visio installed.

If OfficeIMO saves you time, please consider supporting the work through [GitHub Sponsors](https://github.com/sponsors/PrzemyslawKlys) or [PayPal](https://paypal.me/PrzemyslawKlys). PowerShell users should use [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice) for the PowerShell-facing experience.

## Install

```powershell
dotnet add package OfficeIMO.Visio
```

## Quick start

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;

var document = VisioDocument.Create("diagram.vsdx");
document.AsFluent()
    .Info(info => info.Title("Demo").Author("OfficeIMO"))
    .Page("Page-1", page => page
        .Title("Demo Flow")
        .Rect("start", 1, 1, 2, 1, "Start")
        .Diamond("decision", 4, 1.5, 2, 2, "Decision")
        .Ellipse("end", 7, 1.5, 2, 1, "End")
        .Connect("start", "decision", VisioSide.Right, VisioSide.Left,
            connector => connector.RightAngle().ArrowEnd(EndArrow.Triangle))
        .Connect("decision", "end", VisioSide.Right, VisioSide.Left,
            connector => connector.RightAngle().ArrowEnd(EndArrow.Triangle).Label("Yes")))
    .End();
document.Save();
```

For a drawing from an untrusted source, use the bounded load profile:

```csharp
var incoming = VisioDocument.Load("upload.vsdx", VisioLoadOptions.UntrustedDefaults);
```

It rejects macros, embedded payloads, ActiveX, and external relationships before parsing. Ordinary load options retain compatibility with drawings containing those parts; set `PackageSecurity` explicitly for another policy.

## Legacy binary Visio

`VisioDocument.LoadLegacyBinary` reads the bounded binary version 11 profile from
`.vsd` drawings, `.vss` stencils and `.vst` templates. It returns the existing
`VisioDocument` model with an `OfficeLegacyImportReport`; inspect that report before
accepting a conversion. Earlier binary generations fail explicitly.

```csharp
var imported = VisioDocument.LoadLegacyBinary("floorplan.vsd");
foreach (var finding in imported.Report.Findings) {
    Console.WriteLine($"{finding.Code}: {finding.Message}");
}

// After accepting the reported import losses:
imported.Value.Save("floorplan.vsdx");
string svg = imported.Value.Pages[0].ToSvg();
```

The path overload infers the family from its extension. For caller-owned streams,
pass `VisioPackageType.Stencil` or `VisioPackageType.Template` when appropriate.
Seekable streams are read from the beginning and their position is restored;
nonseekable streams are read from their current position. The caller's stream
remains open, including on failure. Binary input is never associated with `Save`:
choose a modern output path explicitly. The matching modern families are VSDX,
VSSX and VSTX.

The profile reconstructs cached transforms, group nesting, master identities and
references, basic line/fill and character styles, MoveTo/LineTo/ArcTo/Ellipse
geometry and cached string text fields. The existing SVG and
[PDF adapter](../OfficeIMO.Visio.Pdf/README.md) project imported pages; retain both
the import report and the subsequent projection report. To render a page-less
stencil, place a recovered master on a drawing page first.

For diagram PDF output, select the adapter's page projection explicitly:

```csharp
using OfficeIMO.Visio.Pdf;

var pdf = imported.Value.ToPdfDocumentResult(new VisioToPdfOptions {
    Mode = VisioPdfProjectionMode.DiagramPages
});
pdf.Save("floorplan.pdf");
```

Cached paint styles and transparency feed the shared model. Geometryless native
master children remain text-only in previews. Non-solid fill patterns use a
foreground-color approximation in managed SVG/PDF; the modern package retains
the cached pattern cells. Exact native appearance remains unqualified.

This is a lossy conversion. Native page/master/shape names use stable fallback
names. Formulas, recalculation, unresolved or numeric fields, custom data,
advanced styling, unsupported curves and foreign objects are omitted or
approximated with diagnostics. Active content is never executed. Original binary
carriers and binary save-back are not supported.

`VisioLegacyBinaryImportOptions` supplies independent source, expansion, record,
shape, text and depth budgets. Defaults are 64 MiB input, 128 MiB cumulative
decoded native streams, one million records/pointers, 250,000 shapes, four million
text characters, 512 compound streams and depth 64. Shared pointer tables use one
cached child list, while every reference consumes its logical subtree budget and
must satisfy the depth limit at that position. Exceeding a budget aborts import;
partial results are not returned.

The [independent fixture manifest](../OfficeIMO.Visio.Tests/Fixtures/LegacyBinary/producer-manifest.json)
pins Microsoft Visio version 11 binary files and paired XML drawings, stencils and
templates. Checks cover identities, groups, cached transforms and text, modern
package reopening, Reader dispatch and SVG/PDF output. Microsoft Visio application
acceptance, broader producers and earlier generations remain unqualified. No
external converter is a runtime dependency.

## Legacy Visio XML

`LoadLegacyXml` imports Visio 2002 and 2003 XML drawings (`.vdx`), stencils (`.vsx`), and templates (`.vtx`) into the existing editable model. `ToLegacyXmlResult` and `SaveLegacyXml` export those families with operation-level fidelity diagnostics.

```csharp
using OfficeIMO.Visio;

var imported = VisioDocument.LoadLegacyXml("workflow.vdx");
foreach (var diagnostic in imported.Report.FidelityDiagnostics)
    Console.WriteLine(diagnostic.Message);

VisioDocument document = imported.Value;
document.Pages[0].Shapes[0].Text = "Reviewed";
document.Save("workflow.vsdx");

var legacy = document.ToLegacyXmlResult();
foreach (var diagnostic in legacy.Report.FidelityDiagnostics)
    Console.WriteLine(diagnostic.Message);
// After accepting these specific omissions, write the inspected payload.
File.WriteAllBytes("workflow-reviewed.vdx", legacy.Value);
```

Create new legacy documents with `VisioDocument.Create()` or `Create(VisioPackageType.Stencil)` / `Create(VisioPackageType.Template)` and the same shapes, masters, and fluent builders. `SaveLegacyXml(path)` writes atomically and rejects omitted content by default; set `allowOmissions: true` only after reviewing an export report. Approximations remain reported. Even a basic newly authored document can contain modern theme or resize settings with no legacy equivalent.

| Contract | Legacy XML profile |
| --- | --- |
| Editable content | Pages, masters, shapes, nested groups, connectors, geometry rows, text, supported styles, Shape Data, User cells, hyperlinks and layers through the shared Visio model |
| Preservation | Supported ShapeSheet cells, formulas, row identities, rich-text XML and native fragments retained by the shared model; not a byte-preserving XML editor |
| Conversion | Legacy XML to/from the corresponding Open XML family; loaded diagrams can use the existing inspection, SVG/raster and PDF paths within their documented profiles |
| Loss reporting | Unmapped modern cells/sections, unsupported package-only content and unsupported legacy details are reported; arbitrary formula evaluation and native layout equivalence are outside the model |
| Safety | Input byte, XML character, nesting and element/attribute budgets; DTDs and XML entities prohibited; external links are not fetched; legacy VBA is rejected |
| Source lifecycle | Legacy import has no associated save path. Ordinary `Save` writes an Open XML package. Use `SaveLegacyXml` explicitly for XML output. Caller streams remain open. |

Embedded foreign bytes and their XML metadata survive XML/Open XML round trips, including nested shapes and masters. Resources follow moved and duplicated shapes, including moves between documents; deleted content is excluded from saved packages once its last reference is removed. The shared Visio model keeps their internal image/object relationships and content types. Opaque or unrecognized bytes are treated as embedded objects for the untrusted-input policy even when the source labels them as images. External foreign resources and resources with dependent parts are rejected. Preservation is bounded to 1,024 resources, 64 MiB per resource, and 128 MiB total.

SVG and raster export render supported embedded bitmap images, including master-owned images, crop offsets, group rotation and mirroring. Supported placement formulas evaluate against the current shape dimensions; unsupported formulas use cached values with a diagnostic. Decoding uses the shared Drawing codecs with a 64 MiB encoded limit and a 32-million-pixel decoded limit. Legacy BMP payloads with a zero file-size header receive that value in a decoding copy; saved bytes remain unchanged. `ExportImage` returns the decoding and placement diagnostics, and its policy can reject omissions. Opaque objects and unsupported images produce visible placeholders; opaque objects are never sent to caller image codecs. Metafiles require a caller codec when the shared decoder cannot read them. The default PDF projection remains semantic text and topology. Optional [PNG page previews](../OfficeIMO.Visio.Pdf/README.md#page-previews-and-legacy-xml) embed this supported raster profile, fit the PDF content area proportionally, and retain preview diagnostics and omission categories in the neutral model, Reader and PDF report. SVG previews are listed as metadata by PDF projection. These previews do not establish native diagram-layout fidelity.

Connectors can have zero, one, or two attached shapes. `From` and `To` are nullable; `StartPoint` and `EndPoint` resolve page coordinates in inches. Setting either point detaches that endpoint. Use `ReconnectConnectorStart` or `ReconnectConnectorEnd` to attach it again. Page creation accepts the page's default unit or an explicit unit:

```csharp
using OfficeIMO.Drawing;

var route = page.AddConnector("route", new OfficePoint(1, 2), new OfficePoint(6, 4),
    unit: VisioMeasurementUnit.Inches);
page.ReconnectConnectorEnd(route, targetShape, VisioSide.Left);
route.StartPoint = new OfficePoint(2, 2); // The start remains free.
```

Free points, loaded attachment positions, and explicit glue points survive XML/Open XML saves. Shape movement updates attached ends, including children of rotated groups. Page duplication includes free routes; selection duplication includes a partial connector when its attached shape is copied. Native local transforms keep saved path and label positions consistent with previews. Label placement uses page coordinates; `TextStyle.TextAngle` sets a connector label's page rotation in radians. Explicit waypoints remain fixed in page coordinates when an endpoint moves, including after reopening.

Imported single-path curves use the shared geometry parser in SVG/raster previews and retain their native geometry rows on save. Endpoint edits translate, rotate, and uniformly scale the preserved path. Replacing a route with `RouteThrough` selects explicit page waypoints. Unsupported scaled geometry and stretching a closed connector by separating its endpoints throw `NotSupportedException`; replace the route explicitly in those cases. Multi-path artwork, arbitrary ShapeSheet recalculation, and native automatic rerouting remain outside this editing profile. Missing free coordinates or unresolved attachment IDs remain preserved XML with a diagnostic.

Circular `ArcTo` previews use the signed bow to the right of the directed chord in Visio's upward Y coordinates. Positive and negative, reversed, diagonal and major-arc controls agree with two independent readers. Connector labels placed with `PlaceLabel` and preview arrowheads follow that corrected route. Native geometry values, formulas and row identities remain unchanged by rendering, reopening and copying; this arc evidence does not qualify broader stored text frames or exact Microsoft Visio arrow appearance; the [cached connector text-frame profile](#cached-connector-text-frame-placement) covers retained frames during endpoint editing.

Open XML saves write readable extended and custom property parts and valid page-window metadata. Packages without a generated thumbnail omit its part and relationship. These defaults do not provide preservation of arbitrary imported package properties or window layouts.

For stream imports, specify `packageType` for a stencil or template; the XML does not reliably distinguish every family. The Visio 2002 and 2003 core namespaces are accepted; XML export uses the 2003 namespace and reports normalization of older input. Legacy `FontEntry` names and charsets feed the shared text model while their original metadata remains preserved. Binary `.vsd`, `.vss`, and `.vst` files and namespace-less XML exports are outside this API.

Loaded stencil and template masters retain their page scale, catalog metadata, named ShapeSheet rows, formulas, and text whitespace. For XML masters, including those with supported embedded foreign resources, changes to modeled shape properties, text, custom rows, and the nested child collection are applied without replacing untouched native cells. Unchanged formulas within edited Character and Connection rows remain preserved; changed numeric cells replace their old formulas. Changing a master extent retains the model's local pin when the native cell was absent, so reopening does not shift its child tree. Deleted foreign content is excluded from the saved master once its last reference is removed. New string child identifiers receive numeric package IDs and retain their original API identifiers. Templates retain unused registered masters for later authoring. This does not evaluate dependent formulas.

For a single Character or Paragraph row supported by `TextStyle`, saving preserves the native row identity, cell order, units, metadata and formulas on untouched cells. Formatting edits retain the `cp` and `pp` references in unchanged labels, including nonzero row indices; changed values replace their old formulas. This applies to shapes, connectors and loaded masters through XML and Open XML saves. Clearing modeled formatting retains an empty row for the existing markers. Master instances receive independent row snapshots and remapped local shape references. Complex rows and multiple formatting runs remain source-preserved.

SVG and raster export render unchanged mixed character runs in shape and connector labels through the shared text layout engine. The cached-value profile includes run fonts, sizes, literal or palette colors, color transparency, bold and italic emphasis, underline and strikethrough patterns, and superscript/subscript placement. Native `pp` markers select independent paragraph rows, including empty paragraphs. Cached paragraph alignment, first/left/right indents, space before/after and absolute or relative line spacing use the same paragraph layout engine. Hanging indents remain inside the text frame; fitting reduces run sizes while keeping paragraph insets and spacing. Unsupported paragraph placement retains the supported character-run projection. Changes applied with `SetShapeSheetSection()` reach rendering; page duplication retains independent formatting rows. Replacing the plain `Text` or `Label` replaces its run boundaries. The [independent fixtures](../OfficeIMO.Visio.Tests/Fixtures/LegacyXml/README.md) qualify color and emphasis changes; the stored font switch also has direct XML and SVG checks, with an independent-reader discrepancy recorded there. The AngularJS paragraph fixture also qualifies mixed left/center alignment and an empty paragraph through independent-reader callbacks and source/reopen/render checks. Negative paragraph spacing and hanging text outside the frame, exact typography, tab/field semantics, vertical writing and remaining character metrics are not qualified.

Native text backgrounds distinguish “none” (`TextBkgnd` values `0` and `255`) from indexed palette colors. A transparent `BackgroundColor` disables the background, and `BackgroundTransparency` uses percentages while its native cache uses a fraction. Untouched background cells retain their caches, formulas and metadata; explicit assignments replace them. Copying into a document with a conflicting palette uses the resolved color. See the [upgrade guidance](../MIGRATION.md#visio-text-background-values) for older percentage caches.

Loaded instances inherit supported master labels and formatting without saving those caches as new local overrides. Existing local rows retain their native cells and formulas; an explicit formatting edit adds or changes the affected cells using the inherited row identity. Label edits, including clearing a label, become explicit overrides. Repeated package reopening retains that distinction.

`GetShapeSheetSections()` returns editable copies of source-preserved sections, including Character/Paragraph rows with current typed edits. Apply an edited or new section with `SetShapeSheetSection()`, or remove it with `RemoveShapeSheetSection()`. Installing a supported text row synchronizes `TextStyle`; removing it clears its local typed formatting. These operations retain the loaded child order and survive XML export and package reopening, including Character overrides on complex master text. Changing a cell formula clears its previous producer error; replacing a Shape Data value with `SetShapeData()` also clears the old formula and error. Untouched formulas and errors remain preserved.

Supported fields in a single Character/Paragraph row remain editable through `TextStyle` when that row also contains other native cells. Saving merges the edit into the existing row, retaining its identity, unknown cells and unmodeled style flags. Complex or unparseable native text sections retain ownership of their formatting; `TextStyle` does not add a conflicting second section. Use `SetShapeSheetSection()` to edit those native rows.

Legacy null-string markers, such as `V="null"`, remain distinct from literal cell values through XML and Open XML round trips. Scoped conversion metadata follows copied shapes, connectors, pages and master instances, including external master registration and stencil import. PageSheet metadata and additional native master roots retain their state during copying and import. Stylesheet import matches numeric style IDs and named rows by their native identity, merges distinct Section, Row and Cell entries, and keeps existing destination values and metadata. Added cells retain their guarded state under the destination row. Existing headers, sections and rows retain their attributes and omitted defaults; newly added entries retain the source attributes. Standard legacy `Err` tokens map to Open XML `E`; other error text is retained with a `VDX_CELL_ERROR` approximation diagnostic. User Prompt units remain preserved.

Restoration checks the current cached value, formula, unit and error against the imported snapshot. Edits before or after copying suppress stale markers. Explicit writes to User `Value`/`Prompt`, Shape Data string properties and `SetShapeData()`, or a detached ShapeSheet cell's `Value`/`SetCell()` replace the null marker even when the cached value stays the same. Install detached edits with `SetShapeSheetSection()`. Matching producer errors remain preserved; `SetShapeData()` clears the replaced value's old formula and error. Reads, unchanged section installation and copies retain native nulls. `VisioShapeDataSchema.ApplyTo(overwriteValues: false)` retains existing values and their formulas/errors while applying supplied field metadata.

Mechanical rebinding of local shape references and imported font IDs updates matching snapshots and retains producer errors. Explicit row-identity changes, same-value formula/unit assignments, and same-value indexer writes through the concrete `Data` dictionary remain outside this preservation profile. Use the typed value APIs when assigning an empty literal over an imported null.

Page and selection duplication remap local numeric `Sheet.<ID>!` formula references to the copied shapes and connectors, including nested shapes, text-formatting rows, User cells, Shape Data, PageSheet cells and retained source sections. Assigned native IDs remain stable when the page is later reordered or receives new shapes. Quoted text, explicit cross-page references and references to elements outside the copied graph retain their original targets.

`AddShape` creates independent editable child trees for grouped masters and preserves imported symbols' default text immediately. Authored single-shape masters with a `TextStyle` also supply independent formatting and text-frame values to their instances and saved master content. Children retain their master-shape links; editing one instance does not change its master or sibling instances. Creation scales child positions, local pins, connection points, and supported cached geometry to the requested size. Local numeric `Sheet.<id>!` references are remapped to stable page IDs. Omitting the label retains these masters' default text; assigning an empty or null `Text` to an instance explicitly clears it when saved. Imported font IDs are remapped when they collide with fonts already in the destination document.

```csharp
var stencil = VisioDocument.LoadLegacyXml("rules.vsx").Value;
var page = stencil.AddPage("Rules");
var master = stencil.GetMaster("Atom");
var instance = page.AddShape("rule-1", master, 3, 5,
    master.Shape.Width * 2, master.Shape.Height * 2);
instance.Children[0].Text = "Customer is eligible";
instance.Children[0].SetUserCell("Review", "approved");
var edited = stencil.ToLegacyXmlResult(); // Inspect edited.Report before saving.
```

Resizing during creation supports straight and relative geometry, ellipses, infinite lines, Bezier/spline rows, evaluable NURBS/polyline formulas, and circular/elliptical arcs. Unequal positive width and height scaling converts circular arcs to editable native `EllipticalArcTo` rows and updates the axis angle and ratio of absolute or Open XML relative elliptical arcs. Endpoints, control points, row identities and geometry flags remain native; the curve is not replaced by a polyline. This profile requires readable finite coordinates and a resulting principal-axis ratio of at most 1000. Invalid or unsupported scaled profiles throw `NotSupportedException` before insertion; creation at the original master dimensions preserves those rows.

Geometry expressions that produce the requested coordinates at the final instance placement remain formulas, including unchanged axes and normalized fractions. Other resized coordinates become cached instance values while the source master retains its original geometry and formulas. A converted bow cell becomes a control-point cell and clears its old formula and producer error. Relative rows retain their coordinate fractions; their native editable format is Open XML, since the legacy Visio 2003 XML schema has no relative-row mapping. Direct assignments to an instance's `Width` or `Height` do not automatically recalculate its child tree or dependent cells.

### Replace an existing shape or group master

`ReplaceMaster` changes the artwork of a page-owned shape or modeled group while retaining its existing shape and child objects, local placement, labels, cached text formatting, style, data, layers, hyperlinks and connector attachments. Every descendant binds to the corresponding child of the replacement master. Replacing a leaf with a group materializes new descendants while retaining the root object. Reordering live children retains their original master slots; an unbound standalone group uses child order. Generated descendants use their own linked blueprint's primitive, including before their geometry is serialized.

```csharp
var stencil = VisioDocument.LoadLegacyXml("rules.vsx").Value;
VisioPage page = document.Pages[0];
VisioShape symbol = page.FindShapeById("rule-1")!;
page.ReplaceMaster(symbol, stencil.GetMaster("Negative Atom"));
document.Save("updated.vsdx");
```

Omitting `resizeToMaster` keeps the existing frames. Setting it to `true` resizes the whole instance through the [existing resize contract](#resize-an-existing-shape-or-group). Referenced connection-point objects remain attached and scale with their shapes. Retained inherited text and effective cached text rows become local; replacing a master does not substitute its default label or formatting. Newly inherited text cells translate references from the old master to the corresponding page shapes; existing local formulas retain their page references. New native artwork is materialized at the instance size without editing the source master. Embedded-image artwork follows the replacement definition, including its image rectangle; obsolete local foreign payloads no longer override it.

Page-backed selections prepare all selected trees before changing any shape or registering a replacement. Selecting an ancestor and its descendant applies one tree update. Incompatible child trees, multiple-root definitions, shared live-instance blueprints, nonpositive root frames, unreadable scaled geometry and frame scaling that requires shear throw before applying changes. Package-backed stencil candidates retain their source relationships and font scope and are imported only after successful preparation. For a populated group, replacement preserves the modeled child tree; restructuring or removing its existing descendants and native Microsoft Visio open/edit/save qualification remain outside this profile. Cached text and geometry preservation does not evaluate arbitrary ShapeSheet formulas.

### Resize an existing shape or group

`ResizeShape` resizes an existing page-owned shape, including a nested shape, in its local axes. It retains the same shape, child, text-style and connection-point objects, pin, angle, labels, formatting and master links. Child positions, local pins, cached text frames, supported geometry and image placement follow the resize. Attached connector endpoints follow their shapes; free endpoints and explicit page-coordinate waypoints stay in place. Source masters and sibling instances retain their original content.

```csharp
VisioPage page = document.Pages[0];
VisioShape symbol = page.FindShapeById("review-symbol")!;
page.ResizeShape(symbol, symbol.Width * 2, symbol.Height * 0.5,
    VisioMeasurementUnit.Inches);
document.Save("resized.vsdx");
```

The operation prepares the complete tree and affected native connector paths before applying changes. Uniform resizing supports arbitrary rotations. Unequal width/height scaling supports child and text frames at multiples of 90 degrees; other angles require a shear and throw `NotSupportedException` without changing the tree. Native curves use the creation profile above: required coordinate cells and NURBS parameters must be readable, and incomplete local geometry rows reject resizing. NURBS knots, weights and degree retain their original evaluated values as coordinates scale. Inherited geometry becomes an editable instance override, and embedded images retain their payload and crop fractions. Physical font sizes, margins and line weights stay unchanged. Current and requested dimensions must be finite and positive; omitted units use the page's default unit.

The NxBRE grouped-master fixture qualifies label and User edits, attached endpoints and source-master preservation through XML/Open XML reopening. Controlled cases cover orthogonal text frames, inherited arcs, cropped images and atomic rejection. Arbitrary ShapeSheet recalculation, text-frame formula provenance, sheared frames, master replacement and full rich-text layout remain outside this resize profile. Independent reader/render evidence does not establish native Microsoft Visio open/edit/save acceptance.

### Cached shape text-frame placement

SVG and PNG export use the cached `TextStyle` frame: `TextPinX/Y`, `TextWidth/Height`, `TextLocPinX/Y` and `TextAngle`. Positions and dimensions use drawing inches, margins use physical inches, and angles use radians. The local-pin offset and asymmetric margins follow the native text-frame reflection and rotation order, then the shape and its containing groups. Plain labels and supported rich text use the same placement. Minimum fitting dimensions do not move the cached frame center. Rendering leaves the model unchanged.

Master instances inherit missing frame values and scale cached positions, dimensions and local pins during creation. Simple authored 2D masters default to a centered frame at 87.5% of the shape width and 75% of its height, with explicit fields overriding those defaults. They use the same effective frame when creating an instance before or after reopening, including partially specified frames and labels supplied at creation. Creation materializes the effective frame fields in each instance's independent `TextStyle`. Margins remain physical text metrics. Local caches, including values marked `F="Inh"`, take precedence over the master's numeric cache. XML/Open XML reopening and page or selection copying retain those effective values.

Cached `FlipX` and `FlipY` cells, including master inheritance and explicit zero overrides, reflect outline coordinates about each shape's local pin before its rotation and translation. SVG, PNG and diagram-page scenes use the same transform through containing groups. Text-frame placement follows [Microsoft's text-block coordinate order](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/3adb1f7a-74ff-4bba-bfdb-9e9ce4d79513); supported text stays readable and its angle follows reflection parity. Rendering preserves the native cells. Unsupported reflection formulas use their cache with an approximation diagnostic; an unusable cache omits the affected content with a diagnostic.

Controlled asymmetric outlines have independent-reader coordinate evidence. Text controls cover local pins, asymmetric margins, reflection combinations, plain/rich labels and XML/Open XML reopening. Off-centre reflected text placement differs from the independent reader, so native Microsoft Visio appearance acceptance remains open. Decorative stencil artwork and database fallbacks report reflection approximations. This profile does not evaluate text-frame formulas or retain their full formula provenance when saving modeled frame cells. Sheared frames and exact typography remain unqualified.

### Page coordinates and reflected diagrams

`VisioShape.GetAbsolutePoint(x, y)`, `GetBounds()`, `GetShapeBounds()` and `GetPageShapeBounds()` use page coordinates in drawing inches. They apply cached shape/master reflections and every containing group's transform. Bounds describe the transformed shape rectangle, not the exact painted outline. Spatial queries, connector glue and routing obstacles use the same frame. Native `PinX` and `PinY` remain coordinates in the containing group.

Selection alignment, distribution and grid placement use page bounds and convert movement back into each shape's parent coordinates. Overlapping selections position each distinct shape once while retaining the selection entries. Grid cells size themselves from those bounds; shapes share the cell centers. Container refitting uses the inverse containing-group transform. These operations retain native reflection cells and do not evaluate dynamic formulas. Dynamic reflection formulas use their usable cache. A reflection with neither a finite cache nor a constant numeric formula causes coordinate operations to throw `InvalidDataException`.

```csharp
var origin = shape.GetAbsolutePoint(0, 0);
var bounds = shape.GetPageShapeBounds();
var neighbors = page.ShapesIntersecting(bounds);
```

### Cached connector text-frame placement

Imported connectors with an explicit cached text frame retain its position, dimensions, local pin and angle relative to the preserved native connector frame. Translating endpoints moves the label with the curve; rotating or uniformly scaling the retained curve transforms the cached text box. Font sizes and margins keep their physical size. Saving materializes the projected frame in native text cells without changing the imported caches, source formatting or text markers. SVG, PNG, label layout and resolved inspection pins use the same projection.

`PlaceLabelAt()` and explicit `LabelPlacement.PinX/PinY` assignments select fixed page placement when both pins are set. Clearing a pin restores path placement. `PlaceLabel()` retains its path position and page-coordinate offsets when a file is reopened. These placement choices survive VDX/VSDX reopening and copying. An explicit change to `TextStyle.TextAngle` retains its page angle while an imported native pin continues to follow the connector. The placement properties retain imported cached values; `CreateInspectionSnapshot()` exposes the currently resolved page pin separately.

`ResizeLabelToText()` measures the box at the current connector scale, and later native scaling starts from that fitted size. Label overlap cleanup retains the native pin and angle binding when it moves a label; an unsuccessful search leaves the placement untouched. Reopened native previews use stored placement without implicit collision cleanup. A collapsed native frame requires explicit label placement before fitting or layout editing.

Saving stores API placement intent in the reserved `OfficeIMOConnectorLabelPlacement` Shape Data row. Loading consumes recognized values internally. Unknown producer values remain ordinary data; saving an explicit placement with a conflicting row name throws instead of overwriting that data.

```csharp
VisioConnector connector = document.Pages[0].Connectors[0];
connector.EndPoint = new OfficePoint(5, 6);
connector.PlaceLabel(0.25, offsetY: 0.2, width: 1.4, height: 0.35);
document.Save("updated.vsdx");
```

The NxBRE callout provides independent native input. Reference-reader controls qualify translation, rotation and uniform scale of its cached text frame through VDX and VSDX. Additional cases cover explicit overrides, fitting, layout cleanup, repeated reopening and page copying. SVG and raster previews use the same native rotation direction for plain and rich text. This profile follows the retained native geometry transform; replacing its route or kind requires explicit label placement. It does not evaluate ShapeSheet text formulas, qualify reflected or sheared frames, or establish exact typography and native Microsoft Visio open/edit/save acceptance.

### Physical drawing scale

SVG and raster export apply `PageScale / DrawingScale` after converting both scale settings to inches. Page bounds, shape and group geometry, connector routes, cached text frames and embedded-image placement follow that ratio. Font sizes, line weights, text margins, paragraph indents and paragraph spacing retain their physical size. Renderer fitting limits and connector-label collision distances also use physical inches. Native page values and shape caches remain in drawing units through rendering, copying and XML/Open XML reopening.

Native PageSheet distance caches use [Visio's internal inch units](https://learn.microsoft.com/en-us/office/vba/api/visio.cell.resultiu), while `U` supplies a display hint. Page dimensions, margins, grid and routing distances follow this contract in XML and Open XML. Untouched valid length cells retain their native caches, formulas and metadata; changing a modeled value or its display unit replaces the affected cell. See the [upgrade guidance](../MIGRATION.md#visio-native-page-lengths) for older OfficeIMO files with metric values in these caches.

`VisioSvgSaveOptions.PixelsPerInch` and `VisioPngSaveOptions.PixelsPerInch` describe physical page inches. The shared `ExportImage` API uses physical dimensions for target sizes, result metadata and raster pixel budgets. For example, a 1728-by-1152-inch drawing with `PageScale = 0.25 in` and `DrawingScale = 12 in` occupies a 36-by-24-inch page; at 48 pixels per inch its SVG and PNG are 1728 by 1152 pixels.

The scale and physical-metric rules follow Microsoft's [drawing-scale contract](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/1a60ffcb-969c-48f1-aa02-ff2228718043), [text-margin contract](https://learn.microsoft.com/en-us/office/client-developer/visio/leftmargin-cell-text-block-format-section) and [paragraph-indent contract](https://learn.microsoft.com/en-us/office/client-developer/visio/indleft-cell-paragraph-section). Controlled reduced, enlarged, mixed-unit and 1:1 drawings cover SVG/raster projection; the PRONOM template supplies independent scaled input. Exact typography and native Microsoft Visio rendering remain unqualified.

### Native geometry flags

Cached geometry paths honor section-level `NoShow`, `NoFill` and `NoLine` values in SVG and raster export. `NoShow` omits the path's painted geometry, `NoFill` omits its fill, and `NoLine` omits its outline. Text remains independent of these geometry flags. The section-level cells take precedence over older OfficeIMO `Row T="Geometry"` headers, which remain readable when the corresponding section cell is absent.

Filled open contours close implicitly for filling. SVG and raster export keep their outlines open unless the native path returns to its start or uses a closed ellipse primitive. Compound contours within a geometry section share an even-odd fill and retain their own stroke closure. Rendering does not add closing rows to the saved native geometry.

Supported single-path native connectors retain their structural route and label anchoring when their geometry is hidden or has no outline. Rendering omits that line and its arrowheads, and connector-label collision search excludes the invisible line. When auxiliary paths are hidden or have no outline, the unique outlined route remains available for rendering and label placement. Saving and copying retain the native geometry cells, formulas and row identities.

Generated builtin shape, master and connector geometry writes section-level settings and unique one-based path-row identities, following the [native geometry-section contract](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/c6f4364f-5fb7-49f3-993e-49d4d709aa02). Loaded geometry sections retain their source structure. Controlled XML/Open XML examples have independent libvisio outline and flag checks; broader producer coverage, multi-path connector artwork and native Microsoft Visio appearance remain unqualified.

The [legacy fixture evidence](../OfficeIMO.Visio.Tests/Fixtures/LegacyXml/README.md) covers independent drawings, stencils and a template, embedded bitmap rendering, schema validation, and libvisio reading/rendering. Native Microsoft Visio acceptance and broader producer coverage remain unqualified. No external reader or converter is a runtime dependency.

Run the [legacy XML example](../OfficeIMO.Examples/Visio/LegacyXml.cs) with `OfficeIMO.Examples --visio-xml` to create, reopen, edit, and convert a drawing.

### Native layers and hyperlinks

Loaded page layers keep their native row indices through XML/Open XML saving and page duplication. New rows use zero-based indices and skip indices already owned by imported rows. Copies remap local shape references in layer formulas and retain guarded native cell state. [Layer membership](https://learn.microsoft.com/en-us/office/client-developer/visio/layer-membership-section) uses the layer's zero-based position in the page list, independently of those row indices. Reordering `page.Layers` writes updated membership positions for shapes, nested children and connectors. Unchanged membership cells retain their source cache, formula, units and error state. Saving changed or cleared membership replaces its cache and removes the old formula and error; restoring the original membership later does not revive that discarded state.

`VisioLayer` and `VisioHyperlink` property assignments replace the corresponding imported formula, native null condition and producer error, including assignments of the current value. Units and untouched neighboring cells remain preserved. Hyperlink row names, numeric identities and localized names survive XML/Open XML saves. Page copies retain independent rows and assignment intent; imported master hyperlink edits leave the source unchanged. OfficeIMO preserves hyperlink targets as data and does not fetch them or evaluate their formulas.

New unnamed hyperlinks use unused `Row_n` names, including when explicit rows appear later in the list or a master supplies inherited rows. Explicit names remain unchanged: callers must keep them unique within the local section, and a name matching a master row intentionally overrides that row. Reservations follow the effective shape or dynamic-connector master, including default master lookup and masters retained from loaded documents. Page and selection copies retain the source's allocated row names without changing the source object. Newly registered simple and grouped masters retain their hyperlink rows when saved.

```csharp
VisioPage page = document.Pages[0];
VisioLayer? review = page.FindLayer("Review");
if (review != null) review.Visible = true;

VisioHyperlink link = page.Shapes[0].AddHyperlink(
    "https://example.org/review", "Review details");
link.SubAddress = "Detail";
```

The AngularJS and NxBRE producer fixtures qualify page layer rows, values and memberships. Jon Breen's independently produced Visio 2002 sample qualifies a named shape hyperlink, its cached values and `No Formula` state, page copying, target edits and repeated XML/Open XML reopening. Controlled inputs cover multiple rows, name collisions, inherited root and child rows, shapes, connectors and masters across the three XML families, copying and imports. Broader independent hyperlink producers, dynamic hyperlink formulas and native Microsoft Visio open/edit/save remain unqualified.

### Layer selection in previews

SVG and raster previews use `VisioLayerRenderMode.Visible` by default. Set `LayerMode` on `VisioSvgSaveOptions`, `VisioPngSaveOptions` or `VisioImageExportOptions`, or call `.LayerMode(...)` on a fluent image builder:

```csharp
string screenPreview = page.ToSvg();
byte[] printPreview = page.ToPng(new VisioPngSaveOptions {
    LayerMode = VisioLayerRenderMode.Printable
});
OfficeImageExportResult allLayers = page.ToImage()
    .LayerMode(VisioLayerRenderMode.All)
    .AsSvg()
    .Export();
```

| Mode | Layer flag used |
| --- | --- |
| `Visible` | `Visible`, independently of `Print` |
| `Printable` | `Print`, independently of `Visible` |
| `All` | Includes all memberships |

A shape or connector appears when any of its layers enables the selected mode. Unassigned shapes and undeclared layer names remain visible. Names match display or universal names without case sensitivity. Group children use their own memberships; a hidden parent does not hide a visible child. Connectors use their own layers, independently of endpoint visibility. Hidden geometry, text, images and stencil artwork are omitted from the preview and from label collision avoidance. Font and image diagnostics describe only included content.

Preview selection leaves the document unchanged. Reader and neutral-model extraction retain hidden text, links, shape data and topology; semantic PDF output retains that content too. Optional PDF raster previews use the selected `PngOptions.LayerMode`. Controlled XML/Open XML inputs and independent libvisio display callbacks qualify the layer-selection profile. Native Microsoft Visio screen/print appearance, dynamic visibility formulas and layer color overrides remain unqualified.

### Cached text-style inheritance

SVG and raster export resolve explicit native `TextStyle` chains for character and paragraph properties. Sparse local rows override individual inherited cells. A shape instance with an explicit `TextStyle` uses that style instead of master text formatting; otherwise it uses the corresponding master row, or the master’s first row when that index is absent. Implicit row positions and padded numeric indices select the corresponding rows. Cached values apply even with `F="Inh"`; rendering does not evaluate formulas or materialize inherited properties into saved shapes. Modeled single-row edits and `SetShapeSheetSection()` edits reach rendering and XML/Open XML reopening. Replacing plain text removes native `cp`/`pp` selections; default row 0 properties still apply in this inherited-style profile.

Loaded shape and connector references retain explicit values and omitted attributes through saving and same-document duplication. Native stylesheet headers and document defaults are retained. `DefaultTextStyle` describes the next shape created by a drawing tool; it does not bind an existing unbound shape. The independent Shorewall drawing qualifies its explicit style 6→7→0 chain: Arial 8 pt with local hanging-indent and bullet rows. Legacy XML outputs pass the Microsoft schema check and retain this label in libvisio callbacks. Open XML output uses a package-relative root document relationship, while loading also accepts existing package-rooted targets. Open XML output uses the native font-name contract described below; libvisio retains the same Arial 8 pt label and three list elements. Generated cases cover sparse rows, master precedence, edits and reopening. Master first-row fallback in independently produced Open XML files and native Microsoft Visio appearance remain unqualified. Serialized legacy style 0 is used as a cached approximation; native Visio’s built-in style 0 can differ.

Inheritance stops after 64 links. Cyclic, dangling or ambiguous references, uncached formulas and themed values emit `VISIO_TEXT_STYLE_INHERITANCE` approximation diagnostics through `ExportImage`; strict image policy can reject them. An `EnableTextProps=0` style excludes its own text properties. Traversal through that disabled style’s parent is unqualified, so the renderer stops and reports the fallback. Native row and section deletion block inherited text properties. Collection-level deletion cannot be represented by legacy XML’s separate row elements and produces a `VDX_SECTION_DELETION` loss report on export. Dynamic themes, arbitrary formula evaluation, inherited text-block geometry, tabs and exact typography remain outside this cached profile.

### Native font names

Open XML files declare fonts with `FaceName.NameU` and store family names in cached `Character.Font`, `AsianFont`, `ComplexScriptFont`, and `Paragraph.BulletFont` cells. Loading accepts these native names and older OfficeIMO numeric references. Saving converts resolvable numeric references and constant formulas, including `GUARD(...)`, to the native family and `FONT("family")` form. Other formulas remain unevaluated. Auxiliary font values `0` and empty strings retain their fallback meaning.

Legacy XML export uses numeric font identities. Distinct charset entries for the same family, legacy font-table metadata, numeric formulas, and producer errors survive an unchanged Open XML save and reopen. Explicit `FontFamily` assignments replace the old font formula and error state, including assignments of the same family. Literal family names assigned through `SetShapeSheetSection()` receive a native table entry.

Loaded masters retain their source font context when registered in another document or used to create instances. Destination font collisions are rebound during saving; repeated saves into different documents leave the shared source master unchanged. New children authored in the destination use its font table.

The independent Shorewall fixture qualifies Arial 8 pt and three list elements through legacy and Open XML conversion and reopening with libvisio. A `FontFamily = "Consolas"` edit retains that size and list structure. Legacy output passes the Microsoft XML schema; native `FaceNames` passes its documented element-schema check. Full Open XML schema validation, native Microsoft Visio acceptance, theme fonts, complex-script typography and broader independent producer coverage remain unqualified.

### Native paragraph bullets

SVG and raster export project cached native bullet rows onto shared paragraph labels. Standard values 1–7 use Unicode round, diamond, filled square, empty square, four-diamond, arrow and check markers. A single-line `BulletStr` supplies a custom marker. Custom `BulletFont` values resolve through the document font table; zero uses the first character's font. `BulletFontSize` accepts absolute native inch sizes, zero for the first character's size, and negative relative sizes (`-1` means 100%). Font fitting reduces label and body sizes together while keeping their declared positions. Raster export uses the [shared glyph fallback](../OfficeIMO.Core/README.md#image-export-density) when no selected font covers a marker; font substitution remains an approximation.

The label starts at `IndLeft + IndFirst`; `TextPosAfterBullet` sets the first text position relative to that anchor. Zero or an absent value uses the quarter-inch minimum label width recorded in the independent round-bullet fixture. Wrapped continuation lines use `IndLeft` and paint no additional marker. Labels wider than the declared gap move the first text position to avoid overlap. Source-section edits, direct master labels, page copies and XML/Open XML reopening retain this behavior. Rendering leaves native rows and plain `Text`/`Label` unchanged; replacing plain text replaces its native bullet boundaries.

The [Shorewall fixture](../OfficeIMO.Visio.Tests/Fixtures/LegacyXml/README.md) qualifies round-bullet cells and paragraph markers against an independent reader. Other marker kinds, custom font/size overrides and fitting have generated artifact checks. Native Microsoft Visio appearance, symbol-font character remapping, RTL lists and exact marker typography remain unqualified. Invalid label values retain the supported character projection and preserved source. Labels that cannot fit the frame can be omitted by layout; fitting does not change native source values. The shared 100,000-character and 4,096-run projection limits include bullet labels.

## Diagram page scenes

`ToDrawings` projects every document page into a detached `OfficeDrawing` at
its physical dimensions in points. `ToDrawing` projects a single page. These
operations use cached source values and leave native XML unchanged.

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Visio;

VisioDocument document = VisioDocument.LoadLegacyXml("workflow.vdx").Value;
var options = new VisioDrawingOptions { LayerMode = VisioLayerRenderMode.Printable };
var projected = document.ToDrawings(options);
for (int index = 0; index < projected.Value.Count; index++) {
    OfficeDrawing page = projected.Value[index];
    File.WriteAllText($"page-{index + 1}.svg",
        OfficeDrawingSvgExporter.ToSvg(page, 1, OfficeSvgSizeUnit.Point));
}
foreach (var diagnostic in projected.Report.FidelityDiagnostics)
    Console.WriteLine($"{diagnostic.Location}: {diagnostic.Message}");
```

The cached profile includes nested group placement, master outlines, native
geometry visibility flags, multiple connector outlines, physical drawing scale,
styled runs and paragraphs, cached text frames, and supported bitmap crop,
rotation and mirroring. Fonts and an optional text shaper belong to
`VisioDrawingOptions` and travel with the returned scenes. For raster export,
set the shared export scale to `targetDpi / 72` because these scenes use points.

The default layer policy selects printable content independently of screen
visibility; children and connectors have their own membership. Blank and
background pages remain in document order. Associated backgrounds are composed
behind each page, from the deepest background to the foreground. Each page
keeps its own physical scale and layer declarations; cached artwork shares the
physical lower-left origin and is clipped by the foreground surface. SVG and
PNG previews use the same composition. Native Visio rendering acceptance for
background chains with differing page sizes or scales remains open.
Stencil masters must be instantiated
on a page first. The operation limits pages, visited shapes/connectors
(including excluded layers), projected geometry points, and retained image
pixels/payload bytes across all pages, and accepts a cancellation token.
Repeated background instances consume those object and image budgets; the
page limit also bounds a single page's composition chain. Image
defaults allow 64 million decoded pixels and 128 MiB of projected PNG bytes;
each decode also uses the existing 32-million-pixel and 64-MiB source limits.

The source-qualified report distinguishes policy exclusions from fidelity loss.
Missing or cyclic background references, connector fills, metadata and opaque
content report omissions. Curve flattening, builtin outline fallback, unsupported
reflection formulas, triangular arrows and text layout report approximations.
Cached outline and shape text-frame reflections follow the placement contract above. Shared
fitting reduces text to fit its cached frame and reports an omission
when the selected fonts still cannot retain all content. Native
font placement and fitting, themes, data graphics and arbitrary ShapeSheet
recalculation are not qualified. `RequireNoLoss` rejects reported loss before
returning scenes. The [PDF diagram-page mode](../OfficeIMO.Visio.Pdf/README.md#diagram-pages)
uses this same owner and preserves its report in the PDF result.

## What it does

- Creates and edits Visio pages, shapes, connectors, text, styles, Shape Data, layers, hyperlinks, containers, comments, and metadata.
- Provides fluent diagram builders for common flowchart, block, dependency, architecture, network, topology, swimlane, org chart, sequence, timeline, and generic graph scenarios.
- Supports loaded-diagram editing, shape selection, topology queries, stencil replacement/migration planning, and container maintenance.
- Edits nested container topology, swimlane metadata and geometric assignment, threaded comments and authors, generated data graphics and legends, and source-preserving ShapeSheet sections and formulas.
- Provides rotation- and connector-aware resize-to-content plus deterministic topology-aware whole-page relayout for dense imported diagrams.
- Preserves opaque VBA project payloads in macro-enabled drawing, template, and stencil packages without executing or rewriting VBA.
- Exports headless PNG, JPEG, TIFF, SVG, and lossless WebP previews for proof and review workflows.
- Includes validation and quality analysis for generated and loaded diagrams, including connector-label collisions with unrelated connector paths.
- Carries caller-supplied stencil license, attribution, and unsupported-master state through shapes, catalogs, and manifests without inferring redistribution rights.

## Editing existing diagrams

`Load` materializes an editable diagram. File and stream entry points accept the
same `VisioLoadOptions`. New asynchronous calls should use the options-first
shape; token-first overloads remain available for source and binary compatibility.

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;
using Color = OfficeIMO.Drawing.OfficeColor;

VisioDocument.Load("operations.vsdx")
    .AsFluent()
    .ExistingPage("Operations", page => page
        .ShapesWithData("Owner", "Ops", selection => selection
            .Fill(Color.LightBlue)
            .ShapeData("Reviewed", "Yes", "Reviewed", VisioShapeDataType.Boolean))
        .ShapesContainingText("Legacy", selection => selection
            .Text(shape => shape.Text!.Replace("Legacy", "Production", StringComparison.Ordinal))))
    .End()
    .Save("operations.updated.vsdx");
```

### Loaded-diagram compatibility boundary

Loaded Open XML Visio editing covers pages, shapes, connectors, text, styles,
Shape Data, hyperlinks, layers, nested containers, swimlanes, threaded comments,
data graphics, legends, typed ShapeSheet sections/formulas, topology queries,
resize-to-content, and whole-diagram relayout. Template and stencil packages use
the same model, including page-less stencils with masters. Macro-enabled variants
retain their VBA project as an opaque bounded payload.

The boundary is deliberate: OfficeIMO does not execute VBA, evaluate arbitrary
ShapeSheet formulas, or claim native Visio layout equivalence. Typed edits retain
unmodeled ShapeSheet rows, cells, attributes, and supported opaque payloads in
their existing preservation stores. Arbitrary producer package parts outside
those stores are not advertised as editable. Ambiguous swimlane geometry is
reported instead of assigned, container cycles are rejected, and semantic
relayout keeps containers and generated adornments fixed unless the caller opts in.

### Signed diagrams

`InspectSignatures()` detects Open XML signature-origin relationships and XML signature parts. Saving a
loaded signed diagram is blocked by default because rebuilding the package would invalidate that evidence. Set
`SignatureMutationPolicy = VisioSignatureMutationPolicy.RemoveInvalidatedSignatures` only when removing the stale
signature carrier is the intended result. `SignPackageSignature(...)` and `ValidatePackageSignatures(...)` create
and cryptographically validate OPC signatures through an explicitly supplied `IOfficeSecurityProvider`.

## Examples

The quick start shows the fluent page API. These examples show the higher-level builders and editing surfaces that belong in `OfficeIMO.Visio`.

### Flowchart builder

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

VisioDocument.Create("flowchart.vsdx")
    .Flowchart("Property buying flowchart", flow => flow
        .Title()
        .Layout(VisioFlowchartLayout.TwoColumnContinuation)
        .RouteBranches(laneSpacing: 0.5)
        .Start("start", "Start with an agent\nyou trust")
        .Step("consult", "Consult with agent to\ndetermine needs")
        .Decision("agreement", "Agreement?")
        .Step("contract", "Accept the contract")
        .End("close", "Close on the property")
        .Branch("agreement", "No", "consult")
        .Branch("agreement", "Yes", "contract")
        .Callout("agreement", "retry-note", "Loop back if rejected", VisioSide.Right))
    .Save();
```

### Network topology builder

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

VisioDocument.Create("network-topology.vsdx")
    .NetworkTopologyDiagram("Branch topology", topology => topology
        .Title()
        .Root("internet", "Internet", VisioNetworkNodeKind.Internet)
        .Firewall("firewall", "Firewall")
        .Switch("core", "Core Switch")
        .Server("app", "App Server")
        .Database("db", "Database")
        .Workstation("finance", "Finance PC")
        .Subnet("edge", "Edge", "internet", "firewall", "core")
        .Subnet("servers", "Server Zone", "app", "db")
        .Ethernet("internet", "firewall", "WAN")
        .Trunk("firewall", "core", "uplink")
        .Trunk("core", "app", "10Gb")
        .Ethernet("app", "db"))
    .Save();
```

### Sequence diagram builder

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

VisioDocument.Create("sequence.vsdx")
    .SequenceDiagram("Checkout sequence", sequence => sequence
        .Title()
        .Theme(VisioStyleTheme.Fluent())
        .Actor("customer", "Customer")
        .Participant("web", "Web App")
        .Control("api", "Orders API")
        .Database("db", "Orders DB")
        .Call("customer", "web", "Checkout")
        .Call("web", "api", "POST /orders")
        .Async("api", "db", "Persist order")
        .Return("api", "web", "201 Created")
        .SelfMessage("web", "Render receipt"))
    .Save();
```

### Timeline roadmap

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

VisioDocument.Create("roadmap.vsdx")
    .TimelineDiagram("Product roadmap", timeline => timeline
        .Title()
        .Theme(VisioStyleTheme.Modern())
        .Range(new DateTime(2026, 1, 1), new DateTime(2026, 6, 30))
        .Span("discovery", new DateTime(2026, 1, 8), new DateTime(2026, 2, 20), "Discovery")
        .Span("build", new DateTime(2026, 2, 21), new DateTime(2026, 5, 15), "Build", lane: 1)
        .Release("preview", new DateTime(2026, 5, 20), "Public preview", VisioTimelinePlacement.Below)
        .Milestone("ga", new DateTime(2026, 6, 25), "GA"))
    .Save();
```

### Layers and Shape Data

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.Drawing.OfficeColor;

var document = VisioDocument.Create("architecture.vsdx");
var page = document.AddPage("Architecture");
page.AddLayer("Infrastructure");
page.AddLayer("Annotations").Print = false;

var server = page.AddStencilShape(VisioStencils.Network.Get("server"),
    "server", 2, 5, "Server");
server.SetShapeData("Owner", "Platform", "Owner",
    VisioShapeDataType.String, "Owning support team");

page.AddToLayer("Infrastructure", server);
page.SelectWithShapeData("Owner", "Platform")
    .Fill(Color.LightBlue)
    .ShapeData("Reviewed", "Yes", "Reviewed",
        VisioShapeDataType.Boolean, "Architecture review complete");

document.Save();
```

When a catalog comes from an external package, make its provenance explicit:

```csharp
var options = new VisioStencilPackageLoadOptions {
    SourceLicense = "License identifier or notice supplied by the caller",
    SourceAttribution = "Required source attribution",
    IncludeUnsupportedMasters = true
};
```

Unsupported masters remain marked as unsupported even when included for inventory or migration planning. Including them does not turn them into a fully supported authoring contract or grant redistribution rights.

### Headless image export

```csharp
using OfficeIMO.Visio;

var document = VisioDocument.Create("pipeline.vsdx");
var page = document.AddPage("Pipeline").Size(8, 4);
var build = page.AddProcess(1.5, 2, 1.4, 0.7, "Build");
var ship = page.AddProcess(5.5, 2, 1.4, 0.7, "Ship");
page.AddConnector(build, ship, ConnectorKind.RightAngle, VisioSide.Right, VisioSide.Left)
    .EndArrow = EndArrow.Arrow;

document.SaveAsSvg("pipeline.svg", new VisioSvgSaveOptions {
    PixelsPerInch = 96,
    BackgroundColor = null
});

document.SaveAsPng("pipeline.png", new VisioPngSaveOptions {
    PixelsPerInch = 144,
    Supersampling = 3
});

OfficeImageExportResult webp = document
    .ToImage()
    .AtDpi(144)
    .FitWithin(1600, 1200)
    .ResolveConnectorLabelOverlaps()
    .AsWebp()
    .Save("pipeline.webp");

IReadOnlyList<OfficeImageExportResult> pages = document
    .ToImages()
    .AllPages()
    .IncludeStencilArtwork()
    .IncludeConnectorLabels()
    .AsJpeg()
    .Save("pipeline-pages");
```

## Content provenance

`VisioDocument.InspectProvenance("input.vsdx")` reports C2PA and AI-specific IPTC metadata in the drawing and its supported embedded images. `VisioDocument.RemoveProvenance("input.vsdx", "clean.vsdx")` removes the selected carriers. Signed-package mutation is blocked unless `OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures` is selected explicitly. Optional cryptographic C2PA verification remains in `OfficeIMO.Security`.

## Related packages and limits

- `OfficeIMO.Visio` generates and edits drawing, template, stencil, and macro-enabled Open XML Visio packages without requiring desktop Visio at runtime.
- `ToOfficeDocumentModel(...)` projects Visio content into the dependency-free model in `OfficeIMO.Core`; direct converters can consume it without taking Reader dependencies.
- External stencil packages retain their package and licensing requirements; OfficeIMO records caller-supplied provenance but never infers licensing terms.
- Use [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice) for PowerShell workflows.
- Open Visio product work is listed in the repository [roadmap](../Docs/ROADMAP.md).

## Deeper docs

- [Repository roadmap](../Docs/ROADMAP.md)
- [Reader package family](../Docs/officeimo.reader.md)
- [Examples](../OfficeIMO.Examples)

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`; `net472` is included when building on Windows.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** `System.IO.Packaging`; Microsoft BCL compatibility packages are used on older targets.
- **OfficeIMO:** `OfficeIMO.Core`. The VSDX model, builders, editing, topology, validation, and PNG/JPEG/TIFF/SVG/WebP renderers are first-party.
- **Security:** Open XML signature carriers are inspected and signed-diagram mutations fail safely without a cryptographic dependency. Signature creation and validation accept an explicit `IOfficeSecurityProvider`; `OfficeIMO.Security` is not pulled transitively.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 2 | 0 | 0 | 0 | 0 | 0 |
| Read | 1 | 0 | 0 | 0 | 0 | 1 |
| Edit | 1 | 0 | 0 | 1 | 0 | 0 |
| Preserve | 1 | 0 | 0 | 0 | 0 | 0 |
| Inspect | 3 | 0 | 0 | 0 | 0 | 0 |
| Validate | 3 | 0 | 0 | 0 | 0 | 0 |
| Remove | 2 | 0 | 0 | 0 | 0 | 0 |
| Convert | 0 | 3 | 0 | 0 | 0 | 0 |
| Export | 5 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Visio` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
