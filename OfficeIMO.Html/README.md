# OfficeIMO.Html

`OfficeIMO.Html` contains the shared HTML parser, resource policy, layout scene, and direct PNG/JPEG/TIFF/SVG/WebP rendering APIs used by OfficeIMO converters.

It owns the reusable parts that should behave consistently across HTML-to-Markdown, HTML-to-Office, image, and PDF workflows:

- trust-aware parsing profiles and shared source, DOM, CSS, selector, responsive-image, and semantic-metadata limits
- shared HTML, CSS, and data-URI charset handling, including legacy web encodings
- URL policy evaluation and base URI resolution
- owned document snapshots, node queries and edits, with replaceable parser and charset providers
- DOM traversal facts and node/depth limit tracking
- image source discovery and deterministic responsive candidate selection for `img`, lazy-loading attributes, `srcset`, `sizes`, and `picture/source`
- image data URI parsing and media-type extension mapping
- deterministic accessible-name, ARIA heading, EPUB structural-semantic, and logical quote/code/footnote projection
- browser-free HTML layout for continuous and paged output
- structured Presentation MathML routed through the shared OfficeIMO.Core expression and vector-rendering model
- bounded CSS length math, caller stylesheets, deterministic media preferences, running strings and elements, and Unicode-range-aware font fallback packs
- first-party TrueType/WOFF 1, .NET 8+ single-face WOFF 2, CFF/CFF2, and variable-font programs, plus an optional complete OpenType shaping provider
- direct PNG, JPEG, TIFF, SVG, and lossless WebP export over `OfficeIMO.Drawing`
- one typed semantic document projection shared by Excel, PowerPoint, and OneNote importers
- executable target capability contracts, preflight analysis, and source-to-target diagnostic provenance
- one operation-scoped resource session for policy, resolution, MIME checks, caching, deduplication, budgets, timeouts, and digests
- fidelity scoring across structure, text, styles, resources, annotations, formulas, charts, geometry, and reopened native artifacts

Markdown, Word, Excel, PowerPoint, RTF, Email, MHTML, and PDF models remain in their owning packages. Those projections are explicit: for example, HTML becomes a `WordDocument` through `OfficeIMO.Word.Html` and a `MarkdownDoc` through `OfficeIMO.Markdown.Html`.

To save a source HTML copy in another directory while retaining its relative resource paths:

```csharp
HtmlConversionDocument report = HtmlConversionDocument.Load("reports/service-review.html");
File.WriteAllText("exports/service-review.html", report.ExportSourceHtml());
```

`ExportSourceHtml()` records the original effective base URI in the document. It preserves source markup; referenced files remain at their original locations. Use `SourceHtml` for the exact original text or `HtmlForConversion` for policy-normalized conversion HTML.

## Inspect and edit owned HTML

```csharp
using OfficeIMO.Html;
using OfficeIMO.Html.Dom;
using OfficeIMO.Markdown.Html;

HtmlDocument source = HtmlDocumentEngine.Default.ParseDocument(
    "<table><tbody><tr id='items'><th>Item</th></tr></tbody></table>");
HtmlElement row = source.QuerySelector("#items")!;
HtmlDocumentFragment cells = HtmlDocumentEngine.Default.ParseFragment(
    "<td>Quarterly report</td><td>Approved</td>", row);

HtmlDocument edited = source.Edit(document => {
    HtmlElement targetRow = document.QuerySelector("#items")!;
    targetRow.AppendChild(document.ImportNode(cells));
});
HtmlConversionDocument conversion = HtmlConversionDocument.FromDocument(edited);
string markdown = conversion.ToMarkdown();
byte[] preview = conversion.ToPng();
```

`HtmlDocumentEngine` is the provider-neutral document entry point. The default engine
uses the packaged AngleSharp provider, while its constructor accepts any
`IHtmlParserProvider`. Full documents and contextual fragments are immutable owned
snapshots. Fragment parsing uses the supplied element and ancestor context, including
table insertion modes, foreign namespaces and ancestor forms. Import a returned
fragment into a mutable destination before insertion; appending it splices its children.

`HtmlConversionDocument.Document` is also an immutable source snapshot. `Edit` clones it and freezes the result;
the original document, retained node handles and cached conversion results remain
unchanged. `Clone` returns a mutable tree for a single owner, while
`HtmlConversionDocument.FromDocument` captures the attached tree, including template
contents, as an immutable conversion snapshot.
`CreateDocumentForConversion` returns a mutable, policy-normalized owned tree.
Use `NodeId` within an edit lineage and `SnapshotId` to distinguish tree instances.
General-purpose `HtmlDocument.Clone` retains detached nodes and original source
offsets. Conversion capture and conversion edits omit detached editing history;
attached nodes retain their IDs and source offsets. Edited HTML is serialized
from the tree; it does not preserve the original spelling of entities or whitespace
inside tags. Template contents are exposed separately through `TemplateContent`.

The lightweight contracts live in `OfficeIMO.Html.Core`, which has no parser,
graphics, font or browser dependency. `OfficeIMO.Html.AngleSharp` supplies the
current parser, selectors, serializer and web charset labels. Applications needing
only HTML inspection can use those two packages. The conversion/rendering package
still uses AngleSharp and AngleSharp.Css internally; selecting another parser does
not remove those runtime dependencies.

Set `HtmlConversionDocumentOptions.ParserProvider` for an alternative inert parser
and `InputEncodingProvider` for source-byte charset resolution. Explicit `Encoding`
arguments on `Load` override sniffing. Resource decoding is configured separately:
`HtmlResourcePipeline.TryDecodeStylesheet` and `HtmlDataUri.DecodeText` accept a charset
provider. The default provider retains legacy code-page support and its existing
process-wide .NET code-page registration. HarfBuzz remains an optional typography
provider through `OfficeIMO.Drawing.HarfBuzz`.

The current structural adapter retains a native tree alongside an owned tree when
owned nodes are requested. Conversion-only parsing keeps the owned projection lazy.
Source length is checked before parsing; node/depth limits, including template
contents, are checked after the native parser constructs its tree. For owned input,
`MaxInputCharacters` also bounds aggregate attached names, attribute values, text,
comments and doctype identifiers before copying or native projection. Canonical
source serialization stops when its expanded output exceeds that limit. These are
separate checks: an owned tree can contain data that HTML serialization omits, such
as children of void elements. Cancellation is cooperative, not a hard worker
memory or execution-time limit. Context reconstruction needed by the retained provider
is isolated inside `OfficeIMO.Html.AngleSharp`; consumers retain the same owned API when
the provider changes. Provider-independent CSS execution is being adopted in qualified
vertical slices while the retained CSS provider covers the remaining grammar.

The owned property grammar covers CSS-wide keywords and selected `display`, `visibility`,
`opacity`, `color`, physical width and height constraints, and physical margin and padding
longhands. It represents constant number/percentage calculations, contextual length-percentage
expressions, and legacy or modern sRGB, HSL, and HWB functions as typed values. Computed
opacity is converted to a clamped number, while functional colors use the shared `OfficeColor`
conversion. Width and spacing percentages remain typed at computed-value time and resolve
against the layout reference at used-value time. Inline declarations enter the managed
cascade through the lossless owned style-block parser.

Qualified rules whose declarations stay inside that property slice retain their original
OfficeIMO syntax nodes and selector AST through the cascade, including nested qualified rules
inside style rules and supported `@media`, `@supports`, `@layer`, and `@container` groups.
Direct declaration runs preserve their authored position around nested rules. The owned matcher
handles atomic selector lists, stylesheet namespace bindings, type, universal, id, class,
and attribute selectors, the four structural combinators, selected child/of-type and An+B
pseudo-classes, a single-range `:lang()`, and `:is()`, `:where()`, and `:not()`. Grouped rules use the owned path only
when every selector and declaration is qualified, so a partially understood list cannot
silently change which elements match. For a namespace-qualified selector that also needs a
wider pseudo-class, the owned matcher retains the stylesheet namespace and combinator envelope
while the provider evaluates that pseudo-class against the current element. `:is()` and `:where()`
use forgiving lists, while `:not()` stays strict. `:empty` follows deployed browser behavior:
whitespace text makes an element non-empty and comments do not.

AngleSharp.Css remains in the default package for wider property grammars, dynamic and
relational pseudo-classes, filtered `:nth-child(... of S)`, conditional-group evaluation,
pseudo-elements, unknown at-rules, and other fallback cases. Retained declarations can still use an owned
selector match, including namespace-qualified selectors. OfficeIMO keeps rule order, cascade
layers, computed-style projection, and fallback selection stable while the owned subset grows.

Cascade explanations are opt-in so normal rendering does not retain candidate graphs for
every element:

```csharp
using OfficeIMO.Html.Css;

HtmlConversionDocument source = HtmlConversionDocument.Parse(
    "<style>@layer theme { .status { color: blue; width:calc(24px + 25%) } }</style>" +
    "<p class='status' style='color:lime'>Ready</p>");
var styles = HtmlComputedStyleEngine.Compute(source, new HtmlComputedStyleOptions {
    IncludeCascadeTraces = true
});
HtmlCssCascadeTrace trace = styles[source.Document.QuerySelector(".status")!]
    .GetCascadeTrace("color")!;
HtmlCssCascadeCandidate winner = trace.Candidates.Single(candidate =>
    candidate.Decision == HtmlCssCascadeDecision.Selected);

HtmlComputedStyle statusStyle = styles[source.Document.QuerySelector(".status")!];
if (statusStyle.TryGetTypedValue("width", out HtmlCssPropertyValue? width)) {
    HtmlCssLengthResolutionResult used = HtmlCssMathResolver.ResolveLength(
        width.MathExpression!,
        new HtmlCssLengthResolutionContext { PercentageReference = 640 });
}
```

The trace uses OfficeIMO types only. It reports computed value, inheritance/reset state,
source kind, selector, layer, specificity, importance, source order, winner decision, and
the owned grammar status of each retained candidate. `IsEffective` identifies the authored
declaration that won the local cascade. `InvalidAtComputedValue` identifies a failed custom-property
substitution whose result fell back to inheritance or the property's initial value. A retained provider may normalize a
stylesheet declaration before the cascade sees it; inspect `HtmlCssStyleSheet` when exact
authored spelling and source spans are required. Traces are currently available for the
properties in `HtmlCssPropertyCatalog`. Inline syntax parsing uses the conversion operation's
CSS byte, token, syntax-node, declaration, nesting, and resolved-selector expansion limits; the explicitly unbounded document overload does
not introduce the standalone CSS parser's default input ceiling.

See [the migration guide](../MIGRATION.md#owned-html-documents-and-callbacks) for the
replaced public DOM signatures.

## Shared Office HTML document shell

Office adapters use `OfficeHtmlDocumentShell` and `OfficeVisualThemeKind` for consistent semantic, editable round-trip, positioned-review, and print-review output. The embedded stylesheet supplies explicit palettes, readable typography, responsive page regions, tables, forms, figures, code, adapter panels, and print rules without making each adapter maintain a separate CSS implementation.

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Html;

string review = OfficeHtmlDocumentShell.WrapBody(
    "<main class=\"officeimo-document\"><h1>Review</h1></main>",
    new OfficeHtmlDocumentOptions {
        EmitDocumentShell = true,
        Title = "Conversion review",
        Language = "en-GB",
        Theme = OfficeVisualThemeKind.Report,
        IncludeDefaultStyles = true,
        BodyClass = "customer-review",
        NewLine = "\n"
    });
```

Set `EmitDocumentShell = false` to return a fragment. Adapter save options expose this object through `DocumentOutput`; their older title, language, theme, style, fragment, and newline properties remain synchronized aliases. Adapter-required body classes are retained and `BodyClass` values are appended with stable de-duplication. The shell is presentation only. Parsing, resource policy, layout interpretation, conversion diagnostics, and static execution boundaries remain owned by the managed HTML engine and the destination adapter.

`HtmlTargetCapabilityContracts` describes conversion routes directionally. `HtmlToTarget` and `TargetToHtml` have independent entry points, result contracts, I/O boundaries, diagnostics, profiles, and feature classifications; a missing reverse route is represented by `TargetToHtml == null`. Catalog collections and finalized gallery or conversion results are defensive snapshots; mutable reports remain available only while callers or converters assemble a result.

## Responsive image selection

Use the DOM-independent selector when a crawler, resource broker, or another host
needs the same candidate decision as the renderer:

```csharp
HtmlResponsiveImageSelection selected = HtmlResponsiveImageSelector.Select(
    "small.webp 400w, medium.webp 800w, large.webp 1200w",
    "(max-width: 600px) 100vw, 50vw",
    defaultSource: "fallback.webp",
    new HtmlResponsiveImageSelectionOptions {
        ViewportWidth = 800,
        ViewportHeight = 600,
        DevicePixelRatio = 2
    });

string selectedUrl = selected.Candidate.Url; // medium.webp
```

The selector supports density descriptors and width descriptors normalized by a
bounded `sizes` list. `MaxSizesCharacters` bounds direct selector calls, while
`HtmlConversionLimits.MaxResponsiveImageSizesCharacters` applies the shared conversion
boundary. It evaluates supported media conditions through the shared CSS
media engine, resolves supported CSS lengths and math, keeps the first duplicate
density, reports whether a default `src` supplied the selected candidate, and falls back
to `100vw` when no supported size matches. Resource policy is applied before candidate
selection. Static rendering derives device density from
`HtmlRenderMediaFeatures.ResolutionDpi`.

## Direct HTML rendering

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Html;

string html = "<h1>Status</h1><p>Generated by OfficeIMO.</p>";
HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
var options = new HtmlRenderOptions {
    ViewportWidth = 720,
    Margins = HtmlRenderMargins.All(24),
    Scale = 1.5,
    MediaFeatures = new HtmlRenderMediaFeatures {
        PreferredColorScheme = HtmlPreferredColorScheme.Light,
        ReducedMotion = HtmlReducedMotionPreference.Reduce
    }
};
options.AdditionalStylesheets.Add("body { color: hsl(215 35% 20%); }");

byte[] png = source.ToPng(options);
byte[] jpeg = source.ToJpeg(options);
byte[] tiff = source.ToTiff(options);
string svg = source.ToSvg(options);
byte[] webp = source.ToWebp(options);

OfficeImageExportResult pngSave = source.SaveAsPng("status.png", options);
IReadOnlyList<OfficeImageExportResult> webpPages = source
    .ToImages(options)
    .Paged()
    .AsWebp()
    .Save("status-pages");
```

Use a named render request when CSS media, viewport behavior, pagination, page
selection, and output format must be independently reviewable. The request takes an
immutable options snapshot. The retained result records the exact profile, surfaces,
source offsets, clipping, requested scale and background, declared providers, and
loss diagnostics before an encoder consumes it.

```csharp
var request = HtmlRenderRequest.Create(
        HtmlRenderIntentProfile.ScreenSnapshotPaged,
        HtmlRenderEncoder.Png,
        new HtmlRenderOptions {
            ViewportWidth = 816,
            PageSize = OfficePageSizes.A4,
            Scale = 1.5
        },
        HtmlRenderDocumentState.EditedSnapshot)
    .WithPageSet(HtmlRenderPageSet.Pages(firstPageIndex: 0, pageCount: 2));

HtmlRenderResult retained = HtmlRenderEngine.Execute(source, request);
IReadOnlyList<OfficeImageExportResult> pages = retained.ExportImages();

HtmlRenderSurface firstSurface = retained.GetSurface(0);
OfficeDrawing preview = firstSurface.CreateDrawing();
HtmlRenderHitTestReport hits = firstSurface.HitTest(120, 80, new HtmlRenderHitTestOptions {
    InteractiveOnly = true,
    MaximumResults = 8
});
firstSurface.TryMapToSource(120, 80, out HtmlRenderSourcePoint? sourcePoint);

HtmlRenderArchiveResult archive = retained.ExportArchive(new HtmlRenderArchiveOptions {
    MaximumArchiveBytes = 128 * 1024 * 1024
});
File.WriteAllBytes("screen-pages.zip", archive.Bytes);
```

Named profiles are immutable defaults. Use `WithCssMedia(...)`,
`WithLayoutSurface(...)`, and `WithPagination(...)` to form a coherent custom
combination; use `WithLayout(...)` or `WithAxes(...)` when two coupled geometry
axes must change atomically. `MatchesNamedProfile` becomes false and `Coverage` becomes
`Unqualified` when those effective axes no longer match the named profile.
The encoder and page-set admission rules still apply.

The built-in profiles are `ScreenViewport`, `ScreenFullPage`, `PrintPaged`,
`ScreenMediaPaged`, `ScreenSnapshotPaged`, and `ContinuousVector`.
`ScreenMediaPaged` applies screen CSS and performs paged reflow.
`ScreenSnapshotPaged` completes one continuous screen layout and slices it into
fixed page canvases, so it may split elements. `PrintPaged` applies print CSS and
normal fragmentation. Separate pages, one selected page, a range, and stitched
output are available now. `SourcePlacements` records every source page or slice
and its output offset in a stitched surface. Projection is cancellable and bounded
by `MaxProjectedVisuals`, `MaxPageCount`, `MaxSurfaceWidth`, and
`MaxSurfaceHeight`.

`HtmlRenderResult.OutputSurfaces` exposes immutable executable surface views for
preview drawings, output-to-source coordinate mapping, and bounded topmost-first
hit testing. Hit testing applies retained transforms and rectangle, rounded, and
path clips, and reports whether each result used transformed bounds or clip-aware
transformed bounds. It does not claim exact glyph, stroke, or arbitrary shape
paint containment.

`ExportArchive()` packages the already selected, ranged, separate, or stitched
PNG/SVG surfaces. The deterministic ZIP contains ordered `pages/page-NNNN.*`
entries plus `manifest.json`; the manifest records request axes and range values, qualification,
dimensions, encoded hashes, clipping, source placements, requested scale and
background, provider IDs, retained HTML diagnostics with source-to-target
provenance, and per-page scale, font, and codec diagnostics. Page encoding loss
participates in both the page and manifest `HasLoss` values. Container adapters
can attach their own evidence without mutating the retained result through
`WithAdditionalDiagnostics(...)`. Image byte limits continue
to come from `HtmlRenderOptions`, while `HtmlRenderArchiveOptions` independently
bounds the final ZIP and manifest. Element-aware placement remains an explicit
unsupported boundary. Inspect `HtmlRenderProfileContracts.All`
or `officeimo html capabilities --format json` for current qualification and encoder
availability.

Set `FidelityPolicy` when diagnosed fallback is not acceptable. The renderer collects the complete report, then rejects any warning, error, approximation, omission, or failure instead of returning a silently simplified scene.

```csharp
var strict = new HtmlRenderOptions {
    ViewportWidth = 720,
    FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
};

OfficeImageExportResult image = source.ExportImage(OfficeImageExportFormat.Png, strict);
```

The static contract includes normal-flow, flex, grid with column and row subgrid, deterministic stacking, basic-shape `clip-path`, paged fragmentation, named pages, running strings and elements, SVG, tagged-PDF semantics, and CSS-controlled PDF bookmarks. Browser-only execution such as JavaScript, animation timelines, live scroll state, and interactive layout is not attempted. Unsupported values that reach the declared feature handlers produce stable diagnostics; selectors outside the bounded selector subset simply do not match. Inspect `HtmlRenderCapabilityCatalog.All`, `HtmlRenderProfileContracts.All`, or the generated support matrix for the exact declared subset.

Inline SVG without intrinsic dimensions uses its resolved painted object size as the
vector viewport, including `object-fit` sizing and cropping. The element's CSS
background and the SVG's supported strokes remain in the same retained scene
for screen, print, and screen-to-page output. The
[named-page evidence](../Build/Project/Evidence/2026-09-19/html-svg-css-viewport/README.md)
records the qualified reference and the remaining browser-default typography and
body-margin differences on adjacent SVG text pages.
When a caller supplies an SVG raster codec, inline SVG reaches that fallback only
if the shared SVG safety predicate accepts its authored dimensions and resource
work. Rejected inline content receives an omission diagnostic.

Iframe `srcdoc` content is laid out as an independent replaced viewport with the
browser default 300 by 150 CSS-pixel intrinsic size, overridden by CSS or `width`
and `height` attributes. The child uses the selected screen or print media context,
keeps its own base URI and styles, clips overflow at the iframe content box, and
retains searchable text in PDF output. `MaxFrameDepth` defaults to eight and may be
lowered for stricter workloads. Layout-operation and repeated-background tile limits
apply cumulatively to the root and all rendered frame viewports. Static rendering does not fetch an iframe `src`;
the runtime application workflow supplies captured same-origin frame bodies through
an isolated render clone.

Paged tables use the same layout and retained scene for PDF, SVG, and raster output. Auto layout considers cell text, replaced images, column spans, and fixed-width visible descendants without feeding percentage widths back into intrinsic track sizing. Rowspans suppress unsafe page boundaries. `break-inside: avoid` on rows and row groups, `break-before` and `break-after` on rows, and aligned line breaks inside oversized multi-cell rows participate in pagination. `thead`/`tfoot` use their CSS table-group defaults; any row group can opt into or out of repetition with `display: table-header-group`, `table-footer-group`, or `table-row-group`. Repeated fragments retain the original table, row, and cell structure identity for tagged PDF. When an avoided or otherwise atomic row is taller than an empty page, the paginator makes bounded progress and reports `HtmlRenderForcedFragment` against the table source.

The PDF adapter uses the shared scene's resolved superscript/subscript scale and vertical offset,
including nested scripts. Logical replacement text owns its painted content once, so independent
PDF text extraction does not repeat the visible glyphs alongside their replacement. Page image
exports retain the same scene geometry and page count; font rasterization can differ between viewers.
Raster decoding preserves encoded color channels and does not automatically apply embedded ICC or
PNG gamma conversion. Normalize source colors explicitly when color-managed output is required.

The same managed path renders inline or block Presentation MathML as vector content. Fractions, roots, scripts, limits, fences, matrices, enclosures, and annotations retain logical text in the shared scene and searchable PDF output; unsupported structures use a diagnosed child-content fallback.

For documents that opt into `hyphens:auto`, supply the language-appropriate break points used by the application. The same immutable lexicon can be shared with the PDF text engine:

```csharp
options.UseTextHyphenationLexicon(new OfficeTextHyphenationLexicon(new[] {
    "ty-pog-ra-phy",
    "de-ter-min-is-tic"
}));
```

The managed renderer also honors author soft hyphens and the bounded CSS controls `hyphenate-character`, `hyphenate-limit-chars`, `hyphenate-limit-lines`, `hyphenate-limit-last: always`, and `hyphenate-limit-zone`. Inserted hyphen glyphs do not replace the source word in logical text.

## Dependency footprint

- **External:** AngleSharp and AngleSharp.Css for DOM and CSS parsing; System.Text.Encoding.CodePages for legacy web encodings.
- **OfficeIMO:** `OfficeIMO.Core`. Resource policy, layout, rendering, and diagnostics are first-party.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

`ToPng()`, `ToJpeg()`, `ToTiff()`, `ToSvg()`, and `ToWebp()` return in-memory output. `ExportImage()` and `ExportImages()` return encoded output plus dimensions and diagnostics. Format-specific save methods and the shared `ToImage()` / `ToImages()` fluent builders write to files or caller-owned streams and return the same structured evidence.

Add `OfficeIMO.Html.Pdf` for direct PDF output. `HtmlToPdfOptions` derives from `HtmlRenderOptions`, so the same configured instance can be used for PDF and all five image formats.

```csharp
using OfficeIMO.Html.Pdf;

var options = new HtmlToPdfOptions {
    Margins = HtmlRenderMargins.All(32)
};

byte[] pdf = source.ToPdfBytes(options);
byte[] png = source.ToPng(options);
string svg = source.ToSvg(options);
```

## Optional bridges

Install only the bridge required by the application:

- `OfficeIMO.Html.Rtf` for semantic HTML/RTF conversion.
- `OfficeIMO.Mhtml` for MHT/MHTML archives and embedded MIME resources.
- `OfficeIMO.Email.Image` for email-body image rendering.
- `OfficeIMO.Html.Pdf` for plain HTML/PDF conversion.
- `OfficeIMO.Mhtml.Pdf` for MHTML/PDF conversion.

The public HTML/RTF and email-image namespaces remain familiar; the package references now describe the actual cross-format capability.

## URL Policy

```csharp
var policy = HtmlUrlPolicy.CreateWebOnlyProfile();
string href = HtmlUrlPolicyEvaluator.ResolveUrl(
    "/docs/start.html",
    new Uri("https://example.com/"),
    policy);
```

## Parsing And Base URIs

```csharp
HtmlConversionDocument document = HtmlConversionDocument.Parse(
    html,
    new HtmlConversionDocumentOptions {
        BaseUri = new Uri("https://example.com/articles/")
    });
Uri? baseUri = document.BaseUri;
```

`HtmlConversionDocument.Load` and `LoadAsync` detect byte-order marks and HTML `meta charset` declarations. Pass the optional `Encoding` argument when transport metadata or application configuration is authoritative. CSS resources use their content-type charset, BOM, or `@charset`; textual data URIs use their declared `charset` and default to UTF-8.

## Trust and conversion limits

```csharp
var options = HtmlConversionDocumentOptions.CreateUntrustedProfile();
options.Limits.MaxInputCharacters = 2_000_000;
options.Limits.MaxHtmlNodes = 50_000;
options.Limits.MaxSelectorEvaluations = 1_000_000;

HtmlConversionDocument source = HtmlConversionDocument.Parse(html, options);
```

The untrusted profile is the default. It rejects local-file navigation, does not fetch external resources by itself, and applies one shared set of limits before adapters allocate native Office objects. Embedded `data:` resources remain available through the separate resource policy and are still subject to renderer or adapter byte budgets. Use `CreateTrustedProfile()` only when the caller controls the HTML and resource locations.

`HtmlConversionLimits` is the common source for parser and CSS complexity decisions. CSS byte volume, rules, declarations, inline lexical tokens, inline syntax nodes, nesting depth, selectors per rule, resolved selector characters, and selector evaluations have separate limits and stable diagnostics. Word forwards its compatibility limit properties to this object; Excel, PowerPoint, and OneNote use `HtmlImportLimits` for native artifact counts, image bytes, chart dimensions, table cells, and geometry. This keeps shared HTML decisions in `OfficeIMO.Html` while leaving format-specific constraints with the target model.

## Shared Diagnostics And Gallery Contracts

Renderer diagnostics carry `OfficeConversionLossKind` independently of severity.
Unsupported transforms and failed image resources remain loss-bearing even when
reported at informational severity. Both `FidelityPolicy.RequireNoLoss` and image
export `Policy.RequireNoLoss` reject these cases. Harmless informational diagnostics
remain accepted.

Gallery JSON uses schema version `1.1`. Artifact `evidence` records observed page
geometry, loss diagnostics, and executed checks when the producer supplies them;
it is `null` for descriptors without those observations. The `expectations` array
contains caller declarations, marked `declared-not-executed`, and is separate from
the artifact checks.

```csharp
var report = new HtmlDiagnosticReport();
report.Add("OfficeIMO.Word.Html", "HtmlCommentSkipped", "Comment skipped");

var scenario = new HtmlCapabilityGalleryScenario(
    "quarterly-report",
    "Quarterly Report",
    "Word HTML",
    "HTML import, DOCX validation, and round-trip export proof");
```

`HtmlDiagnosticReport` and the capability-gallery contracts provide a common shape for HTML converters, PDF bridges, readers, tests, and documentation generators. Existing gallery producers can keep using `HtmlCapabilityGalleryResult.AddArtifact` and `Diagnostics.Add` while assembling a scenario. Constructing an `HtmlCapabilityGalleryManifest` takes a defensive, read-only snapshot; the sequence-based result constructor also creates a frozen snapshot directly.

Native adapters use `HtmlConversionResult<TArtifact>` when callers need conversion evidence. Each diagnostic has a stable code, severity, and `LossKind` (`Approximation`, `Omission`, or `Failure`). Convenience methods still return the native artifact directly; they throw `HtmlConversionException` when required semantic content is missing. Result methods retain the artifact and diagnostics so applications can decide how to handle the failure.

## Conversion Document And Normalized HTML

```csharp
var conversion = HtmlConversionDocumentBuilder.Build(html, new HtmlConversionDocumentOptions {
    Profile = HtmlConversionProfile.Document,
    Trust = HtmlInputTrust.Untrusted,
    BaseUri = new Uri("https://example.com/reports/"),
    UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
});

string normalized = conversion.NormalizedHtml;
var resources = conversion.ResourcePlan.GetSummary(HtmlResourceKind.Image);
var styles = conversion.StyleSummary;
```

`HtmlConversionDocument` is the shared conversion contract for OfficeIMO HTML workflows. It parses once and retains one source DOM. Adapter DOMs, logical structure, computed styles, resources, and normalized output are created lazily when requested, then reused. This avoids paying for visual analysis during a semantic-only conversion and avoids retaining multiple eager document copies.

Target packages accept this shared document while keeping target-specific conversion in their owning packages. The prepared DOM can be sent to Word, Markdown, RTF, Excel, PowerPoint, OneNote, PDF, PNG, JPEG, TIFF, SVG, and WebP without inventing adapter-specific parsing rules. Excel and PowerPoint default to their versioned semantic envelopes for round trips and expose generic import mode for ordinary HTML. OneNote imports ordinary document sections directly.

Reuse the same document for analysis too: `HtmlComputedStyleEngine.Compute(conversion)` and `HtmlRoundTripScorer.Compare(source, target)` accept retained conversion documents. Their string overloads enter through the same bounded parser, so low-level helpers do not create competing trust or limit defaults.

## One report, several output formats

Prepare the report once and reuse the same source document with the format adapters:

```csharp
using OfficeIMO.Html;

var report = HtmlConversionDocument.Load("service-review.html");
File.WriteAllText("service-review-copy.html", report.SourceHtml);
```

See the owning packages for [PDF export](../OfficeIMO.Html.Pdf/README.md#html-to-pdf), [editable Word output](../OfficeIMO.Word.Html/README.md#quick-start), and [typed Excel tables](../OfficeIMO.Excel.Html/README.md#typed-values-in-ordinary-report-tables). The [HTML support matrix](../Docs/officeimo.html-support-matrix.md) describes target coverage.

Run the complete [service review example](../OfficeIMO.Examples/Converters/Html/HtmlMultiFormatReport.cs) with `--multi-format-report`. Its [HTML source](../OfficeIMO.Examples/Converters/Html/Content/Reports/service-review.html) includes grouped rows, totals, leading-zero references, dates, approval values, links, and a second table. Report content stays in ordinary HTML; the adapters own conversion behavior.

Use `--report-source <path.html>` to run another source through the same example. The output index links all exported files and every preview page, and records conversion notes. Saved PDF previews are labelled separately from editable Word and Excel layouts. The [conversion consistency tool](../Build/ConversionConsistency/README.md) checks native exports with an independent PDF rasterizer, required content, and page geometry.

## Semantic IR and target preflight

```csharp
const string semanticHtml = """
    <h1>Quarterly checklist</h1>
    <ol start="3"><li>Publish report</li></ol>
    <input name="owner" value="Finance">
    """;

HtmlConversionDocument source = HtmlConversionDocument.Parse(semanticHtml);
HtmlSemanticDocument semantics = source.SemanticDocument;
HtmlConversionPreflight excel = source.AnalyzeFor(HtmlConversionTarget.Excel);

HtmlSemanticBlock list = semantics.Sections
    .SelectMany(section => section.Blocks)
    .First(block => block.List != null);
Console.WriteLine($"{list.List!.Kind}: start={list.List.Start}, reversed={list.List.IsReversed}");

HtmlSemanticFormControl control = semantics.Sections
    .SelectMany(section => section.Blocks)
    .First(block => block.FormControl != null)
    .FormControl!;
foreach (string value in control.Values) Console.WriteLine(value);

foreach (HtmlFeaturePreflightResult feature in excel.Features) {
    Console.WriteLine($"{feature.Feature}: {feature.Outcome} ({feature.OccurrenceCount})");
}
```

`HtmlSemanticDocument` is the single interpretation of sections, rich runs, nested lists, tables, links, forms, notes, resources, computed styles, and source locations. List state retains ordered, unordered, and definition kinds, effective start and reverse direction, and each item's effective or explicit ordinal. Form controls expose specification-normalized types, one or more selected values, checked/disabled/required/readonly state, and resolved form ownership. Generic Excel, PowerPoint, and OneNote importers consume this model instead of independently deciding what the same DOM means. `AnalyzeFor(target)` uses the executable target registry to report `Supported`, `Approximated`, or `Omitted` before artifact creation. Diagnostics carry source and target provenance so applications can map a warning back to both sides of a conversion.

The generated [HTML support matrix](../Docs/officeimo.html-support-matrix.md) is checked against those executable contracts in the test suite. Run `Build/Export-HtmlSupportMatrix.ps1 -Check` to verify it or omit `-Check` to regenerate it.

Applications can inspect the same versioned contract through `HtmlRenderCapabilityCatalog.ProfileManifests` and `HtmlRenderCapabilityCatalog.All`. A profile manifest pins its admitted providers, specifications, evidence, platforms, outputs, version, and promotion state. Each capability reports the stages where the claim applies, its exact feature subset, and separate coverage, handling, maturity, required or optional provider, specification, evidence, limitation, and diagnostic fields.

```csharp
HtmlCapabilityProfileManifest profile = HtmlRenderCapabilityCatalog.GetProfile(
    HtmlCapabilityProfileIds.StaticScreenV1);

HtmlRenderCapability grid = HtmlRenderCapabilityCatalog.Get("layout-grid");
HtmlCapabilityProfileBinding support = grid.GetProfileBinding(profile.Id);

Console.WriteLine($"{profile.Id} {profile.Version}: {profile.Promotion}");
Console.WriteLine($"{grid.Stages}: {support.Coverage}/{support.Handling}");
Console.WriteLine(string.Join(", ", support.ProviderIds));
```

`Qualified` describes the listed subset and stages; it does not mean an entire HTML or CSS specification. `Fallback`, `Ignored`, and `Rejected` are handling outcomes and carry stable diagnostics when content is changed or refused. `HtmlRenderCapabilityCatalog.Validate()` checks manifest references, ordering, promotion consistency, diagnostic coverage, and release evidence for stable-default bindings.

## Resource sessions

```csharp
var renderOptions = new HtmlRenderOptions {
    ResourceResolver = async (request, cancellationToken) =>
        await ResolveApprovedResourceAsync(request, cancellationToken)
};

HtmlResourceSession session = await HtmlResourceSession.ResolveAsync(
    source.ResourceManifest,
    renderOptions,
    cancellationToken: cancellationToken);

foreach (HtmlResourceSessionEntry resource in session.Resources) {
    Console.WriteLine($"{resource.CanonicalSource} {resource.ContentType} {resource.Sha256}");
}
```

The session owns one immutable policy and limit snapshot for the operation. It deduplicates canonical requests, validates MIME types, enforces request/count/per-resource/total-byte/import-depth budgets, and records accepted resource digests. Synchronous rendering uses the configured synchronous package resolver; application/network resolution remains an explicit asynchronous boundary.

## Semantic envelope v2 and fidelity scoring

Current OfficeIMO semantic exports identify schema `2` and public-safe restoration metadata. Public-safe envelopes can be imported from untrusted input. A target-specific envelope marked `trusted-target` restores private target metadata only when the prepared input is trusted; otherwise the adapter uses the shared generic semantic path and reports the boundary. Schema `1` remains readable.

```csharp
HtmlRoundTripScore score = HtmlRoundTripScorer.Compare(sourceHtml, exportedHtml);
Console.WriteLine(score.Dimensions["structure"]);
Console.WriteLine(score.Dimensions["styles"]);

// After the caller saves, reopens, and exports a native artifact:
HtmlArtifactReloadEvidence reload = HtmlArtifactReloadEvidence.Succeeded("DOCX", reopenedDocument.ToHtml());
HtmlRoundTripScore verified = HtmlRoundTripScorer.Compare(sourceHtml, exportedHtml, reload);
Console.WriteLine(verified.Dimensions["artifact-reload"]);
```

Version 2 scores top-level fidelity dimensions independently. A dimension absent from both inputs is omitted rather than counted as a perfect result. Artifact reload evidence is intentionally caller-supplied: the score is marked verified only when a native artifact was successfully reopened and its re-exported HTML was compared with the original source.

Capability-gallery JSON and Markdown manifests retain the score schema version, top-level dimensions, artifact kind, and reload-verification flag as durable review evidence.

Conversion profile and trust are separate decisions. A `Document` or `HighFidelityPrint` profile does not make external resources trusted. Leave `Trust` as `Untrusted` for user-supplied HTML; set it to `Trusted` only when the caller controls the document and resource locations.

Normalized HTML output is policy-aware: hyperlink and resource URLs are evaluated separately, URL-bearing attributes are resolved against the configured base URI, disallowed URLs are removed, boolean attributes are normalized, event-handler attributes are stripped by default, and non-document executable elements are skipped. External bytes are loaded only through a caller-supplied bounded resolver. Normalized output is intended for clean review, gallery proof, and downstream adapter input selection, not as a browser sandbox.

## Image Sources

```csharp
string source = HtmlImageSourceResolver.ResolveImageSource(
    imageElement,
    baseUri,
    HtmlUrlPolicy.CreateOfficeIMOProfile());
```

## Image Data URIs

```csharp
if (HtmlImageDataUri.TryParse(source, out var dataUri) && dataUri.IsBase64) {
    byte[] bytes = dataUri.DecodeBytes();
    string extension = dataUri.FileExtension;
}
```

Use `HtmlDataUri` when the payload is textual. Its `DecodeText()` method honors a declared `charset` and uses UTF-8 only when the data URI does not declare one.

## Content provenance

`HtmlProvenance.Inspect(html)` reports embedded `<script type="application/c2pa">` carriers, external `<link rel="c2pa-manifest">` references, and provenance inside supported embedded image data URIs. `HtmlProvenance.Remove(html)` removes only selected, structurally valid carriers by default. Inspection and removal never fetch external resources. Optional cryptographic C2PA verification remains in `OfficeIMO.Security`.

## Concealed-content inspection and cleanup

`HtmlContentSafety.Inspect(html)` evaluates the bounded CSS cascade and reports hidden, transparent, tiny, zero-size, clipped, off-canvas, and low-contrast text together with comments, scripts/styles, templates, metadata, alternative text, ARIA labels, and hidden form values. `HtmlContentSafety.RemoveSelected(...)` removes only reviewed current findings—including exact Unicode ranges inside machine-only text nodes—and reinspects the serialized HTML. An empty file selection preserves the source bytes and encoding. No script runs and no external resource is fetched.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 1 | 0 | 0 | 0 | 0 | 0 |
| Read | 1 | 0 | 0 | 0 | 0 | 0 |
| Edit | 1 | 0 | 0 | 0 | 0 | 0 |
| Preserve | 0 | 1 | 0 | 0 | 0 | 0 |
| Inspect | 2 | 0 | 0 | 0 | 0 | 0 |
| Validate | 1 | 0 | 0 | 0 | 0 | 0 |
| Remove | 1 | 0 | 0 | 0 | 0 | 0 |
| Export | 5 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Html` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
