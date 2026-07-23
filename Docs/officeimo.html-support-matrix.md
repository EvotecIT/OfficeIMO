# OfficeIMO HTML support matrix

This file is generated from `HtmlConversionProfileContracts`, `HtmlTargetCapabilityContracts`, and `HtmlDiagnosticCatalog`. Profile and target entries are tested contracts; diagnostic entries describe bounded fallbacks, policy decisions, and safety limits.

## Conversion profiles

### Semantic

- Intended use: Editable office documents, clean HTML export, accessible reports, and deterministic round-trips.
- Fidelity goal: Preserve document structure and meaning first; simplify browser-only layout where needed.
- Supported HTML: headings, paragraphs, lists, tables, links, images, forms-as-document-fields, language-and-direction-metadata
- Supported CSS: inline-style, document-style-rules, text-formatting, table-borders, spacing, colors, direction
- Resource guarantees: URL policy enforcement, responsive image candidate selection, resource manifest reporting
- Diagnostic guarantees: unsupported CSS warnings, resource policy diagnostics, accessibility advisories, round-trip score evidence

### Document

- Intended use: Business documents, invoices, contracts, generated reports, and HTML intended to become DOCX/PDF artifacts.
- Fidelity goal: Balance editability and visual fidelity for common document layouts.
- Supported HTML: semantic sections, headers-and-footers, tables-with-spans, captions, figures, form-controls, comments
- Supported CSS: cascade snapshot, selector matching, font-and-color inheritance, table layout hints, print-friendly spacing
- Resource guarantees: bounded downloads, content-type validation, byte budgets, base URI resolution, blocked resource reporting
- Diagnostic guarantees: diagnostic catalog lookup, shared report aggregation, gallery manifest diagnostics

### High Fidelity Print

- Intended use: Print/PDF lanes, visual review, and workflows where page appearance matters more than editable structure.
- Fidelity goal: Expose layout preservation as an explicit ambition while reporting fallbacks and unsupported browser features.
- Supported HTML: print sections, positioning hints, backgrounds, page-breaks, complex tables, media-heavy content
- Supported CSS: computed-style capture, media intent metadata, layout-affecting declarations, resource dependency graph
- Resource guarantees: complete resource inventory, policy outcome per resource, external dependency diagnostics
- Diagnostic guarantees: fidelity score, layout fallback diagnostics, unsupported high-fidelity feature diagnostics

### Positioned Review

- Intended use: PDF readback, page previews, and diagnostic review lanes where source geometry needs to remain inspectable in HTML.
- Fidelity goal: Preserve review geometry and source anchors while clearly avoiding editable document reconstruction claims.
- Supported HTML: page wrappers, positioned text blocks, positioned images, link frames, form field frames, source anchors
- Supported CSS: absolute positioning, page dimensions, safe overlay styles, review-only visual hints
- Resource guarantees: resource inventory, safe link handling, image placeholder or embedding policy, source coordinate reporting
- Diagnostic guarantees: geometry simplification diagnostics, unsafe link diagnostics, missing resource diagnostics, no-editable-reconstruction boundary

## Target adapter API contracts

| Target | Package | Artifact | HTML import | Result contract | Reverse HTML | Reverse result | Profiles | I/O and async boundary |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Word | `OfficeIMO.Word.Html` | WordDocument | `HtmlConversionDocument.ToWordDocument` | `HtmlToWordResult` | `WordDocument.ToHtml` | `HtmlTextConversionResult` | OfficeIMO, UntrustedHtml, TrustedDocument | Load or LoadAsync the shared document; synchronous and asynchronous Word import are available; path and stream HTML saves preserve caller-owned streams. |
| Excel | `OfficeIMO.Excel.Html` | ExcelDocument | `HtmlConversionDocument.ToExcelDocument` | `HtmlToExcelResult` | `ExcelDocument.ToHtml` | `HtmlTextConversionResult` | Semantic, Auto, Generic, SemanticTables, VisualReview | Load or LoadAsync the shared document, then import synchronously; path and stream HTML saves have synchronous and asynchronous forms. |
| PowerPoint | `OfficeIMO.PowerPoint.Html` | PowerPointPresentation | `HtmlConversionDocument.ToPowerPointPresentation` | `HtmlToPowerPointResult` | `PowerPointPresentation.ToHtml` | `PowerPointToHtmlResult` | Semantic, Auto, Generic, SemanticSlides, VisualReview | Load or LoadAsync the shared document, then import synchronously; path and stream HTML saves have synchronous and asynchronous forms. |
| OneNote | `OfficeIMO.OneNote.Html` | OneNoteSection / OneNoteNotebook | `HtmlConversionDocument.ToOneNoteSection` | `HtmlToOneNoteSectionResult / HtmlToOneNoteNotebookResult` | `OneNoteSection.ToHtmlDocument` | `HtmlTextConversionResult` | GenericSemantic, SemanticHtml, VisualHtml | Load or LoadAsync the shared document, then import synchronously; semantic and visual HTML exports support path, stream, and asynchronous saves. |
| Markdown | `OfficeIMO.Markdown.Html` | MarkdownDoc / Markdown text | `HtmlConversionDocument.ToMarkdownDocument` | `HtmlToMarkdownResult` | `MarkdownDoc.ToHtmlDocument` | — | OfficeIMO, GitHubFlavoredMarkdown, CommonMark, Portable | Load or LoadAsync the shared document; Markdown conversion is synchronous and path or stream saves have synchronous and asynchronous forms. |
| Rtf | `OfficeIMO.Html / OfficeIMO.Rtf` | RtfDocument | `HtmlConversionDocument.ToRtfDocument` | `HtmlToRtfResult` | `RtfDocument.ToHtml` | `RtfToHtmlResult` | OfficeIMO, UntrustedHtml, WebSafe, RoundTrip | Load or LoadAsync the shared document; semantic conversion is synchronous and RTF/HTML path or stream saves have synchronous and asynchronous forms. |
| Pdf | `OfficeIMO.Html.Pdf` | PdfDocument / PDF bytes | `HtmlConversionDocument.ToPdfDocument` | `PdfDocumentConversionResult` | — | — | Continuous, Paged, Screen, Print | Synchronous and asynchronous conversion resolve through the shared render resource pipeline; byte, document, path, and stream outputs are available. |
| Image | `OfficeIMO.Html` | PNG / JPEG / TIFF / SVG / WebP | `HtmlConversionDocument.ToPng / ToSvg / ToJpeg / ToTiff / ToWebp` | `OfficeImageExportResult` | — | — | Continuous, Paged, Screen, Print | Synchronous and asynchronous render APIs share one resource pipeline; in-memory, path, stream, and paged fluent outputs are available. |
| Reader | `OfficeIMO.Reader.Html` | OfficeDocumentReadResult / ReaderChunk | `OfficeDocumentReader.ReadDocument (after AddHtmlHandler)` | `OfficeDocumentReadResult` | — | — | Default, Portable, UntrustedHtml, Mhtml | Registered Reader handlers support path and caller-owned stream input with cancellation and asynchronous document reads. |

## Target semantic capability contracts

| Target | Supported | Approximated | Unsupported |
| --- | --- | --- | --- |
| Word | Metadata, Sections, Headings, Paragraphs, RichText, Links, Lists, Tables, Images, Forms, Notes, Comments, Annotations, Css, Resources | Media, Geometry, PagedLayout | Formulas, Charts |
| Excel | Metadata, Sections, Tables, Images, Comments, Annotations, Formulas, Charts, Geometry, Resources | Headings, Paragraphs, RichText, Links, Lists, Css | Media, Forms, Notes, PagedLayout |
| PowerPoint | Metadata, Sections, Headings, Paragraphs, Tables, Images, Notes, Charts, Geometry, Resources | RichText, Links, Lists, Css | Media, Forms, Comments, Annotations, Formulas, PagedLayout |
| OneNote | Metadata, Sections, Headings, Paragraphs, RichText, Links, Lists, Tables, Images, Notes, Resources | Geometry, Css | Media, Forms, Comments, Annotations, Formulas, Charts, PagedLayout |
| Markdown | Sections, Headings, Paragraphs, RichText, Links, Lists, Tables, Images, Notes, Annotations, Resources | Metadata, Media, Forms, Comments, Css | Formulas, Charts, Geometry, PagedLayout |
| Rtf | Metadata, Sections, Headings, Paragraphs, RichText, Links, Lists, Tables, Images, Forms, Notes, Comments, Annotations, Resources | Media, Geometry, Css, PagedLayout | Formulas, Charts |
| Pdf | Metadata, Sections, Headings, Paragraphs, RichText, Links, Lists, Tables, Images, Geometry, Css, Resources, PagedLayout | Media, Forms, Notes, Comments, Annotations, Formulas, Charts | None |
| Image | Images, Geometry, Css, Resources, PagedLayout | Metadata, Sections, Headings, Paragraphs, RichText, Links, Lists, Tables, Media, Forms, Notes, Comments, Annotations, Formulas, Charts | None |
| Reader | Metadata, Sections, Headings, Paragraphs, RichText, Links, Lists, Tables, Images, Media, Forms, Notes, Resources | Comments, Annotations, Formulas, Charts, Geometry, Css, PagedLayout | None |

## Diagnostic boundaries

| Category | Code | Severity | Meaning | Remediation |
| --- | --- | --- | --- | --- |
| ContentSimplification | `HtmlCommentSkipped` | Info | An HTML comment was omitted from generated document content. | Enable HTML comment import when comments are part of the expected document contract, or keep comments as source-only metadata. |
| Conversion | `ArtifactCreationFailed` | Error | The target artifact could not be constructed. | Inspect the diagnostic detail and validate the source and target-specific constraints. |
| ConversionFidelity | `ContentApproximated` | Warning | Content was represented using a documented approximation. | Use semantic HTML emitted by the matching adapter when exact round-trip fidelity is required. |
| ConversionFidelity | `ContentOmitted` | Warning | Content could not be represented by the target and was omitted. | Simplify the source construct or choose a target that supports it. |
| CssFidelity | `HtmlRenderStylesheetEncodingUnsupported` | Warning | A resolved stylesheet could not be decoded as supported CSS text. | Return UTF-8 CSS or UTF-16 CSS with a byte-order mark. |
| CssFidelity | `HtmlRenderStylesheetImportCycle` | Warning | A recursive stylesheet import cycle was suppressed. | Remove the cyclic @import relationship. |
| CssFidelity | `HtmlRenderStylesheetUrlResourcesPending` | Warning | An external stylesheet contains URL resources that are not active in the current paint model. | Inline those assets or use currently supported image and font resource paths until CSS URL painting is enabled. |
| CssFidelity | `MediaFilterFailed` | Warning | An active stylesheet could not be filtered safely for the selected media. | Correct invalid CSS or simplify nested media rules. |
| CssFidelity | `UnsupportedCssDeclaration` | Warning | A CSS declaration could not be mapped to the target document model. | Prefer document-friendly CSS or route visual-first workloads through the high-fidelity print profile. |
| ImageFidelity | `HtmlRenderSvgContentUnsupported` | Warning | SVG content could not be represented completely by the bounded shared vector scene. | Use supported primitives or paths, bounded local shape/group references, local object-bounding-box linear/radial paint servers, positioned tspan text, and affine transform attributes while broader SVG scene support is being completed. |
| LayoutFidelity | `HtmlRenderEmptyTable` | Info | A table contained no renderable rows or cells. | Add table rows and cells or remove the empty table. |
| LayoutFidelity | `HtmlRenderFlexLayoutPending` | Warning | A flex formatting case currently falls back to normal flow. | Use row or column flex directions with static or relatively positioned items; absolute, fixed, sticky, and nested generated flex formatting contexts remain pending. |
| LayoutFidelity | `HtmlRenderFlexValueUnsupported` | Warning | A flex property value used a deterministic fallback. | Use supported length or percentage bases, flex-start/start, flex-end/end, center, stretch, space-between, space-around, or space-evenly. |
| LayoutFidelity | `HtmlRenderFloatValueUnsupported` | Warning | A float or clear property value used a deterministic fallback. | Use none, left, right, inline-start, or inline-end for float; use those values or both for clear. |
| LayoutFidelity | `HtmlRenderGeneratedContentUnsupported` | Warning | A CSS generated-content expression was omitted. | Use quoted text, attr(), counter(), or counters() with decimal, alphabetic, or Roman counter styles until additional generated-content forms are enabled. |
| LayoutFidelity | `HtmlRenderGeneratedCounterUnsupported` | Warning | A CSS counter declaration was ignored. | Use counter-reset, counter-set, and counter-increment with counter names and optional integer values; reversed counters remain unsupported. |
| LayoutFidelity | `HtmlRenderGridLayoutPending` | Warning | A grid formatting case currently falls back to normal flow. | Use block or inline grid with static or relatively positioned items; absolute/fixed/sticky items and nested generated grid formatting contexts remain pending. |
| LayoutFidelity | `HtmlRenderGridValueUnsupported` | Warning | A grid property value used a deterministic fallback. | Use fixed, percentage, auto, fr, repeat(), or minmax() tracks with numeric lines and span values for the active grid contract. |
| LayoutFidelity | `HtmlRenderMultiColumnValueUnsupported` | Warning | A multi-column property value used a deterministic fallback. | Use supported positive count/width values, balance or auto fill, none or all span, and solid, dashed, dotted, or double rules. |
| LayoutFidelity | `HtmlRenderOverflowClipMarginValueUnsupported` | Warning | An overflow-clip-margin value used its initial fallback. | Use an optional content-box, padding-box, or border-box keyword and one non-negative absolute length. |
| LayoutFidelity | `HtmlRenderOverflowScrollSnapshot` | Info | A scrollable overflow box was captured at its initial static scroll position. | Use visible overflow when all content must remain visible, or hidden/clip when a static clipped export is intended. |
| LayoutFidelity | `HtmlRenderOverflowValueUnsupported` | Warning | An overflow property value used the visible fallback. | Use visible, hidden, clip, auto, or scroll for overflow, overflow-x, and overflow-y. |
| LayoutFidelity | `HtmlRenderPositionInsetUnsupported` | Warning | A positioned inset could not be resolved and used its documented fallback. | Use a supported CSS length or percentage with a definite containing-block dimension. |
| LayoutFidelity | `HtmlRenderPositionStaticAnchorFallback` | Warning | An automatic positioned inset could not use the element's hypothetical normal-flow anchor. | Use an explicit inset or place the positioned element in a supported block, flex, or grid static-position context. |
| LayoutFidelity | `HtmlRenderPositionStickyStatic` | Info | A sticky-positioned element was captured at its stable static document position. | Use fixed positioning for repeated page overlays; sticky scroll-state changes are not meaningful in a static document snapshot. |
| LayoutFidelity | `HtmlRenderPositioningModeUnsupported` | Warning | A CSS positioning mode currently falls back to normal flow. | Use static, relative, absolute, fixed, or sticky positioning without unsupported containing-block features. |
| LayoutFidelity | `HtmlRenderReplacedElementValueUnsupported` | Warning | A replaced-element sizing or object-placement value used a deterministic fallback. | Use a positive aspect ratio, fill, contain, cover, none, or scale-down object fitting, and a supported keyword, length, or percentage object position. |
| LayoutFidelity | `HtmlRenderTableValueUnsupported` | Warning | A table formatting property used its documented fallback. | Use top or bottom for caption-side, auto or fixed for table-layout, separate or collapse for border-collapse, and one or two non-negative absolute lengths for border-spacing. |
| PagedMedia | `HtmlRenderForcedFragment` | Warning | Content had no safe break opportunity within one page. | Add break opportunities or reduce the size of the unbreakable content. |
| PagedMedia | `HtmlRenderPageMarginContentUnsupported` | Warning | A page-margin generated-content expression could not be represented. | Use quoted text with counter(page) or counter(pages) until richer generated content is enabled. |
| PagedMedia | `HtmlRenderPageMarginPositionUnsupported` | Warning | A page-margin position is not recognized by the direct renderer. | Use one of the standard CSS top, bottom, left, right, or corner page-margin box names. |
| PagedMedia | `HtmlRenderPagePseudoGeometryPending` | Warning | A pseudo-page size or margin declaration requires page-by-page body reflow. | Keep body geometry in the generic @page rule until per-page reflow is enabled; pseudo-page margin content is still applied. |
| PagedMedia | `HtmlRenderPageSelectorPending` | Warning | A complex page selector could not be applied per page. | Use a generic, named, :first, :left, or :right @page selector, optionally combining one name with one supported pseudo-page. |
| PagedMedia | `HtmlRenderPageSizeUnsupported` | Warning | An @page size declaration could not be mapped. | Use a supported named size or two absolute physical lengths. |
| PagedMedia | `HtmlRenderTableFooterRepeatSuppressed` | Warning | A repeated table footer left no safe body-row break on an empty page. | Reduce the footer or row height, increase the page content area, or allow the body row to move without a repeated footer. |
| PagedMedia | `HtmlRenderTableHeaderRepeatSuppressed` | Warning | A repeated table header left no safe body-row break on an empty page. | Reduce the header or row height, increase the page content area, or allow the body row to move without a repeated header. |
| PagedMedia | `HtmlRenderVisualFragmentUnsupported` | Warning | A visual could not cross a forced page boundary safely. | Resize the visual or add a safe break before it. |
| PaintFidelity | `HtmlRenderBackgroundImageRepeatUnsupported` | Warning | A CSS background-repeat value used a single-image fallback. | Use repeat, no-repeat, space, round, repeat-x, repeat-y, or a supported two-axis combination. |
| PaintFidelity | `HtmlRenderBackgroundImageValueUnsupported` | Warning | A CSS background image value used a deterministic supported fallback or was omitted. | Use URL backgrounds, opaque linear gradients, or opaque radial circles and ellipses with keyword, length, or percentage geometry and percentage or implicit color stops until additional image functions are enabled. |
| PaintFidelity | `HtmlRenderBorderPaintValueUnsupported` | Warning | A CSS border paint declaration used no-border fallback. | Use one to four non-negative widths, supported colors, and solid, dashed, dotted, double, none, or hidden side styles. |
| PaintFidelity | `HtmlRenderBorderRadiusValueUnsupported` | Warning | A CSS border radius used square-corner fallback. | Use one to four non-negative length or percentage radii, an optional slash-separated vertical axis, or valid one- or two-axis corner longhands. |
| PaintFidelity | `HtmlRenderBoxShadowValueUnsupported` | Warning | A CSS box shadow was omitted. | Use comma-separated inset or outer shadows with two offsets, an optional non-negative blur radius, an optional signed spread radius, and a supported color. |
| PaintFidelity | `HtmlRenderInlinePaintEffectUnsupported` | Warning | A paint effect on a non-atomic inline box used normal inline paint. | Use a block, inline-block, inline-flex, or inline-grid wrapper when an isolated transform or opacity group is required. |
| PaintFidelity | `HtmlRenderOpacityValueUnsupported` | Warning | A CSS opacity value used the opaque fallback. | Use a finite number or percentage; values outside the visible range are clamped. |
| PaintFidelity | `HtmlRenderOutlinePaintValueUnsupported` | Warning | A CSS outline paint declaration was omitted. | Use one non-negative width, supported color, signed offset, and solid, dashed, dotted, double, none, or hidden style. |
| PaintFidelity | `HtmlRenderPositionZIndexPending` | Warning | A positioned element's z-index is not active in the current stacking model. | Keep source order for the current contract until stacking contexts are enabled. |
| PaintFidelity | `HtmlRenderTransformValueUnsupported` | Warning | A CSS transform or transform-origin value used the identity fallback. | Use supported 2D matrix, translate, scale, rotate, or skew functions and a two-dimensional transform origin. |
| ResourceLimit | `HtmlRenderBackgroundImageLayerLimit` | Warning | CSS background-image layers beyond the configured per-element limit were omitted. | Increase MaxBackgroundImageLayers only for trusted documents or reduce the number of declared background layers. |
| ResourceLimit | `HtmlRenderBackgroundImageTileLimitExceeded` | Error | Repeated CSS background images exceeded the configured operation-wide tile limit. | Increase MaxBackgroundImageTiles only for trusted documents or use a larger background tile. |
| ResourceLimit | `HtmlRenderBoxShadowLayerLimit` | Warning | CSS box-shadow layers beyond the configured per-element limit were omitted. | Increase MaxBoxShadowLayers only for trusted documents or reduce the number of declared shadows. |
| ResourceLimit | `HtmlRenderGradientStopLimitExceeded` | Error | CSS gradients exceeded the configured color-stop limit. | Increase MaxGradientStops only for trusted documents or reduce the number of gradient color stops. |
| ResourcePolicy | `FontResourceRejectedByPolicy` | Warning | A font dependency was rejected before loading because its URI is not allowed by policy. | Use packaged fonts from trusted locations or allow approved font hosts in the URL policy. |
| ResourcePolicy | `HtmlRenderExternalImagePending` | Warning | An external image requires asynchronous resource resolution. | Use RenderAsync with an application-supplied resource resolver or embed the image as a data URI. |
| ResourcePolicy | `HtmlRenderExternalStylesheetPending` | Warning | An external stylesheet requires asynchronous resource resolution. | Use RenderAsync with an application-supplied resource resolver or place trusted CSS in a style element. |
| ResourcePolicy | `HtmlRenderResourceByteLimitExceeded` | Warning | A resource exceeded the configured per-resource byte limit. | Reduce the resource or raise the explicit limit for trusted input. |
| ResourcePolicy | `HtmlRenderResourceContentTypeRejected` | Warning | A resolver returned an incompatible media type. | Return bytes whose declared media type matches the requested image or stylesheet kind. |
| ResourcePolicy | `HtmlRenderResourceCountLimitExceeded` | Error | Resolved resources exceeded the configured count limit. | Reduce the resource graph or raise the explicit count limit for trusted input. |
| ResourcePolicy | `HtmlRenderResourceLoadFailed` | Warning | The configured resource resolver failed. | Inspect the resolver boundary and return null for intentionally unavailable resources. |
| ResourcePolicy | `HtmlRenderResourceRequestLimitExceeded` | Error | Resource resolver invocations exceeded the configured request limit. | Reduce broken or unavailable references, or raise the explicit request limit for trusted input. |
| ResourcePolicy | `HtmlRenderResourceTimeout` | Warning | Resource resolution exceeded its timeout. | Reduce resolver latency or raise the bounded timeout for trusted workloads. |
| ResourcePolicy | `HtmlRenderResourceUnavailable` | Warning | The configured resolver returned no resource. | Provide the resource or accept the diagnosed placeholder. |
| ResourcePolicy | `HtmlRenderResourceUriInvalid` | Warning | A resource URI could not be represented as an absolute URI. | Provide a valid base URI and resource reference. |
| ResourcePolicy | `HtmlRenderStylesheetImportDepthExceeded` | Error | Stylesheet imports exceeded the configured recursion depth. | Flatten the import graph or raise the explicit depth limit for trusted input. |
| ResourcePolicy | `HtmlRenderTotalResourceByteLimitExceeded` | Error | Resolved resources exceeded the total byte budget. | Reduce resource volume or raise the explicit total limit for trusted input. |
| ResourcePolicy | `HtmlResourceRejectedByPolicy` | Warning | A resource dependency was rejected before loading because its URI is not allowed by policy. | Adjust the URL policy only for trusted sources, or package the dependency with the HTML input. |
| ResourcePolicy | `HyperlinkRejectedByPolicy` | Warning | A hyperlink target was rejected because its URI is not allowed by policy. | Use http, https, mailto, or a caller-approved scheme instead of script or local file targets. |
| ResourcePolicy | `ImageResourceRejectedByPolicy` | Warning | An image candidate was rejected before loading because its URI is not allowed by policy. | Allow the URI scheme or host for trusted inputs, embed the image as data URI, or provide a local resource resolver. |
| ResourcePolicy | `MediaResourceRejectedByPolicy` | Warning | A media dependency was rejected before loading because its URI is not allowed by policy. | Allow trusted media hosts explicitly, package approved media with the input, or provide a local resource resolver. |
| ResourcePolicy | `ResourceDecodeFailed` | Warning | An embedded resource could not be decoded. | Provide a valid, bounded data URI or use an approved resource resolver. |
| ResourcePolicy | `ResourceTypeUnsupported` | Warning | A resource media type is unsupported by the target adapter. | Convert the resource to a media type supported by the target adapter. |
| ResourcePolicy | `ScriptResourceRejectedByPolicy` | Warning | A script dependency was rejected before loading because its URI is not allowed by policy. | Use caller-provided script handling for trusted automation scenarios, or remove script dependencies from document-oriented HTML inputs. |
| ResourcePolicy | `StylesheetResourceRejectedByPolicy` | Warning | A stylesheet was rejected before loading because its URI is not allowed by policy. | Use caller-provided stylesheet contents for untrusted HTML, or allow the stylesheet scheme and host for trusted documents. |
| Safety | `CssDeclarationLimitExceeded` | Error | CSS declarations exceeded the shared complexity budget. | Reduce CSS declaration volume or raise MaxCssDeclarations only for trusted input. |
| Safety | `CssRuleLimitExceeded` | Error | Active CSS rules exceeded the shared complexity budget. | Reduce CSS rule volume or raise MaxCssRules only for trusted input. |
| Safety | `CssSelectorEvaluationLimitExceeded` | Error | Selector matching exceeded the shared evaluation budget. | Simplify selectors or raise MaxSelectorEvaluations only for trusted input. |
| Safety | `CssSizeLimitExceeded` | Error | One embedded stylesheet exceeded the shared byte budget. | Reduce the stylesheet or raise MaxCssBytes only for trusted input. |
| Safety | `CssTotalSizeLimitExceeded` | Error | Embedded stylesheets exceeded the operation-wide byte budget. | Reduce embedded CSS or raise MaxTotalCssBytes only for trusted input. |
| Safety | `HtmlDepthLimitExceeded` | Error | HTML nesting exceeded the shared pre-analysis depth budget. | Reduce nesting or raise MaxHtmlDepth only for trusted input. |
| Safety | `HtmlNodeLimitExceeded` | Error | The parsed HTML document exceeded the configured DOM node budget before styling or layout. | Reduce repeated markup or raise MaxHtmlNodes only for trusted input. |
| Safety | `HtmlRenderCollapsedTableBorderLimitExceeded` | Error | Collapsed table-border resolution exceeded the configured segment budget. | Reduce table border complexity or raise MaxCollapsedTableBorderSegments only for trusted input. |
| Safety | `HtmlRenderDepthLimitExceeded` | Error | HTML layout exceeded the configured nesting-depth limit. | Reduce nesting or raise the explicit layout-depth limit for trusted input. |
| Safety | `HtmlRenderGridTrackLimitExceeded` | Error | Grid track expansion exceeded the configured limit. | Reduce explicit, implicit, or repeat()-generated tracks, or raise MaxGridTracks for trusted input. |
| Safety | `HtmlRenderInputCharacterLimitExceeded` | Error | HTML source text exceeded the configured character budget before parsing. | Reduce or split the document, move large payloads behind a bounded resolver, or raise MaxInputCharacters only for trusted input. |
| Safety | `HtmlRenderLayoutOperationLimitExceeded` | Error | HTML layout exceeded the configured operation budget. | Simplify the layout or raise MaxLayoutOperations only for trusted input. |
| Safety | `HtmlRenderMultiColumnLimitExceeded` | Error | Multi-column generation exceeded the configured column limit. | Reduce column-count, increase column-width or available height, or raise MaxColumnCount only for trusted documents. |
| Safety | `HtmlRenderTableLimitExceeded` | Error | A table exceeded the configured row or column limit. | Reduce table dimensions or raise MaxTableRows or MaxTableColumns only for trusted input. |
| Safety | `SemanticMetadataLimitExceeded` | Error | A semantic metadata field exceeded its shared limit. | Reduce oversized metadata or raise the explicit metadata limit only for trusted input. |
| Safety | `TargetLimitExceeded` | Error | Input exceeded a target-native or shared import limit. | Reduce or split the document, or raise the explicit limit only for trusted input. |
| SemanticImport | `SemanticBlockMissing` | Warning | An expected semantic content block was not present. | Regenerate the semantic HTML from the source adapter or supply the missing block. |
| SemanticImport | `SemanticContentMissing` | Error | The expected format-specific semantic HTML envelope was not present. | Use generic import mode for ordinary HTML or export with the matching OfficeIMO semantic profile. |
| SemanticImport | `SemanticRestorationTrustRequired` | Warning | The envelope requested target-specific restoration but the caller marked the HTML as untrusted. | Parse the document with a trusted profile only after authenticating its source, or use public-safe semantic restoration. |
| SemanticImport | `SemanticSchemaLegacy` | Info | A legacy semantic envelope without a schema version was accepted. | Re-export the content to add the current semantic schema version. |
| SemanticImport | `SemanticSchemaUnsupported` | Error | The semantic source or schema version is unsupported. | Use a matching OfficeIMO adapter and supported semantic schema version. |
| SemanticImport | `SemanticValueInvalid` | Warning | A semantic value could not be parsed safely. | Use finite, target-valid values in OfficeIMO semantic metadata. |
| TableFidelity | `TableSpanInvalid` | Warning | An invalid or overlapping HTML table span was normalized. | Use positive, non-overlapping rowspan and colspan values. |
| Typography | `HtmlRenderBidiLayoutUnsupported` | Warning | Explicit Unicode bidi controls require an embedding or isolate stage that is not active yet. | Prefer semantic dir attributes for supported simple RTL layout until explicit bidi embedding and isolate controls are enabled. |
| Typography | `HtmlRenderComplexTextShapingUnsupported` | Warning | A joining alphabet outside the bounded core-Arabic contextual shaper used scalar glyphs. | Use the PDF shaping-provider seam for host-managed glyph shaping, or treat broader joining alphabets and OpenType mark positioning as unsupported until the shared managed stage expands. |
| Typography | `HtmlRenderFontDataUriInvalid` | Warning | A font data URI could not be decoded. | Provide a valid percent-encoded or base64 font data URI. |
| Typography | `HtmlRenderFontFaceInvalid` | Warning | An @font-face rule has no usable family descriptor. | Provide a font-family descriptor and at least one usable src entry. |
| Typography | `HtmlRenderFontFaceUnavailable` | Warning | No source from an @font-face rule was available. | Use an allowed data URI or resolve the external font through RenderAsync. |
| Typography | `HtmlRenderFontFormatUnsupported` | Warning | A font source is not a supported TrueType glyf-outline font. | Provide a TTF or TrueType-flavored OpenType face; WOFF, WOFF2, and CFF outlines require pre-conversion. |
