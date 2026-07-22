namespace OfficeIMO.Html;

/// <summary>
/// Catalog of stable OfficeIMO HTML diagnostics and support remediation text.
/// </summary>
public static class HtmlDiagnosticCatalog {
    private static readonly IReadOnlyDictionary<string, HtmlDiagnosticDefinition> Definitions = new Dictionary<string, HtmlDiagnosticDefinition>(StringComparer.OrdinalIgnoreCase) {
        [HtmlConversionDiagnosticCodes.SemanticContentMissing] = ConversionDefinition(HtmlConversionDiagnosticCodes.SemanticContentMissing, "SemanticImport", HtmlDiagnosticSeverity.Error, "The expected format-specific semantic HTML envelope was not present.", "Use generic import mode for ordinary HTML or export with the matching OfficeIMO semantic profile."),
        [HtmlConversionDiagnosticCodes.SemanticBlockMissing] = ConversionDefinition(HtmlConversionDiagnosticCodes.SemanticBlockMissing, "SemanticImport", HtmlDiagnosticSeverity.Warning, "An expected semantic content block was not present.", "Regenerate the semantic HTML from the source adapter or supply the missing block."),
        [HtmlConversionDiagnosticCodes.SemanticValueInvalid] = ConversionDefinition(HtmlConversionDiagnosticCodes.SemanticValueInvalid, "SemanticImport", HtmlDiagnosticSeverity.Warning, "A semantic value could not be parsed safely.", "Use finite, target-valid values in OfficeIMO semantic metadata."),
        [HtmlConversionDiagnosticCodes.ResourceDecodeFailed] = ConversionDefinition(HtmlConversionDiagnosticCodes.ResourceDecodeFailed, "ResourcePolicy", HtmlDiagnosticSeverity.Warning, "An embedded resource could not be decoded.", "Provide a valid, bounded data URI or use an approved resource resolver."),
        [HtmlConversionDiagnosticCodes.ResourceTypeUnsupported] = ConversionDefinition(HtmlConversionDiagnosticCodes.ResourceTypeUnsupported, "ResourcePolicy", HtmlDiagnosticSeverity.Warning, "A resource media type is unsupported by the target adapter.", "Convert the resource to a media type supported by the target adapter."),
        [HtmlConversionDiagnosticCodes.ContentApproximated] = ConversionDefinition(HtmlConversionDiagnosticCodes.ContentApproximated, "ConversionFidelity", HtmlDiagnosticSeverity.Warning, "Content was represented using a documented approximation.", "Use semantic HTML emitted by the matching adapter when exact round-trip fidelity is required."),
        [HtmlConversionDiagnosticCodes.ContentOmitted] = ConversionDefinition(HtmlConversionDiagnosticCodes.ContentOmitted, "ConversionFidelity", HtmlDiagnosticSeverity.Warning, "Content could not be represented by the target and was omitted.", "Simplify the source construct or choose a target that supports it."),
        [HtmlConversionDiagnosticCodes.ArtifactCreationFailed] = ConversionDefinition(HtmlConversionDiagnosticCodes.ArtifactCreationFailed, "Conversion", HtmlDiagnosticSeverity.Error, "The target artifact could not be constructed.", "Inspect the diagnostic detail and validate the source and target-specific constraints."),
        [HtmlConversionDiagnosticCodes.TargetLimitExceeded] = ConversionDefinition(HtmlConversionDiagnosticCodes.TargetLimitExceeded, "Safety", HtmlDiagnosticSeverity.Error, "Input exceeded a target-native or shared import limit.", "Reduce or split the document, or raise the explicit limit only for trusted input."),
        [HtmlConversionDiagnosticCodes.TableSpanInvalid] = ConversionDefinition(HtmlConversionDiagnosticCodes.TableSpanInvalid, "TableFidelity", HtmlDiagnosticSeverity.Warning, "An invalid or overlapping HTML table span was normalized.", "Use positive, non-overlapping rowspan and colspan values."),
        [HtmlConversionDiagnosticCodes.HtmlDepthLimitExceeded] = ConversionDefinition(HtmlConversionDiagnosticCodes.HtmlDepthLimitExceeded, "Safety", HtmlDiagnosticSeverity.Error, "HTML nesting exceeded the shared pre-analysis depth budget.", "Reduce nesting or raise MaxHtmlDepth only for trusted input."),
        [HtmlConversionDiagnosticCodes.CssSizeLimitExceeded] = ConversionDefinition(HtmlConversionDiagnosticCodes.CssSizeLimitExceeded, "Safety", HtmlDiagnosticSeverity.Error, "One embedded stylesheet exceeded the shared byte budget.", "Reduce the stylesheet or raise MaxCssBytes only for trusted input."),
        [HtmlConversionDiagnosticCodes.CssTotalSizeLimitExceeded] = ConversionDefinition(HtmlConversionDiagnosticCodes.CssTotalSizeLimitExceeded, "Safety", HtmlDiagnosticSeverity.Error, "Embedded stylesheets exceeded the operation-wide byte budget.", "Reduce embedded CSS or raise MaxTotalCssBytes only for trusted input."),
        [HtmlConversionDiagnosticCodes.CssRuleLimitExceeded] = ConversionDefinition(HtmlConversionDiagnosticCodes.CssRuleLimitExceeded, "Safety", HtmlDiagnosticSeverity.Error, "Active CSS rules exceeded the shared complexity budget.", "Reduce CSS rule volume or raise MaxCssRules only for trusted input."),
        [HtmlConversionDiagnosticCodes.CssDeclarationLimitExceeded] = ConversionDefinition(HtmlConversionDiagnosticCodes.CssDeclarationLimitExceeded, "Safety", HtmlDiagnosticSeverity.Error, "CSS declarations exceeded the shared complexity budget.", "Reduce CSS declaration volume or raise MaxCssDeclarations only for trusted input."),
        [HtmlConversionDiagnosticCodes.CssSelectorEvaluationLimitExceeded] = ConversionDefinition(HtmlConversionDiagnosticCodes.CssSelectorEvaluationLimitExceeded, "Safety", HtmlDiagnosticSeverity.Error, "Selector matching exceeded the shared evaluation budget.", "Simplify selectors or raise MaxSelectorEvaluations only for trusted input."),
        [HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded] = ConversionDefinition(HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded, "Safety", HtmlDiagnosticSeverity.Error, "A semantic metadata field exceeded its shared limit.", "Reduce oversized metadata or raise the explicit metadata limit only for trusted input."),
        [HtmlConversionDiagnosticCodes.SemanticSchemaUnsupported] = ConversionDefinition(HtmlConversionDiagnosticCodes.SemanticSchemaUnsupported, "SemanticImport", HtmlDiagnosticSeverity.Error, "The semantic source or schema version is unsupported.", "Use a matching OfficeIMO adapter and supported semantic schema version."),
        [HtmlConversionDiagnosticCodes.SemanticSchemaLegacy] = ConversionDefinition(HtmlConversionDiagnosticCodes.SemanticSchemaLegacy, "SemanticImport", HtmlDiagnosticSeverity.Info, "A legacy semantic envelope without a schema version was accepted.", "Re-export the content to add the current semantic schema version."),
        [HtmlConversionDiagnosticCodes.SemanticRestorationTrustRequired] = ConversionDefinition(HtmlConversionDiagnosticCodes.SemanticRestorationTrustRequired, "SemanticImport", HtmlDiagnosticSeverity.Warning, "The envelope requested target-specific restoration but the caller marked the HTML as untrusted.", "Parse the document with a trusted profile only after authenticating its source, or use public-safe semantic restoration."),
        [HtmlConversionDiagnosticCodes.MediaFilterFailed] = ConversionDefinition(HtmlConversionDiagnosticCodes.MediaFilterFailed, "CssFidelity", HtmlDiagnosticSeverity.Warning, "An active stylesheet could not be filtered safely for the selected media.", "Correct invalid CSS or simplify nested media rules."),
        [HtmlRenderDiagnosticCodes.BackgroundImageLayerLimit] = RenderDefinition(HtmlRenderDiagnosticCodes.BackgroundImageLayerLimit, "ResourceLimit", HtmlDiagnosticSeverity.Warning, "CSS background-image layers beyond the configured per-element limit were omitted.", "Increase MaxBackgroundImageLayers only for trusted documents or reduce the number of declared background layers."),
        [HtmlRenderDiagnosticCodes.BackgroundImageRepeatUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.BackgroundImageRepeatUnsupported, "PaintFidelity", HtmlDiagnosticSeverity.Warning, "A CSS background-repeat value used a single-image fallback.", "Use repeat, no-repeat, space, round, repeat-x, repeat-y, or a supported two-axis combination."),
        [HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported, "PaintFidelity", HtmlDiagnosticSeverity.Warning, "A CSS background image value used a deterministic supported fallback or was omitted.", "Use URL backgrounds, opaque linear gradients, or opaque radial circles and ellipses with keyword, length, or percentage geometry and percentage or implicit color stops until additional image functions are enabled."),
        [HtmlRenderDiagnosticCodes.BackgroundImageTileLimitExceeded] = RenderDefinition(HtmlRenderDiagnosticCodes.BackgroundImageTileLimitExceeded, "ResourceLimit", HtmlDiagnosticSeverity.Error, "Repeated CSS background images exceeded the configured operation-wide tile limit.", "Increase MaxBackgroundImageTiles only for trusted documents or use a larger background tile."),
        [HtmlRenderDiagnosticCodes.GradientStopLimitExceeded] = RenderDefinition(HtmlRenderDiagnosticCodes.GradientStopLimitExceeded, "ResourceLimit", HtmlDiagnosticSeverity.Error, "CSS gradients exceeded the configured color-stop limit.", "Increase MaxGradientStops only for trusted documents or reduce the number of gradient color stops."),
        ["HtmlCommentSkipped"] = new HtmlDiagnosticDefinition(
            "HtmlCommentSkipped",
            "ContentSimplification",
            HtmlDiagnosticSeverity.Info,
            "An HTML comment was omitted from generated document content.",
            "Enable HTML comment import when comments are part of the expected document contract, or keep comments as source-only metadata."),
        ["ImageResourceRejectedByPolicy"] = new HtmlDiagnosticDefinition(
            "ImageResourceRejectedByPolicy",
            "ResourcePolicy",
            HtmlDiagnosticSeverity.Warning,
            "An image candidate was rejected before loading because its URI is not allowed by policy.",
            "Allow the URI scheme or host for trusted inputs, embed the image as data URI, or provide a local resource resolver."),
        ["StylesheetResourceRejectedByPolicy"] = new HtmlDiagnosticDefinition(
            "StylesheetResourceRejectedByPolicy",
            "ResourcePolicy",
            HtmlDiagnosticSeverity.Warning,
            "A stylesheet was rejected before loading because its URI is not allowed by policy.",
            "Use caller-provided stylesheet contents for untrusted HTML, or allow the stylesheet scheme and host for trusted documents."),
        ["HyperlinkRejectedByPolicy"] = new HtmlDiagnosticDefinition(
            "HyperlinkRejectedByPolicy",
            "ResourcePolicy",
            HtmlDiagnosticSeverity.Warning,
            "A hyperlink target was rejected because its URI is not allowed by policy.",
            "Use http, https, mailto, or a caller-approved scheme instead of script or local file targets."),
        ["ScriptResourceRejectedByPolicy"] = new HtmlDiagnosticDefinition(
            "ScriptResourceRejectedByPolicy",
            "ResourcePolicy",
            HtmlDiagnosticSeverity.Warning,
            "A script dependency was rejected before loading because its URI is not allowed by policy.",
            "Use caller-provided script handling for trusted automation scenarios, or remove script dependencies from document-oriented HTML inputs."),
        ["MediaResourceRejectedByPolicy"] = new HtmlDiagnosticDefinition(
            "MediaResourceRejectedByPolicy",
            "ResourcePolicy",
            HtmlDiagnosticSeverity.Warning,
            "A media dependency was rejected before loading because its URI is not allowed by policy.",
            "Allow trusted media hosts explicitly, package approved media with the input, or provide a local resource resolver."),
        ["FontResourceRejectedByPolicy"] = new HtmlDiagnosticDefinition(
            "FontResourceRejectedByPolicy",
            "ResourcePolicy",
            HtmlDiagnosticSeverity.Warning,
            "A font dependency was rejected before loading because its URI is not allowed by policy.",
            "Use packaged fonts from trusted locations or allow approved font hosts in the URL policy."),
        ["UnsupportedCssDeclaration"] = new HtmlDiagnosticDefinition(
            "UnsupportedCssDeclaration",
            "CssFidelity",
            HtmlDiagnosticSeverity.Warning,
            "A CSS declaration could not be mapped to the target document model.",
            "Prefer document-friendly CSS or route visual-first workloads through the high-fidelity print profile."),
        ["HtmlResourceRejectedByPolicy"] = new HtmlDiagnosticDefinition(
            "HtmlResourceRejectedByPolicy",
            "ResourcePolicy",
            HtmlDiagnosticSeverity.Warning,
            "A resource dependency was rejected before loading because its URI is not allowed by policy.",
            "Adjust the URL policy only for trusted sources, or package the dependency with the HTML input."),
        [HtmlRenderDiagnosticCodes.DepthLimitExceeded] = RenderDefinition(HtmlRenderDiagnosticCodes.DepthLimitExceeded, "Safety", HtmlDiagnosticSeverity.Error, "HTML layout exceeded the configured nesting-depth limit.", "Reduce nesting or raise the explicit layout-depth limit for trusted input."),
        [HtmlRenderDiagnosticCodes.LayoutOperationLimitExceeded] = RenderDefinition(HtmlRenderDiagnosticCodes.LayoutOperationLimitExceeded, "Safety", HtmlDiagnosticSeverity.Error, "HTML layout exceeded the configured operation budget.", "Simplify the layout or raise MaxLayoutOperations only for trusted input."),
        [HtmlRenderDiagnosticCodes.InputCharacterLimitExceeded] = RenderDefinition(HtmlRenderDiagnosticCodes.InputCharacterLimitExceeded, "Safety", HtmlDiagnosticSeverity.Error, "HTML source text exceeded the configured character budget before parsing.", "Reduce or split the document, move large payloads behind a bounded resolver, or raise MaxInputCharacters only for trusted input."),
        [HtmlRenderDiagnosticCodes.NodeLimitExceeded] = RenderDefinition(HtmlRenderDiagnosticCodes.NodeLimitExceeded, "Safety", HtmlDiagnosticSeverity.Error, "The parsed HTML document exceeded the configured DOM node budget before styling or layout.", "Reduce repeated markup or raise MaxHtmlNodes only for trusted input."),
        [HtmlRenderDiagnosticCodes.EmptyTable] = RenderDefinition(HtmlRenderDiagnosticCodes.EmptyTable, "LayoutFidelity", HtmlDiagnosticSeverity.Info, "A table contained no renderable rows or cells.", "Add table rows and cells or remove the empty table."),
        [HtmlRenderDiagnosticCodes.TableLimitExceeded] = RenderDefinition(HtmlRenderDiagnosticCodes.TableLimitExceeded, "Safety", HtmlDiagnosticSeverity.Error, "A table exceeded the configured row or column limit.", "Reduce table dimensions or raise MaxTableRows or MaxTableColumns only for trusted input."),
        [HtmlRenderDiagnosticCodes.CollapsedTableBorderLimitExceeded] = RenderDefinition(HtmlRenderDiagnosticCodes.CollapsedTableBorderLimitExceeded, "Safety", HtmlDiagnosticSeverity.Error, "Collapsed table-border resolution exceeded the configured segment budget.", "Reduce table border complexity or raise MaxCollapsedTableBorderSegments only for trusted input."),
        [HtmlRenderDiagnosticCodes.ExternalImagePending] = RenderDefinition(HtmlRenderDiagnosticCodes.ExternalImagePending, "ResourcePolicy", HtmlDiagnosticSeverity.Warning, "An external image requires asynchronous resource resolution.", "Use RenderAsync with an application-supplied resource resolver or embed the image as a data URI."),
        [HtmlRenderDiagnosticCodes.ExternalStylesheetPending] = RenderDefinition(HtmlRenderDiagnosticCodes.ExternalStylesheetPending, "ResourcePolicy", HtmlDiagnosticSeverity.Warning, "An external stylesheet requires asynchronous resource resolution.", "Use RenderAsync with an application-supplied resource resolver or place trusted CSS in a style element."),
        [HtmlRenderDiagnosticCodes.FontDataUriInvalid] = RenderDefinition(HtmlRenderDiagnosticCodes.FontDataUriInvalid, "Typography", HtmlDiagnosticSeverity.Warning, "A font data URI could not be decoded.", "Provide a valid percent-encoded or base64 font data URI."),
        [HtmlRenderDiagnosticCodes.FontFaceInvalid] = RenderDefinition(HtmlRenderDiagnosticCodes.FontFaceInvalid, "Typography", HtmlDiagnosticSeverity.Warning, "An @font-face rule has no usable family descriptor.", "Provide a font-family descriptor and at least one usable src entry."),
        [HtmlRenderDiagnosticCodes.FontFaceUnavailable] = RenderDefinition(HtmlRenderDiagnosticCodes.FontFaceUnavailable, "Typography", HtmlDiagnosticSeverity.Warning, "No source from an @font-face rule was available.", "Use an allowed data URI or resolve the external font through RenderAsync."),
        [HtmlRenderDiagnosticCodes.FontFormatUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.FontFormatUnsupported, "Typography", HtmlDiagnosticSeverity.Warning, "A font source is not a supported TrueType glyf-outline font.", "Provide a TTF or TrueType-flavored OpenType face; WOFF, WOFF2, and CFF outlines require pre-conversion."),
        [HtmlRenderDiagnosticCodes.BidiLayoutUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.BidiLayoutUnsupported, "Typography", HtmlDiagnosticSeverity.Warning, "Explicit Unicode bidi controls require an embedding or isolate stage that is not active yet.", "Prefer semantic dir attributes for supported simple RTL layout until explicit bidi embedding and isolate controls are enabled."),
        [HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported, "Typography", HtmlDiagnosticSeverity.Warning, "A joining alphabet outside the bounded core-Arabic contextual shaper used scalar glyphs.", "Use the PDF shaping-provider seam for host-managed glyph shaping, or treat broader joining alphabets and OpenType mark positioning as unsupported until the shared managed stage expands."),
        [HtmlRenderDiagnosticCodes.FlexLayoutPending] = RenderDefinition(HtmlRenderDiagnosticCodes.FlexLayoutPending, "LayoutFidelity", HtmlDiagnosticSeverity.Warning, "A flex formatting case currently falls back to normal flow.", "Use row or column flex directions with static or relatively positioned items; absolute, fixed, sticky, and nested generated flex formatting contexts remain pending."),
        [HtmlRenderDiagnosticCodes.FlexValueUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.FlexValueUnsupported, "LayoutFidelity", HtmlDiagnosticSeverity.Warning, "A flex property value used a deterministic fallback.", "Use supported length or percentage bases, flex-start/start, flex-end/end, center, stretch, space-between, space-around, or space-evenly."),
        [HtmlRenderDiagnosticCodes.FloatValueUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.FloatValueUnsupported, "LayoutFidelity", HtmlDiagnosticSeverity.Warning, "A float or clear property value used a deterministic fallback.", "Use none, left, right, inline-start, or inline-end for float; use those values or both for clear."),
        [HtmlRenderDiagnosticCodes.MultiColumnLimitExceeded] = RenderDefinition(HtmlRenderDiagnosticCodes.MultiColumnLimitExceeded, "Safety", HtmlDiagnosticSeverity.Error, "Multi-column generation exceeded the configured column limit.", "Reduce column-count, increase column-width or available height, or raise MaxColumnCount only for trusted documents."),
        [HtmlRenderDiagnosticCodes.MultiColumnValueUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.MultiColumnValueUnsupported, "LayoutFidelity", HtmlDiagnosticSeverity.Warning, "A multi-column property value used a deterministic fallback.", "Use supported positive count/width values, balance or auto fill, none or all span, and solid, dashed, dotted, or double rules."),
        [HtmlRenderDiagnosticCodes.ForcedFragment] = RenderDefinition(HtmlRenderDiagnosticCodes.ForcedFragment, "PagedMedia", HtmlDiagnosticSeverity.Warning, "Content had no safe break opportunity within one page.", "Add break opportunities or reduce the size of the unbreakable content."),
        [HtmlRenderDiagnosticCodes.GeneratedContentUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.GeneratedContentUnsupported, "LayoutFidelity", HtmlDiagnosticSeverity.Warning, "A CSS generated-content expression was omitted.", "Use quoted text, attr(), counter(), or counters() with decimal, alphabetic, or Roman counter styles until additional generated-content forms are enabled."),
        [HtmlRenderDiagnosticCodes.GeneratedCounterUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.GeneratedCounterUnsupported, "LayoutFidelity", HtmlDiagnosticSeverity.Warning, "A CSS counter declaration was ignored.", "Use counter-reset, counter-set, and counter-increment with counter names and optional integer values; reversed counters remain unsupported."),
        [HtmlRenderDiagnosticCodes.GridLayoutPending] = RenderDefinition(HtmlRenderDiagnosticCodes.GridLayoutPending, "LayoutFidelity", HtmlDiagnosticSeverity.Warning, "A grid formatting case currently falls back to normal flow.", "Use block or inline grid with static or relatively positioned items; absolute/fixed/sticky items and nested generated grid formatting contexts remain pending."),
        [HtmlRenderDiagnosticCodes.GridTrackLimitExceeded] = RenderDefinition(HtmlRenderDiagnosticCodes.GridTrackLimitExceeded, "Safety", HtmlDiagnosticSeverity.Error, "Grid track expansion exceeded the configured limit.", "Reduce explicit, implicit, or repeat()-generated tracks, or raise MaxGridTracks for trusted input."),
        [HtmlRenderDiagnosticCodes.GridValueUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.GridValueUnsupported, "LayoutFidelity", HtmlDiagnosticSeverity.Warning, "A grid property value used a deterministic fallback.", "Use fixed, percentage, auto, fr, repeat(), or minmax() tracks with numeric lines and span values for the active grid contract."),
        [HtmlRenderDiagnosticCodes.OverflowScrollSnapshot] = RenderDefinition(HtmlRenderDiagnosticCodes.OverflowScrollSnapshot, "LayoutFidelity", HtmlDiagnosticSeverity.Info, "A scrollable overflow box was captured at its initial static scroll position.", "Use visible overflow when all content must remain visible, or hidden/clip when a static clipped export is intended."),
        [HtmlRenderDiagnosticCodes.OverflowClipMarginValueUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.OverflowClipMarginValueUnsupported, "LayoutFidelity", HtmlDiagnosticSeverity.Warning, "An overflow-clip-margin value used its initial fallback.", "Use an optional content-box, padding-box, or border-box keyword and one non-negative absolute length."),
        [HtmlRenderDiagnosticCodes.OverflowValueUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.OverflowValueUnsupported, "LayoutFidelity", HtmlDiagnosticSeverity.Warning, "An overflow property value used the visible fallback.", "Use visible, hidden, clip, auto, or scroll for overflow, overflow-x, and overflow-y."),
        [HtmlRenderDiagnosticCodes.TransformValueUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.TransformValueUnsupported, "PaintFidelity", HtmlDiagnosticSeverity.Warning, "A CSS transform or transform-origin value used the identity fallback.", "Use supported 2D matrix, translate, scale, rotate, or skew functions and a two-dimensional transform origin."),
        [HtmlRenderDiagnosticCodes.OpacityValueUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.OpacityValueUnsupported, "PaintFidelity", HtmlDiagnosticSeverity.Warning, "A CSS opacity value used the opaque fallback.", "Use a finite number or percentage; values outside the visible range are clamped."),
        [HtmlRenderDiagnosticCodes.InlinePaintEffectUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.InlinePaintEffectUnsupported, "PaintFidelity", HtmlDiagnosticSeverity.Warning, "A paint effect on a non-atomic inline box used normal inline paint.", "Use a block, inline-block, inline-flex, or inline-grid wrapper when an isolated transform or opacity group is required."),
        [HtmlRenderDiagnosticCodes.ReplacedElementValueUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.ReplacedElementValueUnsupported, "LayoutFidelity", HtmlDiagnosticSeverity.Warning, "A replaced-element sizing or object-placement value used a deterministic fallback.", "Use a positive aspect ratio, fill, contain, cover, none, or scale-down object fitting, and a supported keyword, length, or percentage object position."),
        [HtmlRenderDiagnosticCodes.BorderRadiusValueUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.BorderRadiusValueUnsupported, "PaintFidelity", HtmlDiagnosticSeverity.Warning, "A CSS border radius used square-corner fallback.", "Use one to four non-negative length or percentage radii, an optional slash-separated vertical axis, or valid one- or two-axis corner longhands."),
        [HtmlRenderDiagnosticCodes.BoxShadowLayerLimit] = RenderDefinition(HtmlRenderDiagnosticCodes.BoxShadowLayerLimit, "ResourceLimit", HtmlDiagnosticSeverity.Warning, "CSS box-shadow layers beyond the configured per-element limit were omitted.", "Increase MaxBoxShadowLayers only for trusted documents or reduce the number of declared shadows."),
        [HtmlRenderDiagnosticCodes.BoxShadowValueUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.BoxShadowValueUnsupported, "PaintFidelity", HtmlDiagnosticSeverity.Warning, "A CSS box shadow was omitted.", "Use comma-separated inset or outer shadows with two offsets, an optional non-negative blur radius, an optional signed spread radius, and a supported color."),
        [HtmlRenderDiagnosticCodes.BorderPaintValueUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.BorderPaintValueUnsupported, "PaintFidelity", HtmlDiagnosticSeverity.Warning, "A CSS border paint declaration used no-border fallback.", "Use one to four non-negative widths, supported colors, and solid, dashed, dotted, double, none, or hidden side styles."),
        [HtmlRenderDiagnosticCodes.OutlinePaintValueUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.OutlinePaintValueUnsupported, "PaintFidelity", HtmlDiagnosticSeverity.Warning, "A CSS outline paint declaration was omitted.", "Use one non-negative width, supported color, signed offset, and solid, dashed, dotted, double, none, or hidden style."),
        [HtmlRenderDiagnosticCodes.PositionInsetUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.PositionInsetUnsupported, "LayoutFidelity", HtmlDiagnosticSeverity.Warning, "A positioned inset could not be resolved and used its documented fallback.", "Use a supported CSS length or percentage with a definite containing-block dimension."),
        [HtmlRenderDiagnosticCodes.PositioningModeUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.PositioningModeUnsupported, "LayoutFidelity", HtmlDiagnosticSeverity.Warning, "A CSS positioning mode currently falls back to normal flow.", "Use static, relative, absolute, fixed, or sticky positioning without unsupported containing-block features."),
        [HtmlRenderDiagnosticCodes.PositionStaticAnchorFallback] = RenderDefinition(HtmlRenderDiagnosticCodes.PositionStaticAnchorFallback, "LayoutFidelity", HtmlDiagnosticSeverity.Warning, "An automatic positioned inset could not use the element's hypothetical normal-flow anchor.", "Use an explicit inset or place the positioned element in a supported block, flex, or grid static-position context."),
        [HtmlRenderDiagnosticCodes.PositionStickyStatic] = RenderDefinition(HtmlRenderDiagnosticCodes.PositionStickyStatic, "LayoutFidelity", HtmlDiagnosticSeverity.Info, "A sticky-positioned element was captured at its stable static document position.", "Use fixed positioning for repeated page overlays; sticky scroll-state changes are not meaningful in a static document snapshot."),
        [HtmlRenderDiagnosticCodes.PositionZIndexPending] = RenderDefinition(HtmlRenderDiagnosticCodes.PositionZIndexPending, "PaintFidelity", HtmlDiagnosticSeverity.Warning, "A positioned element's z-index is not active in the current stacking model.", "Keep source order for the current contract until stacking contexts are enabled."),
        [HtmlRenderDiagnosticCodes.PageMarginContentUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.PageMarginContentUnsupported, "PagedMedia", HtmlDiagnosticSeverity.Warning, "A page-margin generated-content expression could not be represented.", "Use quoted text with counter(page) or counter(pages) until richer generated content is enabled."),
        [HtmlRenderDiagnosticCodes.PageMarginPositionUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.PageMarginPositionUnsupported, "PagedMedia", HtmlDiagnosticSeverity.Warning, "A page-margin position is not recognized by the direct renderer.", "Use one of the standard CSS top, bottom, left, right, or corner page-margin box names."),
        [HtmlRenderDiagnosticCodes.PagePseudoGeometryPending] = RenderDefinition(HtmlRenderDiagnosticCodes.PagePseudoGeometryPending, "PagedMedia", HtmlDiagnosticSeverity.Warning, "A pseudo-page size or margin declaration requires page-by-page body reflow.", "Keep body geometry in the generic @page rule until per-page reflow is enabled; pseudo-page margin content is still applied."),
        [HtmlRenderDiagnosticCodes.PageSelectorPending] = RenderDefinition(HtmlRenderDiagnosticCodes.PageSelectorPending, "PagedMedia", HtmlDiagnosticSeverity.Warning, "A complex page selector could not be applied per page.", "Use a generic, named, :first, :left, or :right @page selector, optionally combining one name with one supported pseudo-page."),
        [HtmlRenderDiagnosticCodes.PageSizeUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.PageSizeUnsupported, "PagedMedia", HtmlDiagnosticSeverity.Warning, "An @page size declaration could not be mapped.", "Use a supported named size or two absolute physical lengths."),
        [HtmlRenderDiagnosticCodes.SvgContentUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.SvgContentUnsupported, "ImageFidelity", HtmlDiagnosticSeverity.Warning, "SVG content could not be represented completely by the bounded shared vector scene.", "Use supported primitives or paths, bounded local shape/group references, local object-bounding-box linear/radial paint servers, positioned tspan text, and affine transform attributes while broader SVG scene support is being completed."),
        [HtmlRenderDiagnosticCodes.ResourceByteLimitExceeded] = RenderDefinition(HtmlRenderDiagnosticCodes.ResourceByteLimitExceeded, "ResourcePolicy", HtmlDiagnosticSeverity.Warning, "A resource exceeded the configured per-resource byte limit.", "Reduce the resource or raise the explicit limit for trusted input."),
        [HtmlRenderDiagnosticCodes.ResourceCountLimitExceeded] = RenderDefinition(HtmlRenderDiagnosticCodes.ResourceCountLimitExceeded, "ResourcePolicy", HtmlDiagnosticSeverity.Error, "Resolved resources exceeded the configured count limit.", "Reduce the resource graph or raise the explicit count limit for trusted input."),
        [HtmlRenderDiagnosticCodes.ResourceRequestLimitExceeded] = RenderDefinition(HtmlRenderDiagnosticCodes.ResourceRequestLimitExceeded, "ResourcePolicy", HtmlDiagnosticSeverity.Error, "Resource resolver invocations exceeded the configured request limit.", "Reduce broken or unavailable references, or raise the explicit request limit for trusted input."),
        [HtmlRenderDiagnosticCodes.ResourceContentTypeRejected] = RenderDefinition(HtmlRenderDiagnosticCodes.ResourceContentTypeRejected, "ResourcePolicy", HtmlDiagnosticSeverity.Warning, "A resolver returned an incompatible media type.", "Return bytes whose declared media type matches the requested image or stylesheet kind."),
        [HtmlRenderDiagnosticCodes.ResourceLoadFailed] = RenderDefinition(HtmlRenderDiagnosticCodes.ResourceLoadFailed, "ResourcePolicy", HtmlDiagnosticSeverity.Warning, "The configured resource resolver failed.", "Inspect the resolver boundary and return null for intentionally unavailable resources."),
        [HtmlRenderDiagnosticCodes.ResourceTimeout] = RenderDefinition(HtmlRenderDiagnosticCodes.ResourceTimeout, "ResourcePolicy", HtmlDiagnosticSeverity.Warning, "Resource resolution exceeded its timeout.", "Reduce resolver latency or raise the bounded timeout for trusted workloads."),
        [HtmlRenderDiagnosticCodes.ResourceUnavailable] = RenderDefinition(HtmlRenderDiagnosticCodes.ResourceUnavailable, "ResourcePolicy", HtmlDiagnosticSeverity.Warning, "The configured resolver returned no resource.", "Provide the resource or accept the diagnosed placeholder."),
        [HtmlRenderDiagnosticCodes.ResourceUriInvalid] = RenderDefinition(HtmlRenderDiagnosticCodes.ResourceUriInvalid, "ResourcePolicy", HtmlDiagnosticSeverity.Warning, "A resource URI could not be represented as an absolute URI.", "Provide a valid base URI and resource reference."),
        [HtmlRenderDiagnosticCodes.StylesheetEncodingUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.StylesheetEncodingUnsupported, "CssFidelity", HtmlDiagnosticSeverity.Warning, "A resolved stylesheet could not be decoded as supported CSS text.", "Return UTF-8 CSS or UTF-16 CSS with a byte-order mark."),
        [HtmlRenderDiagnosticCodes.StylesheetImportCycle] = RenderDefinition(HtmlRenderDiagnosticCodes.StylesheetImportCycle, "CssFidelity", HtmlDiagnosticSeverity.Warning, "A recursive stylesheet import cycle was suppressed.", "Remove the cyclic @import relationship."),
        [HtmlRenderDiagnosticCodes.StylesheetImportDepthExceeded] = RenderDefinition(HtmlRenderDiagnosticCodes.StylesheetImportDepthExceeded, "ResourcePolicy", HtmlDiagnosticSeverity.Error, "Stylesheet imports exceeded the configured recursion depth.", "Flatten the import graph or raise the explicit depth limit for trusted input."),
        [HtmlRenderDiagnosticCodes.StylesheetUrlResourcesPending] = RenderDefinition(HtmlRenderDiagnosticCodes.StylesheetUrlResourcesPending, "CssFidelity", HtmlDiagnosticSeverity.Warning, "An external stylesheet contains URL resources that are not active in the current paint model.", "Inline those assets or use currently supported image and font resource paths until CSS URL painting is enabled."),
        [HtmlRenderDiagnosticCodes.TableFooterRepeatSuppressed] = RenderDefinition(HtmlRenderDiagnosticCodes.TableFooterRepeatSuppressed, "PagedMedia", HtmlDiagnosticSeverity.Warning, "A repeated table footer left no safe body-row break on an empty page.", "Reduce the footer or row height, increase the page content area, or allow the body row to move without a repeated footer."),
        [HtmlRenderDiagnosticCodes.TableHeaderRepeatSuppressed] = RenderDefinition(HtmlRenderDiagnosticCodes.TableHeaderRepeatSuppressed, "PagedMedia", HtmlDiagnosticSeverity.Warning, "A repeated table header left no safe body-row break on an empty page.", "Reduce the header or row height, increase the page content area, or allow the body row to move without a repeated header."),
        [HtmlRenderDiagnosticCodes.TableValueUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.TableValueUnsupported, "LayoutFidelity", HtmlDiagnosticSeverity.Warning, "A table formatting property used its documented fallback.", "Use top or bottom for caption-side, auto or fixed for table-layout, separate or collapse for border-collapse, and one or two non-negative absolute lengths for border-spacing."),
        [HtmlRenderDiagnosticCodes.TotalResourceByteLimitExceeded] = RenderDefinition(HtmlRenderDiagnosticCodes.TotalResourceByteLimitExceeded, "ResourcePolicy", HtmlDiagnosticSeverity.Error, "Resolved resources exceeded the total byte budget.", "Reduce resource volume or raise the explicit total limit for trusted input."),
        [HtmlRenderDiagnosticCodes.VisualFragmentUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.VisualFragmentUnsupported, "PagedMedia", HtmlDiagnosticSeverity.Warning, "A visual could not cross a forced page boundary safely.", "Resize the visual or add a safe break before it.")
    };

    private static readonly IReadOnlyList<HtmlDiagnosticDefinition> OrderedDefinitions = Definitions.Values
        .OrderBy(definition => definition.Category, StringComparer.Ordinal)
        .ThenBy(definition => definition.Code, StringComparer.Ordinal)
        .ToList()
        .AsReadOnly();

    /// <summary>All stable diagnostic definitions in deterministic category/code order.</summary>
    public static IReadOnlyList<HtmlDiagnosticDefinition> Ordered => OrderedDefinitions;

    /// <summary>
    /// Gets all known diagnostic definitions.
    /// </summary>
    public static IReadOnlyDictionary<string, HtmlDiagnosticDefinition> All => Definitions;

    /// <summary>
    /// Looks up support metadata for a diagnostic code.
    /// </summary>
    public static bool TryGet(string code, out HtmlDiagnosticDefinition definition) {
        if (string.IsNullOrWhiteSpace(code)) {
            definition = null!;
            return false;
        }

        HtmlDiagnosticDefinition? found;
        if (Definitions.TryGetValue(code.Trim(), out found)) {
            definition = found;
            return true;
        }

        definition = null!;
        return false;
    }

    /// <summary>
    /// Gets support metadata for a diagnostic code, or a generic definition when the code is unknown.
    /// </summary>
    public static HtmlDiagnosticDefinition GetOrCreateGeneric(string code) {
        if (TryGet(code, out HtmlDiagnosticDefinition definition)) {
            return definition;
        }

        return new HtmlDiagnosticDefinition(
            string.IsNullOrWhiteSpace(code) ? "HtmlDiagnostic" : code.Trim(),
            "General",
            HtmlDiagnosticSeverity.Warning,
            "The HTML workflow emitted a diagnostic that is not yet cataloged.",
            "Use the diagnostic source and detail fields to decide whether input, policy, or converter support should be adjusted.");
    }

    private static HtmlDiagnosticDefinition RenderDefinition(string code, string category, HtmlDiagnosticSeverity severity, string message, string remediation) =>
        new HtmlDiagnosticDefinition(code, category, severity, message, remediation);

    private static HtmlDiagnosticDefinition ConversionDefinition(string code, string category, HtmlDiagnosticSeverity severity, string message, string remediation) =>
        RenderDefinition(code, category, severity, message, remediation);
}
