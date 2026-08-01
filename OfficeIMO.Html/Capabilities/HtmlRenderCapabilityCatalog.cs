namespace OfficeIMO.Html;

/// <summary>
/// Executable source of truth for the direct HTML renderer's standards surface.
/// Entries describe the supported subset, not merely whether a property is parsed.
/// </summary>
public static class HtmlRenderCapabilityCatalog {
    private static readonly IReadOnlyList<HtmlRenderCapability> Capabilities = new[] {
        Full("css-cascade", "CSS cascade", HtmlRenderCapabilityKind.Css,
            Features("author stylesheets", "caller stylesheets", "inline styles", "!important", "inheritance", "custom properties", "@supports"),
            "Applies the bounded author cascade, selector specificity, inherited values, var() substitution, supported @supports conditions, and caller stylesheets appended after document styles."),
        Partial("css-selectors", "CSS selectors", HtmlRenderCapabilityKind.Css,
            Features("type", "class", "id", "attribute", "combinators", "structural pseudo-classes", "::before", "::after"),
            "Matches the documented selector subset and generated before/after content; selectors outside the bounded subset do not match."),
        Full("css-length-units", "CSS values", HtmlRenderCapabilityKind.Css,
            Features("px", "pt", "pc", "in", "cm", "mm", "q", "em", "rem", "%"),
            "Resolves absolute, font-relative, and percentage lengths against the active layout reference."),
        Full("css-length-math", "CSS values", HtmlRenderCapabilityKind.Css,
            Features("calc()", "min()", "max()", "clamp()"),
            "Evaluates bounded nested length arithmetic with dimensional checks for addition, subtraction, multiplication, and division."),
        Partial("css-color", "Color and paint", HtmlRenderCapabilityKind.Css,
            Features("named colors", "hex colors", "rgb()", "rgba()", "hsl()", "hsla()", "transparent"),
            "Resolves the listed legacy and modern color forms through the shared Drawing parser. Broader CSS Color forms remain outside the declared subset.",
            HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported,
            HtmlRenderDiagnosticCodes.BorderPaintValueUnsupported),
        Partial("css-backgrounds", "Color and paint", HtmlRenderCapabilityKind.Css,
            Features("background-color", "background-image", "background-position", "background-repeat", "background-size", "linear-gradient()", "radial-gradient()"),
            "Paints bounded image layers plus linear and radial gradients. Unsupported layers or repeat modes use diagnosed fallbacks.",
            HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported,
            HtmlRenderDiagnosticCodes.BackgroundImageRepeatUnsupported),
        Partial("css-borders-effects", "Color and paint", HtmlRenderCapabilityKind.Css,
            Features("border", "border-radius", "outline", "box-shadow", "opacity", "transform"),
            "Paints supported border, radius, outline, shadow, opacity, and two-dimensional affine transform forms.",
            HtmlRenderDiagnosticCodes.BorderPaintValueUnsupported,
            HtmlRenderDiagnosticCodes.BorderRadiusValueUnsupported,
            HtmlRenderDiagnosticCodes.BoxShadowValueUnsupported,
            HtmlRenderDiagnosticCodes.OutlinePaintValueUnsupported,
            HtmlRenderDiagnosticCodes.TransformValueUnsupported),
        Full("layout-block-inline", "Layout", HtmlRenderCapabilityKind.Css,
            Features("block flow", "inline flow", "inline-block", "box sizing", "margins", "padding", "line boxes"),
            "Builds searchable block and inline layout with box-model sizing and line wrapping."),
        Partial("bidi-text", "Typography", HtmlRenderCapabilityKind.Css,
            Features("dir=ltr", "dir=rtl", "LRE", "RLE", "LRO", "RLO", "PDF", "LRI", "RLI", "FSI", "PDI"),
            "Uses the shared Drawing resolver for bounded Unicode embeddings, overrides, isolates, logical text retention, and deterministic visual positioning. Advanced OpenType shaping outside the managed subset remains diagnosed.",
            HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported),
        Partial("layout-flex", "Layout", HtmlRenderCapabilityKind.Css,
            Features("display:flex", "display:inline-flex", "flex-direction", "flex-wrap", "flex", "gap", "alignment"),
            "Lays out supported row and column flex containers, wrapping, gaps, ordering, and alignment. Unsupported values use deterministic fallbacks.",
            HtmlRenderDiagnosticCodes.FlexLayoutPending,
            HtmlRenderDiagnosticCodes.FlexValueUnsupported),
        Partial("layout-grid", "Layout", HtmlRenderCapabilityKind.Css,
            Features("display:grid", "display:inline-grid", "grid-template-*", "grid-auto-*", "repeat()", "minmax()", "gap", "placement"),
            "Lays out bounded explicit and implicit grids with numeric placement. Subgrid, named-line resolution, and intrinsic track sizing are outside the current contract.",
            HtmlRenderDiagnosticCodes.GridLayoutPending,
            HtmlRenderDiagnosticCodes.GridValueUnsupported),
        Partial("layout-columns", "Layout", HtmlRenderCapabilityKind.Css,
            Features("columns", "column-count", "column-width", "column-fill", "column-gap", "column-rule", "column-span"),
            "Builds bounded multi-column layout with balancing, rules, and spanning blocks. Advanced cross-page fragmentation remains limited.",
            HtmlRenderDiagnosticCodes.MultiColumnValueUnsupported),
        Partial("layout-positioning", "Layout", HtmlRenderCapabilityKind.Css,
            Features("position:relative", "position:absolute", "position:fixed", "position:sticky", "insets", "z-index"),
            "Places relative, absolute, and fixed boxes in the supported containing-block model. Sticky content is captured statically and unsupported anchors are diagnosed.",
            HtmlRenderDiagnosticCodes.PositionInsetUnsupported,
            HtmlRenderDiagnosticCodes.PositioningModeUnsupported,
            HtmlRenderDiagnosticCodes.PositionStaticAnchorFallback,
            HtmlRenderDiagnosticCodes.PositionStickyStatic),
        Partial("layout-tables", "Layout", HtmlRenderCapabilityKind.Html,
            Features("table", "caption", "thead", "tbody", "tfoot", "rowspan", "colspan", "border-collapse", "table-layout"),
            "Builds bounded table grids with spans, collapsed or separate borders, captions, and repeated paged headers and footers.",
            HtmlRenderDiagnosticCodes.TableValueUnsupported),
        Partial("generated-content", "Generated content", HtmlRenderCapabilityKind.Css,
            Features("content", "attr()", "counter()", "counters()", "counter-reset", "counter-set", "counter-increment", "::before", "::after"),
            "Generates quoted text, attributes, and supported counters. Unsupported expressions are omitted with diagnostics.",
            HtmlRenderDiagnosticCodes.GeneratedContentUnsupported,
            HtmlRenderDiagnosticCodes.GeneratedCounterUnsupported),
        Partial("paged-page-rules", "Paged media", HtmlRenderCapabilityKind.PagedMedia,
            Features("@page", "size", "margin", ":first", ":left", ":right", "named pages", "margin boxes", "counter(page)", "counter(pages)"),
            "Applies generic page geometry and supported per-page margin content. Pseudo-page geometry that requires body reflow uses the generic geometry.",
            HtmlRenderDiagnosticCodes.PageMarginContentUnsupported,
            HtmlRenderDiagnosticCodes.PagePseudoGeometryPending,
            HtmlRenderDiagnosticCodes.PageSelectorPending,
            HtmlRenderDiagnosticCodes.PageSizeUnsupported),
        Partial("paged-fragmentation", "Paged media", HtmlRenderCapabilityKind.PagedMedia,
            Features("break-before", "break-after", "break-inside", "orphans", "widows", "table header repetition", "table footer repetition"),
            "Honors supported break constraints, text widows/orphans, and repeated table sections. Oversized atomic visuals may require a diagnosed forced fragment.",
            HtmlRenderDiagnosticCodes.ForcedFragment,
            HtmlRenderDiagnosticCodes.VisualFragmentUnsupported),
        Partial("media-queries", "Media queries", HtmlRenderCapabilityKind.Css,
            Features("screen", "print", "width", "height", "orientation", "resolution", "color", "monochrome", "prefers-color-scheme", "prefers-reduced-motion", "hover", "pointer"),
            "Evaluates media type, surface geometry, orientation, and caller-selected deterministic static-environment features."),
        Partial("web-fonts", "Fonts", HtmlRenderCapabilityKind.Resource,
            Features("@font-face", "font-family", "font-style", "font-weight", "unicode-range", "TrueType glyf", "WOFF 1"),
            "Loads bounded policy-approved direct OpenType and WOFF 1 TrueType glyf faces, selects constrained faces by Unicode scalar range, and preserves the text-shaping provider seam. WOFF 2 transformed tables and unsupported outlines fall back with diagnostics.",
            HtmlRenderDiagnosticCodes.FontFaceUnavailable,
            HtmlRenderDiagnosticCodes.FontFormatUnsupported),
        Partial("images", "Images", HtmlRenderCapabilityKind.Resource,
            Features("img", "picture", "srcset", "PNG", "JPEG", "TIFF", "SVG", "WebP", "object-fit", "object-position"),
            "Resolves bounded responsive image candidates and paints supported raster and vector sources."),
        Partial("svg", "SVG", HtmlRenderCapabilityKind.Resource,
            Features("paths", "basic shapes", "groups", "symbols", "use", "named/hex/rgb/hsl paint", "linearGradient", "radialGradient", "clipPath", "mask", "mix-blend-mode", "text", "tspan", "affine transforms"),
            "Maps the listed bounded SVG subset, local user-space masks, and standard blend modes into the shared Drawing scene. Unsupported filters or external references retain supported geometry and are diagnosed.",
            HtmlRenderDiagnosticCodes.SvgContentUnsupported,
            HtmlRenderDiagnosticCodes.SvgRasterFallback),
        Partial("pdf-metadata", "Output metadata", HtmlRenderCapabilityKind.Output,
            Features("title", "author", "subject", "description", "keywords", "creator", "generator", "language", "reading direction"),
            "Carries normalized document metadata into the shared render result; PDF output maps title, author, subject, keywords, language, and reading direction."),
        Partial("pdf-accessibility", "Output accessibility", HtmlRenderCapabilityKind.Output,
            Features("tagged PDF", "document language", "reading order", "headings", "lists", "tables", "links", "alternate text"),
            "Maps supported semantic groups into tagged PDF structures. Broader validator-backed conformance remains an explicit release gate."),
        Rejected("resource-policy", "Resource safety", HtmlRenderCapabilityKind.Resource,
            Features("local files", "remote resources", "data URIs", "package resources", "hyperlinks"),
            "Rejects resources and links outside the caller-selected URL and host-resource policies before loading or emission.",
            "HtmlResourceRejectedByPolicy",
            "HyperlinkRejectedByPolicy",
            "ImageResourceRejectedByPolicy",
            "FontResourceRejectedByPolicy",
            "StylesheetResourceRejectedByPolicy"),
        Fallback("unsupported-static-features", "Static output boundary", HtmlRenderCapabilityKind.Output,
            Features("scroll state", "sticky state", "interactive controls", "animation", "perspective transforms"),
            "Captures a deterministic static representation when a dynamic feature has a meaningful snapshot; otherwise the feature is omitted or uses its initial value.",
            HtmlRenderDiagnosticCodes.OverflowScrollSnapshot,
            HtmlRenderDiagnosticCodes.PositionStickyStatic,
            HtmlRenderDiagnosticCodes.TransformValueUnsupported),
        Ignored("active-content", "Active content", HtmlRenderCapabilityKind.Html,
            Features("script execution", "event handlers", "embedded active content"),
            "Does not execute active content during parsing, layout, or output generation.",
            "ScriptResourceRejectedByPolicy")
    };

    private static readonly IReadOnlyDictionary<string, HtmlRenderCapability> ById =
        Capabilities.ToDictionary(capability => capability.Id, StringComparer.OrdinalIgnoreCase);

    /// <summary>Gets all built-in renderer capability contracts in stable area and identifier order.</summary>
    public static IReadOnlyList<HtmlRenderCapability> All { get; } = Capabilities
        .OrderBy(capability => capability.Area, StringComparer.Ordinal)
        .ThenBy(capability => capability.Id, StringComparer.Ordinal)
        .ToList()
        .AsReadOnly();

    /// <summary>Gets a renderer capability by stable identifier.</summary>
    public static HtmlRenderCapability Get(string id) {
        if (!TryGet(id, out HtmlRenderCapability capability)) {
            throw new ArgumentOutOfRangeException(nameof(id), id, "Unknown HTML renderer capability.");
        }
        return capability;
    }

    /// <summary>Attempts to get a renderer capability by stable identifier.</summary>
    public static bool TryGet(string? id, out HtmlRenderCapability capability) {
        if (!string.IsNullOrWhiteSpace(id)
            && ById.TryGetValue(id!.Trim(), out HtmlRenderCapability? found)
            && found != null) {
            capability = found;
            return true;
        }
        capability = null!;
        return false;
    }

    private static string[] Features(params string[] values) => values;

    private static HtmlRenderCapability Full(
        string id,
        string area,
        HtmlRenderCapabilityKind kind,
        IEnumerable<string> features,
        string behavior,
        params string[] diagnostics) =>
        Create(id, area, kind, HtmlRenderSupportLevel.Full, features, behavior, diagnostics);

    private static HtmlRenderCapability Partial(
        string id,
        string area,
        HtmlRenderCapabilityKind kind,
        IEnumerable<string> features,
        string behavior,
        params string[] diagnostics) =>
        Create(id, area, kind, HtmlRenderSupportLevel.Partial, features, behavior, diagnostics);

    private static HtmlRenderCapability Fallback(
        string id,
        string area,
        HtmlRenderCapabilityKind kind,
        IEnumerable<string> features,
        string behavior,
        params string[] diagnostics) =>
        Create(id, area, kind, HtmlRenderSupportLevel.Fallback, features, behavior, diagnostics);

    private static HtmlRenderCapability Ignored(
        string id,
        string area,
        HtmlRenderCapabilityKind kind,
        IEnumerable<string> features,
        string behavior,
        params string[] diagnostics) =>
        Create(id, area, kind, HtmlRenderSupportLevel.Ignored, features, behavior, diagnostics);

    private static HtmlRenderCapability Rejected(
        string id,
        string area,
        HtmlRenderCapabilityKind kind,
        IEnumerable<string> features,
        string behavior,
        params string[] diagnostics) =>
        Create(id, area, kind, HtmlRenderSupportLevel.Rejected, features, behavior, diagnostics);

    private static HtmlRenderCapability Create(
        string id,
        string area,
        HtmlRenderCapabilityKind kind,
        HtmlRenderSupportLevel supportLevel,
        IEnumerable<string> features,
        string behavior,
        IEnumerable<string> diagnostics) =>
        new HtmlRenderCapability(id, area, kind, supportLevel, features, behavior, diagnostics);
}
