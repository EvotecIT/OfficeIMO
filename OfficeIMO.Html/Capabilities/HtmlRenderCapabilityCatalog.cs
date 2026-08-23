namespace OfficeIMO.Html;

/// <summary>
/// Executable source of truth for the direct HTML renderer's standards surface.
/// Entries describe the supported subset, not merely whether a property is parsed.
/// </summary>
public static class HtmlRenderCapabilityCatalog {
    private static readonly IReadOnlyList<HtmlRenderCapability> Capabilities = new[] {
        Full("css-cascade", "CSS cascade", HtmlRenderCapabilityKind.Css,
            Features("author stylesheets", "caller stylesheets", "inline styles", "!important", "inheritance", "custom properties", "@supports", "@layer", "revert-layer"),
            "Applies the bounded author cascade, selector specificity, cascade-layer ordering, layer rollback, inherited values, var() substitution, supported @supports conditions, and caller stylesheets appended after document styles."),
        Full("css-selectors", "CSS selectors", HtmlRenderCapabilityKind.Css,
            Features("type", "class", "id", "attribute", "combinators", "structural pseudo-classes", "CSS nesting", "::before", "::after"),
            "Matches the documented selector subset, bounded parent-list and ampersand nesting including nested conditional rules, and generated before/after content; selectors outside the bounded subset do not match."),
        Full("css-length-units", "CSS values", HtmlRenderCapabilityKind.Css,
            Features("px", "pt", "pc", "in", "cm", "mm", "q", "em", "rem", "%", "vw/vh/vmin/vmax", "sv*/lv*/dv* viewport units", "cqw/cqh/cqi/cqb/cqmin/cqmax"),
            "Resolves absolute, font-relative, percentage, static viewport-family, and bounded container-query lengths against the active layout references. Writing-mode-relative and font-metric-relative unit families remain outside this subset."),
        Full("css-length-math", "CSS values", HtmlRenderCapabilityKind.Css,
            Features("calc()", "min()", "max()", "clamp()"),
            "Evaluates bounded nested length arithmetic with dimensional checks for addition, subtraction, multiplication, and division."),
        Full("css-color", "Color and paint", HtmlRenderCapabilityKind.Css,
            Features("named colors", "hex colors", "rgb()/rgba()", "hsl()/hsla()", "hwb()", "lab()/lch()", "oklab()/oklch()", "color() predefined spaces", "color-mix() in sRGB, linear sRGB, or OKLab", "transparent"),
            "Resolves the listed CSS Color 4 forms through the shared Drawing parser, including wide-gamut predefined spaces and premultiplied-alpha color mixing."),
        Fallback("css-color-fallback", "Color and paint", HtmlRenderCapabilityKind.Css,
            Features("relative color syntax", "unlisted color() spaces", "unlisted color-mix() interpolation spaces"),
            "Uses the property initial value and emits a typed diagnostic when a color expression is outside the declared static contract.",
            HtmlRenderDiagnosticCodes.ColorValueUnsupported,
            HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported,
            HtmlRenderDiagnosticCodes.BorderPaintValueUnsupported),
        Full("css-backgrounds", "Color and paint", HtmlRenderCapabilityKind.Css,
            Features("background-color", "background-image", "background-position", "background-repeat", "background-size", "linear-gradient()", "repeating-linear-gradient()", "radial-gradient()", "repeating-radial-gradient()", "conic-gradient()", "repeating-conic-gradient()"),
            "Paints bounded image layers, native vector linear and radial gradients, and bounded vector conic-gradient expansions across raster, SVG, and PDF outputs."),
        Fallback("css-backgrounds-fallback", "Color and paint", HtmlRenderCapabilityKind.Css,
            Features("unsupported background layers", "unsupported repeat modes"),
            "Uses a diagnosed initial layer or repeat value when authored syntax is outside the declared static contract.",
            HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported,
            HtmlRenderDiagnosticCodes.BackgroundImageRepeatUnsupported),
        Full("css-borders-effects", "Color and paint", HtmlRenderCapabilityKind.Css,
            Features("border", "border-radius", "outline", "box-shadow", "opacity", "transform", "clip-path basic shapes", "clip-path geometry boxes"),
            "Paints the declared border, radius, outline, shadow, opacity, two-dimensional affine transform, and basic-shape clip-path forms against margin, border, padding, or content reference boxes."),
        Fallback("css-borders-effects-fallback", "Color and paint", HtmlRenderCapabilityKind.Css,
            Features("unsupported border paint", "unsupported radius", "unsupported shadow", "unsupported outline", "unsupported transform", "unsupported clip-path"),
            "Uses a typed omission or initial-value diagnostic when an effect is outside the declared static contract.",
            HtmlRenderDiagnosticCodes.BorderPaintValueUnsupported,
            HtmlRenderDiagnosticCodes.BorderRadiusValueUnsupported,
            HtmlRenderDiagnosticCodes.BoxShadowValueUnsupported,
            HtmlRenderDiagnosticCodes.OutlinePaintValueUnsupported,
            HtmlRenderDiagnosticCodes.TransformValueUnsupported,
            HtmlRenderDiagnosticCodes.ClipPathValueUnsupported),
        Full("layout-block-inline", "Layout", HtmlRenderCapabilityKind.Css,
            Features("block flow", "inline flow", "inline-block", "box sizing", "aspect-ratio", "margins", "padding", "line boxes"),
            "Builds searchable block and inline layout with box-model and preferred-aspect-ratio sizing and line wrapping."),
        Full("bidi-text", "Typography", HtmlRenderCapabilityKind.Css,
            Features("dir=ltr", "dir=rtl", "LRE", "RLE", "LRO", "RLO", "PDF", "LRI", "RLI", "FSI", "PDI"),
            "Uses the shared Drawing resolver for bounded Unicode embeddings, overrides, isolates, logical text retention, and deterministic visual positioning."),
        Full("text-flow", "Typography", HtmlRenderCapabilityKind.Css,
            Features("white-space", "nowrap", "Unicode/CJK line breaks", "overflow-wrap", "word-break", "hyphens", "hyphenate-character", "hyphenate-limit-chars", "hyphenate-limit-lines", "hyphenate-limit-last:always", "hyphenate-limit-zone", "letter-spacing", "word-spacing", "text-overflow:ellipsis", "line-clamp", "-webkit-line-clamp", "tab-size"),
            "Builds managed line boxes with preserved or collapsed whitespace, punctuation-safe Unicode/CJK and ordinary hyphen/slash opportunities from the shared typography owner, manual soft-hyphen control, caller-supplied automatic hyphenation, CSS character/line/last-line/zone limits, custom inserted characters without changing logical text, keep-all suppression of CJK-only boundaries, emergency wrapping, glyph and word spacing, inherited numeric tab stops, end ellipsis, and bounded multi-line clamping."),
        Fallback("text-shaping-fallback", "Typography", HtmlRenderCapabilityKind.Css,
            Features("OpenType shaping outside the managed script subset"),
            "Retains logical searchable text and uses deterministic glyph fallback when the configured shaping provider cannot shape a run.",
            HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported),
        Full("layout-flex", "Layout", HtmlRenderCapabilityKind.Css,
            Features("display:flex", "display:inline-flex", "flex-direction", "flex-wrap", "flex", "gap", "alignment"),
            "Lays out bounded row and column flex containers, wrapping, gaps, ordering, intrinsic bases, and alignment."),
        Fallback("layout-flex-fallback", "Layout", HtmlRenderCapabilityKind.Css,
            Features("unsupported flex property values", "unhandled flex formatting contexts"),
            "Uses diagnosed normal-flow or initial-value fallbacks for flex syntax outside the bounded contract.",
            HtmlRenderDiagnosticCodes.FlexLayoutPending,
            HtmlRenderDiagnosticCodes.FlexValueUnsupported),
        Full("layout-grid", "Layout", HtmlRenderCapabilityKind.Css,
            Features("display:grid", "display:inline-grid", "grid-template-*", "grid-auto-*", "repeat()", "minmax()", "min-content", "max-content", "fit-content()", "column and row subgrid", "first-baseline alignment", "auto item minima", "gap", "numeric and named placement"),
            "Lays out bounded explicit and implicit grids with intrinsic and automatic item contributions, responsive repeats, fixed and intrinsic minimum tracks, inherited parent column tracks, first-baseline alignment, numeric placement, named areas, and named lines."),
        Fallback("layout-grid-fallback", "Layout", HtmlRenderCapabilityKind.Css,
            Features("fractional automatic minima exceeding allocated shares", "subgrid without a resolved parent grid", "unsupported track functions", "unsupported auto-flow values"),
            "Uses diagnosed auto tracks or normal flow for grid syntax outside the bounded contract.",
            HtmlRenderDiagnosticCodes.GridLayoutPending,
            HtmlRenderDiagnosticCodes.GridValueUnsupported),
        Full("layout-columns", "Layout", HtmlRenderCapabilityKind.Css,
            Features("columns", "column-count", "column-width", "column-fill", "column-gap", "column-rule", "column-span"),
            "Builds bounded multi-column layout with balancing, rules, spanning blocks, and legal internal break points."),
        Fallback("layout-columns-fallback", "Layout", HtmlRenderCapabilityKind.Css,
            Features("unsupported column values", "cross-page atomic column fragments"),
            "Uses diagnosed initial values or a bounded forced fragment when column content cannot be split safely.",
            HtmlRenderDiagnosticCodes.ForcedFragment,
            HtmlRenderDiagnosticCodes.MultiColumnValueUnsupported),
        Full("layout-positioning", "Layout", HtmlRenderCapabilityKind.Css,
            Features("position:relative", "position:absolute", "position:fixed", "position:sticky", "position:running()", "insets", "z-index"),
            "Places relative, absolute, and fixed boxes in the declared containing-block model, captures named running elements for paged margin boxes, retains deterministic stacking, and captures sticky content at its stable static position."),
        Fallback("layout-positioning-fallback", "Layout", HtmlRenderCapabilityKind.Css,
            Features("unsupported positioning modes", "unresolved insets", "unsupported static anchors"),
            "Uses a diagnosed static anchor, auto inset, or normal-flow placement when authored positioning is outside the declared contract.",
            HtmlRenderDiagnosticCodes.PositionInsetUnsupported,
            HtmlRenderDiagnosticCodes.PositioningModeUnsupported,
            HtmlRenderDiagnosticCodes.PositionStaticAnchorFallback,
            HtmlRenderDiagnosticCodes.PositionStickyStatic),
        Full("layout-tables", "Layout", HtmlRenderCapabilityKind.Html,
            Features("table", "caption", "thead", "tbody", "tfoot", "rowspan", "colspan", "border-collapse", "table-layout"),
            "Builds bounded auto and fixed table grids with spans, collapsed or separate borders, captions, and repeated paged headers and footers."),
        Fallback("layout-tables-fallback", "Layout", HtmlRenderCapabilityKind.Html,
            Features("unsupported table property values", "malformed spanning grids"),
            "Normalizes malformed spans and uses diagnosed initial values outside the bounded table contract.",
            HtmlRenderDiagnosticCodes.TableValueUnsupported),
        Full("generated-content", "Generated content", HtmlRenderCapabilityKind.Css,
            Features("content", "attr()", "counter()", "counters()", "symbols()", "@counter-style", "counter-reset", "counter-set", "counter-increment", "::before", "::after"),
            "Generates quoted text, attributes, standard counters, functional counter styles, and document-scoped named cyclic, fixed, numeric, alphabetic, symbolic, or additive counter styles with bounded range, pad, negative, prefix, suffix, and fallback handling."),
        Fallback("generated-content-fallback", "Generated content", HtmlRenderCapabilityKind.Css,
            Features("unsupported content expressions", "unsupported counter definitions"),
            "Omits an unsupported generated fragment and emits a typed diagnostic without changing source-flow text.",
            HtmlRenderDiagnosticCodes.GeneratedContentUnsupported,
            HtmlRenderDiagnosticCodes.GeneratedCounterUnsupported),
        Full("list-markers", "Generated content", HtmlRenderCapabilityKind.Css,
            Features("list-style-type", "decimal", "decimal-leading-zero", "lower-alpha", "upper-alpha", "lower-roman", "upper-roman", "lower-greek", "cjk-decimal", "cjk-heavenly-stem", "cjk-earthly-branch", "cjk-ideographic", "japanese-informal", "japanese-formal", "korean-hangul-formal", "korean-hanja-informal", "korean-hanja-formal", "simp-chinese-informal", "simp-chinese-formal", "trad-chinese-informal", "trad-chinese-formal", "hiragana", "hiragana-iroha", "katakana", "katakana-iroha", "full-width", "symbols()", "@counter-style", "disc", "circle", "square", "quoted markers", "start", "reversed", "value"),
            "Formats standard, bounded longhand and alphabetic East Asian, functional, author-defined, and unordered markers through the same counter-style owner used by generated content while preserving canonical HTML list ordinals and standard or custom marker affixes."),
        Full("paged-page-rules", "Paged media", HtmlRenderCapabilityKind.PagedMedia,
            Features("@page", "size", "margin", ":first", ":left", ":right", "named pages", "page-local viewport units", "nested block continuation reflow", "inline continuation reflow", "table row continuation reflow", "wrapped flex-line continuation reflow", "margin boxes", "counter(page)", "counter(pages)", "string-set", "string()", "position:running()", "element()"),
            "Applies generic and named page masters, resolves geometry and viewport units per page, reconstructs logical source progress for text, nested blocks, safe table rows, and normal wrapped flex lines when masters change, and emits page counters, running strings, and clipped vector running-element snapshots with first, start, last, and first-except selection."),
        Fallback("paged-page-rules-fallback", "Paged media", HtmlRenderCapabilityKind.PagedMedia,
            Features("unsupported page selectors", "geometry changes across unsupported complex continuations", "unsupported margin content"),
            "Retains source-page layout for a continuation that cannot be reconstructed from logical source progress, or omits unsupported margin content, with a stable diagnostic while preserving document flow.",
            HtmlRenderDiagnosticCodes.PageMarginContentUnsupported,
            HtmlRenderDiagnosticCodes.PagePseudoGeometryPending,
            HtmlRenderDiagnosticCodes.PageSelectorPending,
            HtmlRenderDiagnosticCodes.PageSizeUnsupported),
        Full("paged-fragmentation", "Paged media", HtmlRenderCapabilityKind.PagedMedia,
            Features("break-before", "break-after", "break-inside", "orphans", "widows", "table header repetition", "table footer repetition"),
            "Honors bounded break constraints, text widows/orphans, legal flex/grid/column break points, and repeated table sections."),
        Fallback("paged-fragmentation-fallback", "Paged media", HtmlRenderCapabilityKind.PagedMedia,
            Features("oversized atomic visuals", "unsplittable replaced content"),
            "Uses a diagnosed forced fragment when an atomic visual cannot fit or split within the active page master.",
            HtmlRenderDiagnosticCodes.ForcedFragment,
            HtmlRenderDiagnosticCodes.VisualFragmentUnsupported),
        Full("media-queries", "Media queries", HtmlRenderCapabilityKind.Css,
            Features("screen", "print", "width", "height", "orientation", "resolution", "color", "monochrome", "prefers-color-scheme", "prefers-reduced-motion", "hover", "pointer"),
            "Evaluates media type, surface geometry, orientation, and caller-selected deterministic static-environment features."),
        Full("container-queries", "Container queries", HtmlRenderCapabilityKind.Css,
            Features("container", "container-name", "container-type", "named queries", "size features", "range syntax", "style() queries", "container query units"),
            "Evaluates bounded named or nearest-ancestor inline-size and size queries, chained ranges, computed-equivalent style queries, and container-relative units during managed layout. Layout-dependent container sizing outside the bounded block sizing model remains conservative."),
        Full("web-fonts", "Fonts", HtmlRenderCapabilityKind.Resource,
            Features("@font-face", "font-family", "font-style", "font-weight", "unicode-range", "TrueType glyf", "WOFF 1", "optional WOFF 2", "optional CFF/CFF2", "optional variable-font instances", "fallback packs"),
            "Loads bounded policy-approved font programs, selects constrained faces by Unicode scalar range, and routes measurement, complex layout, rasterization, and accessible PDF outlines through the configured shared provider. The dependency-free core owns TrueType glyf and WOFF 1; optional providers own WOFF 2, CFF/CFF2, and variable-font instances."),
        Fallback("web-fonts-fallback", "Fonts", HtmlRenderCapabilityKind.Resource,
            Features("optional font format without a configured provider", "provider-declined face", "invalid font programs", "unavailable font sources"),
            "Rejects an unusable face with a typed diagnostic and continues through the declared font fallback chain.",
            HtmlRenderDiagnosticCodes.FontFaceUnavailable,
            HtmlRenderDiagnosticCodes.FontFormatUnsupported),
        Full("images", "Images", HtmlRenderCapabilityKind.Resource,
            Features("img", "picture", "srcset", "PNG", "JPEG", "TIFF", "SVG", "WebP", "object-fit", "object-position", "image-orientation", "image-resolution"),
            "Resolves bounded responsive image candidates, normalizes embedded JPEG/TIFF orientation consistently across outputs, honors explicit CSS density, and paints supported raster and vector sources."),
        Full("svg", "SVG", HtmlRenderCapabilityKind.Resource,
            Features("paths", "basic shapes", "groups", "symbols", "use", "named/hex/rgb/hsl paint", "linearGradient", "radialGradient", "clipPath", "mask", "mix-blend-mode", "text", "tspan", "text-anchor", "dominant-baseline", "baseline-shift", "affine transforms"),
            "Maps the listed bounded SVG subset, reusable symbols, gradients, local clips and masks, positioned searchable text with baseline control, transforms, and standard blend modes into the shared Drawing scene."),
        Full("mathml", "MathML", HtmlRenderCapabilityKind.Html,
            Features("math", "mrow", "mtext", "mi", "mn", "mo", "mfrac", "msqrt", "mroot", "msup", "msub", "msubsup", "mmultiscripts", "mprescripts", "mfenced", "mtable", "mtr", "mtd", "menclose", "mphantom", "mover", "munder", "munderover", "semantics", "annotation", "display=block"),
            "Parses the listed Presentation MathML structures through the shared OfficeIMO.Core expression owner, lays them out as managed vectors with measured inline baselines, retains logical text in the shared scene and SVG text, and maps accessible names plus actual text into PDF output."),
        Fallback("mathml-fallback", "MathML", HtmlRenderCapabilityKind.Html,
            Features("unsupported Presentation MathML elements", "Content MathML"),
            "Retains supported descendants or logical fallback text and emits a typed diagnostic for structures outside the declared MathML contract.",
            HtmlRenderDiagnosticCodes.MathMlContentUnsupported),
        Fallback("svg-fallback", "SVG", HtmlRenderCapabilityKind.Resource,
            Features("unsupported SVG filters", "external references", "unsupported paint servers"),
            "Retains supported geometry and uses a diagnosed raster or omission fallback for SVG features outside the bounded vector contract.",
            HtmlRenderDiagnosticCodes.SvgContentUnsupported,
            HtmlRenderDiagnosticCodes.SvgRasterFallback),
        Full("pdf-metadata", "Output metadata", HtmlRenderCapabilityKind.Output,
            Features("title", "author", "subject", "description", "keywords", "creator", "generator", "language", "reading direction"),
            "Carries normalized document metadata into the shared render result; PDF output maps title, author, subject, keywords, language, and reading direction."),
        Full("pdf-accessibility", "Output accessibility", HtmlRenderCapabilityKind.Output,
            Features("tagged PDF", "document language", "reading order", "headings", "lists", "tables", "links", "alternate text", "bounded structural validation"),
            "Maps semantic groups into tagged PDF structures and exposes HtmlPdfAccessibilityValidator for deterministic language, parent-tree, hierarchy, marked-content, table, list, link, and figure checks."),
        Full("pdf-form-controls", "Interactive PDF forms", HtmlRenderCapabilityKind.Output,
            Features("text inputs", "password inputs", "file inputs", "text areas", "check boxes", "radio groups", "single-select", "multi-select", "required", "disabled and readonly", "maxlength", "accessible names", "static-output fallback"),
            "Retains standard HTML control semantics and geometry in the shared scene; PDF output emits positioned AcroForm widgets by default while image/SVG output, explicit static PDF mode, and diagnosed transformed, translucent, or clipped controls paint the same managed fallback visuals."),
        Rejected("resource-policy", "Resource safety", HtmlRenderCapabilityKind.Resource,
            Features("local files", "remote resources", "data URIs", "package resources", "hyperlinks"),
            "Rejects resources and links outside the caller-selected URL and host-resource policies before loading or emission.",
            "HtmlResourceRejectedByPolicy",
            "HyperlinkRejectedByPolicy",
            "ImageResourceRejectedByPolicy",
            "FontResourceRejectedByPolicy",
            "StylesheetResourceRejectedByPolicy"),
        Fallback("unsupported-static-features", "Static output boundary", HtmlRenderCapabilityKind.Output,
            Features("scroll state", "sticky state", "animation", "perspective transforms"),
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
