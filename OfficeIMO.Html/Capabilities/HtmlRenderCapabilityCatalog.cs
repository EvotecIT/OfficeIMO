namespace OfficeIMO.Html;

/// <summary>
/// Executable source of truth for versioned HTML document, rendering, and output compatibility contracts.
/// Entries classify exact subsets by processing stage and profile instead of implying complete specification support.
/// </summary>
public static partial class HtmlRenderCapabilityCatalog {
    /// <summary>Current machine-readable capability schema version.</summary>
    public const int SchemaVersion = 3;

    private static readonly IReadOnlyList<HtmlRenderCapability> Capabilities = new[] {
        Qualified("html-source-decoding", "HTML source", HtmlRenderCapabilityKind.Encoding,
            DocumentScope(HtmlCapabilityStage.SourceAndDecoding, HtmlCapabilitySpecificationIds.Encoding, HtmlCapabilitySpecificationIds.Html),
            Features("byte-order marks", "explicit caller encoding", "HTML meta charset", "web charset labels", "legacy code pages through the configured provider"),
            "Loads bounded source bytes through the configured encoding provider. Explicit caller encoding wins; otherwise the current path applies byte-order-mark and HTML meta-charset detection before producing the immutable source snapshot."),
        Qualified("html-document-parsing", "HTML document", HtmlRenderCapabilityKind.Html,
            DocumentScope(HtmlCapabilityStage.ParseAndPreserve, HtmlCapabilitySpecificationIds.Html),
            Features("full-document parsing", "HTML error recovery", "templates", "doctype and document mode", "source locations", "bounded native-to-owned projection"),
            "Parses a bounded inert full HTML document through the selected provider and projects it into OfficeIMO-owned nodes without exposing provider objects. Contextual fragment parsing, complete foreign-content qualification, and pre-allocation worker isolation remain outside this profile."),
        Qualified("html-owned-dom", "HTML document", HtmlRenderCapabilityKind.Dom,
            DocumentScope(HtmlCapabilityStage.DomAndQuery, HtmlCapabilitySpecificationIds.Dom, HtmlCapabilitySpecificationIds.Html),
            Features("owned nodes", "stable node IDs", "immutable snapshots", "mutable clones", "query selectors", "attributes", "text content", "detached editing history", "template contents"),
            "Exposes provider-neutral document and node contracts for query, inspection, cloning, bounded mutation, import, and immutable capture. Node IDs remain meaningful within an edit lineage and snapshot IDs distinguish independently frozen trees."),
        Qualified("html-serialization", "HTML document", HtmlRenderCapabilityKind.Html,
            DocumentScope(HtmlCapabilityStage.ParseAndPreserve, HtmlCapabilitySpecificationIds.Html),
            Features("owned-tree serialization", "canonical conversion HTML", "template contents", "bounded expanded output"),
            "Serializes the owned tree deterministically within configured limits. Edited output preserves document semantics and source locations where retained, without promising the original spelling of entities, whitespace inside tags, or invalid source syntax."),
        Qualified("css-cascade", "CSS cascade", HtmlRenderCapabilityKind.Css,
            CssScope(HtmlCapabilityStage.CascadeAndCompute, HtmlCapabilitySpecificationIds.CssCascade),
            Features("author stylesheets", "caller stylesheets", "inline styles", "!important", "inheritance", "custom properties", "@supports", "@layer", "revert-layer"),
            "Applies the bounded author cascade, selector specificity, cascade-layer ordering, layer rollback, inherited values, var() substitution, supported @supports conditions, and caller stylesheets appended after document styles."),
        Qualified("css-selectors", "CSS selectors", HtmlRenderCapabilityKind.Css,
            CssScope(HtmlCapabilityStage.DomAndQuery | HtmlCapabilityStage.CascadeAndCompute, HtmlCapabilitySpecificationIds.Selectors, HtmlCapabilitySpecificationIds.CssCascade),
            Features("type", "class", "id", "attribute", "combinators", "structural pseudo-classes", "CSS nesting", "::before", "::after", "::marker"),
            "Matches the documented selector subset, bounded parent-list and ampersand nesting including nested conditional rules, and generated before, after, and list-marker content; selectors outside the bounded subset do not match."),
        Qualified("css-length-units", "CSS values", HtmlRenderCapabilityKind.Css,
            CssScope(HtmlCapabilityStage.CascadeAndCompute | HtmlCapabilityStage.Layout, HtmlCapabilitySpecificationIds.CssValues),
            Features("px", "pt", "pc", "in", "cm", "mm", "q", "em", "rem", "%", "vw/vh/vmin/vmax", "sv*/lv*/dv* viewport units", "cqw/cqh/cqi/cqb/cqmin/cqmax"),
            "Resolves absolute, font-relative, percentage, static viewport-family, and bounded container-query lengths against the active layout references. Writing-mode-relative and font-metric-relative unit families remain outside this subset."),
        Qualified("css-length-math", "CSS values", HtmlRenderCapabilityKind.Css,
            CssScope(HtmlCapabilityStage.CascadeAndCompute | HtmlCapabilityStage.Layout, HtmlCapabilitySpecificationIds.CssValues),
            Features("calc()", "min()", "max()", "clamp()"),
            "Evaluates bounded nested length arithmetic with dimensional checks for addition, subtraction, multiplication, and division."),
        Qualified("css-color", "Color and paint", HtmlRenderCapabilityKind.Css,
            CssScope(HtmlCapabilityStage.CascadeAndCompute | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.CssColor),
            Features("named colors", "hex colors", "rgb()/rgba()", "hsl()/hsla()", "hwb()", "lab()/lch()", "oklab()/oklch()", "color() predefined spaces", "color-mix() in sRGB, linear sRGB, or OKLab", "transparent"),
            "Resolves the listed CSS Color 4 forms through the shared Drawing parser, including wide-gamut predefined spaces and premultiplied-alpha color mixing."),
        Fallback("css-color-fallback", "Color and paint", HtmlRenderCapabilityKind.Css,
            CssScope(HtmlCapabilityStage.CascadeAndCompute | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.CssColor),
            Features("relative color syntax", "unlisted color() spaces", "unlisted color-mix() interpolation spaces"),
            "Uses the property initial value and emits a typed diagnostic when a color expression is outside the declared static contract.",
            HtmlRenderDiagnosticCodes.ColorValueUnsupported,
            HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported,
            HtmlRenderDiagnosticCodes.BorderPaintValueUnsupported),
        Qualified("css-backgrounds", "Color and paint", HtmlRenderCapabilityKind.Css,
            CssScope(HtmlCapabilityStage.CascadeAndCompute | HtmlCapabilityStage.Layout | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.CssBackgrounds, HtmlCapabilitySpecificationIds.CssImages),
            Features("background-color", "background-image", "background-position", "background-repeat", "background-size", "background-origin", "background-clip", "background-attachment:scroll/local", "linear-gradient()", "repeating-linear-gradient()", "radial-gradient()", "repeating-radial-gradient()", "conic-gradient()", "repeating-conic-gradient()"),
            "Paints bounded image layers against independent border, padding, or content positioning and clipping boxes, clips the background color with the final layer contract, and retains native vector linear and radial gradients plus bounded vector conic-gradient expansions across raster, SVG, and PDF outputs."),
        Fallback("css-backgrounds-fallback", "Color and paint", HtmlRenderCapabilityKind.Css,
            CssScope(HtmlCapabilityStage.CascadeAndCompute | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.CssBackgrounds, HtmlCapabilitySpecificationIds.CssImages),
            Features("unsupported background layers", "unsupported repeat modes", "background-attachment:fixed"),
            "Uses a diagnosed initial layer or repeat value when authored syntax is outside the declared static contract.",
            HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported,
            HtmlRenderDiagnosticCodes.BackgroundImageRepeatUnsupported),
        Qualified("css-borders-effects", "Color and paint", HtmlRenderCapabilityKind.Css,
            CssScope(HtmlCapabilityStage.CascadeAndCompute | HtmlCapabilityStage.Layout | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.CssBackgrounds, HtmlCapabilitySpecificationIds.CssEffects),
            Features("border", "solid/dashed/dotted/double/groove/ridge/inset/outset", "collapsed table border conflict resolution", "border-radius", "outline", "outline-color:invert", "box-shadow", "text-shadow", "opacity", "transform", "clip-path basic shapes", "clip-path geometry boxes"),
            "Paints the declared border styles, including two-tone three-dimensional edges and collapsed-table conflict winners, radius, deterministic backdrop-inverted outlines, bounded multi-layer box and text shadows with vector blur approximations, opacity, two-dimensional affine transform, and basic-shape clip-path forms against margin, border, padding, or content reference boxes."),
        Fallback("css-borders-effects-fallback", "Color and paint", HtmlRenderCapabilityKind.Css,
            CssScope(HtmlCapabilityStage.CascadeAndCompute | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.CssBackgrounds, HtmlCapabilitySpecificationIds.CssEffects),
            Features("unsupported border paint", "unsupported radius", "unsupported shadow", "unsupported outline", "unsupported transform", "unsupported clip-path"),
            "Uses a typed omission or initial-value diagnostic when an effect is outside the declared static contract.",
            HtmlRenderDiagnosticCodes.BorderPaintValueUnsupported,
            HtmlRenderDiagnosticCodes.BorderRadiusValueUnsupported,
            HtmlRenderDiagnosticCodes.BoxShadowValueUnsupported,
            HtmlRenderDiagnosticCodes.TextShadowValueUnsupported,
            HtmlRenderDiagnosticCodes.TextShadowLayerLimit,
            HtmlRenderDiagnosticCodes.OutlinePaintValueUnsupported,
            HtmlRenderDiagnosticCodes.TransformValueUnsupported,
            HtmlRenderDiagnosticCodes.ClipPathValueUnsupported),
        Qualified("layout-block-inline", "Layout", HtmlRenderCapabilityKind.Css,
            CssScope(HtmlCapabilityStage.Layout | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.CssBoxLayout, HtmlCapabilitySpecificationIds.CssText),
            Features("block flow", "inline flow", "inline-block", "box sizing", "aspect-ratio", "margins", "padding", "line boxes", "box-decoration-break:slice/clone"),
            "Builds searchable block and inline layout with box-model and preferred-aspect-ratio sizing, line wrapping, and fragment-aware inline background, border, radius, shadow, and outline painting. Slice retains only document-outer inline edges across line and page continuation fragments; clone repeats the complete decoration on each fragment."),
        Qualified("bidi-text", "Typography", HtmlRenderCapabilityKind.Css,
            TypographyScope(HtmlCapabilityStage.Layout | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.CssWritingModes, HtmlCapabilitySpecificationIds.Unicode17),
            Features("dir=ltr", "dir=rtl", "LRE", "RLE", "LRO", "RLO", "PDF", "LRI", "RLI", "FSI", "PDI"),
            "Uses the shared Drawing resolver for bounded Unicode embeddings, overrides, isolates, logical text retention, and deterministic visual positioning."),
        Qualified("text-flow", "Typography", HtmlRenderCapabilityKind.Css,
            TypographyScope(HtmlCapabilityStage.Layout | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.CssText, HtmlCapabilitySpecificationIds.Unicode17),
            Features("white-space", "nowrap", "Unicode/CJK line breaks", "overflow-wrap", "word-break", "hyphens", "hyphenate-character", "hyphenate-limit-chars", "hyphenate-limit-lines", "hyphenate-limit-last:always", "hyphenate-limit-zone", "letter-spacing", "word-spacing", "text-decoration-line", "text-decoration-style", "text-decoration-color", "text-overflow:ellipsis", "line-clamp", "-webkit-line-clamp", "tab-size"),
            "Builds managed line boxes with preserved or collapsed whitespace, punctuation-safe Unicode/CJK and ordinary hyphen/slash opportunities from the shared typography owner, manual soft-hyphen control, caller-supplied automatic hyphenation, CSS character/line/last-line/zone limits, custom inserted characters without changing logical text, keep-all suppression of CJK-only boundaries, emergency wrapping, glyph and word spacing, independently colored native decoration patterns, inherited numeric tab stops, end ellipsis, and bounded multi-line clamping."),
        Qualified("vertical-text", "Typography", HtmlRenderCapabilityKind.Css,
            TypographyScope(HtmlCapabilityStage.Layout | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.CssWritingModes, HtmlCapabilitySpecificationIds.Unicode17),
            Features("writing-mode:vertical-rl", "writing-mode:vertical-lr", "writing-mode:sideways-rl", "writing-mode:sideways-lr", "vertical block flow", "text-orientation:mixed", "text-orientation:upright", "text-orientation:sideways", "logical margin/padding/border/inset/size properties", "ruby", "ruby-position", "ruby-align", "::first-letter", "::first-line"),
            "Maps logical box geometry through the active writing mode, advances ordinary vertical block children in right-to-left or left-to-right columns, lays out searchable vertical inline text with upright CJK and emoji plus sideways Latin glyphs, preserves logical extraction through positioned glyph groups, places scaled ruby annotations over or under horizontal and vertical bases, and applies punctuation-aware first-letter plus actual-wrap-aware first-line styling."),
        Fallback("text-shaping-fallback", "Typography", HtmlRenderCapabilityKind.Css,
            TypographyScope(HtmlCapabilityStage.Layout | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.CssFonts, HtmlCapabilitySpecificationIds.Unicode17),
            Features("OpenType shaping outside the managed script and lookup subset", "small-caps glyph synthesis"),
            "Retains logical searchable text and uses deterministic glyph fallback when the configured shaping provider cannot shape a run; managed small caps use diagnosed uppercase glyphs while browser output retains native CSS rendering.",
            HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported,
            HtmlRenderDiagnosticCodes.OpenTypeFeatureUnsupported,
            HtmlRenderDiagnosticCodes.FontVariantApproximated),
        Qualified("layout-flex", "Layout", HtmlRenderCapabilityKind.Css,
            CssScope(HtmlCapabilityStage.Layout, HtmlCapabilitySpecificationIds.CssFlexbox),
            Features("display:flex", "display:inline-flex", "flex-direction", "flex-wrap", "flex", "gap", "alignment"),
            "Lays out bounded row and column flex containers, wrapping, gaps, ordering, intrinsic bases, and alignment."),
        Fallback("layout-flex-fallback", "Layout", HtmlRenderCapabilityKind.Css,
            CssScope(HtmlCapabilityStage.Layout, HtmlCapabilitySpecificationIds.CssFlexbox),
            Features("unsupported flex property values", "unhandled flex formatting contexts"),
            "Uses diagnosed normal-flow or initial-value fallbacks for flex syntax outside the bounded contract.",
            HtmlRenderDiagnosticCodes.FlexLayoutPending,
            HtmlRenderDiagnosticCodes.FlexValueUnsupported),
        Qualified("layout-grid", "Layout", HtmlRenderCapabilityKind.Css,
            CssScope(HtmlCapabilityStage.Layout, HtmlCapabilitySpecificationIds.CssGrid),
            Features("display:grid", "display:inline-grid", "grid-template-*", "grid-auto-*", "repeat()", "minmax()", "min-content", "max-content", "fit-content()", "column and row subgrid", "first-baseline alignment", "auto item minima", "gap", "numeric and named placement"),
            "Lays out bounded explicit and implicit grids with intrinsic and automatic item contributions, responsive repeats, fixed and intrinsic minimum tracks, inherited parent column tracks, first-baseline alignment, numeric placement, named areas, and named lines."),
        Fallback("layout-grid-fallback", "Layout", HtmlRenderCapabilityKind.Css,
            CssScope(HtmlCapabilityStage.Layout, HtmlCapabilitySpecificationIds.CssGrid),
            Features("fractional automatic minima exceeding allocated shares", "subgrid without a resolved parent grid", "unsupported track functions", "unsupported auto-flow values"),
            "Uses diagnosed auto tracks or normal flow for grid syntax outside the bounded contract.",
            HtmlRenderDiagnosticCodes.GridLayoutPending,
            HtmlRenderDiagnosticCodes.GridValueUnsupported),
        Qualified("layout-columns", "Layout", HtmlRenderCapabilityKind.Css,
            CssScope(HtmlCapabilityStage.Layout | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.CssMulticol),
            Features("columns", "column-count", "column-width", "column-fill", "column-gap", "column-rule", "column-span"),
            "Builds bounded multi-column layout with balancing, rules, spanning blocks, and legal internal break points."),
        Fallback("layout-columns-fallback", "Layout", HtmlRenderCapabilityKind.Css,
            CssScope(HtmlCapabilityStage.Layout, HtmlCapabilitySpecificationIds.CssMulticol),
            Features("unsupported column values", "cross-page atomic column fragments"),
            "Uses diagnosed initial values or a bounded forced fragment when column content cannot be split safely.",
            HtmlRenderDiagnosticCodes.ForcedFragment,
            HtmlRenderDiagnosticCodes.MultiColumnValueUnsupported),
        Qualified("layout-positioning", "Layout", HtmlRenderCapabilityKind.Css,
            CssScope(HtmlCapabilityStage.Layout | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.CssPosition),
            Features("position:relative", "position:absolute", "position:fixed", "position:sticky", "position:running()", "insets", "z-index"),
            "Places relative, absolute, and fixed boxes in the declared containing-block model, captures named running elements for paged margin boxes, retains deterministic stacking, and captures sticky content at its stable static position."),
        Fallback("layout-positioning-fallback", "Layout", HtmlRenderCapabilityKind.Css,
            CssScope(HtmlCapabilityStage.Layout, HtmlCapabilitySpecificationIds.CssPosition),
            Features("unsupported positioning modes", "unresolved insets", "unsupported static anchors"),
            "Uses a diagnosed static anchor, auto inset, or normal-flow placement when authored positioning is outside the declared contract.",
            HtmlRenderDiagnosticCodes.PositionInsetUnsupported,
            HtmlRenderDiagnosticCodes.PositioningModeUnsupported,
            HtmlRenderDiagnosticCodes.PositionStaticAnchorFallback,
            HtmlRenderDiagnosticCodes.PositionStickyStatic),
        Qualified("layout-tables", "Layout", HtmlRenderCapabilityKind.Html,
            HtmlLayoutScope(HtmlCapabilityStage.Layout | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.Html, HtmlCapabilitySpecificationIds.CssTables),
            Features("table", "caption", "thead", "tbody", "tfoot", "CSS table row groups", "rowspan", "colspan", "border-collapse", "table-layout", "authored descendant intrinsic widths"),
            "Builds bounded auto and fixed table grids with intrinsic cell and fixed-width descendant contributions, spans, collapsed or separate borders, captions, CSS-classified row groups, and repeated paged headers and footers."),
        Fallback("layout-tables-fallback", "Layout", HtmlRenderCapabilityKind.Html,
            HtmlLayoutScope(HtmlCapabilityStage.Layout, HtmlCapabilitySpecificationIds.Html, HtmlCapabilitySpecificationIds.CssTables),
            Features("unsupported table property values", "malformed spanning grids"),
            "Normalizes malformed spans and uses diagnosed initial values outside the bounded table contract.",
            HtmlRenderDiagnosticCodes.TableValueUnsupported),
        Qualified("generated-content", "Generated content", HtmlRenderCapabilityKind.Css,
            CssScope(HtmlCapabilityStage.CascadeAndCompute | HtmlCapabilityStage.Layout | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.CssGeneratedContent),
            Features("content", "url() images", "mixed text and images", "attr()", "counter()", "counters()", "symbols()", "target-text()", "target-counter(page)", "target-counter(list-item)", "leader()", "quotes", "open-quote", "close-quote", "no-open-quote", "no-close-quote", "@counter-style", "counter-reset", "counter-set", "counter-increment", "::before", "::after", "::marker", "::footnote-call", "::footnote-marker", "counter(list-item)"),
            "Generates mixed secured images and quoted text, authored nested quote pairs, attributes, standard and cross-reference counters, artifact-safe leaders, page targets through bounded paginator reflow, functional counter styles, and document-scoped named cyclic, fixed, numeric, alphabetic, symbolic, or additive counter styles with bounded range, pad, negative, prefix, suffix, and fallback handling."),
        Fallback("generated-content-fallback", "Generated content", HtmlRenderCapabilityKind.Css,
            CssScope(HtmlCapabilityStage.CascadeAndCompute | HtmlCapabilityStage.Layout | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.CssGeneratedContent),
            Features("unsupported content expressions", "unsupported counter definitions"),
            "Omits an unsupported generated fragment and emits a typed diagnostic without changing source-flow text.",
            HtmlRenderDiagnosticCodes.GeneratedContentUnsupported,
            HtmlRenderDiagnosticCodes.GeneratedCounterUnsupported),
        Qualified("list-markers", "Generated content", HtmlRenderCapabilityKind.Css,
            CssScope(HtmlCapabilityStage.Layout | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.CssGeneratedContent),
            Features("list-style-type", "list-style-position", "inside", "outside", "list-style-image", "::marker", "content", "counter(list-item)", "decimal", "decimal-leading-zero", "lower-alpha", "upper-alpha", "lower-roman", "upper-roman", "lower-greek", "cjk-decimal", "cjk-heavenly-stem", "cjk-earthly-branch", "cjk-ideographic", "japanese-informal", "japanese-formal", "korean-hangul-formal", "korean-hanja-informal", "korean-hanja-formal", "simp-chinese-informal", "simp-chinese-formal", "trad-chinese-informal", "trad-chinese-formal", "hiragana", "hiragana-iroha", "katakana", "katakana-iroha", "full-width", "symbols()", "@counter-style", "disc", "circle", "square", "quoted markers", "start", "reversed", "value"),
            "Formats standard, bounded longhand and alphabetic East Asian, functional, author-defined, image, styled generated, and unordered markers through the same counter-style and secured resource owners used by generated content and images, with inside or hanging outside geometry and canonical HTML list ordinals."),
        Qualified("paged-page-rules", "Paged media", HtmlRenderCapabilityKind.PagedMedia,
            PagedScope(HtmlCapabilityStage.CascadeAndCompute | HtmlCapabilityStage.Layout | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.CssPagedMedia),
            Features("@page", "size", "margin", "bleed", "marks:crop/cross", "per-page TrimBox/BleedBox", ":first", ":left", ":right", "named pages", "page-local viewport units", "nested block continuation reflow", "inline continuation reflow", "table row continuation reflow", "wrapped flex-line continuation reflow", "margin boxes", "counter(page)", "counter(pages)", "string-set", "string()", "position:running()", "element()"),
            "Applies generic and named page masters, expands each print sheet around its trim page for resolved bleed and vector crop/registration marks, retains per-page TrimBox/BleedBox metadata for the PDF adapter, resolves geometry and viewport units per page, reconstructs logical source progress for text, nested blocks, safe table rows, and normal wrapped flex lines when masters change, and emits page counters, running strings, and clipped vector running-element snapshots with first, start, last, and first-except selection."),
        Fallback("paged-page-rules-fallback", "Paged media", HtmlRenderCapabilityKind.PagedMedia,
            PagedScope(HtmlCapabilityStage.CascadeAndCompute | HtmlCapabilityStage.Layout | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.CssPagedMedia),
            Features("unsupported page selectors", "geometry changes across unsupported complex continuations", "unsupported margin content"),
            "Retains source-page layout for a continuation that cannot be reconstructed from logical source progress, or omits unsupported margin content, with a stable diagnostic while preserving document flow.",
            HtmlRenderDiagnosticCodes.PageMarginContentUnsupported,
            HtmlRenderDiagnosticCodes.PagePseudoGeometryPending,
            HtmlRenderDiagnosticCodes.PageSelectorPending,
            HtmlRenderDiagnosticCodes.PageSizeUnsupported,
            HtmlRenderDiagnosticCodes.PageBleedUnsupported,
            HtmlRenderDiagnosticCodes.PageMarksUnsupported),
        Qualified("paged-footnotes", "Paged media", HtmlRenderCapabilityKind.PagedMedia,
            PagedScope(HtmlCapabilityStage.Layout | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.CssPagedMedia, HtmlCapabilitySpecificationIds.CssGeneratedContent),
            Features("float:footnote", "inline footnotes", "block footnotes", "::footnote-call", "::footnote-marker", "mixed marker text and images", "page-area body layout", "counter(footnote)", "same-page reservation", "next-page deferral", "long-note continuation", "named pages", "table/flex/grid/column descendants", "PDF named destinations", "PDF Note structure"),
            "Extracts notes from supported layout containers, lays note bodies against the actual call-page footnote area, preserves mixed secured-image and text markers, reserves bounded page-bottom areas, converges boundary reflow through deterministic next-page deferral, continues long notes at legal line breaks, preserves named page geometry, emits bidirectional internal PDF navigation, and tags note content in reading order. Continuous output keeps authored note content in normal flow."),
        Qualified("paged-fragmentation", "Paged media", HtmlRenderCapabilityKind.PagedMedia,
            PagedScope(HtmlCapabilityStage.Layout | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.CssPagedMedia),
            Features("break-before", "break-after", "break-inside", "orphans", "widows", "box-decoration-break across page continuations", "table row and row-group break avoidance", "rowspan-safe table breaks", "aligned oversized-row line breaks", "table header repetition", "table footer repetition", "stable repeated table semantics"),
            "Honors bounded break constraints, text widows/orphans, inline decoration continuity, legal flex/grid/column break points, rowspan-safe and aligned line-level table breaks, and repeated table sections with stable logical structure identity."),
        Fallback("paged-fragmentation-fallback", "Paged media", HtmlRenderCapabilityKind.PagedMedia,
            PagedScope(HtmlCapabilityStage.Layout, HtmlCapabilitySpecificationIds.CssPagedMedia),
            Features("oversized atomic visuals", "unsplittable replaced content"),
            "Uses a diagnosed forced fragment when an atomic visual cannot fit or split within the active page master.",
            HtmlRenderDiagnosticCodes.ForcedFragment,
            HtmlRenderDiagnosticCodes.VisualFragmentUnsupported),
        Qualified("media-queries", "Media queries", HtmlRenderCapabilityKind.Css,
            CssScope(HtmlCapabilityStage.CascadeAndCompute, HtmlCapabilitySpecificationIds.CssConditional),
            Features("screen", "print", "width", "height", "orientation", "resolution", "color", "monochrome", "prefers-color-scheme", "prefers-reduced-motion", "hover", "pointer"),
            "Evaluates media type, surface geometry, orientation, and caller-selected deterministic static-environment features."),
        Qualified("container-queries", "Container queries", HtmlRenderCapabilityKind.Css,
            CssScope(HtmlCapabilityStage.CascadeAndCompute | HtmlCapabilityStage.Layout, HtmlCapabilitySpecificationIds.CssConditional),
            Features("container", "container-name", "container-type", "named queries", "size features", "range syntax", "style() queries", "container query units"),
            "Evaluates bounded named or nearest-ancestor inline-size and size queries, chained ranges, computed-equivalent style queries, and container-relative units during managed layout. Layout-dependent container sizing outside the bounded block sizing model remains conservative."),
        Qualified("web-fonts", "Fonts", HtmlRenderCapabilityKind.Resource,
            TypographyScope(HtmlCapabilityStage.SourceAndDecoding | HtmlCapabilityStage.CascadeAndCompute | HtmlCapabilityStage.Layout | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.CssFonts, HtmlCapabilitySpecificationIds.Unicode17),
            Features("@font-face", "font-family", "font-style italic/oblique angles", "font-stretch", "font-weight 1-1000", "nearest-face matching", "unicode-range", "TrueType glyf", "WOFF 1", "single-face WOFF 2 on .NET 8+", "CFF/CFF2", "standardized variable-font instances with avar 1.0", "fallback packs"),
            "Loads bounded policy-approved font programs, selects constrained faces by Unicode scalar range plus CSS-compatible numeric weight, stretch, and italic/oblique-angle matching, treats variation selectors and join controls as shaping inputs, and routes logical text through measurement, complete layout, rasterization, and accessible PDF outlines. OfficeIMO owns TrueType glyf, WOFF 1, CFF/CFF2, and standardized variable-font instances with avar 1.0 segment maps directly; single-face WOFF 2 decoding is built in on .NET 8 and newer. Every outline program uses the bounded contract with cancellation, per-run character limits, and an operation-wide path-command budget."),
        Fallback("web-fonts-fallback", "Fonts", HtmlRenderCapabilityKind.Resource,
            TypographyScope(HtmlCapabilityStage.SourceAndDecoding | HtmlCapabilityStage.Layout | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.CssFonts, HtmlCapabilitySpecificationIds.Unicode17),
            Features("experimental avar 2.0 cross-axis mappings", "WOFF 2 on earlier target frameworks", "unsupported font containers or outline programs", "invalid font programs", "unavailable font sources"),
            "Rejects an unusable face with a typed diagnostic and continues through the declared font fallback chain.",
            HtmlRenderDiagnosticCodes.FontFaceUnavailable,
            HtmlRenderDiagnosticCodes.FontFormatUnsupported),
        Qualified("images", "Images", HtmlRenderCapabilityKind.Resource,
            ResourceScope(HtmlCapabilityStage.SourceAndDecoding | HtmlCapabilityStage.Layout | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.Html, HtmlCapabilitySpecificationIds.CssImages),
            Features("img", "picture", "srcset", "PNG", "JPEG", "TIFF", "SVG", "WebP", "object-fit", "object-position", "image-orientation", "image-resolution"),
            "Resolves bounded responsive image candidates, normalizes embedded JPEG/TIFF orientation consistently across outputs, honors explicit CSS density, and paints supported raster and vector sources."),
        Qualified("svg", "SVG", HtmlRenderCapabilityKind.Resource,
            ResourceScope(HtmlCapabilityStage.ParseAndPreserve | HtmlCapabilityStage.Layout | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.Svg2, HtmlCapabilitySpecificationIds.CssEffects),
            Features("inline svg", "paths", "basic shapes", "groups", "symbols", "use", "named/hex/rgb/hsl paint", "linearGradient", "radialGradient", "pattern fill", "context-fill/context-stroke", "marker-start", "marker-mid", "marker-end", "clipPath", "mask", "mix-blend-mode", "feDropShadow", "feGaussianBlur", "feOffset", "foreignObject inline XHTML", "text", "gradient/pattern text paint", "text stroke", "logical ActualText", "tspan", "textPath", "text-anchor", "dominant-baseline", "baseline-shift", "writing-mode", "text-orientation", "affine transforms"),
            "Treats inline SVG elements and SVG image resources as replaced content backed by the shared Drawing engine. Maps the listed bounded SVG subset, reusable symbols, gradients, object-bounding-box or user-space vector pattern fills, reusable start/middle/end marker scenes, local clips and masks, static drop-shadow/blur/offset filters, and caller-owned inline foreign-object viewports into the shared Drawing scene. OfficeIMO.Html supplies a depth- and node-bounded inline XHTML renderer with external resource resolution disabled. Positioned searchable text, referenced path placement, mixed-orientation vertical glyph layout, transforms, and standard blend modes remain vector. Blur is a bounded viewport-clipped vector approximation. Gradient-, pattern-, or stroke-painted glyphs remain vector outlines while one logical ActualText value is retained for SVG accessibility and PDF extraction."),
        Qualified("mathml", "MathML", HtmlRenderCapabilityKind.Html,
            HtmlLayoutScope(HtmlCapabilityStage.ParseAndPreserve | HtmlCapabilityStage.Layout | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.MathMlCore),
            Features("math", "mrow", "mtext", "mi", "mn", "mo", "mfrac", "msqrt", "mroot", "msup", "msub", "msubsup", "mmultiscripts", "mprescripts", "mfenced", "mtable", "mtr", "mtd", "menclose", "mphantom", "mover", "munder", "munderover", "semantics", "annotation", "display=block"),
            "Parses the listed Presentation MathML structures through the shared OfficeIMO.Core expression owner, lays them out as managed vectors with measured inline baselines, retains logical text in the shared scene and SVG text, and maps accessible names plus actual text into PDF output."),
        Fallback("mathml-fallback", "MathML", HtmlRenderCapabilityKind.Html,
            HtmlLayoutScope(HtmlCapabilityStage.ParseAndPreserve | HtmlCapabilityStage.Layout | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.MathMlCore),
            Features("unsupported Presentation MathML elements", "Content MathML"),
            "Retains supported descendants or logical fallback text and emits a typed diagnostic for structures outside the declared MathML contract.",
            HtmlRenderDiagnosticCodes.MathMlContentUnsupported),
        Fallback("svg-fallback", "SVG", HtmlRenderCapabilityKind.Resource,
            ResourceScope(HtmlCapabilityStage.ParseAndPreserve | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.Svg2),
            Features("dynamic or multi-input SVG filter graphs", "external references", "external foreignObject resources", "unsupported pattern strokes", "unsupported paint servers"),
            "Retains supported geometry and uses a diagnosed raster or omission fallback for SVG features outside the bounded vector contract.",
            HtmlRenderDiagnosticCodes.SvgContentUnsupported,
            HtmlRenderDiagnosticCodes.SvgRasterFallback),
        Qualified("pdf-metadata", "Output metadata", HtmlRenderCapabilityKind.Output,
            OutputScope(HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.Pdf20),
            Features("title", "author", "subject", "description", "keywords", "creator", "generator", "language", "reading direction"),
            "Carries normalized document metadata into the shared render result; PDF output maps title, author, subject, keywords, language, and reading direction."),
        Qualified("pdf-accessibility", "Output accessibility", HtmlRenderCapabilityKind.Output,
            OutputScope(HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.Pdf20),
            Features("tagged PDF", "document language", "reading order", "headings", "lists", "tables", "notes", "links", "named destinations", "alternate text", "bounded structural validation"),
            "Maps semantic groups, including notes, into tagged PDF structures, emits internal links through named destinations, and exposes HtmlPdfAccessibilityValidator for deterministic language, parent-tree, hierarchy, marked-content, table, list, link, and figure checks."),
        Qualified("pdf-form-controls", "Interactive PDF forms", HtmlRenderCapabilityKind.Output,
            OutputScope(HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.Pdf20, HtmlCapabilitySpecificationIds.Html),
            Features("text inputs", "password inputs", "file inputs", "text areas", "check boxes", "radio groups", "single-select", "multi-select", "required", "disabled and readonly", "maxlength", "accessible names", "static-output fallback"),
            "Retains standard HTML control semantics and geometry in the shared scene; PDF output emits positioned AcroForm widgets by default while image/SVG output, explicit static PDF mode, and diagnosed transformed, translucent, or clipped controls paint the same managed fallback visuals."),
        Rejected("resource-policy", "Resource safety", HtmlRenderCapabilityKind.Resource,
            ResourceScope(HtmlCapabilityStage.SourceAndDecoding | HtmlCapabilityStage.PaintAndOutput, HtmlCapabilitySpecificationIds.Html, HtmlCapabilitySpecificationIds.Url),
            Features("local files", "remote resources", "data URIs", "package resources", "hyperlinks"),
            "Rejects resources and links outside the caller-selected URL and host-resource policies before loading or emission.",
            "HtmlResourceRejectedByPolicy",
            "HyperlinkRejectedByPolicy",
            HtmlRenderDiagnosticCodes.HyperlinkTargetUnavailable,
            "ImageResourceRejectedByPolicy",
            "FontResourceRejectedByPolicy",
            "StylesheetResourceRejectedByPolicy"),
        Fallback("unsupported-static-features", "Static output boundary", HtmlRenderCapabilityKind.Output,
            StaticBoundaryScope(HtmlCapabilityStage.Layout | HtmlCapabilityStage.PaintAndOutput | HtmlCapabilityStage.RuntimeAndInteraction, HtmlCapabilitySpecificationIds.Html, HtmlCapabilitySpecificationIds.CssEffects, HtmlCapabilitySpecificationIds.CssPosition),
            Features("scroll state", "sticky state", "animation", "perspective transforms"),
            "Captures a deterministic static representation when a dynamic feature has a meaningful snapshot; otherwise the feature is omitted or uses its initial value.",
            HtmlRenderDiagnosticCodes.OverflowScrollSnapshot,
            HtmlRenderDiagnosticCodes.PositionStickyStatic,
            HtmlRenderDiagnosticCodes.TransformValueUnsupported),
        Ignored("active-content", "Active content", HtmlRenderCapabilityKind.Html,
            StaticBoundaryScope(HtmlCapabilityStage.ParseAndPreserve | HtmlCapabilityStage.RuntimeAndInteraction, HtmlCapabilitySpecificationIds.Html),
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

    private static HtmlRenderCapability Qualified(
        string id,
        string area,
        HtmlRenderCapabilityKind kind,
        CapabilityScope scope,
        IEnumerable<string> features,
        string behavior,
        params string[] diagnostics) =>
        Create(id, area, kind, scope, HtmlCapabilityCoverage.Qualified, HtmlCapabilityHandling.Native, features, behavior, diagnostics);

    private static HtmlRenderCapability Fallback(
        string id,
        string area,
        HtmlRenderCapabilityKind kind,
        CapabilityScope scope,
        IEnumerable<string> features,
        string behavior,
        params string[] diagnostics) =>
        Create(id, area, kind, scope, HtmlCapabilityCoverage.Partial, HtmlCapabilityHandling.Fallback, features, behavior, diagnostics);

    private static HtmlRenderCapability Ignored(
        string id,
        string area,
        HtmlRenderCapabilityKind kind,
        CapabilityScope scope,
        IEnumerable<string> features,
        string behavior,
        params string[] diagnostics) =>
        Create(id, area, kind, scope, HtmlCapabilityCoverage.Unsupported, HtmlCapabilityHandling.Ignored, features, behavior, diagnostics);

    private static HtmlRenderCapability Rejected(
        string id,
        string area,
        HtmlRenderCapabilityKind kind,
        CapabilityScope scope,
        IEnumerable<string> features,
        string behavior,
        params string[] diagnostics) =>
        Create(id, area, kind, scope, HtmlCapabilityCoverage.Qualified, HtmlCapabilityHandling.Rejected, features, behavior, diagnostics);

    private static HtmlRenderCapability Create(
        string id,
        string area,
        HtmlRenderCapabilityKind kind,
        CapabilityScope scope,
        HtmlCapabilityCoverage coverage,
        HtmlCapabilityHandling handling,
        IEnumerable<string> features,
        string behavior,
        IEnumerable<string> diagnostics) {
        string[] featureArray = features.ToArray();
        IEnumerable<string> limitations = handling == HtmlCapabilityHandling.Native
            ? Array.Empty<string>()
            : featureArray;
        HtmlCapabilityProfileBinding[] bindings = scope.ProfileIds.Select(profileId =>
            new HtmlCapabilityProfileBinding(
                profileId,
                coverage,
                handling,
                HtmlCapabilityMaturity.Required,
                PromotionFor(profileId),
                scope.ProviderIds,
                scope.SpecificationIds,
                scope.GetEvidenceIds(profileId, id),
                scope.OptionalProviderIds)).ToArray();
        return new HtmlRenderCapability(id, area, kind, scope.Stages, bindings, featureArray, behavior, limitations, diagnostics);
    }

    private static CapabilityScope DocumentScope(HtmlCapabilityStage stages, params string[] specifications) =>
        ScopeByProfile(stages,
            new[] { HtmlCapabilityProfileIds.WebDocumentV1, HtmlCapabilityProfileIds.StaticScreenV1, HtmlCapabilityProfileIds.PagedPrintV1 },
            new[] { HtmlCapabilityProviderIds.OfficeIMOHtmlCore, HtmlCapabilityProviderIds.AngleSharpHtml },
            specifications,
            DocumentEvidenceIds);

    private static CapabilityScope CssScope(HtmlCapabilityStage stages, params string[] specifications) =>
        ScopeByProfile(stages,
            StaticAndPagedProfiles(),
            new[] { HtmlCapabilityProviderIds.OfficeIMOHtml, HtmlCapabilityProviderIds.AngleSharpHtml, HtmlCapabilityProviderIds.AngleSharpCss },
            WithCssSnapshot(specifications),
            StaticEvidenceIds);

    private static CapabilityScope TypographyScope(HtmlCapabilityStage stages, params string[] specifications) =>
        ScopeByProfile(stages,
            StaticAndPagedProfiles(),
            new[] { HtmlCapabilityProviderIds.OfficeIMOHtml, HtmlCapabilityProviderIds.OfficeIMOCore, HtmlCapabilityProviderIds.AngleSharpCss },
            WithCssSnapshot(specifications),
            StaticEvidenceIds,
            new[] { HtmlCapabilityProviderIds.CallerTextShaper });

    private static CapabilityScope HtmlLayoutScope(HtmlCapabilityStage stages, params string[] specifications) =>
        ScopeByProfile(stages,
            StaticAndPagedProfiles(),
            new[] { HtmlCapabilityProviderIds.OfficeIMOHtml, HtmlCapabilityProviderIds.OfficeIMOCore, HtmlCapabilityProviderIds.AngleSharpHtml, HtmlCapabilityProviderIds.AngleSharpCss },
            specifications,
            StaticEvidenceIds);

    private static CapabilityScope PagedScope(HtmlCapabilityStage stages, params string[] specifications) =>
        ScopeByProfile(stages,
            new[] { HtmlCapabilityProfileIds.PagedPrintV1 },
            new[] { HtmlCapabilityProviderIds.OfficeIMOHtml, HtmlCapabilityProviderIds.OfficeIMOHtmlPdf, HtmlCapabilityProviderIds.OfficeIMOPdf, HtmlCapabilityProviderIds.AngleSharpHtml, HtmlCapabilityProviderIds.AngleSharpCss },
            WithCssSnapshot(specifications),
            StaticEvidenceIds);

    private static CapabilityScope ResourceScope(HtmlCapabilityStage stages, params string[] specifications) =>
        ScopeByProfile(stages,
            StaticAndPagedProfiles(),
            new[] { HtmlCapabilityProviderIds.OfficeIMOHtml, HtmlCapabilityProviderIds.OfficeIMOCore, HtmlCapabilityProviderIds.AngleSharpHtml },
            specifications,
            StaticEvidenceIds);

    private static CapabilityScope OutputScope(HtmlCapabilityStage stages, params string[] specifications) =>
        ScopeByProfile(stages,
            new[] { HtmlCapabilityProfileIds.PagedPrintV1 },
            new[] { HtmlCapabilityProviderIds.OfficeIMOHtmlPdf, HtmlCapabilityProviderIds.OfficeIMOPdf },
            specifications,
            StaticEvidenceIds);

    private static CapabilityScope StaticBoundaryScope(HtmlCapabilityStage stages, params string[] specifications) =>
        ScopeByProfile(stages,
            StaticAndPagedProfiles(),
            new[] { HtmlCapabilityProviderIds.OfficeIMOHtml, HtmlCapabilityProviderIds.AngleSharpHtml, HtmlCapabilityProviderIds.AngleSharpCss },
            specifications,
            StaticEvidenceIds);

    private static CapabilityScope Scope(
        HtmlCapabilityStage stages,
        IEnumerable<string> profiles,
        IEnumerable<string> providers,
        IEnumerable<string> specifications,
        IEnumerable<string> evidence,
        IEnumerable<string>? optionalProviders = null) =>
        new CapabilityScope(stages, profiles, providers, specifications, (_, _) => evidence, optionalProviders);

    private static CapabilityScope ScopeByProfile(
        HtmlCapabilityStage stages,
        IEnumerable<string> profiles,
        IEnumerable<string> providers,
        IEnumerable<string> specifications,
        Func<string, string, IEnumerable<string>> evidence,
        IEnumerable<string>? optionalProviders = null) =>
        new CapabilityScope(stages, profiles, providers, specifications, evidence, optionalProviders);

    private static string[] StaticAndPagedProfiles() =>
        new[] { HtmlCapabilityProfileIds.StaticScreenV1, HtmlCapabilityProfileIds.PagedPrintV1 };

    private static string[] DocumentEvidenceIds(string profileId, string capabilityId) {
        if (string.Equals(profileId, HtmlCapabilityProfileIds.WebDocumentV1, StringComparison.OrdinalIgnoreCase)) {
            return new[] { HtmlCapabilityEvidenceIds.OfficeIMOHtmlTests, HtmlCapabilityEvidenceIds.DocumentV1 };
        }
        HtmlCapabilityEvidenceSelection[] selections;
        string selectedEvidence;
        if (string.Equals(profileId, HtmlCapabilityProfileIds.PagedPrintV1, StringComparison.OrdinalIgnoreCase)) {
            selections = CreateH4V2PagedSelections();
            selectedEvidence = HtmlCapabilityEvidenceIds.H4PagedV2;
        } else {
            selections = CreateH4V2ScreenSelections();
            selectedEvidence = HtmlCapabilityEvidenceIds.H4ScreenV2;
        }
        return selections.Any(selection => string.Equals(selection.CapabilityId, capabilityId, StringComparison.OrdinalIgnoreCase))
            ? new[] { HtmlCapabilityEvidenceIds.OfficeIMOHtmlTests, HtmlCapabilityEvidenceIds.DocumentV1, selectedEvidence }
            : new[] { HtmlCapabilityEvidenceIds.OfficeIMOHtmlTests, HtmlCapabilityEvidenceIds.DocumentV1 };
    }

    private static string[] StaticEvidenceIds(string profileId, string capabilityId) {
        HtmlCapabilityEvidenceSelection[] selections;
        string selectedEvidence;
        if (string.Equals(profileId, HtmlCapabilityProfileIds.PagedPrintV1, StringComparison.OrdinalIgnoreCase)) {
            selections = CreateH4V2PagedSelections();
            selectedEvidence = HtmlCapabilityEvidenceIds.H4PagedV2;
        } else {
            selections = CreateH4V2ScreenSelections();
            selectedEvidence = HtmlCapabilityEvidenceIds.H4ScreenV2;
        }
        if (selections.Any(selection => string.Equals(selection.CapabilityId, capabilityId, StringComparison.OrdinalIgnoreCase))) {
            return new[] { HtmlCapabilityEvidenceIds.OfficeIMOHtmlTests, selectedEvidence };
        }
        return new[] {
            HtmlCapabilityEvidenceIds.OfficeIMOHtmlTests,
            string.Equals(profileId, HtmlCapabilityProfileIds.PagedPrintV1, StringComparison.OrdinalIgnoreCase)
                ? HtmlCapabilityEvidenceIds.H4PagedV1
                : HtmlCapabilityEvidenceIds.H4ScreenV1
        };
    }

    private static string[] WithCssSnapshot(IEnumerable<string> specifications) =>
        new[] { HtmlCapabilitySpecificationIds.CssSnapshot }.Concat(specifications).ToArray();

    private static HtmlCapabilityPromotionState PromotionFor(string profileId) =>
        string.Equals(profileId, HtmlCapabilityProfileIds.WebDocumentV1, StringComparison.OrdinalIgnoreCase)
            ? HtmlCapabilityPromotionState.QualifiedOptIn
            : HtmlCapabilityPromotionState.StableDefault;

    private sealed class CapabilityScope {
        internal CapabilityScope(
            HtmlCapabilityStage stages,
            IEnumerable<string> profileIds,
            IEnumerable<string> providerIds,
            IEnumerable<string> specificationIds,
            Func<string, string, IEnumerable<string>> evidenceIds,
            IEnumerable<string>? optionalProviderIds) {
            if (stages == HtmlCapabilityStage.None) throw new ArgumentOutOfRangeException(nameof(stages));
            Stages = stages;
            ProfileIds = HtmlCapabilityContractValue.Normalize(profileIds, nameof(profileIds));
            ProviderIds = HtmlCapabilityContractValue.Normalize(providerIds, nameof(providerIds));
            OptionalProviderIds = HtmlCapabilityContractValue.Normalize(optionalProviderIds ?? Array.Empty<string>(), nameof(optionalProviderIds));
            SpecificationIds = HtmlCapabilityContractValue.Normalize(specificationIds, nameof(specificationIds));
            _evidenceIds = evidenceIds ?? throw new ArgumentNullException(nameof(evidenceIds));
        }

        internal HtmlCapabilityStage Stages { get; }
        internal IReadOnlyList<string> ProfileIds { get; }
        internal IReadOnlyList<string> ProviderIds { get; }
        internal IReadOnlyList<string> OptionalProviderIds { get; }
        internal IReadOnlyList<string> SpecificationIds { get; }
        private readonly Func<string, string, IEnumerable<string>> _evidenceIds;

        internal IReadOnlyList<string> GetEvidenceIds(string profileId, string capabilityId) =>
            HtmlCapabilityContractValue.Normalize(_evidenceIds(profileId, capabilityId), nameof(_evidenceIds));
    }
}
