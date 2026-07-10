namespace OfficeIMO.Html;

/// <summary>
/// Catalog of stable OfficeIMO HTML diagnostics and support remediation text.
/// </summary>
public static class HtmlDiagnosticCatalog {
    private static readonly IReadOnlyDictionary<string, HtmlDiagnosticDefinition> Definitions = new Dictionary<string, HtmlDiagnosticDefinition>(StringComparer.OrdinalIgnoreCase) {
        [HtmlRenderDiagnosticCodes.BackgroundImageLayerLimit] = RenderDefinition(HtmlRenderDiagnosticCodes.BackgroundImageLayerLimit, "PaintFidelity", HtmlDiagnosticSeverity.Warning, "Only the first CSS background-image layer was painted.", "Flatten multiple background layers into one image until shared layered paint is enabled."),
        [HtmlRenderDiagnosticCodes.BackgroundImageRepeatUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.BackgroundImageRepeatUnsupported, "PaintFidelity", HtmlDiagnosticSeverity.Warning, "A CSS background-repeat value used a single-image fallback.", "Use repeat, no-repeat, repeat-x, repeat-y, or a two-axis repeat/no-repeat value until space and round distribution are enabled."),
        [HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported, "PaintFidelity", HtmlDiagnosticSeverity.Warning, "A CSS background image value used a deterministic supported fallback.", "Use no-repeat with auto, contain, cover, or one/two absolute or percentage size values."),
        [HtmlRenderDiagnosticCodes.BackgroundImageTileLimitExceeded] = RenderDefinition(HtmlRenderDiagnosticCodes.BackgroundImageTileLimitExceeded, "ResourceLimit", HtmlDiagnosticSeverity.Error, "Repeated CSS background images exceeded the configured operation-wide tile limit.", "Increase MaxBackgroundImageTiles only for trusted documents or use a larger background tile."),
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
        [HtmlRenderDiagnosticCodes.EmptyTable] = RenderDefinition(HtmlRenderDiagnosticCodes.EmptyTable, "LayoutFidelity", HtmlDiagnosticSeverity.Info, "A table contained no renderable rows or cells.", "Add table rows and cells or remove the empty table."),
        [HtmlRenderDiagnosticCodes.ExternalImagePending] = RenderDefinition(HtmlRenderDiagnosticCodes.ExternalImagePending, "ResourcePolicy", HtmlDiagnosticSeverity.Warning, "An external image requires asynchronous resource resolution.", "Use RenderAsync with an application-supplied resource resolver or embed the image as a data URI."),
        [HtmlRenderDiagnosticCodes.ExternalStylesheetPending] = RenderDefinition(HtmlRenderDiagnosticCodes.ExternalStylesheetPending, "ResourcePolicy", HtmlDiagnosticSeverity.Warning, "An external stylesheet requires asynchronous resource resolution.", "Use RenderAsync with an application-supplied resource resolver or place trusted CSS in a style element."),
        [HtmlRenderDiagnosticCodes.FontDataUriInvalid] = RenderDefinition(HtmlRenderDiagnosticCodes.FontDataUriInvalid, "Typography", HtmlDiagnosticSeverity.Warning, "A font data URI could not be decoded.", "Provide a valid percent-encoded or base64 font data URI."),
        [HtmlRenderDiagnosticCodes.FontFaceInvalid] = RenderDefinition(HtmlRenderDiagnosticCodes.FontFaceInvalid, "Typography", HtmlDiagnosticSeverity.Warning, "An @font-face rule has no usable family descriptor.", "Provide a font-family descriptor and at least one usable src entry."),
        [HtmlRenderDiagnosticCodes.FontFaceUnavailable] = RenderDefinition(HtmlRenderDiagnosticCodes.FontFaceUnavailable, "Typography", HtmlDiagnosticSeverity.Warning, "No source from an @font-face rule was available.", "Use an allowed data URI or resolve the external font through RenderAsync."),
        [HtmlRenderDiagnosticCodes.FontFormatUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.FontFormatUnsupported, "Typography", HtmlDiagnosticSeverity.Warning, "A font source is not a supported TrueType glyf-outline font.", "Provide a TTF or TrueType-flavored OpenType face; WOFF, WOFF2, and CFF outlines require pre-conversion."),
        [HtmlRenderDiagnosticCodes.FlexLayoutPending] = RenderDefinition(HtmlRenderDiagnosticCodes.FlexLayoutPending, "LayoutFidelity", HtmlDiagnosticSeverity.Warning, "Flex layout currently falls back to normal flow.", "Use normal document flow for the current contract or wait for the dedicated flex formatting context."),
        [HtmlRenderDiagnosticCodes.ForcedFragment] = RenderDefinition(HtmlRenderDiagnosticCodes.ForcedFragment, "PagedMedia", HtmlDiagnosticSeverity.Warning, "Content had no safe break opportunity within one page.", "Add break opportunities or reduce the size of the unbreakable content."),
        [HtmlRenderDiagnosticCodes.GridLayoutPending] = RenderDefinition(HtmlRenderDiagnosticCodes.GridLayoutPending, "LayoutFidelity", HtmlDiagnosticSeverity.Warning, "Grid layout currently falls back to normal flow.", "Use tables or normal flow for the current contract or wait for the dedicated grid formatting context."),
        [HtmlRenderDiagnosticCodes.InlineImageFallback] = RenderDefinition(HtmlRenderDiagnosticCodes.InlineImageFallback, "LayoutFidelity", HtmlDiagnosticSeverity.Warning, "An inline image was represented by alternative text.", "Place the image as a block for the current contract or provide meaningful alternative text."),
        [HtmlRenderDiagnosticCodes.PageMarginContentUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.PageMarginContentUnsupported, "PagedMedia", HtmlDiagnosticSeverity.Warning, "A page-margin generated-content expression could not be represented.", "Use quoted text with counter(page) or counter(pages) until richer generated content is enabled."),
        [HtmlRenderDiagnosticCodes.PageMarginPositionUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.PageMarginPositionUnsupported, "PagedMedia", HtmlDiagnosticSeverity.Warning, "A page-margin position is not recognized by the direct renderer.", "Use one of the standard CSS top, bottom, left, right, or corner page-margin box names."),
        [HtmlRenderDiagnosticCodes.PagePseudoGeometryPending] = RenderDefinition(HtmlRenderDiagnosticCodes.PagePseudoGeometryPending, "PagedMedia", HtmlDiagnosticSeverity.Warning, "A pseudo-page size or margin declaration requires page-by-page body reflow.", "Keep body geometry in the generic @page rule until per-page reflow is enabled; pseudo-page margin content is still applied."),
        [HtmlRenderDiagnosticCodes.PageSelectorPending] = RenderDefinition(HtmlRenderDiagnosticCodes.PageSelectorPending, "PagedMedia", HtmlDiagnosticSeverity.Warning, "A complex page selector could not be applied per page.", "Use a generic, named, :first, :left, or :right @page selector, optionally combining one name with one supported pseudo-page."),
        [HtmlRenderDiagnosticCodes.PageSizeUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.PageSizeUnsupported, "PagedMedia", HtmlDiagnosticSeverity.Warning, "An @page size declaration could not be mapped.", "Use a supported named size or two absolute physical lengths."),
        [HtmlRenderDiagnosticCodes.RasterDecoderUnavailable] = RenderDefinition(HtmlRenderDiagnosticCodes.RasterDecoderUnavailable, "ImageFidelity", HtmlDiagnosticSeverity.Warning, "The dependency-free PNG backend cannot decode an image format retained for SVG or PDF.", "Use PNG, uncompressed BMP, first-frame GIF, or an application-provided pre-conversion."),
        [HtmlRenderDiagnosticCodes.ResourceByteLimitExceeded] = RenderDefinition(HtmlRenderDiagnosticCodes.ResourceByteLimitExceeded, "ResourcePolicy", HtmlDiagnosticSeverity.Warning, "A resource exceeded the configured per-resource byte limit.", "Reduce the resource or raise the explicit limit for trusted input."),
        [HtmlRenderDiagnosticCodes.ResourceCountLimitExceeded] = RenderDefinition(HtmlRenderDiagnosticCodes.ResourceCountLimitExceeded, "ResourcePolicy", HtmlDiagnosticSeverity.Error, "Resolved resources exceeded the configured count limit.", "Reduce the resource graph or raise the explicit count limit for trusted input."),
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
        [HtmlRenderDiagnosticCodes.TotalResourceByteLimitExceeded] = RenderDefinition(HtmlRenderDiagnosticCodes.TotalResourceByteLimitExceeded, "ResourcePolicy", HtmlDiagnosticSeverity.Error, "Resolved resources exceeded the total byte budget.", "Reduce resource volume or raise the explicit total limit for trusted input."),
        [HtmlRenderDiagnosticCodes.VisualFragmentUnsupported] = RenderDefinition(HtmlRenderDiagnosticCodes.VisualFragmentUnsupported, "PagedMedia", HtmlDiagnosticSeverity.Warning, "A visual could not cross a forced page boundary safely.", "Resize the visual or add a safe break before it.")
    };

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
}
