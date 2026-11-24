namespace OfficeIMO.Html;

internal sealed class HtmlExternalStylesheetAnalysis {
    internal HtmlExternalStylesheetAnalysis(
        string css,
        IReadOnlyList<HtmlExternalStylesheetImport> imports,
        IReadOnlyList<HtmlResourceReference> fontResources,
        IReadOnlyList<HtmlResourceReference> imageResources,
        IReadOnlyList<HtmlResourceReference> deferredImageResources) {
        Css = css;
        Imports = imports;
        FontResources = fontResources;
        ImageResources = imageResources;
        DeferredImageResources = deferredImageResources;
    }

    internal string Css { get; }
    internal IReadOnlyList<HtmlExternalStylesheetImport> Imports { get; }
    internal IReadOnlyList<HtmlResourceReference> FontResources { get; }
    internal IReadOnlyList<HtmlResourceReference> ImageResources { get; }
    internal IReadOnlyList<HtmlResourceReference> DeferredImageResources { get; }
}

internal sealed class HtmlExternalStylesheetImport {
    internal HtmlExternalStylesheetImport(int start, int end, HtmlResourceReference reference, bool isApplicable,
        bool hasLayerCondition, bool hasUnknownSupportsCondition, bool hasUnknownMediaCondition) {
        Start = start;
        End = end;
        Reference = reference;
        IsApplicable = isApplicable;
        HasLayerCondition = hasLayerCondition;
        HasUnknownSupportsCondition = hasUnknownSupportsCondition;
        HasUnknownMediaCondition = hasUnknownMediaCondition;
    }

    internal int Start { get; }
    internal int End { get; }
    internal HtmlResourceReference Reference { get; }
    internal bool IsApplicable { get; }
    internal bool HasLayerCondition { get; }
    internal bool HasUnknownSupportsCondition { get; }
    internal bool HasUnknownMediaCondition { get; }
}

public static partial class HtmlResourcePipeline {
    /// <summary>
    /// Attempts to decode stylesheet bytes using CSS encoding precedence: byte-order mark,
    /// transport content-type charset, <c>@charset</c>, then strict UTF-8.
    /// </summary>
    /// <param name="bytes">Encoded stylesheet bytes.</param>
    /// <param name="contentType">Optional transport content type containing a charset.</param>
    /// <param name="css">Decoded stylesheet text when successful.</param>
    /// <param name="encodingProvider">Optional web charset provider for this decoding operation.</param>
    /// <returns><see langword="true"/> when the stylesheet encoding is supported and valid.</returns>
    public static bool TryDecodeStylesheet(byte[] bytes, string? contentType, out string css, Dom.IHtmlEncodingProvider? encodingProvider = null) =>
        (encodingProvider == null ? HtmlTextEncodingResolver.Default : new HtmlTextEncodingResolver(encodingProvider)).TryDecodeCss(bytes, contentType, out css);

    /// <summary>
    /// Parses an external stylesheet and builds the same policy-aware resource manifest used by
    /// the HTML renderer for imports, fonts, and images.
    /// </summary>
    /// <param name="css">Stylesheet source text.</param>
    /// <param name="baseUri">Absolute URI used to resolve relative stylesheet references.</param>
    /// <param name="options">Optional resource policy and media-context settings.</param>
    /// <returns>A manifest containing applicable imports and referenced font and image resources.</returns>
    public static HtmlResourceManifest BuildStylesheetManifest(
        string css,
        Uri baseUri,
        HtmlResourcePipelineOptions? options = null) {
        ArgumentNullException.ThrowIfNull(css);
        ArgumentNullException.ThrowIfNull(baseUri);
        if (!baseUri.IsAbsoluteUri) throw new ArgumentException("The stylesheet base URI must be absolute.", nameof(baseUri));

        HtmlResourcePipelineOptions resolved = options ?? new HtmlResourcePipelineOptions();
        HtmlConversionLimits limits = resolved.Limits ?? HtmlConversionLimits.UntrustedDefaults;
        limits.Validate();
        new HtmlCssByteBudget(limits).ReserveOrThrow(css);
        HtmlExternalStylesheetAnalysis analysis = AnalyzeExternalStylesheet(css, baseUri, resolved);
        var manifest = new HtmlResourceManifest();
        foreach (HtmlExternalStylesheetImport import in analysis.Imports) {
            if (import.IsApplicable) manifest.Add(import.Reference);
        }
        foreach (HtmlResourceReference resource in analysis.FontResources) manifest.Add(resource);
        foreach (HtmlResourceReference resource in analysis.ImageResources) manifest.Add(resource);
        return manifest;
    }

    internal static HtmlExternalStylesheetAnalysis AnalyzeExternalStylesheet(string css, Uri baseUri, HtmlResourcePipelineOptions options,
        bool includeInactiveResources = false) => AnalyzeExternalStylesheet(css, baseUri, options, includeInactiveResources, null);

    private static HtmlExternalStylesheetAnalysis AnalyzeExternalStylesheet(string css, Uri baseUri, HtmlResourcePipelineOptions options,
        bool includeInactiveResources, ExternalStylesheetSyntax? syntax) {
        string normalized = syntax?.Text ?? StripCssCommentsOutsideStrings(css ?? string.Empty);
        var imports = new List<HtmlExternalStylesheetImport>();
        var fontResources = new List<HtmlResourceReference>();
        var imageResources = new List<HtmlResourceReference>();
        var deferredImageResources = new List<HtmlResourceReference>();
        HtmlUrlPolicy resourcePolicy = GetResourceUrlPolicy(options);
        foreach (CssImportReference import in syntax?.Imports ?? ExtractCssImports(normalized)) {
            string source = DecodeCssEscapes(import.Source);
            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(source, baseUri, resourcePolicy);
            bool allowed = !string.IsNullOrWhiteSpace(resolved) && IsResourceKindSchemeAllowed(HtmlResourceKind.Stylesheet, resolved);
            var reference = new HtmlResourceReference(
                HtmlResourceKind.Stylesheet,
                "style",
                "css-import",
                source,
                resolved,
                allowed,
                allowed ? string.Empty : GetDiagnosticCode(HtmlResourceKind.Stylesheet), baseUri, resourcePolicy.ResolvedUrlTransform);
            bool isApplicable = IsApplicableCssImport(
                import.ConditionText,
                options,
                out bool hasUnknownSupportsCondition,
                out bool hasUnknownMediaCondition);
            imports.Add(new HtmlExternalStylesheetImport(
                import.Start,
                import.End,
                reference,
                isApplicable,
                HasCssImportLayerCondition(import.ConditionText),
                hasUnknownSupportsCondition,
                hasUnknownMediaCondition));
        }

        IReadOnlyList<HtmlCssFontFaceDefinition> definitions;
        if (syntax == null) definitions = ExtractFontFaces(normalized, options, includeInactiveResources);
        else {
            var activeFaces = new List<HtmlCssFontFaceDefinition>();
            syntax.FontFaces.AddApplicable(options, includeInactiveResources, activeFaces);
            definitions = activeFaces;
        }
        foreach (HtmlCssFontFaceDefinition definition in definitions) {
            foreach (string source in ExtractFontFaceUrls(definition.Source)) {
                string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(source, baseUri, resourcePolicy);
                bool allowed = !string.IsNullOrWhiteSpace(resolved) && IsResourceKindSchemeAllowed(HtmlResourceKind.Font, resolved);
                fontResources.Add(new HtmlResourceReference(
                    HtmlResourceKind.Font,
                    "style",
                    "font-face-src",
                    source,
                    resolved,
                    allowed,
                    allowed ? string.Empty : GetDiagnosticCode(HtmlResourceKind.Font), baseUri, resourcePolicy.ResolvedUrlTransform));
            }
        }

        List<SourceRange> inactiveRanges = includeInactiveResources ? new List<SourceRange>()
            : syntax?.GetInactiveRanges(options) ?? GetInactiveCssRuleRanges(normalized, options);
        List<SourceRange> importRanges = imports.Select(import => new SourceRange(import.Start, import.End)).ToList();
        foreach (StylesheetImageReference image in syntax?.Images ?? ReadStylesheetImages(normalized, importRanges, options.DeferPageMarginBackgrounds)) {
            if (IsInRanges(image.Start, inactiveRanges)) continue;
            AddImage(image.Source, image.Attribute, image.CustomProperty || options.DeferPageMarginBackgrounds && image.PageRule);
        }

        return new HtmlExternalStylesheetAnalysis(normalized, imports.AsReadOnly(), fontResources.AsReadOnly(), imageResources.AsReadOnly(), deferredImageResources.AsReadOnly());

        void AddImage(string source, string attribute, bool defer) {
            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(source, baseUri, resourcePolicy);
            bool allowed = !string.IsNullOrWhiteSpace(resolved) && IsResourceKindSchemeAllowed(HtmlResourceKind.Image, resolved);
            var reference = new HtmlResourceReference(HtmlResourceKind.Image, "style", attribute, source, resolved, allowed,
                allowed ? string.Empty : GetDiagnosticCode(HtmlResourceKind.Image), baseUri, resourcePolicy.ResolvedUrlTransform);
            // The final document cascade decides which generated box exists and which declaration wins.
            if (defer) deferredImageResources.Add(reference);
            else imageResources.Add(reference);
        }
    }

    private static IEnumerable<StylesheetImageReference> ReadStylesheetImages(string css, List<SourceRange> importRanges, bool includePageRule = true) {
        foreach (System.Text.RegularExpressions.Match match in CssUrlExpression.Matches(css)) {
            if (!IsValidCssUrlMatch(css, match)
                || !IsCssFunctionNameAt(css, match.Index, "url")
                || IsInsideCssString(css, match.Index)
                || IsInRanges(match.Index, importRanges)) continue;
            bool custom = IsCustomPropertyUrl(css, match.Index);
            if (!custom && ClassifyCssUrl(css, match.Index) != HtmlResourceKind.Image) continue;
            string source = DecodeCssEscapes(match.Groups["url"].Value.Trim().Trim('\'', '"'));
            if (!string.IsNullOrWhiteSpace(source) && !IsFragmentOnlyReference(source))
                yield return new StylesheetImageReference(source, "css-url", match.Index, custom, includePageRule && !custom && IsInCssPageRule(css, match.Index));
        }
        foreach (CssStringUrlReference reference in ExtractImageSetStringUrls(css)) {
            if (IsInRanges(reference.Start, importRanges)) continue;
            bool custom = TryGetCustomPropertyName(css, reference.Start, out _);
            if (!custom && ClassifyCssUrl(css, reference.Start) != HtmlResourceKind.Image) continue;
            string source = DecodeCssEscapes(reference.Source);
            if (!string.IsNullOrWhiteSpace(source) && !IsFragmentOnlyReference(source))
                yield return new StylesheetImageReference(source, "css-image-set", reference.Start, custom, includePageRule && !custom && IsInCssPageRule(css, reference.Start));
        }
    }

    private static bool HasCssImportLayerCondition(string conditionText) {
        string remaining = conditionText.TrimStart();
        return TryConsumeCssImportFunctionCondition(remaining, "layer", out _, out _)
            || StartsWithCssIdentifier(remaining, "layer");
    }
}
