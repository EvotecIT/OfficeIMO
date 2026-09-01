namespace OfficeIMO.Html;

internal sealed class HtmlExternalStylesheetAnalysis {
    internal HtmlExternalStylesheetAnalysis(
        string css,
        IReadOnlyList<HtmlExternalStylesheetImport> imports,
        IReadOnlyList<HtmlResourceReference> fontResources,
        IReadOnlyList<HtmlResourceReference> imageResources) {
        Css = css;
        Imports = imports;
        FontResources = fontResources;
        ImageResources = imageResources;
    }

    internal string Css { get; }
    internal IReadOnlyList<HtmlExternalStylesheetImport> Imports { get; }
    internal IReadOnlyList<HtmlResourceReference> FontResources { get; }
    internal IReadOnlyList<HtmlResourceReference> ImageResources { get; }
}

internal sealed class HtmlExternalStylesheetImport {
    internal HtmlExternalStylesheetImport(int start, int end, HtmlResourceReference reference, bool isApplicable) {
        Start = start;
        End = end;
        Reference = reference;
        IsApplicable = isApplicable;
    }

    internal int Start { get; }
    internal int End { get; }
    internal HtmlResourceReference Reference { get; }
    internal bool IsApplicable { get; }
}

public static partial class HtmlResourcePipeline {
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
        if (css == null) throw new ArgumentNullException(nameof(css));
        if (baseUri == null) throw new ArgumentNullException(nameof(baseUri));
        if (!baseUri.IsAbsoluteUri) throw new ArgumentException("The stylesheet base URI must be absolute.", nameof(baseUri));

        HtmlResourcePipelineOptions resolved = options ?? new HtmlResourcePipelineOptions();
        HtmlConversionLimits limits = resolved.Limits ?? HtmlConversionLimits.CreateUntrustedProfile();
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

    internal static HtmlExternalStylesheetAnalysis AnalyzeExternalStylesheet(string css, Uri baseUri, HtmlResourcePipelineOptions options) {
        string normalized = StripCssCommentsOutsideStrings(css ?? string.Empty);
        var imports = new List<HtmlExternalStylesheetImport>();
        var fontResources = new List<HtmlResourceReference>();
        var imageResources = new List<HtmlResourceReference>();
        HtmlUrlPolicy resourcePolicy = GetResourceUrlPolicy(options);
        foreach (CssImportReference import in ExtractCssImports(normalized)) {
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
                allowed ? string.Empty : GetDiagnosticCode(HtmlResourceKind.Stylesheet));
            imports.Add(new HtmlExternalStylesheetImport(
                import.Start,
                import.End,
                reference,
                IsApplicableCssImport(import.ConditionText, options)));
        }

        foreach (HtmlCssFontFaceDefinition definition in ExtractFontFaces(normalized, options)) {
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
                    allowed ? string.Empty : GetDiagnosticCode(HtmlResourceKind.Font)));
            }
        }

        List<SourceRange> inactiveRanges = GetInactiveCssRuleRanges(normalized, options);
        List<SourceRange> importRanges = imports.Select(import => new SourceRange(import.Start, import.End)).ToList();
        foreach (System.Text.RegularExpressions.Match match in CssUrlExpression.Matches(normalized)) {
            if (!IsValidCssUrlMatch(normalized, match)
                || !IsCssFunctionNameAt(normalized, match.Index, "url")
                || IsInsideCssString(normalized, match.Index)
                || IsInRanges(match.Index, importRanges)
                || IsInRanges(match.Index, inactiveRanges)
                || ClassifyCssUrl(normalized, match.Index) != HtmlResourceKind.Image) {
                continue;
            }

            string source = DecodeCssEscapes(match.Groups["url"].Value.Trim().Trim('\'', '"'));
            if (string.IsNullOrWhiteSpace(source) || IsFragmentOnlyReference(source)) {
                continue;
            }

            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(source, baseUri, resourcePolicy);
            bool allowed = !string.IsNullOrWhiteSpace(resolved) && IsResourceKindSchemeAllowed(HtmlResourceKind.Image, resolved);
            imageResources.Add(new HtmlResourceReference(
                HtmlResourceKind.Image,
                "style",
                "css-url",
                source,
                resolved,
                allowed,
                allowed ? string.Empty : GetDiagnosticCode(HtmlResourceKind.Image)));
        }

        return new HtmlExternalStylesheetAnalysis(normalized, imports.AsReadOnly(), fontResources.AsReadOnly(), imageResources.AsReadOnly());
    }
}
