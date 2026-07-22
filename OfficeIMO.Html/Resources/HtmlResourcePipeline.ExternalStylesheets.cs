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
            if (!IsCssFunctionNameAt(normalized, match.Index, "url")
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
