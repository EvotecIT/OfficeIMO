namespace OfficeIMO.Html;

internal sealed class HtmlExternalStylesheetAnalysis {
    internal HtmlExternalStylesheetAnalysis(string css, IReadOnlyList<HtmlExternalStylesheetImport> imports, IReadOnlyList<HtmlResourceReference> fontResources) {
        Css = css;
        Imports = imports;
        FontResources = fontResources;
    }

    internal string Css { get; }
    internal IReadOnlyList<HtmlExternalStylesheetImport> Imports { get; }
    internal IReadOnlyList<HtmlResourceReference> FontResources { get; }
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
        HtmlUrlPolicy resourcePolicy = HtmlResourceUrlPolicy.Create(options.UrlPolicy);
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
                IsApplicableCssImport(import.ConditionText, options.MediaContext)));
        }

        foreach (HtmlCssFontFaceDefinition definition in ExtractFontFaces(normalized, options.MediaContext)) {
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

        return new HtmlExternalStylesheetAnalysis(normalized, imports.AsReadOnly(), fontResources.AsReadOnly());
    }
}
