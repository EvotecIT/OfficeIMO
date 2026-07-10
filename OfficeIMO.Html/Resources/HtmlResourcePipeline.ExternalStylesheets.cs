namespace OfficeIMO.Html;

internal sealed class HtmlExternalStylesheetAnalysis {
    internal HtmlExternalStylesheetAnalysis(string css, IReadOnlyList<HtmlExternalStylesheetImport> imports) {
        Css = css;
        Imports = imports;
    }

    internal string Css { get; }
    internal IReadOnlyList<HtmlExternalStylesheetImport> Imports { get; }
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

        return new HtmlExternalStylesheetAnalysis(normalized, imports.AsReadOnly());
    }
}
