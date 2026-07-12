namespace OfficeIMO.Html;

/// <summary>RTF document plus structured diagnostics from HTML import.</summary>
public sealed class HtmlToRtfResult : HtmlConversionResult<RtfDocument> {
    internal HtmlToRtfResult(RtfDocument document, IEnumerable<HtmlDiagnostic> diagnostics) : base(document) {
        Diagnostics.AddRange(diagnostics);
    }

    /// <summary>Imported RTF document.</summary>
    public RtfDocument Document => Artifact;
}

/// <summary>Semantic HTML plus structured diagnostics from RTF export.</summary>
public sealed class RtfToHtmlResult : HtmlConversionResult<string> {
    internal RtfToHtmlResult(string html, IEnumerable<HtmlDiagnostic> diagnostics) : base(html) {
        Diagnostics.AddRange(diagnostics);
    }

    /// <summary>Exported semantic HTML.</summary>
    public string Html => Artifact;
}
