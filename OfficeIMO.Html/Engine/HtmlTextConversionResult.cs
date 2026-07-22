namespace OfficeIMO.Html;

/// <summary>HTML text paired with the shared structured conversion report.</summary>
public sealed class HtmlTextConversionResult : HtmlConversionResult<string> {
    /// <summary>Creates an HTML text result and copies diagnostics in emission order.</summary>
    public HtmlTextConversionResult(string value, IEnumerable<HtmlDiagnostic>? diagnostics = null) : base(value) {
        if (diagnostics != null) AddDiagnostics(diagnostics);
    }
}
