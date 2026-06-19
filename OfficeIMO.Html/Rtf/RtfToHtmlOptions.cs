namespace OfficeIMO.Html;

/// <summary>
/// Controls RTF to semantic HTML conversion.
/// </summary>
public sealed partial class RtfToHtmlOptions {
    /// <summary>Writes only the body fragment instead of a complete HTML document.</summary>
    public bool FragmentOnly { get; set; } = true;

    /// <summary>Includes document metadata when a full HTML document is requested.</summary>
    public bool IncludeMetadata { get; set; } = true;

    /// <summary>Optional HTML document title. When unset, the RTF title is used.</summary>
    public string? Title { get; set; }

    /// <summary>Embeds PNG and JPEG images as data URI values.</summary>
    public bool EmbedImagesAsDataUri { get; set; } = true;

    /// <summary>Newline sequence used by the generated HTML.</summary>
    public string NewLine { get; set; } = Environment.NewLine;

    /// <summary>
    /// Diagnostics produced while converting the RTF document model into HTML.
    /// </summary>
    public List<HtmlRtfConversionDiagnostic> Diagnostics { get; } = new List<HtmlRtfConversionDiagnostic>();

    /// <summary>
    /// Optional callback invoked whenever a conversion diagnostic is produced.
    /// </summary>
    public Action<HtmlRtfConversionDiagnostic>? DiagnosticHandler { get; set; }

    /// <summary>
    /// Creates a reusable copy of the current save options.
    /// </summary>
    /// <returns>A new <see cref="RtfToHtmlOptions"/> with the same configuration values.</returns>
    public RtfToHtmlOptions Clone() => new RtfToHtmlOptions {
        FragmentOnly = FragmentOnly,
        IncludeMetadata = IncludeMetadata,
        Title = Title,
        EmbedImagesAsDataUri = EmbedImagesAsDataUri,
        NewLine = NewLine,
        DiagnosticHandler = DiagnosticHandler
    };

    internal string GetNewLine() => string.IsNullOrEmpty(NewLine) ? Environment.NewLine : NewLine;

    internal void AddDiagnostic(string code, string message, string? source = null, Exception? exception = null, HtmlRtfConversionDiagnosticSeverity severity = HtmlRtfConversionDiagnosticSeverity.Warning) {
        string? detail = exception == null ? null : exception.GetType().Name + ": " + exception.Message;
        var diagnostic = new HtmlRtfConversionDiagnostic(code, message, severity, source, detail);
        Diagnostics.Add(diagnostic);
        DiagnosticHandler?.Invoke(diagnostic);
    }
}
