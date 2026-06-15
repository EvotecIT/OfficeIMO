namespace OfficeIMO.Rtf.Html;

/// <summary>
/// Controls semantic HTML to RTF conversion.
/// </summary>
public sealed partial class RtfHtmlReadOptions {
    /// <summary>
    /// Creates the default OfficeIMO RTF HTML import profile.
    /// </summary>
    /// <returns>A new <see cref="RtfHtmlReadOptions"/> instance using the default semantic bridge behavior.</returns>
    public static RtfHtmlReadOptions CreateOfficeIMOProfile() => new RtfHtmlReadOptions();

    /// <summary>
    /// Creates a bounded offline profile for untrusted HTML ingestion.
    /// </summary>
    /// <remarks>
    /// The RTF HTML bridge does not fetch external resources. This profile adds conservative
    /// structural limits while preserving the same dependency-free semantic conversion path.
    /// Callers can relax individual limits when their ingestion boundary is more trusted.
    /// </remarks>
    /// <returns>A new <see cref="RtfHtmlReadOptions"/> instance configured for untrusted HTML.</returns>
    public static RtfHtmlReadOptions CreateUntrustedHtmlProfile() => new RtfHtmlReadOptions {
        MaxHtmlNodes = 10000,
        MaxHtmlDepth = 64,
        IgnoreInsignificantWhitespace = true,
        PreserveUnknownTagsAsText = false
    };

    /// <summary>Base URI used to resolve relative hyperlinks and image sources.</summary>
    public Uri? BaseUri { get; set; }

    /// <summary>Preserves unknown element names as bracketed text markers instead of treating them as transparent containers.</summary>
    public bool PreserveUnknownTagsAsText { get; set; }

    /// <summary>When enabled, text nodes made only of whitespace are ignored outside preformatted elements.</summary>
    public bool IgnoreInsignificantWhitespace { get; set; } = true;

    /// <summary>
    /// Optional maximum number of parsed HTML element and text nodes allowed for a conversion operation.
    /// When exceeded, conversion stops with <see cref="RtfHtmlConversionLimitException"/> and an error diagnostic.
    /// </summary>
    public int? MaxHtmlNodes { get; set; }

    /// <summary>
    /// Optional maximum parsed HTML element nesting depth allowed for a conversion operation.
    /// When exceeded, conversion stops with <see cref="RtfHtmlConversionLimitException"/> and an error diagnostic.
    /// </summary>
    public int? MaxHtmlDepth { get; set; }

    /// <summary>
    /// Diagnostics produced while converting HTML into the RTF document model.
    /// </summary>
    public List<RtfHtmlConversionDiagnostic> Diagnostics { get; } = new List<RtfHtmlConversionDiagnostic>();

    /// <summary>
    /// Optional callback invoked whenever a conversion diagnostic is produced.
    /// </summary>
    public Action<RtfHtmlConversionDiagnostic>? DiagnosticHandler { get; set; }

    /// <summary>
    /// Creates a reusable copy of the current options without carrying runtime diagnostics into the clone.
    /// </summary>
    /// <returns>A new <see cref="RtfHtmlReadOptions"/> with the same configuration values.</returns>
    public RtfHtmlReadOptions Clone() => new RtfHtmlReadOptions {
        BaseUri = BaseUri,
        PreserveUnknownTagsAsText = PreserveUnknownTagsAsText,
        IgnoreInsignificantWhitespace = IgnoreInsignificantWhitespace,
        MaxHtmlNodes = MaxHtmlNodes,
        MaxHtmlDepth = MaxHtmlDepth,
        DiagnosticHandler = DiagnosticHandler
    };
}
