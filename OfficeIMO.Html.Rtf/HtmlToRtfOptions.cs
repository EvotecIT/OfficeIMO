namespace OfficeIMO.Html;

/// <summary>
/// Controls semantic HTML to RTF conversion.
/// </summary>
public sealed partial class HtmlToRtfOptions {
    /// <summary>
    /// Creates the default OfficeIMO HTML to RTF import profile.
    /// </summary>
    /// <returns>A new <see cref="HtmlToRtfOptions"/> instance using the default semantic bridge behavior.</returns>
    public static HtmlToRtfOptions CreateOfficeIMOProfile() => new HtmlToRtfOptions();

    /// <summary>
    /// Creates a bounded offline profile for untrusted HTML ingestion.
    /// </summary>
    /// <remarks>
    /// The HTML/RTF bridge does not fetch external resources. This profile adds conservative
    /// structural limits while preserving the same shared OfficeIMO HTML conversion path.
    /// Callers can relax individual limits when their ingestion boundary is more trusted.
    /// </remarks>
    /// <returns>A new <see cref="HtmlToRtfOptions"/> instance configured for untrusted HTML.</returns>
    public static HtmlToRtfOptions CreateUntrustedHtmlProfile() => new HtmlToRtfOptions {
        MaxHtmlNodes = 10000,
        MaxHtmlDepth = 64,
        IgnoreInsignificantWhitespace = true,
        PreserveUnknownTagsAsText = false
    };

    /// <summary>Base URI used to resolve relative hyperlinks and image sources.</summary>
    public Uri? BaseUri { get; set; }

    /// <summary>URL policy used before hyperlinks are materialized into RTF content.</summary>
    public HtmlUrlPolicy UrlPolicy { get; set; } = HtmlUrlPolicy.CreateOfficeIMOProfile();

    /// <summary>Optional separate policy for image sources. When omitted, <see cref="UrlPolicy"/> is used.</summary>
    public HtmlUrlPolicy? ResourceUrlPolicy { get; set; }

    /// <summary>Preserves unknown element names as bracketed text markers instead of treating them as transparent containers.</summary>
    public bool PreserveUnknownTagsAsText { get; set; }

    /// <summary>When enabled, text nodes made only of whitespace are ignored outside preformatted elements.</summary>
    public bool IgnoreInsignificantWhitespace { get; set; } = true;

    /// <summary>
    /// Optional maximum number of parsed HTML element and text nodes allowed for a conversion operation.
    /// When exceeded, conversion stops with <see cref="HtmlRtfConversionLimitException"/> and an error diagnostic.
    /// </summary>
    public int? MaxHtmlNodes { get; set; }

    /// <summary>
    /// Optional maximum parsed HTML element nesting depth allowed for a conversion operation.
    /// When exceeded, conversion stops with <see cref="HtmlRtfConversionLimitException"/> and an error diagnostic.
    /// </summary>
    public int? MaxHtmlDepth { get; set; }

    internal List<HtmlRtfConversionDiagnostic> Diagnostics { get; } = new List<HtmlRtfConversionDiagnostic>();

    /// <summary>Shared cross-adapter fidelity and policy report for this conversion.</summary>
    internal RtfConversionReport ConversionReport { get; } = new RtfConversionReport();

    /// <summary>Shared HTML diagnostic report for cross-format aggregation.</summary>
    internal HtmlDiagnosticReport HtmlDiagnostics { get; } = new HtmlDiagnosticReport();

    /// <summary>
    /// Creates a reusable copy of the current options without carrying runtime diagnostics into the clone.
    /// </summary>
    /// <returns>A new <see cref="HtmlToRtfOptions"/> with the same configuration values.</returns>
    public HtmlToRtfOptions Clone() => new HtmlToRtfOptions {
        BaseUri = BaseUri,
        UrlPolicy = (UrlPolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile()).Clone(),
        ResourceUrlPolicy = ResourceUrlPolicy?.Clone(),
        PreserveUnknownTagsAsText = PreserveUnknownTagsAsText,
        IgnoreInsignificantWhitespace = IgnoreInsignificantWhitespace,
        MaxHtmlNodes = MaxHtmlNodes,
        MaxHtmlDepth = MaxHtmlDepth
    };

    internal HtmlUrlPolicy GetResourceUrlPolicy() => ResourceUrlPolicy ?? UrlPolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile();
}
