namespace OfficeIMO.Html;

/// <summary>
/// Controls RTF to semantic HTML conversion.
/// </summary>
public sealed partial class RtfToHtmlOptions {
    /// <summary>
    /// Creates options for publishing untrusted RTF as semantic HTML. Private OfficeIMO round-trip
    /// metadata and inline data URI images are disabled, and only web and mail hyperlinks are allowed.
    /// </summary>
    public static RtfToHtmlOptions CreateWebSafeProfile() => new RtfToHtmlOptions();

    /// <summary>
    /// Creates options for a trusted OfficeIMO HTML round trip. The output can contain private
    /// metadata and binary payloads and must not be published without sanitization.
    /// </summary>
    public static RtfToHtmlOptions CreateRoundTripProfile() => new RtfToHtmlOptions {
        UrlPolicy = HtmlUrlPolicy.CreateOfficeIMOProfile(),
        IncludeRoundTripMetadata = true,
        EmbedImagesAsDataUri = true,
        MaxEmbeddedImageBytes = int.MaxValue
    };

    /// <summary>Writes only the body fragment instead of a complete HTML document.</summary>
    public bool FragmentOnly { get; set; } = true;

    /// <summary>Includes document metadata when a full HTML document is requested.</summary>
    public bool IncludeMetadata { get; set; } = true;

    /// <summary>
    /// Prefers Outlook/Exchange HTML encapsulated in the RTF transport over its plain-text fallback.
    /// The encapsulated HTML is always reparsed through the bounded HTML reader and current URL policy.
    /// </summary>
    public bool PreferEncapsulatedHtml { get; set; } = true;

    /// <summary>Optional HTML document title. When unset, the RTF title is used.</summary>
    public string? Title { get; set; }

    /// <summary>
    /// URL policy applied to every hyperlink and caller-supplied image source written to HTML.
    /// The default is the restrictive web-only policy.
    /// </summary>
    public HtmlUrlPolicy UrlPolicy { get; set; } = HtmlUrlPolicy.CreateWebOnlyProfile();

    /// <summary>
    /// Includes private <c>data-officeimo-rtf-*</c> metadata used for trusted fidelity round trips.
    /// This can include encoded object and image payloads and is disabled by default.
    /// </summary>
    public bool IncludeRoundTripMetadata { get; set; }

    /// <summary>Embeds supported images as data URI values. Disabled by default for web-safe output.</summary>
    public bool EmbedImagesAsDataUri { get; set; }

    /// <summary>Maximum size of one image that may be embedded as a data URI.</summary>
    public int MaxEmbeddedImageBytes { get; set; } = 1_000_000;

    /// <summary>
    /// Optional callback that stores or maps an RTF image and returns its HTML source URL.
    /// Returned URLs are validated using <see cref="UrlPolicy"/>.
    /// </summary>
    public Func<RtfImage, string?>? ImageSourceResolver { get; set; }

    /// <summary>Newline sequence used by the generated HTML.</summary>
    public string NewLine { get; set; } = Environment.NewLine;

    /// <summary>
    /// Diagnostics produced while converting the RTF document model into HTML.
    /// </summary>
    public List<HtmlRtfConversionDiagnostic> Diagnostics { get; } = new List<HtmlRtfConversionDiagnostic>();

    /// <summary>Shared cross-adapter fidelity and policy report for this conversion.</summary>
    public RtfConversionReport ConversionReport { get; } = new RtfConversionReport();

    /// <summary>Shared HTML diagnostic report for cross-format aggregation.</summary>
    public HtmlDiagnosticReport HtmlDiagnostics { get; } = new HtmlDiagnosticReport();

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
        PreferEncapsulatedHtml = PreferEncapsulatedHtml,
        Title = Title,
        UrlPolicy = (UrlPolicy ?? HtmlUrlPolicy.CreateWebOnlyProfile()).Clone(),
        IncludeRoundTripMetadata = IncludeRoundTripMetadata,
        EmbedImagesAsDataUri = EmbedImagesAsDataUri,
        MaxEmbeddedImageBytes = MaxEmbeddedImageBytes,
        ImageSourceResolver = ImageSourceResolver,
        NewLine = NewLine,
        DiagnosticHandler = DiagnosticHandler
    };

    internal string GetNewLine() => string.IsNullOrEmpty(NewLine) ? Environment.NewLine : NewLine;

    internal HtmlUrlPolicy GetUrlPolicy() => UrlPolicy ?? HtmlUrlPolicy.CreateWebOnlyProfile();

    internal void AddDiagnostic(string code, string message, string? source = null, Exception? exception = null, HtmlRtfConversionDiagnosticSeverity severity = HtmlRtfConversionDiagnosticSeverity.Warning) {
        string? detail = exception == null ? null : exception.GetType().Name + ": " + exception.Message;
        var diagnostic = new HtmlRtfConversionDiagnostic(code, message, severity, source, detail);
        Diagnostics.Add(diagnostic);
        HtmlRtfConversionReportMapper.Add(ConversionReport, diagnostic);
        HtmlRtfConversionReportMapper.Add(HtmlDiagnostics, diagnostic);
        DiagnosticHandler?.Invoke(diagnostic);
    }
}
