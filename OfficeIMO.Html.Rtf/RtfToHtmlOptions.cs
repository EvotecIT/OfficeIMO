namespace OfficeIMO.Html;

/// <summary>
/// Controls RTF to semantic HTML conversion.
/// </summary>
public sealed partial class RtfToHtmlOptions {
    private RtfHtmlExportProfile _exportProfile = RtfHtmlExportProfile.SemanticDocument;
    private OfficeHtmlDocumentOptions _documentOutput = new() {
        EmitDocumentShell = false,
        Title = null,
        Language = null,
        Theme = OfficeVisualThemeKind.WordLike,
        IncludeDefaultStyles = false,
        BodyClass = "officeimo-html officeimo-rtf-html",
        NewLine = Environment.NewLine
    };

    /// <summary>
    /// Creates options for publishing untrusted RTF as semantic HTML. Private OfficeIMO round-trip
    /// metadata and inline data URI images are disabled, and only web and mail hyperlinks are allowed.
    /// </summary>
    public static RtfToHtmlOptions CreateWebSafeProfile() => new RtfToHtmlOptions {
        ExportProfile = RtfHtmlExportProfile.SemanticDocument
    };

    /// <summary>
    /// Creates options for a trusted OfficeIMO HTML round trip. The output can contain private
    /// metadata and binary payloads and must not be published without sanitization.
    /// </summary>
    public static RtfToHtmlOptions CreateRoundTripProfile() => new RtfToHtmlOptions {
        ExportProfile = RtfHtmlExportProfile.DocumentRoundTrip,
        UrlPolicy = HtmlUrlPolicy.CreateOfficeIMOProfile(),
        IncludeRoundTripMetadata = true,
        EmbedImagesAsDataUri = true,
        MaxEmbeddedImageBytes = int.MaxValue
    };

    /// <summary>
    /// Creates a complete, styled HTML document for browser and print review. The profile remains
    /// static and never enables script execution or remote browser behavior.
    /// </summary>
    public static RtfToHtmlOptions CreatePrintReviewProfile(OfficeVisualThemeKind theme = OfficeVisualThemeKind.WordLike) => new RtfToHtmlOptions {
        ExportProfile = RtfHtmlExportProfile.PrintReview,
        Theme = theme,
        FragmentOnly = false,
        IncludeDefaultStyles = true
    };

    /// <summary>Named RTF-to-HTML fidelity contract represented by this export.</summary>
    public RtfHtmlExportProfile ExportProfile {
        get => _exportProfile;
        set {
            if (!Enum.IsDefined(typeof(RtfHtmlExportProfile), value)) {
                throw new ArgumentOutOfRangeException(nameof(value), value, "RTF HTML export profile is not supported.");
            }
            _exportProfile = value;
        }
    }

    /// <summary>
    /// Compatibility bridge to the shared profile catalog. New code should use
    /// <see cref="ExportProfile"/> so only RTF profiles are representable.
    /// </summary>
    public OfficeHtmlConversionProfile Profile {
        get => ExportProfile switch {
            RtfHtmlExportProfile.DocumentRoundTrip => OfficeHtmlConversionProfile.RtfDocumentRoundTrip,
            RtfHtmlExportProfile.PrintReview => OfficeHtmlConversionProfile.RtfPrintReview,
            _ => OfficeHtmlConversionProfile.RtfSemanticDocument
        };
        set => ExportProfile = value switch {
            OfficeHtmlConversionProfile.RtfSemanticDocument => RtfHtmlExportProfile.SemanticDocument,
            OfficeHtmlConversionProfile.RtfDocumentRoundTrip => RtfHtmlExportProfile.DocumentRoundTrip,
            OfficeHtmlConversionProfile.RtfPrintReview => RtfHtmlExportProfile.PrintReview,
            _ => throw new ArgumentOutOfRangeException(nameof(value), value, "The selected HTML conversion profile is not an RTF profile.")
        };
    }

    /// <summary>Shared engine profile used by the selected RTF export lane.</summary>
    public HtmlConversionProfile SharedProfile => ExportProfile switch {
        RtfHtmlExportProfile.DocumentRoundTrip => HtmlConversionProfile.Document,
        RtfHtmlExportProfile.PrintReview => HtmlConversionProfile.HighFidelityPrint,
        _ => HtmlConversionProfile.Semantic
    };

    /// <summary>Composed document-versus-fragment, theme, title, language, style, and newline settings.</summary>
    public OfficeHtmlDocumentOptions DocumentOutput {
        get => _documentOutput;
        set => _documentOutput = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Compatibility alias for <see cref="OfficeHtmlDocumentOptions.Theme"/>.</summary>
    public OfficeVisualThemeKind Theme { get => DocumentOutput.Theme; set => DocumentOutput.Theme = value; }

    /// <summary>Compatibility alias for <see cref="OfficeHtmlDocumentOptions.IncludeDefaultStyles"/>.</summary>
    public bool IncludeDefaultStyles { get => DocumentOutput.IncludeDefaultStyles; set => DocumentOutput.IncludeDefaultStyles = value; }

    /// <summary>Compatibility alias inverse of <see cref="OfficeHtmlDocumentOptions.EmitDocumentShell"/>.</summary>
    public bool FragmentOnly { get => !DocumentOutput.EmitDocumentShell; set => DocumentOutput.EmitDocumentShell = !value; }

    /// <summary>Includes document metadata when a full HTML document is requested.</summary>
    public bool IncludeMetadata { get; set; } = true;

    /// <summary>
    /// Prefers Outlook/Exchange HTML encapsulated in the RTF transport over its plain-text fallback.
    /// The encapsulated HTML is always reparsed through the bounded HTML reader and current URL policy.
    /// </summary>
    public bool PreferEncapsulatedHtml { get; set; } = true;

    /// <summary>Compatibility alias for <see cref="OfficeHtmlDocumentOptions.Title"/>. When unset, the RTF title is used.</summary>
    public string? Title { get => DocumentOutput.Title; set => DocumentOutput.Title = value; }

    /// <summary>Language override for the HTML root. Null preserves the RTF document language.</summary>
    public string? Language { get => DocumentOutput.Language; set => DocumentOutput.Language = value; }

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

    /// <summary>Compatibility alias for <see cref="OfficeHtmlDocumentOptions.NewLine"/>.</summary>
    public string NewLine { get => DocumentOutput.NewLine; set => DocumentOutput.NewLine = value; }

    internal List<HtmlRtfConversionDiagnostic> Diagnostics { get; } = new List<HtmlRtfConversionDiagnostic>();

    /// <summary>Shared cross-adapter fidelity and policy report for this conversion.</summary>
    internal RtfConversionReport ConversionReport { get; } = new RtfConversionReport();

    /// <summary>Shared HTML diagnostic report for cross-format aggregation.</summary>
    internal HtmlDiagnosticReport HtmlDiagnostics { get; } = new HtmlDiagnosticReport();

    /// <summary>
    /// Creates a reusable copy of the current save options.
    /// </summary>
    /// <returns>A new <see cref="RtfToHtmlOptions"/> with the same configuration values.</returns>
    public RtfToHtmlOptions Clone() => new RtfToHtmlOptions {
        ExportProfile = ExportProfile,
        DocumentOutput = DocumentOutput.Clone(),
        IncludeMetadata = IncludeMetadata,
        PreferEncapsulatedHtml = PreferEncapsulatedHtml,
        UrlPolicy = (UrlPolicy ?? HtmlUrlPolicy.CreateWebOnlyProfile()).Clone(),
        IncludeRoundTripMetadata = IncludeRoundTripMetadata,
        EmbedImagesAsDataUri = EmbedImagesAsDataUri,
        MaxEmbeddedImageBytes = MaxEmbeddedImageBytes,
        ImageSourceResolver = ImageSourceResolver,
    };

    internal string GetNewLine() => string.IsNullOrEmpty(NewLine) ? Environment.NewLine : NewLine;

    internal HtmlUrlPolicy GetUrlPolicy() => UrlPolicy ?? HtmlUrlPolicy.CreateWebOnlyProfile();

    internal void AddDiagnostic(string code, string message, string? source = null, Exception? exception = null, HtmlRtfConversionDiagnosticSeverity severity = HtmlRtfConversionDiagnosticSeverity.Warning, RtfConversionAction? action = null) {
        string? detail = exception == null ? null : exception.GetType().Name + ": " + exception.Message;
        var diagnostic = new HtmlRtfConversionDiagnostic(code, message, severity, source, detail, action);
        Diagnostics.Add(diagnostic);
        HtmlRtfConversionReportMapper.Add(ConversionReport, diagnostic);
        HtmlRtfConversionReportMapper.Add(HtmlDiagnostics, diagnostic);
    }
}
