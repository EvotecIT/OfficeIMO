using DrawingCore = OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

/// <summary>
/// Controls direct HTML layout and PDF generation.
/// Layout, resource, page, and safety settings are inherited from <see cref="HtmlRenderOptions"/>.
/// </summary>
/// <example>
/// <code>
/// var options = new HtmlPdfSaveOptions {
///     PageSize = OfficePageSizes.A4,
///     Margins = HtmlRenderMargins.All(32),
///     DefaultFontFamily = "Arial"
/// };
/// byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(options);
/// </code>
/// </example>
public sealed class HtmlPdfSaveOptions : HtmlRenderOptions {
    private PdfCore.PdfResourcePolicy _resourcePolicy = PdfCore.PdfResourcePolicy.CreateDefault();

    internal HtmlRenderResourceResolver? EmbeddedPackageResourceResolver { get; set; }
    internal HtmlUrlPolicy? EmbeddedPackageHostResourceUrlPolicy { get; set; }
    /// <summary>Creates direct paged HTML-to-PDF options using the standard defaults.</summary>
    public HtmlPdfSaveOptions() {
        Mode = HtmlRenderMode.Paged;
        UrlPolicy = HtmlUrlPolicy.CreateHyperlinkProfile();
    }

    /// <summary>
    /// Creates PDF-capable options from shared HTML rendering settings without changing their layout mode.
    /// PDF conversion enforces paged layout on its own conversion snapshot and applies PDF-safe hyperlink defaults.
    /// </summary>
    /// <param name="renderOptions">Shared settings used by PNG, SVG, and PDF rendering.</param>
    /// <remarks>Set <see cref="HtmlRenderOptions.UrlPolicy"/> on the returned options to explicitly use a different hyperlink policy.</remarks>
    public HtmlPdfSaveOptions(HtmlRenderOptions renderOptions) : base(renderOptions) {
        if (renderOptions is HtmlPdfSaveOptions pdfOptions) {
            CopyPdfSettingsFrom(pdfOptions);
        } else {
            UrlPolicy = HtmlUrlPolicy.CreateHyperlinkProfile();
        }
    }

    /// <summary>OfficeIMO-managed font fallback groups used by generated PDF text.</summary>
    public PdfCore.PdfTextFallbackFeatures TextFallbacks { get; set; } = PdfCore.PdfTextFallbackFeatures.Default;

    /// <summary>Dependency-free shaping mode used by generated PDF text.</summary>
    public PdfCore.PdfTextShapingMode TextShapingMode { get; set; } = PdfCore.PdfTextShapingMode.LatinLigatures;

    /// <summary>Optional caller-supplied embedded font family used by generated PDF text.</summary>
    public PdfCore.PdfEmbeddedFontFamily? FontFamily { get; set; }

    /// <summary>Optional host-provided shaping seam used with caller-supplied or resolved embedded fonts.</summary>
    public DrawingCore.IOfficeTextShapingProvider? TextShapingProvider { get; set; }

    /// <summary>Host-resource policy. Defaults to balanced conversion: system fonts and bounded in-source resources are allowed, while local and remote reads are denied.</summary>
    public PdfCore.PdfResourcePolicy ResourcePolicy {
        get => _resourcePolicy;
        set => _resourcePolicy = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Creates an independent options snapshot for one PDF conversion.</summary>
    public HtmlPdfSaveOptions ClonePdf() {
        return new HtmlPdfSaveOptions(this);
    }

    /// <summary>Creates an independent options snapshot.</summary>
    public override HtmlRenderOptions Clone() => ClonePdf();

    /// <summary>Returns a snapshot of the active HTML resource policy.</summary>
    public HtmlPdfResourcePolicySummary GetResourcePolicySummary() => HtmlPdfResourcePolicySummary.From(this);

    private void CopyPdfSettingsFrom(HtmlPdfSaveOptions source) {
        TextFallbacks = source.TextFallbacks;
        TextShapingMode = source.TextShapingMode;
        FontFamily = source.FontFamily;
        TextShapingProvider = source.TextShapingProvider;
        ResourcePolicy = source.ResourcePolicy.Clone();
        EmbeddedPackageResourceResolver = source.EmbeddedPackageResourceResolver;
        EmbeddedPackageHostResourceUrlPolicy = source.EmbeddedPackageHostResourceUrlPolicy?.Clone();
    }
}
