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
/// byte[] pdf = html.ToPdf(options);
/// </code>
/// </example>
public sealed class HtmlPdfSaveOptions : HtmlRenderOptions {
    /// <summary>Creates direct paged HTML-to-PDF options using the standard defaults.</summary>
    public HtmlPdfSaveOptions() {
        Mode = HtmlRenderMode.Paged;
    }

    /// <summary>
    /// Creates PDF-capable options from shared HTML rendering settings without changing their layout mode.
    /// PDF conversion enforces paged layout on its own conversion snapshot.
    /// </summary>
    /// <param name="renderOptions">Shared settings used by PNG, SVG, and PDF rendering.</param>
    public HtmlPdfSaveOptions(HtmlRenderOptions renderOptions) : base(renderOptions) { }

    /// <summary>OfficeIMO-managed font fallback groups used by generated PDF text.</summary>
    public PdfCore.PdfTextFallbackFeatures TextFallbacks { get; set; } = PdfCore.PdfTextFallbackFeatures.Default;

    /// <summary>Dependency-free shaping mode used by generated PDF text.</summary>
    public PdfCore.PdfTextShapingMode TextShapingMode { get; set; } = PdfCore.PdfTextShapingMode.LatinLigatures;

    /// <summary>Optional caller-supplied embedded font family used by generated PDF text.</summary>
    public PdfCore.PdfEmbeddedFontFamily? FontFamily { get; set; }

    /// <summary>Optional host-provided shaping seam used with caller-supplied or resolved embedded fonts.</summary>
    public PdfCore.IPdfTextShapingProvider? TextShapingProvider { get; set; }

    /// <summary>Creates an independent options snapshot for one PDF conversion.</summary>
    public HtmlPdfSaveOptions ClonePdf() {
        HtmlPdfSaveOptions clone = CopyTo(new HtmlPdfSaveOptions());
        clone.TextFallbacks = TextFallbacks;
        clone.TextShapingMode = TextShapingMode;
        clone.FontFamily = FontFamily;
        clone.TextShapingProvider = TextShapingProvider;
        return clone;
    }

    /// <summary>Creates an independent options snapshot.</summary>
    public override HtmlRenderOptions Clone() => ClonePdf();

    /// <summary>Returns a snapshot of the active HTML resource policy.</summary>
    public HtmlPdfResourcePolicySummary GetResourcePolicySummary() => HtmlPdfResourcePolicySummary.From(this);
}
