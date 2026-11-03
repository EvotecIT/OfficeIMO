using DrawingCore = OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

/// <summary>
/// Controls direct HTML layout and PDF generation.
/// Layout, resource, page, and safety settings are inherited from <see cref="HtmlRenderOptions"/>.
/// </summary>
/// <example>
/// <code>
/// var options = new HtmlToPdfOptions {
///     PageSize = OfficePageSizes.A4,
///     Margins = HtmlRenderMargins.All(32),
///     DefaultFontFamily = "Arial"
/// };
/// using var pdf = HtmlConversionDocument.Parse(html).ToPdf(options);
/// </code>
/// </example>
public sealed class HtmlToPdfOptions : HtmlRenderOptions, IHtmlPrintRenderOptions {
    private int _maxOutlinedTextCharactersPerRun = 16_384;
    private int _maxOutlinedTextPathCommands = 1_000_000;
    private double? _printLayoutWidthCssPixels;
    private PdfCore.PdfResourcePolicy _resourcePolicy = PdfCore.PdfResourcePolicy.CreateLibraryFontDefault();
    private PdfCore.PdfCreateOptions _documentOptions = new PdfCore.PdfCreateOptions {
        CompressContentStreams = true
    }.EnableTaggedPdfCatalogMarkers();

    internal HtmlRenderResourceResolver? EmbeddedPackageResourceResolver { get; set; }
    internal HtmlUrlPolicy? EmbeddedPackageHostResourceUrlPolicy { get; set; }
    // Retain the container resolver identity so a caller's later replacement remains authoritative.
    internal HtmlRenderResourceResolver? EmbeddedPackageConfiguredResourceResolver { get; set; }
    internal HtmlRenderResourceResolver? EmbeddedPackageHostResourceResolver { get; set; }
    // Image export wraps shared options; the retained container policy stays authoritative.
    internal bool PreservesHtmlResourcePolicy { get; private set; }
    internal DrawingCore.OfficePageSize? PrintOutputPageSize { get; set; }
    // One approved library-selected face snapshot supplies both layout and PDF paint.
    internal PdfCore.PdfEmbeddedFontFamily? LibrarySelectedDefaultFontFamily { get; set; }
    /// <summary>Creates direct paged HTML-to-PDF options using the standard defaults.</summary>
    public HtmlToPdfOptions() {
        Mode = HtmlRenderMode.Paged;
        UrlPolicy = HtmlUrlPolicy.CreateHyperlinkProfile();
        UsesBrowserBodyMargin = true;
    }

    /// <summary>
    /// Creates PDF-capable options from shared HTML rendering settings without changing their layout mode.
    /// PDF conversion enforces paged layout on its own conversion snapshot.
    /// </summary>
    /// <param name="renderOptions">Shared settings used by PNG, SVG, and PDF rendering.</param>
    /// <remarks>Generic rendering settings are intersected with the PDF hyperlink policy so caller restrictions are preserved. PDF-specific snapshots preserve their explicitly configured URL policy.</remarks>
    public HtmlToPdfOptions(HtmlRenderOptions renderOptions) : base(renderOptions) {
        if (renderOptions is HtmlToPdfOptions pdfOptions) {
            CopyPdfSettingsFrom(pdfOptions);
        } else {
            UrlPolicy = HtmlUrlPolicy.Intersect(UrlPolicy, HtmlUrlPolicy.CreateHyperlinkProfile());
        }
    }

    /// <summary>OfficeIMO-managed font fallback groups used by generated PDF text.</summary>
    public PdfCore.PdfTextFallbackFeatures TextFallbacks { get; set; } =
        PdfCore.PdfTextFallbackFeatures.Default |
        PdfCore.PdfTextFallbackFeatures.MultilingualFonts;

    /// <summary>Dependency-free shaping mode used by generated PDF text.</summary>
    public PdfCore.PdfTextShapingMode TextShapingMode { get; set; } = PdfCore.PdfTextShapingMode.LatinLigatures;

    /// <summary>Optional caller-supplied embedded font family used by generated PDF text.</summary>
    public PdfCore.PdfEmbeddedFontFamily? FontFamily { get; set; }

    /// <summary>
    /// Converts supported standard HTML inputs, text areas, and selects into interactive AcroForm fields.
    /// Disable this to retain the shared renderer's static control paint in the PDF.
    /// </summary>
    public bool InteractiveFormControls { get; set; } = true;

    /// <summary>Optional caller-supplied modification date for generated Formula MathML associated files.</summary>
    /// <remarks>Ordinary PDF does not invent a date. Archival profiles requiring embedded-file dates
    /// must receive an explicit value. Formula associated files require PDF 2.0 unless the selected
    /// PDF/A-3 profile permits them in PDF 1.7.</remarks>
    public DateTimeOffset? MathMlSourceModificationDate { get; set; }

    /// <summary>
    /// Fits visible layout overflow, including fixed-width descendants and tables, into a
    /// uniform CSS print page. Clipped descendants and paint-only shadows do not widen the
    /// layout. Fitting retains the physical CSS media context and uses bounded reflow.
    /// Named and page-specific rules require
    /// an explicit fitting width. Set to false to retain unscaled print layout. An explicit
    /// <see cref="PrintLayoutWidthCssPixels"/> takes precedence.
    /// </summary>
    public bool AutoFitWidePrintContent { get; set; } = true;

    /// <summary>
    /// Optional wider CSS layout width for print-to-page fitting. The PDF keeps <see cref="HtmlRenderOptions.PageSize"/>
    /// as its physical size while the paged layout is uniformly reduced to fit its width.
    /// CSS media queries and responsive resource selection continue to use <see cref="HtmlRenderOptions.ViewportWidth"/>.
    /// When <see cref="HtmlRenderOptions.HonorCssPageRules"/> is true, the resolved authored page
    /// size and margins remain the physical PDF geometry while the content is reduced to fit its
    /// printable width. The requested width must exceed that printable width.
    /// </summary>
    public double? PrintLayoutWidthCssPixels {
        get => _printLayoutWidthCssPixels;
        set {
            if (value.HasValue && (value.Value <= 0D || double.IsNaN(value.Value) || double.IsInfinity(value.Value)))
                throw new ArgumentOutOfRangeException(nameof(value));
            _printLayoutWidthCssPixels = value;
        }
    }

    /// <summary>
    /// Maximum UTF-16 characters expanded into vector font outlines for one rendered text run.
    /// Raise this only for trusted documents with unusually long unbroken text.
    /// </summary>
    public int MaxOutlinedTextCharactersPerRun {
        get => _maxOutlinedTextCharactersPerRun;
        set {
            if (value <= 0) throw new ArgumentOutOfRangeException(nameof(value));
            _maxOutlinedTextCharactersPerRun = value;
        }
    }

    /// <summary>
    /// Maximum path commands generated by text outlines across one PDF conversion.
    /// </summary>
    public int MaxOutlinedTextPathCommands {
        get => _maxOutlinedTextPathCommands;
        set {
            if (value <= 0) throw new ArgumentOutOfRangeException(nameof(value));
            _maxOutlinedTextPathCommands = value;
        }
    }

    /// <summary>
    /// PDF writer settings used for the generated document, including file version,
    /// forward-only serialization, tagging, compliance groundwork, and output intents.
    /// </summary>
    /// <remarks>
    /// Conversion snapshots this value. HTML-specific font and shaping properties are
    /// then applied as explicit adapter overrides.
    /// </remarks>
    public PdfCore.PdfCreateOptions PdfOptions {
        get => _documentOptions;
        set => _documentOptions = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Optional host-provided shaping seam used with caller-supplied or resolved embedded fonts.</summary>
    public new DrawingCore.IOfficeTextShapingProvider? TextShapingProvider {
        get => base.TextShapingProvider;
        set => base.TextShapingProvider = value;
    }

    /// <summary>Host-resource policy. Defaults to library-selected installed fonts and bounded in-source resources; document-selected host fonts, local files and remote reads are disabled.</summary>
    public PdfCore.PdfResourcePolicy ResourcePolicy {
        get => _resourcePolicy;
        set => _resourcePolicy = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Creates an independent options snapshot for one PDF conversion.</summary>
    public HtmlToPdfOptions ClonePdf() {
        return new HtmlToPdfOptions(this);
    }

    /// <summary>Creates an independent options snapshot.</summary>
    public override HtmlRenderOptions Clone() => ClonePdf();

    /// <summary>Returns a snapshot of the active HTML resource policy.</summary>
    public HtmlPdfResourcePolicySummary GetResourcePolicySummary() => HtmlPdfResourcePolicySummary.From(this);

    /// <summary>
    /// Creates options for PNG, JPEG, TIFF, WebP or SVG render requests from shared render options.
    /// Text is measured and painted through the PDF engine, while the URL and resource policy of
    /// <paramref name="source"/> stays authoritative instead of <see cref="ResourcePolicy"/>.
    /// PDF-capable options are returned unchanged.
    /// </summary>
    /// <param name="source">Shared render options, or null for the defaults.</param>
    public static HtmlToPdfOptions ForImageExport(HtmlRenderOptions? source) {
        if (source is HtmlToPdfOptions pdf) return pdf;
        HtmlRenderOptions shared = source ?? new HtmlRenderOptions();
        return new HtmlToPdfOptions(shared) { UrlPolicy = shared.UrlPolicy.Clone(), PreservesHtmlResourcePolicy = true };
    }

    bool IHtmlPrintRenderOptions.UsesPrintResourcePolicy => !PreservesHtmlResourcePolicy;

    HtmlRenderOptions IHtmlPrintRenderOptions.PrepareRenderOptions(HtmlRenderRequest request) =>
        HtmlPdfRenderedConverter.PrepareNonPdfOptions(request);

    IReadOnlyList<string> IHtmlPrintRenderOptions.RetainConfiguredDefaultFonts(OfficeFontFaceCollection fonts,
        HtmlResourceSession resources, HtmlDiagnosticReport diagnostics, CancellationToken cancellationToken) =>
        HtmlPdfRenderedConverter.RetainConfiguredDefaultFonts(fonts, this, resources, diagnostics, cancellationToken);

    bool IHtmlPrintRenderOptions.AllowsDocumentSystemFontFamilies =>
        ResourcePolicy.AllowSystemFontEmbedding && ResourcePolicy.AllowDocumentFontEmbedding;

    bool IHtmlPrintRenderOptions.TryPrepareLibraryFontMetrics(IEnumerable<HtmlRenderVisual> flow,
        IEnumerable<(string Text, OfficeFontInfo Font)> marginRequirements, OfficeFontFaceCollection fonts,
        HtmlResourceSession resources, HtmlDiagnosticReport diagnostics, CancellationToken cancellationToken) =>
        HtmlPdfRenderedConverter.TryPrepareLibraryFontMetrics(flow, marginRequirements, fonts, this, resources,
            diagnostics, cancellationToken);

    HtmlRenderRequest IHtmlPrintRenderOptions.PrepareSiteBundleRequest(HtmlSiteBundle bundle, HtmlRenderRequest request) =>
        HtmlSiteBundlePdfExtensions.Prepare(bundle, request);

    bool IHtmlPrintRenderOptions.AllowsEmbeddedPackageResources => ResourcePolicy.AllowEmbeddedPackageResources;

    HtmlRenderResourceResolver? IHtmlPrintRenderOptions.EmbeddedPackageResourceResolver {
        get => EmbeddedPackageResourceResolver;
        set => EmbeddedPackageResourceResolver = value;
    }

    HtmlUrlPolicy? IHtmlPrintRenderOptions.EmbeddedPackageHostResourceUrlPolicy {
        get => EmbeddedPackageHostResourceUrlPolicy;
        set => EmbeddedPackageHostResourceUrlPolicy = value;
    }

    HtmlRenderResourceResolver? IHtmlPrintRenderOptions.EmbeddedPackageConfiguredResourceResolver {
        get => EmbeddedPackageConfiguredResourceResolver;
        set => EmbeddedPackageConfiguredResourceResolver = value;
    }

    HtmlRenderResourceResolver? IHtmlPrintRenderOptions.EmbeddedPackageHostResourceResolver {
        get => EmbeddedPackageHostResourceResolver;
        set => EmbeddedPackageHostResourceResolver = value;
    }

    private void CopyPdfSettingsFrom(HtmlToPdfOptions source) {
        FailOnLoss = source.FailOnLoss;
        PreservesHtmlResourcePolicy = source.PreservesHtmlResourcePolicy;
        TextFallbacks = source.TextFallbacks;
        TextShapingMode = source.TextShapingMode;
        FontFamily = source.FontFamily;
        LibrarySelectedDefaultFontFamily = source.LibrarySelectedDefaultFontFamily;
        InteractiveFormControls = source.InteractiveFormControls;
        MathMlSourceModificationDate = source.MathMlSourceModificationDate;
        PrintLayoutWidthCssPixels = source.PrintLayoutWidthCssPixels;
        AutoFitWidePrintContent = source.AutoFitWidePrintContent;
        PrintOutputPageSize = source.PrintOutputPageSize;
        MaxOutlinedTextCharactersPerRun = source.MaxOutlinedTextCharactersPerRun;
        MaxOutlinedTextPathCommands = source.MaxOutlinedTextPathCommands;
        PdfOptions = source.PdfOptions.Clone();
        TextShapingProvider = source.TextShapingProvider;
        ResourcePolicy = source.ResourcePolicy.Clone();
        EmbeddedPackageResourceResolver = source.EmbeddedPackageResourceResolver;
        EmbeddedPackageHostResourceUrlPolicy = source.EmbeddedPackageHostResourceUrlPolicy?.Clone();
        EmbeddedPackageConfiguredResourceResolver = source.EmbeddedPackageConfiguredResourceResolver;
        EmbeddedPackageHostResourceResolver = source.EmbeddedPackageHostResourceResolver;
    }
}
