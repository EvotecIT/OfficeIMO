namespace OfficeIMO.Pdf;

/// <summary>
/// Options controlling page geometry and default typography for a PDF document.
/// </summary>
public sealed class PdfOptions {
    private PdfAlign _headerAlign = PdfAlign.Center;
    private PdfAlign _footerAlign = PdfAlign.Center;
    private System.Collections.Generic.List<FooterSegment>? _headerSegments;
    private System.Collections.Generic.List<FooterSegment>? _firstPageHeaderSegments;
    private System.Collections.Generic.List<FooterSegment>? _evenPageHeaderSegments;
    private System.Collections.Generic.List<FooterSegment>? _footerSegments;
    private System.Collections.Generic.List<FooterSegment>? _firstPageFooterSegments;
    private System.Collections.Generic.List<FooterSegment>? _evenPageFooterSegments;
    private string? _headerLeftFormat;
    private string? _headerCenterFormat;
    private string? _headerRightFormat;
    private string? _firstPageHeaderLeftFormat;
    private string? _firstPageHeaderCenterFormat;
    private string? _firstPageHeaderRightFormat;
    private string? _evenPageHeaderLeftFormat;
    private string? _evenPageHeaderCenterFormat;
    private string? _evenPageHeaderRightFormat;
    private string? _footerLeftFormat;
    private string? _footerCenterFormat;
    private string? _footerRightFormat;
    private string? _firstPageFooterLeftFormat;
    private string? _firstPageFooterCenterFormat;
    private string? _firstPageFooterRightFormat;
    private string? _evenPageFooterLeftFormat;
    private string? _evenPageFooterCenterFormat;
    private string? _evenPageFooterRightFormat;
    private System.Collections.Generic.List<PdfHeaderFooterImage>? _headerImages;
    private System.Collections.Generic.List<PdfHeaderFooterImage>? _firstPageHeaderImages;
    private System.Collections.Generic.List<PdfHeaderFooterImage>? _evenPageHeaderImages;
    private System.Collections.Generic.List<PdfHeaderFooterImage>? _footerImages;
    private System.Collections.Generic.List<PdfHeaderFooterImage>? _firstPageFooterImages;
    private System.Collections.Generic.List<PdfHeaderFooterImage>? _evenPageFooterImages;
    private System.Collections.Generic.List<PdfHeaderFooterShape>? _headerShapes;
    private System.Collections.Generic.List<PdfHeaderFooterShape>? _firstPageHeaderShapes;
    private System.Collections.Generic.List<PdfHeaderFooterShape>? _evenPageHeaderShapes;
    private System.Collections.Generic.List<PdfHeaderFooterShape>? _footerShapes;
    private System.Collections.Generic.List<PdfHeaderFooterShape>? _firstPageFooterShapes;
    private System.Collections.Generic.List<PdfHeaderFooterShape>? _evenPageFooterShapes;
    private PdfStandardFont _defaultFont = PdfStandardFont.Helvetica;
    private PdfStandardFont _headerFont = PdfStandardFont.Helvetica;
    private PdfStandardFont _footerFont = PdfStandardFont.Helvetica;
    private int _pageNumberStart = 1;
    private bool _hasExplicitPageNumberStart;
    private PdfPageNumberStyle _pageNumberStyle = PdfPageNumberStyle.Arabic;
    private string? _pageLabelPrefix;
    private PdfParagraphStyle? _defaultParagraphStyle;
    private PdfTableStyle? _defaultTableStyle = TableStyles.Light();
    private PdfHeadingStyles? _defaultHeadingStyles;
    private PdfListStyle? _defaultListStyle;
    private PanelStyle? _defaultPanelStyle;
    private PdfHorizontalRuleStyle? _defaultHorizontalRuleStyle;
    private PdfImageStyle? _defaultImageStyle;
    private PdfDrawingStyle? _defaultDrawingStyle;
    private PdfRowStyle? _defaultRowStyle;
    private PdfComplianceProfile _complianceProfile;
    private PdfOutputIntent? _outputIntent;
    private string? _language;
    private PdfViewerPreferencesOptions? _viewerPreferences;
    private PdfTextWatermark? _textWatermark;
    private PdfImageWatermark? _imageWatermark;
    private PdfPageBorder? _pageBorder;
    private PdfPageBackgroundImage? _pageBackgroundImage;
    private System.Collections.Generic.List<PdfPageBackgroundShape>? _pageBackgroundShapes;
    private System.Collections.Generic.Dictionary<PdfStandardFont, PdfEmbeddedFont>? _embeddedFonts;
    private System.Collections.Generic.List<PdfEmbeddedFile>? _embeddedFiles;

    /// <summary>Page width in points (1 pt = 1/72 in). Default is 612 (Letter 8.5in).</summary>
    public double PageWidth { get; set; } = 612; // Letter 8.5in * 72
    /// <summary>Page height in points. Default is 792 (Letter 11in).</summary>
    public double PageHeight { get; set; } = 792; // Letter 11in * 72
    /// <summary>Page size in points.</summary>
    public PageSize PageSize {
        get => new PageSize(PageWidth, PageHeight);
        set {
            Guard.Positive(value.Width, nameof(PageSize));
            Guard.Positive(value.Height, nameof(PageSize));
            PageWidth = value.Width;
            PageHeight = value.Height;
        }
    }
    /// <summary>Page orientation inferred from the current page size.</summary>
    public PdfPageOrientation PageOrientation => PageWidth > PageHeight ? PdfPageOrientation.Landscape : PdfPageOrientation.Portrait;
    /// <summary>Optional page background color rendered behind all page content.</summary>
    public PdfColor? BackgroundColor { get; set; }
    /// <summary>Optional reusable text watermark rendered behind all page content.</summary>
    public PdfTextWatermark? TextWatermark {
        get => _textWatermark?.Clone();
        set => _textWatermark = value?.Clone();
    }
    internal PdfTextWatermark? TextWatermarkSnapshot => _textWatermark?.Clone();
    /// <summary>Optional reusable image watermark rendered behind all page content.</summary>
    public PdfImageWatermark? ImageWatermark {
        get => _imageWatermark?.Clone();
        set => _imageWatermark = value?.Clone();
    }
    internal PdfImageWatermark? ImageWatermarkSnapshot => _imageWatermark?.Clone();
    /// <summary>Optional reusable page border rendered as a page decoration.</summary>
    public PdfPageBorder? PageBorder {
        get => _pageBorder?.Clone();
        set => _pageBorder = value?.Clone();
    }
    internal PdfPageBorder? PageBorderSnapshot => _pageBorder?.Clone();
    /// <summary>Optional reusable page background image rendered behind all page content.</summary>
    public PdfPageBackgroundImage? PageBackgroundImage {
        get => _pageBackgroundImage?.Clone();
        set => _pageBackgroundImage = value?.Clone();
    }
    internal PdfPageBackgroundImage? PageBackgroundImageSnapshot => _pageBackgroundImage?.Clone();
    /// <summary>Optional reusable page background shapes rendered behind all page content.</summary>
    public System.Collections.Generic.IReadOnlyList<PdfPageBackgroundShape>? PageBackgroundShapes {
        get => ClonePageBackgroundShapes(_pageBackgroundShapes);
        set => _pageBackgroundShapes = ClonePageBackgroundShapes(value);
    }
    internal System.Collections.Generic.IReadOnlyList<PdfPageBackgroundShape> PageBackgroundShapeSnapshots =>
        (System.Collections.Generic.IReadOnlyList<PdfPageBackgroundShape>?)ClonePageBackgroundShapes(_pageBackgroundShapes) ?? System.Array.Empty<PdfPageBackgroundShape>();
    /// <summary>Left margin in points. Default 72 (1 inch).</summary>
    public double MarginLeft { get; set; } = 72; // 1 in
    /// <summary>Right margin in points. Default 72 (1 inch).</summary>
    public double MarginRight { get; set; } = 72;
    /// <summary>Top margin in points. Default 72 (1 inch).</summary>
    public double MarginTop { get; set; } = 72;
    /// <summary>Bottom margin in points. Default 72 (1 inch).</summary>
    public double MarginBottom { get; set; } = 72;
    /// <summary>Page margins in points.</summary>
    public PageMargins Margins {
        get => new PageMargins(MarginLeft, MarginTop, MarginRight, MarginBottom);
        set {
            MarginLeft = value.Left;
            MarginTop = value.Top;
            MarginRight = value.Right;
            MarginBottom = value.Bottom;
        }
    }
    /// <summary>Default standard font used for paragraphs.</summary>
    public PdfStandardFont DefaultFont {
        get => _defaultFont;
        set {
            Guard.StandardFont(value, nameof(DefaultFont), "PDF default font must be one of the supported standard PDF fonts.");
            _defaultFont = value;
        }
    }
    /// <summary>Default paragraph font size in points. Default 11.</summary>
    public double DefaultFontSize { get; set; } = 11;
    /// <summary>When true, generated page content streams are written with Flate compression.</summary>
    public bool CompressContentStreams { get; set; }
    /// <summary>When true, generated standard-font resources include WinAnsi-to-Unicode CMaps for stronger extraction and compliance groundwork.</summary>
    public bool IncludeStandardFontToUnicodeMaps { get; set; }
    /// <summary>When true, embedded TrueType font file streams are Flate-compressed while preserving their original /Length1 metadata.</summary>
    public bool CompressEmbeddedFonts { get; set; } = true;
    /// <summary>When true, generated PDFs include a catalog XMP metadata stream synchronized with document Info metadata.</summary>
    public bool IncludeXmpMetadata { get; set; }
    /// <summary>When true, generated PDFs include catalog page labels that match the configured page-number style and start number.</summary>
    public bool IncludePageLabels { get; set; }
    /// <summary>Optional catalog page-label prefix, for example "A-" or "Appendix ". Requires <see cref="IncludePageLabels"/> to be emitted.</summary>
    public string? PageLabelPrefix {
        get => _pageLabelPrefix;
        set {
            PdfPageLabelDictionaryBuilder.ValidatePrefix(value, nameof(PageLabelPrefix));
            _pageLabelPrefix = value;
        }
    }
    /// <summary>Requested generated-PDF compliance profile. Non-None profiles are validated strictly and fail until their required primitives are implemented.</summary>
    public PdfComplianceProfile ComplianceProfile {
        get => _complianceProfile;
        set {
            Guard.ComplianceProfile(value, nameof(ComplianceProfile));
            _complianceProfile = value;
        }
    }
    /// <summary>Optional generated catalog output intent backed by an ICC profile.</summary>
    public PdfOutputIntent? OutputIntent {
        get => _outputIntent?.Clone();
        set => _outputIntent = value?.Clone();
    }
    internal PdfOutputIntent? OutputIntentSnapshot => _outputIntent?.Clone();
    /// <summary>Optional document language for the generated catalog /Lang entry, for example "en-US".</summary>
    public string? Language {
        get => _language;
        set {
            ValidateOptionalLanguage(value, nameof(Language));
            _language = value;
        }
    }

    /// <summary>Optional simple viewer preferences emitted through the generated catalog.</summary>
    public PdfViewerPreferencesOptions? ViewerPreferences {
        get => _viewerPreferences?.Clone();
        set => _viewerPreferences = value?.Clone();
    }

    internal PdfViewerPreferencesOptions? ViewerPreferencesSnapshot => _viewerPreferences?.Clone();

    /// <summary>Embedded files associated with the generated document catalog.</summary>
    public System.Collections.Generic.IReadOnlyList<PdfEmbeddedFile> EmbeddedFiles =>
        (System.Collections.Generic.IReadOnlyList<PdfEmbeddedFile>?)CloneEmbeddedFiles(_embeddedFiles) ?? System.Array.Empty<PdfEmbeddedFile>();

    internal System.Collections.Generic.IReadOnlyList<PdfEmbeddedFile> EmbeddedFileSnapshots =>
        (System.Collections.Generic.IReadOnlyList<PdfEmbeddedFile>?)CloneEmbeddedFiles(_embeddedFiles) ?? System.Array.Empty<PdfEmbeddedFile>();

    /// <summary>Embedded TrueType font mappings keyed by generated standard-font slot.</summary>
    public System.Collections.Generic.IReadOnlyDictionary<PdfStandardFont, PdfEmbeddedFont> EmbeddedFonts {
        get {
            if (_embeddedFonts == null || _embeddedFonts.Count == 0) {
                return new System.Collections.Generic.Dictionary<PdfStandardFont, PdfEmbeddedFont>();
            }

            var copy = new System.Collections.Generic.Dictionary<PdfStandardFont, PdfEmbeddedFont>();
            foreach (var embeddedFont in _embeddedFonts) {
                copy[embeddedFont.Key] = embeddedFont.Value.Clone();
            }

            return copy;
        }
    }

    /// <summary>Embeds a TrueType font file for a generated standard-font slot.</summary>
    public PdfOptions EmbedStandardFont(PdfStandardFont font, byte[] data, string? fontName = null) {
        var embeddedFont = new PdfEmbeddedFont(font, data, fontName);
        (_embeddedFonts ??= new System.Collections.Generic.Dictionary<PdfStandardFont, PdfEmbeddedFont>())[font] = embeddedFont;
        return this;
    }

    /// <summary>Embeds a TrueType font file from disk for a generated standard-font slot.</summary>
    public PdfOptions EmbedStandardFont(PdfStandardFont font, string path, string? fontName = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return EmbedStandardFont(font, System.IO.File.ReadAllBytes(path), fontName);
    }

    /// <summary>Removes all embedded standard-font mappings.</summary>
    public PdfOptions ClearEmbeddedStandardFonts() {
        _embeddedFonts?.Clear();
        return this;
    }

    /// <summary>Requests a generated-PDF compliance profile for this document.</summary>
    public PdfOptions RequireCompliance(PdfComplianceProfile profile) {
        ComplianceProfile = profile;
        return this;
    }

    /// <summary>Sets a generated catalog output intent backed by an ICC profile.</summary>
    public PdfOptions SetOutputIntent(PdfOutputIntent? outputIntent) {
        OutputIntent = outputIntent;
        return this;
    }

    /// <summary>Sets a generated catalog output intent from ICC profile bytes.</summary>
    public PdfOptions SetOutputIntent(byte[] iccProfile, string outputConditionIdentifier = "sRGB IEC61966-2.1") {
        OutputIntent = new PdfOutputIntent(iccProfile, outputConditionIdentifier);
        return this;
    }

    /// <summary>Sets or clears the generated catalog document language.</summary>
    public PdfOptions SetLanguage(string? language) {
        Language = language;
        return this;
    }

    /// <summary>Enables or disables generated catalog page labels.</summary>
    public PdfOptions SetPageLabels(bool include = true, string? prefix = null) {
        IncludePageLabels = include;
        PageLabelPrefix = prefix;
        return this;
    }

    /// <summary>Sets or clears generated catalog viewer preferences.</summary>
    public PdfOptions SetViewerPreferences(PdfViewerPreferencesOptions? preferences) {
        ViewerPreferences = preferences;
        return this;
    }

    /// <summary>Configures generated catalog viewer preferences.</summary>
    public PdfOptions ConfigureViewerPreferences(System.Action<PdfViewerPreferencesOptions> configure) {
        Guard.NotNull(configure, nameof(configure));
        var preferences = _viewerPreferences?.Clone() ?? new PdfViewerPreferencesOptions();
        configure(preferences);
        _viewerPreferences = preferences;
        return this;
    }

    /// <summary>Adds an embedded file associated with the generated PDF catalog.</summary>
    public PdfOptions AddEmbeddedFile(PdfEmbeddedFile file) {
        Guard.NotNull(file, nameof(file));
        if (_embeddedFiles != null) {
            foreach (PdfEmbeddedFile existingFile in _embeddedFiles) {
                if (string.Equals(existingFile.FileName, file.FileName, System.StringComparison.Ordinal)) {
                    throw new System.ArgumentException("PDF embedded file names must be unique within a generated document.", nameof(file));
                }
            }
        }

        (_embeddedFiles ??= new System.Collections.Generic.List<PdfEmbeddedFile>()).Add(file.Clone());
        return this;
    }

    /// <summary>Adds an embedded file associated with the generated PDF catalog.</summary>
    public PdfOptions AddEmbeddedFile(
        string fileName,
        byte[] data,
        string? mimeType = null,
        PdfAssociatedFileRelationship relationship = PdfAssociatedFileRelationship.Unspecified,
        string? description = null) {
        return AddEmbeddedFile(new PdfEmbeddedFile(fileName, data, mimeType, relationship, description));
    }

    /// <summary>Removes all embedded files associated with the generated PDF catalog.</summary>
    public PdfOptions ClearEmbeddedFiles() {
        _embeddedFiles?.Clear();
        return this;
    }

    internal bool TryGetEmbeddedStandardFont(PdfStandardFont font, out PdfEmbeddedFont? embeddedFont) {
        if (_embeddedFonts != null && _embeddedFonts.TryGetValue(font, out PdfEmbeddedFont? value)) {
            embeddedFont = value.Clone();
            return true;
        }

        embeddedFont = null;
        return false;
    }

    /// <summary>When true, renders header text using <see cref="HeaderFormat"/>.</summary>
    public bool ShowHeader { get; set; }
    /// <summary>Header text format, supports {page} and {pages}. Default: empty.</summary>
    public string HeaderFormat { get; set; } = string.Empty;
    /// <summary>When true, page 1 uses first-page header/footer content instead of the running header/footer.</summary>
    public bool DifferentFirstPageHeaderFooter { get; set; }
    /// <summary>Header text format used on page 1 when <see cref="DifferentFirstPageHeaderFooter"/> is true. Supports {page} and {pages}.</summary>
    public string FirstPageHeaderFormat { get; set; } = string.Empty;
    /// <summary>When true, even-numbered pages use even-page header/footer content instead of the running odd-page content.</summary>
    public bool DifferentOddAndEvenPagesHeaderFooter { get; set; }
    /// <summary>Header text format used on even-numbered pages when <see cref="DifferentOddAndEvenPagesHeaderFooter"/> is true. Supports {page} and {pages}.</summary>
    public string EvenPageHeaderFormat { get; set; } = string.Empty;
    /// <summary>Header font.</summary>
    public PdfStandardFont HeaderFont {
        get => _headerFont;
        set {
            Guard.StandardFont(value, nameof(HeaderFont), "PDF header font must be one of the supported standard PDF fonts.");
            _headerFont = value;
        }
    }
    /// <summary>Header font size in points.</summary>
    public double HeaderFontSize { get; set; } = 9;
    /// <summary>Header text color. When null, the current PDF fill color is preserved.</summary>
    public PdfColor? HeaderTextColor { get; set; }
    /// <summary>Header alignment.</summary>
    public PdfAlign HeaderAlign {
        get => _headerAlign;
        set {
            Guard.LeftCenterRightAlign(value, nameof(HeaderAlign), "PDF header");
            _headerAlign = value;
        }
    }
    /// <summary>Header baseline Y offset above the top margin (points). Default 18.</summary>
    public double HeaderOffsetY { get; set; } = 18;

    /// <summary>When true, renders page numbers in the footer using <see cref="FooterFormat"/>.</summary>
    public bool ShowPageNumbers { get; set; } // default false
    /// <summary>Footer text format, supports {page} and {pages}. Default: "Page {page}/{pages}".</summary>
    public string FooterFormat { get; set; } = "Page {page}/{pages}";
    /// <summary>Footer text format used on page 1 when <see cref="DifferentFirstPageHeaderFooter"/> is true. Supports {page} and {pages}.</summary>
    public string FirstPageFooterFormat { get; set; } = string.Empty;
    /// <summary>Footer text format used on even-numbered pages when <see cref="DifferentOddAndEvenPagesHeaderFooter"/> is true. Supports {page} and {pages}.</summary>
    public string EvenPageFooterFormat { get; set; } = string.Empty;
    /// <summary>Footer font.</summary>
    public PdfStandardFont FooterFont {
        get => _footerFont;
        set {
            Guard.StandardFont(value, nameof(FooterFont), "PDF footer font must be one of the supported standard PDF fonts.");
            _footerFont = value;
        }
    }
    /// <summary>Footer font size in points.</summary>
    public double FooterFontSize { get; set; } = 9;
    /// <summary>Footer text color. When null, the current PDF fill color is preserved.</summary>
    public PdfColor? FooterTextColor { get; set; }
    /// <summary>Footer alignment.</summary>
    public PdfAlign FooterAlign {
        get => _footerAlign;
        set {
            Guard.LeftCenterRightAlign(value, nameof(FooterAlign), "PDF footer");
            _footerAlign = value;
        }
    }
    /// <summary>Footer baseline Y position from bottom margin (points). Default 18.</summary>
    public double FooterOffsetY { get; set; } = 18;

    /// <summary>First visible page number for this document or section flow. Default 1.</summary>
    public int PageNumberStart {
        get => _pageNumberStart;
        set {
            if (value < 1) {
                throw new System.ArgumentOutOfRangeException(nameof(PageNumberStart), "PDF page number start must be a positive value.");
            }

            _pageNumberStart = value;
            _hasExplicitPageNumberStart = true;
        }
    }
    internal bool HasExplicitPageNumberStart => _hasExplicitPageNumberStart;

    /// <summary>Visible numbering style for generated page tokens. Default Arabic.</summary>
    public PdfPageNumberStyle PageNumberStyle {
        get => _pageNumberStyle;
        set {
            Guard.PageNumberStyle(value, nameof(PageNumberStyle));
            _pageNumberStyle = value;
        }
    }

    /// <summary>Default text color for blocks when none is specified.</summary>
    public PdfColor? DefaultTextColor { get; set; }
    /// <summary>Default paragraph style applied when a paragraph does not specify its own style.</summary>
    public PdfParagraphStyle? DefaultParagraphStyle {
        get => _defaultParagraphStyle?.Clone();
        set => _defaultParagraphStyle = value?.Clone();
    }
    /// <summary>Default table style applied when none is provided.</summary>
    public PdfTableStyle? DefaultTableStyle {
        get => _defaultTableStyle?.Clone();
        set => _defaultTableStyle = value?.Clone();
    }
    /// <summary>Default heading styles applied when H1/H2/H3 blocks do not specify their own style.</summary>
    public PdfHeadingStyles? DefaultHeadingStyles {
        get => _defaultHeadingStyles?.Clone();
        set => _defaultHeadingStyles = value?.Clone();
    }
    /// <summary>Default list style applied when bullet and numbered lists do not specify their own style.</summary>
    public PdfListStyle? DefaultListStyle {
        get => _defaultListStyle?.Clone();
        set => _defaultListStyle = value?.Clone();
    }
    /// <summary>Default panel style applied when panel paragraphs do not specify their own style.</summary>
    public PanelStyle? DefaultPanelStyle {
        get => _defaultPanelStyle?.Clone();
        set => _defaultPanelStyle = value?.Clone();
    }
    /// <summary>Default horizontal rule style applied when horizontal rules do not specify their own style.</summary>
    public PdfHorizontalRuleStyle? DefaultHorizontalRuleStyle {
        get => _defaultHorizontalRuleStyle?.Clone();
        set => _defaultHorizontalRuleStyle = value?.Clone();
    }
    /// <summary>Default image placement style applied when images do not specify their own style.</summary>
    public PdfImageStyle? DefaultImageStyle {
        get => _defaultImageStyle?.Clone();
        set => _defaultImageStyle = value?.Clone();
    }
    /// <summary>Default placement style for OfficeIMO.Drawing-backed flow objects.</summary>
    public PdfDrawingStyle? DefaultDrawingStyle {
        get => _defaultDrawingStyle?.Clone();
        set => _defaultDrawingStyle = value?.Clone();
    }
    /// <summary>Default row/column layout style applied when rows do not specify their own style.</summary>
    public PdfRowStyle? DefaultRowStyle {
        get => _defaultRowStyle?.Clone();
        set => _defaultRowStyle = value?.Clone();
    }
    /// <summary>Optional debug overlays (margins, baselines, boxes).</summary>
    public PdfDebugOptions? Debug { get; set; }

    /// <summary>When true, H1/H2/H3 blocks are written as PDF outline/bookmark entries.</summary>
    public bool CreateOutlineFromHeadings { get; set; }

    /// <summary>Advanced footer template segments. When set, overrides FooterFormat.</summary>
    public System.Collections.Generic.List<FooterSegment>? FooterSegments {
        get => _footerSegments is null ? null : new System.Collections.Generic.List<FooterSegment>(_footerSegments);
        set => _footerSegments = value is null ? null : new System.Collections.Generic.List<FooterSegment>(value);
    }

    /// <summary>Advanced page-1 footer template segments used when <see cref="DifferentFirstPageHeaderFooter"/> is true.</summary>
    public System.Collections.Generic.List<FooterSegment>? FirstPageFooterSegments {
        get => _firstPageFooterSegments is null ? null : new System.Collections.Generic.List<FooterSegment>(_firstPageFooterSegments);
        set => _firstPageFooterSegments = value is null ? null : new System.Collections.Generic.List<FooterSegment>(value);
    }

    /// <summary>Advanced even-page footer template segments used when <see cref="DifferentOddAndEvenPagesHeaderFooter"/> is true.</summary>
    public System.Collections.Generic.List<FooterSegment>? EvenPageFooterSegments {
        get => _evenPageFooterSegments is null ? null : new System.Collections.Generic.List<FooterSegment>(_evenPageFooterSegments);
        set => _evenPageFooterSegments = value is null ? null : new System.Collections.Generic.List<FooterSegment>(value);
    }

    /// <summary>Applies reusable default styles to this options object.</summary>
    public PdfOptions ApplyTheme(PdfTheme theme) {
        Guard.NotNull(theme, nameof(theme));
        theme.Clone().ApplyTo(this);
        return this;
    }

    internal bool HasHeaderContent => (ShowHeader && HeaderFormat != null && HeaderFormat.Length > 0) ||
        (_headerSegments != null && _headerSegments.Count > 0) ||
        HasHeaderZoneContent ||
        HasHeaderImageContent;
    internal bool HasFooterContent => ShowPageNumbers ||
        (_footerSegments != null && _footerSegments.Count > 0) ||
        HasFooterZoneContent ||
        HasFooterImageContent;
    internal bool HasHeaderZoneContent =>
        !string.IsNullOrEmpty(_headerLeftFormat) ||
        !string.IsNullOrEmpty(_headerCenterFormat) ||
        !string.IsNullOrEmpty(_headerRightFormat);
    internal bool HasFirstPageHeaderZoneContent =>
        !string.IsNullOrEmpty(_firstPageHeaderLeftFormat) ||
        !string.IsNullOrEmpty(_firstPageHeaderCenterFormat) ||
        !string.IsNullOrEmpty(_firstPageHeaderRightFormat);
    internal bool HasEvenPageHeaderZoneContent =>
        !string.IsNullOrEmpty(_evenPageHeaderLeftFormat) ||
        !string.IsNullOrEmpty(_evenPageHeaderCenterFormat) ||
        !string.IsNullOrEmpty(_evenPageHeaderRightFormat);
    internal bool HasFooterZoneContent =>
        !string.IsNullOrEmpty(_footerLeftFormat) ||
        !string.IsNullOrEmpty(_footerCenterFormat) ||
        !string.IsNullOrEmpty(_footerRightFormat);
    internal bool HasFirstPageFooterZoneContent =>
        !string.IsNullOrEmpty(_firstPageFooterLeftFormat) ||
        !string.IsNullOrEmpty(_firstPageFooterCenterFormat) ||
        !string.IsNullOrEmpty(_firstPageFooterRightFormat);
    internal bool HasEvenPageFooterZoneContent =>
        !string.IsNullOrEmpty(_evenPageFooterLeftFormat) ||
        !string.IsNullOrEmpty(_evenPageFooterCenterFormat) ||
        !string.IsNullOrEmpty(_evenPageFooterRightFormat);
    internal bool HasHeaderImageContent => _headerImages != null && _headerImages.Count > 0;
    internal bool HasFirstPageHeaderImageContent => _firstPageHeaderImages != null && _firstPageHeaderImages.Count > 0;
    internal bool HasEvenPageHeaderImageContent => _evenPageHeaderImages != null && _evenPageHeaderImages.Count > 0;
    internal bool HasFooterImageContent => _footerImages != null && _footerImages.Count > 0;
    internal bool HasFirstPageFooterImageContent => _firstPageFooterImages != null && _firstPageFooterImages.Count > 0;
    internal bool HasEvenPageFooterImageContent => _evenPageFooterImages != null && _evenPageFooterImages.Count > 0;
    internal bool HasHeaderContentForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return (FirstPageHeaderFormat != null && FirstPageHeaderFormat.Length > 0) ||
                (_firstPageHeaderSegments != null && _firstPageHeaderSegments.Count > 0) ||
                HasFirstPageHeaderZoneContent ||
                HasFirstPageHeaderImageContent;
        }

        if (IsEvenPageVariant(pageNumber)) {
            return (EvenPageHeaderFormat != null && EvenPageHeaderFormat.Length > 0) ||
                (_evenPageHeaderSegments != null && _evenPageHeaderSegments.Count > 0) ||
                HasEvenPageHeaderZoneContent ||
                HasEvenPageHeaderImageContent;
        }

        return HasHeaderContent;
    }

    internal bool HasHeaderTextContentForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return (FirstPageHeaderFormat != null && FirstPageHeaderFormat.Length > 0) ||
                (_firstPageHeaderSegments != null && _firstPageHeaderSegments.Count > 0) ||
                HasFirstPageHeaderZoneContent;
        }

        if (IsEvenPageVariant(pageNumber)) {
            return (EvenPageHeaderFormat != null && EvenPageHeaderFormat.Length > 0) ||
                (_evenPageHeaderSegments != null && _evenPageHeaderSegments.Count > 0) ||
                HasEvenPageHeaderZoneContent;
        }

        return (ShowHeader && HeaderFormat != null && HeaderFormat.Length > 0) ||
            (_headerSegments != null && _headerSegments.Count > 0) ||
            HasHeaderZoneContent;
    }

    internal bool HasFooterContentForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return (FirstPageFooterFormat != null && FirstPageFooterFormat.Length > 0) ||
                (_firstPageFooterSegments != null && _firstPageFooterSegments.Count > 0) ||
                HasFirstPageFooterZoneContent ||
                HasFirstPageFooterImageContent;
        }

        if (IsEvenPageVariant(pageNumber)) {
            return (EvenPageFooterFormat != null && EvenPageFooterFormat.Length > 0) ||
                (_evenPageFooterSegments != null && _evenPageFooterSegments.Count > 0) ||
                HasEvenPageFooterZoneContent ||
                HasEvenPageFooterImageContent;
        }

        return HasFooterContent;
    }

    internal bool HasFooterTextContentForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return (FirstPageFooterFormat != null && FirstPageFooterFormat.Length > 0) ||
                (_firstPageFooterSegments != null && _firstPageFooterSegments.Count > 0) ||
                HasFirstPageFooterZoneContent;
        }

        if (IsEvenPageVariant(pageNumber)) {
            return (EvenPageFooterFormat != null && EvenPageFooterFormat.Length > 0) ||
                (_evenPageFooterSegments != null && _evenPageFooterSegments.Count > 0) ||
                HasEvenPageFooterZoneContent;
        }

        return ShowPageNumbers ||
            (_footerSegments != null && _footerSegments.Count > 0) ||
            HasFooterZoneContent;
    }

    internal string GetHeaderFormatForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return FirstPageHeaderFormat;
        }

        return IsEvenPageVariant(pageNumber) ? EvenPageHeaderFormat : HeaderFormat;
    }

    internal System.Collections.Generic.IReadOnlyList<FooterSegment>? GetHeaderSegmentsForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return _firstPageHeaderSegments;
        }

        return IsEvenPageVariant(pageNumber) ? _evenPageHeaderSegments : _headerSegments;
    }

    internal (string? Left, string? Center, string? Right) GetHeaderZonesForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return (_firstPageHeaderLeftFormat, _firstPageHeaderCenterFormat, _firstPageHeaderRightFormat);
        }

        if (IsEvenPageVariant(pageNumber)) {
            return (_evenPageHeaderLeftFormat, _evenPageHeaderCenterFormat, _evenPageHeaderRightFormat);
        }

        return (_headerLeftFormat, _headerCenterFormat, _headerRightFormat);
    }

    internal System.Collections.Generic.IReadOnlyList<PdfHeaderFooterImage> GetHeaderImagesForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return _firstPageHeaderImages != null ? _firstPageHeaderImages : (System.Collections.Generic.IReadOnlyList<PdfHeaderFooterImage>)System.Array.Empty<PdfHeaderFooterImage>();
        }

        if (IsEvenPageVariant(pageNumber)) {
            return _evenPageHeaderImages != null ? _evenPageHeaderImages : (System.Collections.Generic.IReadOnlyList<PdfHeaderFooterImage>)System.Array.Empty<PdfHeaderFooterImage>();
        }

        return _headerImages != null ? _headerImages : (System.Collections.Generic.IReadOnlyList<PdfHeaderFooterImage>)System.Array.Empty<PdfHeaderFooterImage>();
    }

    internal System.Collections.Generic.IReadOnlyList<PdfHeaderFooterShape> GetHeaderShapesForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return _firstPageHeaderShapes != null ? _firstPageHeaderShapes : (System.Collections.Generic.IReadOnlyList<PdfHeaderFooterShape>)System.Array.Empty<PdfHeaderFooterShape>();
        }

        if (IsEvenPageVariant(pageNumber)) {
            return _evenPageHeaderShapes != null ? _evenPageHeaderShapes : (System.Collections.Generic.IReadOnlyList<PdfHeaderFooterShape>)System.Array.Empty<PdfHeaderFooterShape>();
        }

        return _headerShapes != null ? _headerShapes : (System.Collections.Generic.IReadOnlyList<PdfHeaderFooterShape>)System.Array.Empty<PdfHeaderFooterShape>();
    }

    internal string GetFooterFormatForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return FirstPageFooterFormat;
        }

        return IsEvenPageVariant(pageNumber) ? EvenPageFooterFormat : FooterFormat;
    }

    internal System.Collections.Generic.IReadOnlyList<FooterSegment>? GetFooterSegmentsForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return _firstPageFooterSegments;
        }

        return IsEvenPageVariant(pageNumber) ? _evenPageFooterSegments : _footerSegments;
    }

    internal (string? Left, string? Center, string? Right) GetFooterZonesForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return (_firstPageFooterLeftFormat, _firstPageFooterCenterFormat, _firstPageFooterRightFormat);
        }

        if (IsEvenPageVariant(pageNumber)) {
            return (_evenPageFooterLeftFormat, _evenPageFooterCenterFormat, _evenPageFooterRightFormat);
        }

        return (_footerLeftFormat, _footerCenterFormat, _footerRightFormat);
    }

    internal System.Collections.Generic.IReadOnlyList<PdfHeaderFooterImage> GetFooterImagesForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return _firstPageFooterImages != null ? _firstPageFooterImages : (System.Collections.Generic.IReadOnlyList<PdfHeaderFooterImage>)System.Array.Empty<PdfHeaderFooterImage>();
        }

        if (IsEvenPageVariant(pageNumber)) {
            return _evenPageFooterImages != null ? _evenPageFooterImages : (System.Collections.Generic.IReadOnlyList<PdfHeaderFooterImage>)System.Array.Empty<PdfHeaderFooterImage>();
        }

        return _footerImages != null ? _footerImages : (System.Collections.Generic.IReadOnlyList<PdfHeaderFooterImage>)System.Array.Empty<PdfHeaderFooterImage>();
    }

    internal System.Collections.Generic.IReadOnlyList<PdfHeaderFooterShape> GetFooterShapesForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return _firstPageFooterShapes != null ? _firstPageFooterShapes : (System.Collections.Generic.IReadOnlyList<PdfHeaderFooterShape>)System.Array.Empty<PdfHeaderFooterShape>();
        }

        if (IsEvenPageVariant(pageNumber)) {
            return _evenPageFooterShapes != null ? _evenPageFooterShapes : (System.Collections.Generic.IReadOnlyList<PdfHeaderFooterShape>)System.Array.Empty<PdfHeaderFooterShape>();
        }

        return _footerShapes != null ? _footerShapes : (System.Collections.Generic.IReadOnlyList<PdfHeaderFooterShape>)System.Array.Empty<PdfHeaderFooterShape>();
    }

    private bool IsEvenPageVariant(int pageNumber) =>
        DifferentOddAndEvenPagesHeaderFooter && pageNumber > 0 && pageNumber % 2 == 0;
    internal PdfParagraphStyle? DefaultParagraphStyleSnapshot => _defaultParagraphStyle;
    internal PdfTableStyle? DefaultTableStyleSnapshot => _defaultTableStyle;
    internal PdfHeadingStyles? DefaultHeadingStylesSnapshot => _defaultHeadingStyles;
    internal PdfListStyle? DefaultListStyleSnapshot => _defaultListStyle;
    internal PanelStyle? DefaultPanelStyleSnapshot => _defaultPanelStyle;
    internal PdfHorizontalRuleStyle? DefaultHorizontalRuleStyleSnapshot => _defaultHorizontalRuleStyle;
    internal PdfImageStyle? DefaultImageStyleSnapshot => _defaultImageStyle;
    internal PdfDrawingStyle? DefaultDrawingStyleSnapshot => _defaultDrawingStyle;
    internal PdfRowStyle? DefaultRowStyleSnapshot => _defaultRowStyle;

    /// <summary>Sets the default style for a built-in heading level.</summary>
    public PdfOptions SetDefaultHeadingStyle(int level, PdfHeadingStyle style) {
        Guard.NotNull(style, nameof(style));
        (_defaultHeadingStyles ??= new PdfHeadingStyles()).Set(level, style);
        return this;
    }

    /// <summary>Creates a deep copy of the options.</summary>
    public PdfOptions Clone() {
        var clone = new PdfOptions {
            PageWidth = PageWidth,
            PageHeight = PageHeight,
            BackgroundColor = BackgroundColor,
            TextWatermark = _textWatermark?.Clone(),
            ImageWatermark = _imageWatermark?.Clone(),
            PageBorder = _pageBorder?.Clone(),
            PageBackgroundImage = _pageBackgroundImage?.Clone(),
            PageBackgroundShapes = ClonePageBackgroundShapes(_pageBackgroundShapes),
            MarginLeft = MarginLeft,
            MarginRight = MarginRight,
            MarginTop = MarginTop,
            MarginBottom = MarginBottom,
            DefaultFont = DefaultFont,
            DefaultFontSize = DefaultFontSize,
            CompressContentStreams = CompressContentStreams,
            IncludeStandardFontToUnicodeMaps = IncludeStandardFontToUnicodeMaps,
            CompressEmbeddedFonts = CompressEmbeddedFonts,
            IncludeXmpMetadata = IncludeXmpMetadata,
            IncludePageLabels = IncludePageLabels,
            PageLabelPrefix = PageLabelPrefix,
            ComplianceProfile = ComplianceProfile,
            OutputIntent = _outputIntent?.Clone(),
            Language = Language,
            ViewerPreferences = _viewerPreferences?.Clone(),
            _embeddedFonts = CloneEmbeddedFonts(_embeddedFonts),
            _embeddedFiles = CloneEmbeddedFiles(_embeddedFiles),
            ShowHeader = ShowHeader,
            HeaderFormat = HeaderFormat,
            DifferentFirstPageHeaderFooter = DifferentFirstPageHeaderFooter,
            FirstPageHeaderFormat = FirstPageHeaderFormat,
            DifferentOddAndEvenPagesHeaderFooter = DifferentOddAndEvenPagesHeaderFooter,
            EvenPageHeaderFormat = EvenPageHeaderFormat,
            HeaderFont = HeaderFont,
            HeaderFontSize = HeaderFontSize,
            HeaderTextColor = HeaderTextColor,
            HeaderAlign = HeaderAlign,
            HeaderOffsetY = HeaderOffsetY,
            ShowPageNumbers = ShowPageNumbers,
            FooterFormat = FooterFormat,
            FirstPageFooterFormat = FirstPageFooterFormat,
            EvenPageFooterFormat = EvenPageFooterFormat,
            FooterFont = FooterFont,
            FooterFontSize = FooterFontSize,
            FooterTextColor = FooterTextColor,
            FooterAlign = FooterAlign,
            FooterOffsetY = FooterOffsetY,
            PageNumberStyle = PageNumberStyle,
            DefaultTextColor = DefaultTextColor,
            DefaultParagraphStyle = _defaultParagraphStyle?.Clone(),
            DefaultTableStyle = _defaultTableStyle?.Clone(),
            DefaultHeadingStyles = _defaultHeadingStyles?.Clone(),
            DefaultListStyle = _defaultListStyle?.Clone(),
            DefaultPanelStyle = _defaultPanelStyle?.Clone(),
            DefaultHorizontalRuleStyle = _defaultHorizontalRuleStyle?.Clone(),
            DefaultImageStyle = _defaultImageStyle?.Clone(),
            DefaultDrawingStyle = _defaultDrawingStyle?.Clone(),
            DefaultRowStyle = _defaultRowStyle?.Clone(),
            CreateOutlineFromHeadings = CreateOutlineFromHeadings,
            Debug = Debug is null ? null : new PdfDebugOptions {
                ShowContentArea = Debug.ShowContentArea,
                ShowTableBaselines = Debug.ShowTableBaselines,
                ShowTableRowBoxes = Debug.ShowTableRowBoxes,
                ShowTableColumnGuides = Debug.ShowTableColumnGuides
            },
            _headerSegments = _headerSegments is null ? null : new System.Collections.Generic.List<FooterSegment>(_headerSegments),
            _firstPageHeaderSegments = _firstPageHeaderSegments is null ? null : new System.Collections.Generic.List<FooterSegment>(_firstPageHeaderSegments),
            _evenPageHeaderSegments = _evenPageHeaderSegments is null ? null : new System.Collections.Generic.List<FooterSegment>(_evenPageHeaderSegments),
            FooterSegments = _footerSegments is null ? null : new System.Collections.Generic.List<FooterSegment>(_footerSegments),
            FirstPageFooterSegments = _firstPageFooterSegments is null ? null : new System.Collections.Generic.List<FooterSegment>(_firstPageFooterSegments),
            EvenPageFooterSegments = _evenPageFooterSegments is null ? null : new System.Collections.Generic.List<FooterSegment>(_evenPageFooterSegments),
            _headerLeftFormat = _headerLeftFormat,
            _headerCenterFormat = _headerCenterFormat,
            _headerRightFormat = _headerRightFormat,
            _firstPageHeaderLeftFormat = _firstPageHeaderLeftFormat,
            _firstPageHeaderCenterFormat = _firstPageHeaderCenterFormat,
            _firstPageHeaderRightFormat = _firstPageHeaderRightFormat,
            _evenPageHeaderLeftFormat = _evenPageHeaderLeftFormat,
            _evenPageHeaderCenterFormat = _evenPageHeaderCenterFormat,
            _evenPageHeaderRightFormat = _evenPageHeaderRightFormat,
            _footerLeftFormat = _footerLeftFormat,
            _footerCenterFormat = _footerCenterFormat,
            _footerRightFormat = _footerRightFormat,
            _firstPageFooterLeftFormat = _firstPageFooterLeftFormat,
            _firstPageFooterCenterFormat = _firstPageFooterCenterFormat,
            _firstPageFooterRightFormat = _firstPageFooterRightFormat,
            _evenPageFooterLeftFormat = _evenPageFooterLeftFormat,
            _evenPageFooterCenterFormat = _evenPageFooterCenterFormat,
            _evenPageFooterRightFormat = _evenPageFooterRightFormat,
            _headerImages = CloneHeaderFooterImages(_headerImages),
            _firstPageHeaderImages = CloneHeaderFooterImages(_firstPageHeaderImages),
            _evenPageHeaderImages = CloneHeaderFooterImages(_evenPageHeaderImages),
            _footerImages = CloneHeaderFooterImages(_footerImages),
            _firstPageFooterImages = CloneHeaderFooterImages(_firstPageFooterImages),
            _evenPageFooterImages = CloneHeaderFooterImages(_evenPageFooterImages),
            _headerShapes = CloneHeaderFooterShapes(_headerShapes),
            _firstPageHeaderShapes = CloneHeaderFooterShapes(_firstPageHeaderShapes),
            _evenPageHeaderShapes = CloneHeaderFooterShapes(_evenPageHeaderShapes),
            _footerShapes = CloneHeaderFooterShapes(_footerShapes),
            _firstPageFooterShapes = CloneHeaderFooterShapes(_firstPageFooterShapes),
            _evenPageFooterShapes = CloneHeaderFooterShapes(_evenPageFooterShapes)
        };
        clone._pageNumberStart = _pageNumberStart;
        clone._hasExplicitPageNumberStart = _hasExplicitPageNumberStart;
        return clone;
    }

    private static System.Collections.Generic.Dictionary<PdfStandardFont, PdfEmbeddedFont>? CloneEmbeddedFonts(System.Collections.Generic.Dictionary<PdfStandardFont, PdfEmbeddedFont>? fonts) {
        if (fonts == null) {
            return null;
        }

        var clone = new System.Collections.Generic.Dictionary<PdfStandardFont, PdfEmbeddedFont>();
        foreach (var font in fonts) {
            clone[font.Key] = font.Value.Clone();
        }

        return clone;
    }

    private static System.Collections.Generic.List<PdfEmbeddedFile>? CloneEmbeddedFiles(System.Collections.Generic.IEnumerable<PdfEmbeddedFile>? files) {
        if (files == null) {
            return null;
        }

        var clone = new System.Collections.Generic.List<PdfEmbeddedFile>();
        foreach (PdfEmbeddedFile file in files) {
            Guard.NotNull(file, nameof(EmbeddedFiles));
            clone.Add(file.Clone());
        }

        return clone;
    }

    private static System.Collections.Generic.List<PdfHeaderFooterImage>? CloneHeaderFooterImages(System.Collections.Generic.List<PdfHeaderFooterImage>? images) {
        if (images == null) {
            return null;
        }

        var clone = new System.Collections.Generic.List<PdfHeaderFooterImage>(images.Count);
        foreach (PdfHeaderFooterImage image in images) {
            clone.Add(image.Clone());
        }

        return clone;
    }

    private static System.Collections.Generic.List<PdfHeaderFooterShape>? CloneHeaderFooterShapes(System.Collections.Generic.List<PdfHeaderFooterShape>? shapes) {
        if (shapes == null) {
            return null;
        }

        var clone = new System.Collections.Generic.List<PdfHeaderFooterShape>(shapes.Count);
        foreach (PdfHeaderFooterShape shape in shapes) {
            clone.Add(shape.Clone());
        }

        return clone;
    }

    private static System.Collections.Generic.List<PdfPageBackgroundShape>? ClonePageBackgroundShapes(System.Collections.Generic.IEnumerable<PdfPageBackgroundShape>? shapes) {
        if (shapes == null) {
            return null;
        }

        var clone = new System.Collections.Generic.List<PdfPageBackgroundShape>();
        foreach (PdfPageBackgroundShape shape in shapes) {
            Guard.NotNull(shape, nameof(PageBackgroundShapes));
            clone.Add(shape.Clone());
        }

        return clone;
    }

    internal void AddPageBackgroundShape(PdfPageBackgroundShape shape) {
        Guard.NotNull(shape, nameof(shape));
        (_pageBackgroundShapes ??= new System.Collections.Generic.List<PdfPageBackgroundShape>()).Add(shape.Clone());
    }

    internal void ClearPageBackgroundShapes() {
        _pageBackgroundShapes = null;
    }

    internal void ClearPageNumberStartOverride() {
        _hasExplicitPageNumberStart = false;
    }

    internal System.Collections.Generic.List<FooterSegment> ResetHeaderSegmentsForCompose() {
        _headerSegments = new System.Collections.Generic.List<FooterSegment>();
        ShowHeader = true;
        return _headerSegments;
    }

    internal void ClearHeaderSegmentsForCompose() {
        _headerSegments = null;
    }

    internal void SetHeaderZonesForCompose(string? left, string? center, string? right) {
        ValidateZones(left, center, right, nameof(left));
        ClearHeaderSegmentsForCompose();
        HeaderFormat = string.Empty;
        ShowHeader = true;
        _headerLeftFormat = left;
        _headerCenterFormat = center;
        _headerRightFormat = right;
    }

    internal void ClearHeaderZonesForCompose() {
        _headerLeftFormat = null;
        _headerCenterFormat = null;
        _headerRightFormat = null;
    }

    internal void AddHeaderImageForCompose(PdfHeaderFooterImage image) {
        Guard.NotNull(image, nameof(image));
        ShowHeader = true;
        (_headerImages ??= new System.Collections.Generic.List<PdfHeaderFooterImage>()).Add(image.Clone());
    }

    internal void SetFirstPageHeaderZonesForCompose(string? left, string? center, string? right) {
        ValidateZones(left, center, right, nameof(left));
        ClearFirstPageHeaderSegmentsForCompose();
        DifferentFirstPageHeaderFooter = true;
        FirstPageHeaderFormat = string.Empty;
        _firstPageHeaderLeftFormat = left;
        _firstPageHeaderCenterFormat = center;
        _firstPageHeaderRightFormat = right;
    }

    internal void ClearFirstPageHeaderZonesForCompose() {
        _firstPageHeaderLeftFormat = null;
        _firstPageHeaderCenterFormat = null;
        _firstPageHeaderRightFormat = null;
    }

    internal void AddFirstPageHeaderImageForCompose(PdfHeaderFooterImage image) {
        Guard.NotNull(image, nameof(image));
        DifferentFirstPageHeaderFooter = true;
        (_firstPageHeaderImages ??= new System.Collections.Generic.List<PdfHeaderFooterImage>()).Add(image.Clone());
    }

    internal void SetEvenPageHeaderZonesForCompose(string? left, string? center, string? right) {
        ValidateZones(left, center, right, nameof(left));
        ClearEvenPageHeaderSegmentsForCompose();
        DifferentOddAndEvenPagesHeaderFooter = true;
        EvenPageHeaderFormat = string.Empty;
        _evenPageHeaderLeftFormat = left;
        _evenPageHeaderCenterFormat = center;
        _evenPageHeaderRightFormat = right;
    }

    internal void ClearEvenPageHeaderZonesForCompose() {
        _evenPageHeaderLeftFormat = null;
        _evenPageHeaderCenterFormat = null;
        _evenPageHeaderRightFormat = null;
    }

    internal void AddEvenPageHeaderImageForCompose(PdfHeaderFooterImage image) {
        Guard.NotNull(image, nameof(image));
        DifferentOddAndEvenPagesHeaderFooter = true;
        (_evenPageHeaderImages ??= new System.Collections.Generic.List<PdfHeaderFooterImage>()).Add(image.Clone());
    }

    internal void AddHeaderShapeForCompose(PdfHeaderFooterShape shape) {
        Guard.NotNull(shape, nameof(shape));
        ShowHeader = true;
        (_headerShapes ??= new System.Collections.Generic.List<PdfHeaderFooterShape>()).Add(shape.Clone());
    }

    internal void AddFirstPageHeaderShapeForCompose(PdfHeaderFooterShape shape) {
        Guard.NotNull(shape, nameof(shape));
        ShowHeader = true;
        DifferentFirstPageHeaderFooter = true;
        (_firstPageHeaderShapes ??= new System.Collections.Generic.List<PdfHeaderFooterShape>()).Add(shape.Clone());
    }

    internal void AddEvenPageHeaderShapeForCompose(PdfHeaderFooterShape shape) {
        Guard.NotNull(shape, nameof(shape));
        ShowHeader = true;
        DifferentOddAndEvenPagesHeaderFooter = true;
        (_evenPageHeaderShapes ??= new System.Collections.Generic.List<PdfHeaderFooterShape>()).Add(shape.Clone());
    }

    internal System.Collections.Generic.List<FooterSegment> ResetFirstPageHeaderSegmentsForCompose() {
        _firstPageHeaderSegments = new System.Collections.Generic.List<FooterSegment>();
        DifferentFirstPageHeaderFooter = true;
        return _firstPageHeaderSegments;
    }

    internal void ClearFirstPageHeaderSegmentsForCompose() {
        _firstPageHeaderSegments = null;
    }

    internal System.Collections.Generic.List<FooterSegment> ResetEvenPageHeaderSegmentsForCompose() {
        _evenPageHeaderSegments = new System.Collections.Generic.List<FooterSegment>();
        DifferentOddAndEvenPagesHeaderFooter = true;
        return _evenPageHeaderSegments;
    }

    internal void ClearEvenPageHeaderSegmentsForCompose() {
        _evenPageHeaderSegments = null;
    }

    internal System.Collections.Generic.List<FooterSegment> ResetFooterSegmentsForCompose() {
        _footerSegments = new System.Collections.Generic.List<FooterSegment>();
        return _footerSegments;
    }

    internal void ClearFooterSegmentsForCompose() {
        _footerSegments = null;
    }

    internal void SetFooterZonesForCompose(string? left, string? center, string? right) {
        ValidateZones(left, center, right, nameof(left));
        ClearFooterSegmentsForCompose();
        FooterFormat = string.Empty;
        ShowPageNumbers = true;
        _footerLeftFormat = left;
        _footerCenterFormat = center;
        _footerRightFormat = right;
    }

    internal void ClearFooterZonesForCompose() {
        _footerLeftFormat = null;
        _footerCenterFormat = null;
        _footerRightFormat = null;
    }

    internal void AddFooterImageForCompose(PdfHeaderFooterImage image) {
        Guard.NotNull(image, nameof(image));
        (_footerImages ??= new System.Collections.Generic.List<PdfHeaderFooterImage>()).Add(image.Clone());
    }

    internal void SetFirstPageFooterZonesForCompose(string? left, string? center, string? right) {
        ValidateZones(left, center, right, nameof(left));
        ClearFirstPageFooterSegmentsForCompose();
        DifferentFirstPageHeaderFooter = true;
        FirstPageFooterFormat = string.Empty;
        _firstPageFooterLeftFormat = left;
        _firstPageFooterCenterFormat = center;
        _firstPageFooterRightFormat = right;
    }

    internal void ClearFirstPageFooterZonesForCompose() {
        _firstPageFooterLeftFormat = null;
        _firstPageFooterCenterFormat = null;
        _firstPageFooterRightFormat = null;
    }

    internal void AddFirstPageFooterImageForCompose(PdfHeaderFooterImage image) {
        Guard.NotNull(image, nameof(image));
        DifferentFirstPageHeaderFooter = true;
        (_firstPageFooterImages ??= new System.Collections.Generic.List<PdfHeaderFooterImage>()).Add(image.Clone());
    }

    internal void SetEvenPageFooterZonesForCompose(string? left, string? center, string? right) {
        ValidateZones(left, center, right, nameof(left));
        ClearEvenPageFooterSegmentsForCompose();
        DifferentOddAndEvenPagesHeaderFooter = true;
        EvenPageFooterFormat = string.Empty;
        _evenPageFooterLeftFormat = left;
        _evenPageFooterCenterFormat = center;
        _evenPageFooterRightFormat = right;
    }

    internal void ClearEvenPageFooterZonesForCompose() {
        _evenPageFooterLeftFormat = null;
        _evenPageFooterCenterFormat = null;
        _evenPageFooterRightFormat = null;
    }

    internal void AddEvenPageFooterImageForCompose(PdfHeaderFooterImage image) {
        Guard.NotNull(image, nameof(image));
        DifferentOddAndEvenPagesHeaderFooter = true;
        (_evenPageFooterImages ??= new System.Collections.Generic.List<PdfHeaderFooterImage>()).Add(image.Clone());
    }

    internal void AddFooterShapeForCompose(PdfHeaderFooterShape shape) {
        Guard.NotNull(shape, nameof(shape));
        (_footerShapes ??= new System.Collections.Generic.List<PdfHeaderFooterShape>()).Add(shape.Clone());
    }

    internal void AddFirstPageFooterShapeForCompose(PdfHeaderFooterShape shape) {
        Guard.NotNull(shape, nameof(shape));
        DifferentFirstPageHeaderFooter = true;
        (_firstPageFooterShapes ??= new System.Collections.Generic.List<PdfHeaderFooterShape>()).Add(shape.Clone());
    }

    internal void AddEvenPageFooterShapeForCompose(PdfHeaderFooterShape shape) {
        Guard.NotNull(shape, nameof(shape));
        DifferentOddAndEvenPagesHeaderFooter = true;
        (_evenPageFooterShapes ??= new System.Collections.Generic.List<PdfHeaderFooterShape>()).Add(shape.Clone());
    }

    internal System.Collections.Generic.List<FooterSegment> ResetFirstPageFooterSegmentsForCompose() {
        _firstPageFooterSegments = new System.Collections.Generic.List<FooterSegment>();
        DifferentFirstPageHeaderFooter = true;
        return _firstPageFooterSegments;
    }

    internal void ClearFirstPageFooterSegmentsForCompose() {
        _firstPageFooterSegments = null;
    }

    internal System.Collections.Generic.List<FooterSegment> ResetEvenPageFooterSegmentsForCompose() {
        _evenPageFooterSegments = new System.Collections.Generic.List<FooterSegment>();
        DifferentOddAndEvenPagesHeaderFooter = true;
        return _evenPageFooterSegments;
    }

    internal void ClearEvenPageFooterSegmentsForCompose() {
        _evenPageFooterSegments = null;
    }

    internal void Validate() {
        if (PageWidth <= 0 || double.IsNaN(PageWidth) || double.IsInfinity(PageWidth)) {
            throw new System.ArgumentException("PDF page width must be a positive finite value.");
        }

        if (PageHeight <= 0 || double.IsNaN(PageHeight) || double.IsInfinity(PageHeight)) {
            throw new System.ArgumentException("PDF page height must be a positive finite value.");
        }

        if (MarginLeft < 0 || double.IsNaN(MarginLeft) || double.IsInfinity(MarginLeft)) {
            throw new System.ArgumentException("PDF left margin must be a non-negative finite value.");
        }

        if (MarginRight < 0 || double.IsNaN(MarginRight) || double.IsInfinity(MarginRight)) {
            throw new System.ArgumentException("PDF right margin must be a non-negative finite value.");
        }

        if (MarginTop < 0 || double.IsNaN(MarginTop) || double.IsInfinity(MarginTop)) {
            throw new System.ArgumentException("PDF top margin must be a non-negative finite value.");
        }

        if (MarginBottom < 0 || double.IsNaN(MarginBottom) || double.IsInfinity(MarginBottom)) {
            throw new System.ArgumentException("PDF bottom margin must be a non-negative finite value.");
        }

        if (PageWidth - MarginLeft - MarginRight <= 0) {
            throw new System.ArgumentException("PDF margins must leave a positive content width.");
        }

        if (PageHeight - MarginTop - MarginBottom <= 0) {
            throw new System.ArgumentException("PDF margins must leave a positive content height.");
        }

        Guard.StandardFont(DefaultFont, nameof(DefaultFont), "PDF default font must be one of the supported standard PDF fonts.");
        Guard.StandardFont(HeaderFont, nameof(HeaderFont), "PDF header font must be one of the supported standard PDF fonts.");
        Guard.StandardFont(FooterFont, nameof(FooterFont), "PDF footer font must be one of the supported standard PDF fonts.");
        Guard.PageNumberStyle(PageNumberStyle, nameof(PageNumberStyle));
        Guard.ComplianceProfile(ComplianceProfile, nameof(ComplianceProfile));
        PdfPageLabelDictionaryBuilder.ValidatePrefix(PageLabelPrefix, nameof(PageLabelPrefix));

        if (DefaultFontSize <= 0 || double.IsNaN(DefaultFontSize) || double.IsInfinity(DefaultFontSize)) {
            throw new System.ArgumentException("PDF default font size must be a positive finite value.");
        }

        if (HeaderFontSize <= 0 || double.IsNaN(HeaderFontSize) || double.IsInfinity(HeaderFontSize)) {
            throw new System.ArgumentException("PDF header font size must be a positive finite value.");
        }

        if (FooterFontSize <= 0 || double.IsNaN(FooterFontSize) || double.IsInfinity(FooterFontSize)) {
            throw new System.ArgumentException("PDF footer font size must be a positive finite value.");
        }

        if (HeaderOffsetY < 0 || double.IsNaN(HeaderOffsetY) || double.IsInfinity(HeaderOffsetY)) {
            throw new System.ArgumentException("PDF header offset must be a non-negative finite value.");
        }

        if (HasHeaderContent && HeaderOffsetY > MarginTop) {
            throw new System.ArgumentException("PDF header offset must not exceed the top margin when header content is enabled.");
        }

        if (HasHeaderContentForPage(1) && HeaderOffsetY > MarginTop) {
            throw new System.ArgumentException("PDF header offset must not exceed the top margin when header content is enabled.");
        }

        if (HasHeaderContentForPage(2) && HeaderOffsetY > MarginTop) {
            throw new System.ArgumentException("PDF header offset must not exceed the top margin when header content is enabled.");
        }

        if (FooterOffsetY < 0 || double.IsNaN(FooterOffsetY) || double.IsInfinity(FooterOffsetY)) {
            throw new System.ArgumentException("PDF footer offset must be a non-negative finite value.");
        }

        if (PageNumberStart < 1) {
            throw new System.ArgumentException("PDF page number start must be a positive value.");
        }

        if (HasFooterContent && FooterOffsetY > MarginBottom) {
            throw new System.ArgumentException("PDF footer offset must not exceed the bottom margin when footer content is enabled.");
        }

        if (HasFooterContentForPage(1) && FooterOffsetY > MarginBottom) {
            throw new System.ArgumentException("PDF footer offset must not exceed the bottom margin when footer content is enabled.");
        }

        if (HasFooterContentForPage(2) && FooterOffsetY > MarginBottom) {
            throw new System.ArgumentException("PDF footer offset must not exceed the bottom margin when footer content is enabled.");
        }

        if (HeaderFormat == null) {
            throw new System.ArgumentException("PDF header format cannot be null.");
        }

        if (FirstPageHeaderFormat == null) {
            throw new System.ArgumentException("PDF first-page header format cannot be null.");
        }

        if (EvenPageHeaderFormat == null) {
            throw new System.ArgumentException("PDF even-page header format cannot be null.");
        }

        if (FooterFormat == null) {
            throw new System.ArgumentException("PDF footer format cannot be null.");
        }

        if (FirstPageFooterFormat == null) {
            throw new System.ArgumentException("PDF first-page footer format cannot be null.");
        }

        if (EvenPageFooterFormat == null) {
            throw new System.ArgumentException("PDF even-page footer format cannot be null.");
        }

        ValidatePageTextSegments(_headerSegments, "header");
        ValidatePageTextSegments(_firstPageHeaderSegments, "header");
        ValidatePageTextSegments(_evenPageHeaderSegments, "header");
        ValidateFooterSegments(_footerSegments);
        ValidateFooterSegments(_firstPageFooterSegments);
        ValidateFooterSegments(_evenPageFooterSegments);
        ValidateZoneString(_headerLeftFormat, "header");
        ValidateZoneString(_headerCenterFormat, "header");
        ValidateZoneString(_headerRightFormat, "header");
        ValidateZoneString(_firstPageHeaderLeftFormat, "header");
        ValidateZoneString(_firstPageHeaderCenterFormat, "header");
        ValidateZoneString(_firstPageHeaderRightFormat, "header");
        ValidateZoneString(_evenPageHeaderLeftFormat, "header");
        ValidateZoneString(_evenPageHeaderCenterFormat, "header");
        ValidateZoneString(_evenPageHeaderRightFormat, "header");
        ValidateZoneString(_footerLeftFormat, "footer");
        ValidateZoneString(_footerCenterFormat, "footer");
        ValidateZoneString(_footerRightFormat, "footer");
        ValidateZoneString(_firstPageFooterLeftFormat, "footer");
        ValidateZoneString(_firstPageFooterCenterFormat, "footer");
        ValidateZoneString(_firstPageFooterRightFormat, "footer");
        ValidateZoneString(_evenPageFooterLeftFormat, "footer");
        ValidateZoneString(_evenPageFooterCenterFormat, "footer");
        ValidateZoneString(_evenPageFooterRightFormat, "footer");
        ValidateOptionalLanguage(Language, nameof(Language));
    }

    private static void ValidateZones(string? left, string? center, string? right, string paramName) {
        if (left == null && center == null && right == null) {
            throw new System.ArgumentException("At least one PDF header/footer zone must contain text.", paramName);
        }

        ValidateZoneString(left, "header/footer");
        ValidateZoneString(center, "header/footer");
        ValidateZoneString(right, "header/footer");
    }

    private static void ValidateZoneString(string? value, string scope) {
        if (value == null) {
            return;
        }

        if (value.Length == 0) {
            throw new System.ArgumentException("PDF " + scope + " zone text cannot be empty.");
        }
    }

    private static void ValidateOptionalLanguage(string? value, string paramName) {
        if (value == null) {
            return;
        }

        if (string.IsNullOrWhiteSpace(value)) {
            throw new System.ArgumentException("PDF document language cannot be empty or whitespace.", paramName);
        }

        for (int i = 0; i < value.Length; i++) {
            if (char.IsControl(value[i])) {
                throw new System.ArgumentException("PDF document language cannot contain control characters.", paramName);
            }
        }
    }

    private static void ValidatePageTextSegments(System.Collections.Generic.List<FooterSegment>? segments, string scope) {
        if (segments != null) {
            for (int i = 0; i < segments.Count; i++) {
                var segment = segments[i];
                if (segment == null) {
                    throw new System.ArgumentException("PDF " + scope + " segments cannot contain null entries.");
                }

                if (segment.Kind == FooterSegmentKind.Text && segment.Text == null) {
                    throw new System.ArgumentException("PDF " + scope + " text segments cannot be null.");
                }

                if (segment.Kind != FooterSegmentKind.Text &&
                    segment.Kind != FooterSegmentKind.PageNumber &&
                    segment.Kind != FooterSegmentKind.TotalPages) {
                    throw new System.ArgumentException("PDF " + scope + " segments must use a supported segment kind.");
                }
            }
        }
    }

    private static void ValidateFooterSegments(System.Collections.Generic.List<FooterSegment>? segments) {
        if (segments != null) {
            for (int i = 0; i < segments.Count; i++) {
                var segment = segments[i];
                if (segment == null) {
                    throw new System.ArgumentException("PDF footer segments cannot contain null entries.");
                }

                if (segment.Kind == FooterSegmentKind.Text && segment.Text == null) {
                    throw new System.ArgumentException("PDF footer text segments cannot be null.");
                }

                if (segment.Kind != FooterSegmentKind.Text &&
                    segment.Kind != FooterSegmentKind.PageNumber &&
                    segment.Kind != FooterSegmentKind.TotalPages) {
                    throw new System.ArgumentException("PDF footer segments must use a supported segment kind.");
                }
            }
        }
    }
}


