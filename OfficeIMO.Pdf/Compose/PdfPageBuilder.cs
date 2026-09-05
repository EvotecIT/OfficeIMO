using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Page-level configuration (size, margins, default styles) and content/footers.
/// </summary>
public sealed class PdfPageBuilder {
    private readonly PdfDocument _doc;
    private readonly PdfOptions _options;
    internal PdfOptions Options => _options;
    internal PdfPageBuilder(PdfDocument doc, PdfOptions options) { _doc = doc; _options = options; }

    /// <summary>Sets page size using a predefined <see cref="PageSize"/>.</summary>
    public PdfPageBuilder Size(PageSize size) {
        Guard.Positive(size.Width, nameof(size));
        Guard.Positive(size.Height, nameof(size));
        Options.PageSize = size;
        return this;
    }
    /// <summary>Sets custom page size in points.</summary>
    public PdfPageBuilder Size(double width, double height) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Options.PageWidth = width;
        Options.PageHeight = height;
        return this;
    }
    /// <summary>Sets page orientation while preserving the current page size dimensions.</summary>
    public PdfPageBuilder Orientation(OfficePageOrientation orientation) {
        var oriented = new PageSize(Options.PageWidth, Options.PageHeight).WithOrientation(orientation);
        Options.PageWidth = oriented.Width;
        Options.PageHeight = oriented.Height;
        return this;
    }
    /// <summary>Sets or clears the page background color.</summary>
    public PdfPageBuilder Background(PdfColor? color) {
        Options.BackgroundColor = color;
        return this;
    }
    /// <summary>Sets or clears the page-scoped text watermark rendered behind page content.</summary>
    public PdfPageBuilder Watermark(PdfTextWatermark? watermark) {
        Options.TextWatermark = watermark;
        return this;
    }
    /// <summary>Sets a page-scoped text watermark rendered behind page content.</summary>
    public PdfPageBuilder Watermark(string text, double? fontSize = null, PdfColor? color = null, double? opacity = null, double? rotationAngle = null, PdfStandardFont? font = null, bool bold = true, bool italic = false) {
        var watermark = new PdfTextWatermark(text) {
            Bold = bold,
            Italic = italic
        };
        if (fontSize.HasValue) watermark.FontSize = fontSize.Value;
        if (color.HasValue) watermark.Color = color.Value;
        if (opacity.HasValue) watermark.Opacity = opacity.Value;
        if (rotationAngle.HasValue) watermark.RotationAngle = rotationAngle.Value;
        if (font.HasValue) watermark.Font = font.Value;
        Options.TextWatermark = watermark;
        return this;
    }
    /// <summary>Sets or clears the first-page text watermark rendered behind page content.</summary>
    public PdfPageBuilder FirstPageWatermark(PdfTextWatermark? watermark) {
        Options.FirstPageTextWatermark = watermark;
        return this;
    }
    /// <summary>Suppresses the inherited text watermark on the first page.</summary>
    public PdfPageBuilder SuppressFirstPageTextWatermark() {
        Options.SuppressFirstPageTextWatermark();
        return this;
    }
    /// <summary>Suppresses inherited text and image watermarks on the first page.</summary>
    public PdfPageBuilder SuppressFirstPageWatermark() {
        return SuppressFirstPageTextWatermark().SuppressFirstPageImageWatermark();
    }
    /// <summary>Sets a first-page text watermark rendered behind page content.</summary>
    public PdfPageBuilder FirstPageWatermark(string text, double? fontSize = null, PdfColor? color = null, double? opacity = null, double? rotationAngle = null, PdfStandardFont? font = null, bool bold = true, bool italic = false) {
        var watermark = new PdfTextWatermark(text) {
            Bold = bold,
            Italic = italic
        };
        if (fontSize.HasValue) watermark.FontSize = fontSize.Value;
        if (color.HasValue) watermark.Color = color.Value;
        if (opacity.HasValue) watermark.Opacity = opacity.Value;
        if (rotationAngle.HasValue) watermark.RotationAngle = rotationAngle.Value;
        if (font.HasValue) watermark.Font = font.Value;
        Options.FirstPageTextWatermark = watermark;
        return this;
    }
    /// <summary>Sets or clears the even-page text watermark rendered behind page content.</summary>
    public PdfPageBuilder EvenPagesWatermark(PdfTextWatermark? watermark) {
        Options.EvenPageTextWatermark = watermark;
        return this;
    }
    /// <summary>Suppresses the inherited text watermark on even pages.</summary>
    public PdfPageBuilder SuppressEvenPagesTextWatermark() {
        Options.SuppressEvenPageTextWatermark();
        return this;
    }
    /// <summary>Suppresses inherited text and image watermarks on even pages.</summary>
    public PdfPageBuilder SuppressEvenPagesWatermark() {
        return SuppressEvenPagesTextWatermark().SuppressEvenPagesImageWatermark();
    }
    /// <summary>Sets an even-page text watermark rendered behind page content.</summary>
    public PdfPageBuilder EvenPagesWatermark(string text, double? fontSize = null, PdfColor? color = null, double? opacity = null, double? rotationAngle = null, PdfStandardFont? font = null, bool bold = true, bool italic = false) {
        var watermark = new PdfTextWatermark(text) {
            Bold = bold,
            Italic = italic
        };
        if (fontSize.HasValue) watermark.FontSize = fontSize.Value;
        if (color.HasValue) watermark.Color = color.Value;
        if (opacity.HasValue) watermark.Opacity = opacity.Value;
        if (rotationAngle.HasValue) watermark.RotationAngle = rotationAngle.Value;
        if (font.HasValue) watermark.Font = font.Value;
        Options.EvenPageTextWatermark = watermark;
        return this;
    }
    /// <summary>Sets or clears the page-scoped image watermark rendered behind page content.</summary>
    public PdfPageBuilder ImageWatermark(PdfImageWatermark? watermark) {
        Options.ImageWatermark = watermark;
        return this;
    }
    /// <summary>Sets a page-scoped image watermark rendered behind page content.</summary>
    public PdfPageBuilder ImageWatermark(byte[] imageBytes, double width, double height, double? opacity = null, double? rotationAngle = null) {
        var watermark = new PdfImageWatermark(imageBytes, width, height);
        if (opacity.HasValue) watermark.Opacity = opacity.Value;
        if (rotationAngle.HasValue) watermark.RotationAngle = rotationAngle.Value;
        Options.ImageWatermark = watermark;
        return this;
    }
    /// <summary>Sets or clears the first-page image watermark rendered behind page content.</summary>
    public PdfPageBuilder FirstPageImageWatermark(PdfImageWatermark? watermark) {
        Options.FirstPageImageWatermark = watermark;
        return this;
    }
    /// <summary>Suppresses the inherited image watermark on the first page.</summary>
    public PdfPageBuilder SuppressFirstPageImageWatermark() {
        Options.SuppressFirstPageImageWatermark();
        return this;
    }
    /// <summary>Sets a first-page image watermark rendered behind page content.</summary>
    public PdfPageBuilder FirstPageImageWatermark(byte[] imageBytes, double width, double height, double? opacity = null, double? rotationAngle = null) {
        var watermark = new PdfImageWatermark(imageBytes, width, height);
        if (opacity.HasValue) watermark.Opacity = opacity.Value;
        if (rotationAngle.HasValue) watermark.RotationAngle = rotationAngle.Value;
        Options.FirstPageImageWatermark = watermark;
        return this;
    }
    /// <summary>Sets or clears the even-page image watermark rendered behind page content.</summary>
    public PdfPageBuilder EvenPagesImageWatermark(PdfImageWatermark? watermark) {
        Options.EvenPageImageWatermark = watermark;
        return this;
    }
    /// <summary>Suppresses the inherited image watermark on even pages.</summary>
    public PdfPageBuilder SuppressEvenPagesImageWatermark() {
        Options.SuppressEvenPageImageWatermark();
        return this;
    }
    /// <summary>Sets an even-page image watermark rendered behind page content.</summary>
    public PdfPageBuilder EvenPagesImageWatermark(byte[] imageBytes, double width, double height, double? opacity = null, double? rotationAngle = null) {
        var watermark = new PdfImageWatermark(imageBytes, width, height);
        if (opacity.HasValue) watermark.Opacity = opacity.Value;
        if (rotationAngle.HasValue) watermark.RotationAngle = rotationAngle.Value;
        Options.EvenPageImageWatermark = watermark;
        return this;
    }
    /// <summary>Sets or clears the page-scoped page border.</summary>
    public PdfPageBuilder PageBorder(PdfPageBorder? border) {
        Options.PageBorder = border;
        return this;
    }
    /// <summary>Sets a page-scoped page border.</summary>
    public PdfPageBuilder PageBorder(PdfColor? color = null, double? width = null, double? inset = null, double? opacity = null, OfficeIMO.Drawing.OfficeStrokeDashStyle dashStyle = OfficeIMO.Drawing.OfficeStrokeDashStyle.Solid) {
        var border = new PdfPageBorder {
            DashStyle = dashStyle
        };
        if (color.HasValue) border.Color = color.Value;
        if (width.HasValue) border.Width = width.Value;
        if (inset.HasValue) border.Inset = inset.Value;
        if (opacity.HasValue) border.Opacity = opacity.Value;
        Options.PageBorder = border;
        return this;
    }
    /// <summary>Sets or clears the page-scoped background image.</summary>
    public PdfPageBuilder BackgroundImage(PdfPageBackgroundImage? image) {
        Options.PageBackgroundImage = image;
        return this;
    }
    /// <summary>Sets a page-scoped background image.</summary>
    public PdfPageBuilder BackgroundImage(byte[] imageBytes, OfficeIMO.Drawing.OfficeImageFit fit = OfficeIMO.Drawing.OfficeImageFit.Cover, double? opacity = null) {
        var image = new PdfPageBackgroundImage(imageBytes) {
            Fit = fit
        };
        if (opacity.HasValue) image.Opacity = opacity.Value;
        Options.PageBackgroundImage = image;
        return this;
    }
    /// <summary>Adds a page-scoped background shape rendered behind page content.</summary>
    public PdfPageBuilder BackgroundShape(PdfPageBackgroundShape shape) {
        Options.AddPageBackgroundShape(shape);
        return this;
    }
    /// <summary>Replaces or clears page-scoped background shapes.</summary>
    public PdfPageBuilder BackgroundShapes(System.Collections.Generic.IEnumerable<PdfPageBackgroundShape>? shapes) {
        Options.PageBackgroundShapes = shapes?.ToList();
        return this;
    }
    /// <summary>Clears page-scoped background shapes.</summary>
    public PdfPageBuilder ClearBackgroundShapes() {
        Options.ClearPageBackgroundShapes();
        return this;
    }
    /// <summary>Adds a page-scoped rectangle background shape.</summary>
    public PdfPageBuilder BackgroundRectangle(double x, double y, double width, double height, PdfColor? fill = null, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeIMO.Drawing.OfficeLinearGradient? fillGradient = null) =>
        BackgroundShape(PdfPageBackgroundShape.Rectangle(x, y, width, height, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient));
    /// <summary>Adds a page-scoped rounded rectangle background shape.</summary>
    public PdfPageBuilder BackgroundRoundedRectangle(double x, double y, double width, double height, double cornerRadius, PdfColor? fill = null, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeIMO.Drawing.OfficeLinearGradient? fillGradient = null) =>
        BackgroundShape(PdfPageBackgroundShape.RoundedRectangle(x, y, width, height, cornerRadius, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient));
    /// <summary>Adds a page-scoped ellipse background shape.</summary>
    public PdfPageBuilder BackgroundEllipse(double x, double y, double width, double height, PdfColor? fill = null, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeIMO.Drawing.OfficeLinearGradient? fillGradient = null) =>
        BackgroundShape(PdfPageBackgroundShape.Ellipse(x, y, width, height, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient));
    /// <summary>Adds a page-scoped top background band using the current page size.</summary>
    public PdfPageBuilder BackgroundTopBand(double height, PdfColor? fill = null, double insetX = 0D, double offsetY = 0D, double cornerRadius = 0D, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeIMO.Drawing.OfficeLinearGradient? fillGradient = null) =>
        BackgroundShape(PdfPageBackgroundShape.TopBand(Options.PageWidth, Options.PageHeight, height, fill, insetX, offsetY, cornerRadius, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient));
    /// <summary>Adds a page-scoped bottom background band using the current page size.</summary>
    public PdfPageBuilder BackgroundBottomBand(double height, PdfColor? fill = null, double insetX = 0D, double offsetY = 0D, double cornerRadius = 0D, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeIMO.Drawing.OfficeLinearGradient? fillGradient = null) =>
        BackgroundShape(PdfPageBackgroundShape.BottomBand(Options.PageWidth, Options.PageHeight, height, fill, insetX, offsetY, cornerRadius, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient));
    /// <summary>Adds a page-scoped left background band using the current page size.</summary>
    public PdfPageBuilder BackgroundLeftBand(double width, PdfColor? fill = null, double insetY = 0D, double offsetX = 0D, double cornerRadius = 0D, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeIMO.Drawing.OfficeLinearGradient? fillGradient = null) =>
        BackgroundShape(PdfPageBackgroundShape.LeftBand(Options.PageWidth, Options.PageHeight, width, fill, insetY, offsetX, cornerRadius, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient));
    /// <summary>Adds a page-scoped right background band using the current page size.</summary>
    public PdfPageBuilder BackgroundRightBand(double width, PdfColor? fill = null, double insetY = 0D, double offsetX = 0D, double cornerRadius = 0D, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeIMO.Drawing.OfficeLinearGradient? fillGradient = null) =>
        BackgroundShape(PdfPageBackgroundShape.RightBand(Options.PageWidth, Options.PageHeight, width, fill, insetY, offsetX, cornerRadius, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient));
    /// <summary>Sets page orientation to portrait while preserving the current page size dimensions.</summary>
    public PdfPageBuilder Portrait() => Orientation(OfficePageOrientation.Portrait);
    /// <summary>Sets page orientation to landscape while preserving the current page size dimensions.</summary>
    public PdfPageBuilder Landscape() => Orientation(OfficePageOrientation.Landscape);
    /// <summary>Sets uniform page margins (all sides in points).</summary>
    public PdfPageBuilder Margin(double all) {
        Guard.NonNegative(all, nameof(all));
        Options.MarginLeft = Options.MarginRight = Options.MarginTop = Options.MarginBottom = all;
        return this;
    }
    /// <summary>Sets page margins from a reusable margin value.</summary>
    public PdfPageBuilder Margin(PageMargins margins) {
        Options.Margins = margins;
        return this;
    }
    /// <summary>Sets page margins (left, top, right, bottom in points).</summary>
    public PdfPageBuilder Margin(double left, double top, double right, double bottom) {
        Guard.NonNegative(left, nameof(left));
        Guard.NonNegative(top, nameof(top));
        Guard.NonNegative(right, nameof(right));
        Guard.NonNegative(bottom, nameof(bottom));
        Options.MarginLeft = left;
        Options.MarginTop = top;
        Options.MarginRight = right;
        Options.MarginBottom = bottom;
        return this;
    }

    /// <summary>Sets or clears page-scoped TrimBox and BleedBox geometry.</summary>
    public PdfPageBuilder PrintProductionPageBoxes(PdfPrintProductionPageBoxes? boxes) {
        Options.PrintProductionPageBoxes = boxes;
        return this;
    }

    /// <summary>Sets the first visible page number for this page or section flow.</summary>
    public PdfPageBuilder PageNumberStart(int start) {
        Options.PageNumberStart = start;
        return this;
    }

    /// <summary>Sets the visible page-number style for this page or section flow.</summary>
    public PdfPageBuilder PageNumberStyle(PdfPageNumberStyle style) {
        Options.PageNumberStyle = style;
        return this;
    }

    /// <summary>Applies reusable page-scoped default styles.</summary>
    public PdfPageBuilder Theme(PdfTheme theme) { Guard.NotNull(theme, nameof(theme)); theme.Clone().ApplyTo(Options); return this; }
    /// <summary>Applies a shared, page-scoped font, fallback, language, and shaping profile.</summary>
    public PdfPageBuilder Typography(OfficeIMO.Drawing.OfficeRenderingProfile profile, OfficeIMO.Drawing.OfficeRenderingProfileApplyMode mode = OfficeIMO.Drawing.OfficeRenderingProfileApplyMode.Replace) { Options.UseRenderingProfile(profile, mode); return this; }
    /// <summary>Uses a caller-supplied TrueType font family for this composed page or section.</summary>
    public PdfPageBuilder UseFontFamily(PdfEmbeddedFontFamily fontFamily) { Options.UseFontFamily(fontFamily); return this; }
    /// <summary>Registers a planned embedded-font fallback set for generated rich text runs on this composed page or section.</summary>
    public PdfPageBuilder RegisterEmbeddedFontFallbacks(PdfEmbeddedFontFallbackSet fallbackSet) { Options.RegisterEmbeddedFontFallbacks(fallbackSet); return this; }
    /// <summary>Applies OfficeIMO's built-in generated-text fallback groups for this composed page or section.</summary>
    public PdfPageBuilder UseTextFallbacks(PdfTextFallbackFeatures features = PdfTextFallbackFeatures.Default) { Options.UseTextFallbacks(features); return this; }
    /// <summary>Registers generated-text fallback fonts from installed system font families without requiring callers to choose PDF font slots.</summary>
    public PdfPageBuilder UseEmbeddedFontFallbacksFromSystem(string? familyNames, int maxFallbackFonts = 2) { Options.UseEmbeddedFontFallbacksFromSystem(familyNames, maxFallbackFonts); return this; }
    /// <summary>Uses caller-supplied TrueType font files for this composed page or section.</summary>
    public PdfPageBuilder UseFontFamily(string familyName, byte[] regular, byte[]? bold = null, byte[]? italic = null, byte[]? boldItalic = null) { Options.UseFontFamily(familyName, regular, bold, italic, boldItalic); return this; }
    /// <summary>Uses caller-supplied TrueType font files for this composed page or section.</summary>
    public PdfPageBuilder UseFontFamily(string familyName, string regularPath, string? boldPath = null, string? italicPath = null, string? boldItalicPath = null) { Options.UseFontFamily(familyName, regularPath, boldPath, italicPath, boldItalicPath); return this; }
    /// <summary>Sets or clears the page-scoped generated text line-break callback used for long unspaced tokens.</summary>
    public PdfPageBuilder TextLineBreaks(Func<string, IReadOnlyList<int>>? callback) { Options.SetTextLineBreaks(callback); return this; }
    /// <summary>Sets or clears the page-scoped generated text hyphenation callback used for long unspaced tokens.</summary>
    public PdfPageBuilder TextHyphenation(PdfTextHyphenationCallback? callback) { Options.SetTextHyphenation(callback); return this; }
    /// <summary>Uses or clears an immutable first-party word hyphenation dictionary.</summary>
    public PdfPageBuilder TextHyphenationDictionary(PdfHyphenationLexicon? dictionary) { Options.UseTextHyphenationDictionary(dictionary); return this; }
    /// <summary>Configures default text style for the page.</summary>
    public PdfPageBuilder DefaultTextStyle(System.Action<PdfTextStyleBuilder> style) { Guard.NotNull(style, nameof(style)); var s = new PdfTextStyleBuilder(Options); style(s); return this; }
    /// <summary>Configures default text style for the page from a reusable text style object.</summary>
    public PdfPageBuilder DefaultTextStyle(PdfTextStyle style) { Guard.NotNull(style, nameof(style)); style.Clone().ApplyTo(Options); return this; }
    /// <summary>Configures the default paragraph style for page paragraphs that do not provide an explicit style.</summary>
    public PdfPageBuilder DefaultParagraphStyle(PdfParagraphStyle style) { Guard.NotNull(style, nameof(style)); Options.DefaultParagraphStyle = style; return this; }
    /// <summary>Configures the default table style for page tables that do not provide an explicit style.</summary>
    public PdfPageBuilder DefaultTableStyle(PdfTableStyle style) { Guard.NotNull(style, nameof(style)); Options.DefaultTableStyle = style; return this; }
    /// <summary>Configures the default table style for the page from a supported Word table style name.</summary>
    public PdfPageBuilder DefaultTableStyle(string wordTableStyleName) { Options.DefaultTableStyle = TableStyles.FromWordTableStyle(wordTableStyleName); return this; }
    /// <summary>Configures the default style for a built-in heading level on the page.</summary>
    public PdfPageBuilder DefaultHeadingStyle(int level, PdfHeadingStyle style) { Guard.NotNull(style, nameof(style)); Options.SetDefaultHeadingStyle(level, style); return this; }
    /// <summary>Configures the default style for page bullet and numbered lists.</summary>
    public PdfPageBuilder DefaultListStyle(PdfListStyle style) { Guard.NotNull(style, nameof(style)); Options.DefaultListStyle = style; return this; }
    /// <summary>Configures the default style for page panel paragraphs.</summary>
    public PdfPageBuilder DefaultPanelStyle(PdfPanelStyle style) { Guard.NotNull(style, nameof(style)); Options.DefaultPanelStyle = style; return this; }
    /// <summary>Configures the default style for page horizontal rules.</summary>
    public PdfPageBuilder DefaultHorizontalRuleStyle(PdfHorizontalRuleStyle style) { Guard.NotNull(style, nameof(style)); Options.DefaultHorizontalRuleStyle = style; return this; }
    /// <summary>Configures the default style for page images.</summary>
    public PdfPageBuilder DefaultImageStyle(PdfImageStyle style) { Guard.NotNull(style, nameof(style)); Options.DefaultImageStyle = style; return this; }
    /// <summary>Configures the default placement style for page drawing objects.</summary>
    public PdfPageBuilder DefaultDrawingStyle(PdfDrawingStyle style) { Guard.NotNull(style, nameof(style)); Options.DefaultDrawingStyle = style; return this; }
    /// <summary>Configures the default row/column layout style for the page.</summary>
    public PdfPageBuilder DefaultRowStyle(PdfRowStyle style) { Guard.NotNull(style, nameof(style)); Options.DefaultRowStyle = style; return this; }
    /// <summary>Builds the page content through the universal content receiver.</summary>
    public PdfPageBuilder Content(System.Action<PdfContentBuilder> build) {
        Guard.NotNull(build, nameof(build));
        build(new PdfContentBuilder(_doc));
        return this;
    }
    /// <summary>Adds foreground page content at absolute top-left page coordinates.</summary>
    public PdfPageBuilder Canvas(System.Action<PdfPageCanvas> build) { _doc.Canvas(build); return this; }
    /// <summary>Defines the header layout and content.</summary>
    public PdfPageBuilder Header(System.Action<PdfHeaderBuilder> build) { Guard.NotNull(build, nameof(build)); var h = new PdfHeaderBuilder(Options); build(h); return this; }
    /// <summary>Defines the footer layout and content.</summary>
    public PdfPageBuilder Footer(System.Action<PdfFooterBuilder> build) { Guard.NotNull(build, nameof(build)); var f = new PdfFooterBuilder(Options); build(f); return this; }
}
