using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Page-level configuration (size, margins, default styles) and content/footers.
/// </summary>
public class PdfPageCompose {
    private readonly PdfDocument _doc;
    private readonly PdfOptions _options;
    internal PdfOptions Options => _options;
    internal PdfPageCompose(PdfDocument doc, PdfOptions options) { _doc = doc; _options = options; }

    /// <summary>Sets page size using a predefined <see cref="PageSize"/>.</summary>
    public PdfPageCompose Size(PageSize size) {
        Guard.Positive(size.Width, nameof(size));
        Guard.Positive(size.Height, nameof(size));
        Options.PageSize = size;
        return this;
    }
    /// <summary>Sets custom page size in points.</summary>
    public PdfPageCompose Size(double width, double height) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Options.PageWidth = width;
        Options.PageHeight = height;
        return this;
    }
    /// <summary>Sets page orientation while preserving the current page size dimensions.</summary>
    public PdfPageCompose Orientation(PdfPageOrientation orientation) {
        var oriented = new PageSize(Options.PageWidth, Options.PageHeight).WithOrientation(orientation);
        Options.PageWidth = oriented.Width;
        Options.PageHeight = oriented.Height;
        return this;
    }
    /// <summary>Sets or clears the page background color.</summary>
    public PdfPageCompose Background(PdfColor? color) {
        Options.BackgroundColor = color;
        return this;
    }
    /// <summary>Sets or clears the page-scoped text watermark rendered behind page content.</summary>
    public PdfPageCompose Watermark(PdfTextWatermark? watermark) {
        Options.TextWatermark = watermark;
        return this;
    }
    /// <summary>Sets a page-scoped text watermark rendered behind page content.</summary>
    public PdfPageCompose Watermark(string text, double? fontSize = null, PdfColor? color = null, double? opacity = null, double? rotationAngle = null, PdfStandardFont? font = null, bool bold = true, bool italic = false) {
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
    public PdfPageCompose FirstPageWatermark(PdfTextWatermark? watermark) {
        Options.FirstPageTextWatermark = watermark;
        return this;
    }
    /// <summary>Suppresses the inherited text watermark on the first page.</summary>
    public PdfPageCompose SuppressFirstPageTextWatermark() {
        Options.SuppressFirstPageTextWatermark();
        return this;
    }
    /// <summary>Suppresses inherited text and image watermarks on the first page.</summary>
    public PdfPageCompose SuppressFirstPageWatermark() {
        return SuppressFirstPageTextWatermark().SuppressFirstPageImageWatermark();
    }
    /// <summary>Sets a first-page text watermark rendered behind page content.</summary>
    public PdfPageCompose FirstPageWatermark(string text, double? fontSize = null, PdfColor? color = null, double? opacity = null, double? rotationAngle = null, PdfStandardFont? font = null, bool bold = true, bool italic = false) {
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
    public PdfPageCompose EvenPagesWatermark(PdfTextWatermark? watermark) {
        Options.EvenPageTextWatermark = watermark;
        return this;
    }
    /// <summary>Suppresses the inherited text watermark on even pages.</summary>
    public PdfPageCompose SuppressEvenPagesTextWatermark() {
        Options.SuppressEvenPageTextWatermark();
        return this;
    }
    /// <summary>Suppresses inherited text and image watermarks on even pages.</summary>
    public PdfPageCompose SuppressEvenPagesWatermark() {
        return SuppressEvenPagesTextWatermark().SuppressEvenPagesImageWatermark();
    }
    /// <summary>Sets an even-page text watermark rendered behind page content.</summary>
    public PdfPageCompose EvenPagesWatermark(string text, double? fontSize = null, PdfColor? color = null, double? opacity = null, double? rotationAngle = null, PdfStandardFont? font = null, bool bold = true, bool italic = false) {
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
    public PdfPageCompose ImageWatermark(PdfImageWatermark? watermark) {
        Options.ImageWatermark = watermark;
        return this;
    }
    /// <summary>Sets a page-scoped image watermark rendered behind page content.</summary>
    public PdfPageCompose ImageWatermark(byte[] imageBytes, double width, double height, double? opacity = null, double? rotationAngle = null) {
        var watermark = new PdfImageWatermark(imageBytes, width, height);
        if (opacity.HasValue) watermark.Opacity = opacity.Value;
        if (rotationAngle.HasValue) watermark.RotationAngle = rotationAngle.Value;
        Options.ImageWatermark = watermark;
        return this;
    }
    /// <summary>Sets or clears the first-page image watermark rendered behind page content.</summary>
    public PdfPageCompose FirstPageImageWatermark(PdfImageWatermark? watermark) {
        Options.FirstPageImageWatermark = watermark;
        return this;
    }
    /// <summary>Suppresses the inherited image watermark on the first page.</summary>
    public PdfPageCompose SuppressFirstPageImageWatermark() {
        Options.SuppressFirstPageImageWatermark();
        return this;
    }
    /// <summary>Sets a first-page image watermark rendered behind page content.</summary>
    public PdfPageCompose FirstPageImageWatermark(byte[] imageBytes, double width, double height, double? opacity = null, double? rotationAngle = null) {
        var watermark = new PdfImageWatermark(imageBytes, width, height);
        if (opacity.HasValue) watermark.Opacity = opacity.Value;
        if (rotationAngle.HasValue) watermark.RotationAngle = rotationAngle.Value;
        Options.FirstPageImageWatermark = watermark;
        return this;
    }
    /// <summary>Sets or clears the even-page image watermark rendered behind page content.</summary>
    public PdfPageCompose EvenPagesImageWatermark(PdfImageWatermark? watermark) {
        Options.EvenPageImageWatermark = watermark;
        return this;
    }
    /// <summary>Suppresses the inherited image watermark on even pages.</summary>
    public PdfPageCompose SuppressEvenPagesImageWatermark() {
        Options.SuppressEvenPageImageWatermark();
        return this;
    }
    /// <summary>Sets an even-page image watermark rendered behind page content.</summary>
    public PdfPageCompose EvenPagesImageWatermark(byte[] imageBytes, double width, double height, double? opacity = null, double? rotationAngle = null) {
        var watermark = new PdfImageWatermark(imageBytes, width, height);
        if (opacity.HasValue) watermark.Opacity = opacity.Value;
        if (rotationAngle.HasValue) watermark.RotationAngle = rotationAngle.Value;
        Options.EvenPageImageWatermark = watermark;
        return this;
    }
    /// <summary>Sets or clears the page-scoped page border.</summary>
    public PdfPageCompose PageBorder(PdfPageBorder? border) {
        Options.PageBorder = border;
        return this;
    }
    /// <summary>Sets a page-scoped page border.</summary>
    public PdfPageCompose PageBorder(PdfColor? color = null, double? width = null, double? inset = null, double? opacity = null, OfficeIMO.Drawing.OfficeStrokeDashStyle dashStyle = OfficeIMO.Drawing.OfficeStrokeDashStyle.Solid) {
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
    public PdfPageCompose BackgroundImage(PdfPageBackgroundImage? image) {
        Options.PageBackgroundImage = image;
        return this;
    }
    /// <summary>Sets a page-scoped background image.</summary>
    public PdfPageCompose BackgroundImage(byte[] imageBytes, OfficeIMO.Drawing.OfficeImageFit fit = OfficeIMO.Drawing.OfficeImageFit.Cover, double? opacity = null) {
        var image = new PdfPageBackgroundImage(imageBytes) {
            Fit = fit
        };
        if (opacity.HasValue) image.Opacity = opacity.Value;
        Options.PageBackgroundImage = image;
        return this;
    }
    /// <summary>Adds a page-scoped background shape rendered behind page content.</summary>
    public PdfPageCompose BackgroundShape(PdfPageBackgroundShape shape) {
        Options.AddPageBackgroundShape(shape);
        return this;
    }
    /// <summary>Replaces or clears page-scoped background shapes.</summary>
    public PdfPageCompose BackgroundShapes(System.Collections.Generic.IEnumerable<PdfPageBackgroundShape>? shapes) {
        Options.PageBackgroundShapes = shapes?.ToList();
        return this;
    }
    /// <summary>Clears page-scoped background shapes.</summary>
    public PdfPageCompose ClearBackgroundShapes() {
        Options.ClearPageBackgroundShapes();
        return this;
    }
    /// <summary>Adds a page-scoped rectangle background shape.</summary>
    public PdfPageCompose BackgroundRectangle(double x, double y, double width, double height, PdfColor? fill = null, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeIMO.Drawing.OfficeLinearGradient? fillGradient = null) =>
        BackgroundShape(PdfPageBackgroundShape.Rectangle(x, y, width, height, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient));
    /// <summary>Adds a page-scoped rounded rectangle background shape.</summary>
    public PdfPageCompose BackgroundRoundedRectangle(double x, double y, double width, double height, double cornerRadius, PdfColor? fill = null, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeIMO.Drawing.OfficeLinearGradient? fillGradient = null) =>
        BackgroundShape(PdfPageBackgroundShape.RoundedRectangle(x, y, width, height, cornerRadius, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient));
    /// <summary>Adds a page-scoped ellipse background shape.</summary>
    public PdfPageCompose BackgroundEllipse(double x, double y, double width, double height, PdfColor? fill = null, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeIMO.Drawing.OfficeLinearGradient? fillGradient = null) =>
        BackgroundShape(PdfPageBackgroundShape.Ellipse(x, y, width, height, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient));
    /// <summary>Adds a page-scoped top background band using the current page size.</summary>
    public PdfPageCompose BackgroundTopBand(double height, PdfColor? fill = null, double insetX = 0D, double offsetY = 0D, double cornerRadius = 0D, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeIMO.Drawing.OfficeLinearGradient? fillGradient = null) =>
        BackgroundShape(PdfPageBackgroundShape.TopBand(Options.PageWidth, Options.PageHeight, height, fill, insetX, offsetY, cornerRadius, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient));
    /// <summary>Adds a page-scoped bottom background band using the current page size.</summary>
    public PdfPageCompose BackgroundBottomBand(double height, PdfColor? fill = null, double insetX = 0D, double offsetY = 0D, double cornerRadius = 0D, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeIMO.Drawing.OfficeLinearGradient? fillGradient = null) =>
        BackgroundShape(PdfPageBackgroundShape.BottomBand(Options.PageWidth, Options.PageHeight, height, fill, insetX, offsetY, cornerRadius, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient));
    /// <summary>Adds a page-scoped left background band using the current page size.</summary>
    public PdfPageCompose BackgroundLeftBand(double width, PdfColor? fill = null, double insetY = 0D, double offsetX = 0D, double cornerRadius = 0D, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeIMO.Drawing.OfficeLinearGradient? fillGradient = null) =>
        BackgroundShape(PdfPageBackgroundShape.LeftBand(Options.PageWidth, Options.PageHeight, width, fill, insetY, offsetX, cornerRadius, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient));
    /// <summary>Adds a page-scoped right background band using the current page size.</summary>
    public PdfPageCompose BackgroundRightBand(double width, PdfColor? fill = null, double insetY = 0D, double offsetX = 0D, double cornerRadius = 0D, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeIMO.Drawing.OfficeLinearGradient? fillGradient = null) =>
        BackgroundShape(PdfPageBackgroundShape.RightBand(Options.PageWidth, Options.PageHeight, width, fill, insetY, offsetX, cornerRadius, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient));
    /// <summary>Sets page orientation to portrait while preserving the current page size dimensions.</summary>
    public PdfPageCompose Portrait() => Orientation(PdfPageOrientation.Portrait);
    /// <summary>Sets page orientation to landscape while preserving the current page size dimensions.</summary>
    public PdfPageCompose Landscape() => Orientation(PdfPageOrientation.Landscape);
    /// <summary>Sets uniform page margins (all sides in points).</summary>
    public PdfPageCompose Margin(double all) {
        Guard.NonNegative(all, nameof(all));
        Options.MarginLeft = Options.MarginRight = Options.MarginTop = Options.MarginBottom = all;
        return this;
    }
    /// <summary>Sets page margins from a reusable margin value.</summary>
    public PdfPageCompose Margin(PageMargins margins) {
        Options.Margins = margins;
        return this;
    }
    /// <summary>Sets page margins (left, top, right, bottom in points).</summary>
    public PdfPageCompose Margin(double left, double top, double right, double bottom) {
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

    /// <summary>Sets the first visible page number for this page or section flow.</summary>
    public PdfPageCompose PageNumberStart(int start) {
        Options.PageNumberStart = start;
        return this;
    }

    /// <summary>Sets the visible page-number style for this page or section flow.</summary>
    public PdfPageCompose PageNumberStyle(PdfPageNumberStyle style) {
        Options.PageNumberStyle = style;
        return this;
    }

    /// <summary>Applies reusable page-scoped default styles.</summary>
    public PdfPageCompose Theme(PdfTheme theme) { Guard.NotNull(theme, nameof(theme)); theme.Clone().ApplyTo(Options); return this; }
    /// <summary>Uses a caller-supplied TrueType font family for this composed page or section.</summary>
    public PdfPageCompose UseFontFamily(PdfEmbeddedFontFamily fontFamily) { Options.UseFontFamily(fontFamily); return this; }
    /// <summary>Registers a planned embedded-font fallback set for generated rich text runs on this composed page or section.</summary>
    public PdfPageCompose RegisterEmbeddedFontFallbacks(PdfEmbeddedFontFallbackSet fallbackSet) { Options.RegisterEmbeddedFontFallbacks(fallbackSet); return this; }
    /// <summary>Applies OfficeIMO's built-in generated-text fallback groups for this composed page or section.</summary>
    public PdfPageCompose UseTextFallbacks(PdfTextFallbackFeatures features = PdfTextFallbackFeatures.Default) { Options.UseTextFallbacks(features); return this; }
    /// <summary>Registers generated-text fallback fonts from installed system font families without requiring callers to choose PDF font slots.</summary>
    public PdfPageCompose UseEmbeddedFontFallbacksFromSystem(string? familyNames, int maxFallbackFonts = 2) { Options.UseEmbeddedFontFallbacksFromSystem(familyNames, maxFallbackFonts); return this; }
    /// <summary>Uses caller-supplied TrueType font files for this composed page or section.</summary>
    public PdfPageCompose UseFontFamily(string familyName, byte[] regular, byte[]? bold = null, byte[]? italic = null, byte[]? boldItalic = null) { Options.UseFontFamily(familyName, regular, bold, italic, boldItalic); return this; }
    /// <summary>Uses caller-supplied TrueType font files for this composed page or section.</summary>
    public PdfPageCompose UseFontFamily(string familyName, string regularPath, string? boldPath = null, string? italicPath = null, string? boldItalicPath = null) { Options.UseFontFamily(familyName, regularPath, boldPath, italicPath, boldItalicPath); return this; }
    /// <summary>Sets or clears the page-scoped generated text line-break callback used for long unspaced tokens.</summary>
    public PdfPageCompose TextLineBreaks(Func<string, IReadOnlyList<int>>? callback) { Options.SetTextLineBreaks(callback); return this; }
    /// <summary>Sets or clears the page-scoped generated text hyphenation callback used for long unspaced tokens.</summary>
    public PdfPageCompose TextHyphenation(PdfTextHyphenationCallback? callback) { Options.SetTextHyphenation(callback); return this; }
    /// <summary>Uses or clears an immutable first-party word hyphenation dictionary.</summary>
    public PdfPageCompose TextHyphenationDictionary(PdfHyphenationLexicon? dictionary) { Options.UseTextHyphenationDictionary(dictionary); return this; }
    /// <summary>Configures default text style for the page.</summary>
    public PdfPageCompose DefaultTextStyle(System.Action<PdfTextStyleCompose> style) { Guard.NotNull(style, nameof(style)); var s = new PdfTextStyleCompose(Options); style(s); return this; }
    /// <summary>Configures default text style for the page from a reusable text style object.</summary>
    public PdfPageCompose DefaultTextStyle(PdfTextStyle style) { Guard.NotNull(style, nameof(style)); style.Clone().ApplyTo(Options); return this; }
    /// <summary>Configures the default paragraph style for page paragraphs that do not provide an explicit style.</summary>
    public PdfPageCompose DefaultParagraphStyle(PdfParagraphStyle style) { Guard.NotNull(style, nameof(style)); Options.DefaultParagraphStyle = style; return this; }
    /// <summary>Configures the default table style for page tables that do not provide an explicit style.</summary>
    public PdfPageCompose DefaultTableStyle(PdfTableStyle style) { Guard.NotNull(style, nameof(style)); Options.DefaultTableStyle = style; return this; }
    /// <summary>Configures the default table style for the page from a supported Word table style name.</summary>
    public PdfPageCompose DefaultTableStyle(string wordTableStyleName) { Options.DefaultTableStyle = TableStyles.FromWordTableStyle(wordTableStyleName); return this; }
    /// <summary>Configures the default style for a built-in heading level on the page.</summary>
    public PdfPageCompose DefaultHeadingStyle(int level, PdfHeadingStyle style) { Guard.NotNull(style, nameof(style)); Options.SetDefaultHeadingStyle(level, style); return this; }
    /// <summary>Configures the default style for page bullet and numbered lists.</summary>
    public PdfPageCompose DefaultListStyle(PdfListStyle style) { Guard.NotNull(style, nameof(style)); Options.DefaultListStyle = style; return this; }
    /// <summary>Configures the default style for page panel paragraphs.</summary>
    public PdfPageCompose DefaultPanelStyle(PanelStyle style) { Guard.NotNull(style, nameof(style)); Options.DefaultPanelStyle = style; return this; }
    /// <summary>Configures the default style for page horizontal rules.</summary>
    public PdfPageCompose DefaultHorizontalRuleStyle(PdfHorizontalRuleStyle style) { Guard.NotNull(style, nameof(style)); Options.DefaultHorizontalRuleStyle = style; return this; }
    /// <summary>Configures the default style for page images.</summary>
    public PdfPageCompose DefaultImageStyle(PdfImageStyle style) { Guard.NotNull(style, nameof(style)); Options.DefaultImageStyle = style; return this; }
    /// <summary>Configures the default placement style for page drawing objects.</summary>
    public PdfPageCompose DefaultDrawingStyle(PdfDrawingStyle style) { Guard.NotNull(style, nameof(style)); Options.DefaultDrawingStyle = style; return this; }
    /// <summary>Configures the default row/column layout style for the page.</summary>
    public PdfPageCompose DefaultRowStyle(PdfRowStyle style) { Guard.NotNull(style, nameof(style)); Options.DefaultRowStyle = style; return this; }
    /// <summary>Builds the page content using a column/row flow.</summary>
    public PdfPageCompose Content(System.Action<PdfContentCompose> build) { Guard.NotNull(build, nameof(build)); var c = new PdfContentCompose(_doc); build(c); return this; }
    /// <summary>Adds foreground page content at absolute top-left page coordinates.</summary>
    public PdfPageCompose Canvas(System.Action<PdfPageCanvas> build) { _doc.Canvas(build); return this; }
    /// <summary>Defines the header layout and content.</summary>
    public PdfPageCompose Header(System.Action<PdfHeaderCompose> build) { Guard.NotNull(build, nameof(build)); var h = new PdfHeaderCompose(Options); build(h); return this; }
    /// <summary>Defines the footer layout and content.</summary>
    public PdfPageCompose Footer(System.Action<PdfFooterCompose> build) { Guard.NotNull(build, nameof(build)); var f = new PdfFooterCompose(Options); build(f); return this; }
}
