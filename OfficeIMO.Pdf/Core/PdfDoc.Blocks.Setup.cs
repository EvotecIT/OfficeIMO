using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDoc {
    /// <summary>Adds a level-1 heading.</summary>
    public PdfDoc H1(string text, PdfAlign align = PdfAlign.Left, PdfColor? color = null, string? linkUri = null, PdfHeadingStyle? style = null, string? linkContents = null, string? linkDestinationName = null) {
        Guard.NotNullOrWhiteSpace(text, nameof(text));
        Guard.OptionalAbsoluteUri(linkUri, nameof(linkUri));
        AddBlock(new HeadingBlock(1, text, align, color, linkUri, style, linkContents, linkDestinationName)); return this;
    }
    /// <summary>Adds a level-2 heading.</summary>
    public PdfDoc H2(string text, PdfAlign align = PdfAlign.Left, PdfColor? color = null, string? linkUri = null, PdfHeadingStyle? style = null, string? linkContents = null, string? linkDestinationName = null) {
        Guard.NotNullOrWhiteSpace(text, nameof(text));
        Guard.OptionalAbsoluteUri(linkUri, nameof(linkUri));
        AddBlock(new HeadingBlock(2, text, align, color, linkUri, style, linkContents, linkDestinationName)); return this;
    }
    /// <summary>Adds a level-3 heading.</summary>
    public PdfDoc H3(string text, PdfAlign align = PdfAlign.Left, PdfColor? color = null, string? linkUri = null, PdfHeadingStyle? style = null, string? linkContents = null, string? linkDestinationName = null) {
        Guard.NotNullOrWhiteSpace(text, nameof(text));
        Guard.OptionalAbsoluteUri(linkUri, nameof(linkUri));
        AddBlock(new HeadingBlock(3, text, align, color, linkUri, style, linkContents, linkDestinationName)); return this;
    }

    /// <summary>Inserts a page break.</summary>
    public PdfDoc PageBreak() { AddBlock(new PageBreakBlock()); return this; }

    /// <summary>Adds invisible vertical space to the current document flow.</summary>
    public PdfDoc Spacer(double height) {
        AddBlock(new SpacerBlock(height));
        return this;
    }

    /// <summary>Configures a page-scoped flow with its own page setup and default styles.</summary>
    public PdfDoc Page(System.Action<PdfPageCompose> configure) {
        AddComposedPage(configure);
        return this;
    }

    /// <summary>Configures a section-scoped flow with its own page setup and default styles.</summary>
    public PdfDoc Section(System.Action<PdfPageCompose> configure) {
        AddComposedPage(configure);
        return this;
    }

    /// <summary>Sets the document-wide default page size used by top-level flow and composed pages.</summary>
    public PdfDoc Size(PageSize size) {
        _options.PageSize = size;
        return this;
    }

    /// <summary>Sets the document-wide default page size in points.</summary>
    public PdfDoc Size(double width, double height) {
        _options.PageSize = new PageSize(width, height);
        return this;
    }

    /// <summary>Sets the document-wide default page orientation while preserving the current page size dimensions.</summary>
    public PdfDoc Orientation(PdfPageOrientation orientation) {
        _options.PageSize = _options.PageSize.WithOrientation(orientation);
        return this;
    }

    /// <summary>Sets or clears the document-wide default page background color.</summary>
    public PdfDoc Background(PdfColor? color) {
        _options.BackgroundColor = color;
        return this;
    }

    /// <summary>Sets or clears the document-wide default text watermark rendered behind page content.</summary>
    public PdfDoc Watermark(PdfTextWatermark? watermark) {
        _options.TextWatermark = watermark;
        return this;
    }

    /// <summary>Sets a document-wide default text watermark rendered behind page content.</summary>
    public PdfDoc Watermark(string text, double? fontSize = null, PdfColor? color = null, double? opacity = null, double? rotationAngle = null, PdfStandardFont? font = null, bool bold = true, bool italic = false) {
        var watermark = new PdfTextWatermark(text) {
            Bold = bold,
            Italic = italic
        };
        if (fontSize.HasValue) watermark.FontSize = fontSize.Value;
        if (color.HasValue) watermark.Color = color.Value;
        if (opacity.HasValue) watermark.Opacity = opacity.Value;
        if (rotationAngle.HasValue) watermark.RotationAngle = rotationAngle.Value;
        if (font.HasValue) watermark.Font = font.Value;
        _options.TextWatermark = watermark;
        return this;
    }

    /// <summary>Sets or clears the document-wide default image watermark rendered behind page content.</summary>
    public PdfDoc ImageWatermark(PdfImageWatermark? watermark) {
        _options.ImageWatermark = watermark;
        return this;
    }

    /// <summary>Sets a document-wide default image watermark rendered behind page content.</summary>
    public PdfDoc ImageWatermark(byte[] imageBytes, double width, double height, double? opacity = null, double? rotationAngle = null) {
        var watermark = new PdfImageWatermark(imageBytes, width, height);
        if (opacity.HasValue) watermark.Opacity = opacity.Value;
        if (rotationAngle.HasValue) watermark.RotationAngle = rotationAngle.Value;
        _options.ImageWatermark = watermark;
        return this;
    }

    /// <summary>Sets or clears the document-wide default page border.</summary>
    public PdfDoc PageBorder(PdfPageBorder? border) {
        _options.PageBorder = border;
        return this;
    }

    /// <summary>Sets a document-wide default page border.</summary>
    public PdfDoc PageBorder(PdfColor? color = null, double? width = null, double? inset = null, double? opacity = null, OfficeStrokeDashStyle dashStyle = OfficeStrokeDashStyle.Solid) {
        var border = new PdfPageBorder {
            DashStyle = dashStyle
        };
        if (color.HasValue) border.Color = color.Value;
        if (width.HasValue) border.Width = width.Value;
        if (inset.HasValue) border.Inset = inset.Value;
        if (opacity.HasValue) border.Opacity = opacity.Value;
        _options.PageBorder = border;
        return this;
    }

    /// <summary>Sets or clears the document-wide default page background image.</summary>
    public PdfDoc BackgroundImage(PdfPageBackgroundImage? image) {
        _options.PageBackgroundImage = image;
        return this;
    }

    /// <summary>Sets a document-wide default page background image.</summary>
    public PdfDoc BackgroundImage(byte[] imageBytes, OfficeImageFit fit = OfficeImageFit.Cover, double? opacity = null) {
        var image = new PdfPageBackgroundImage(imageBytes) {
            Fit = fit
        };
        if (opacity.HasValue) image.Opacity = opacity.Value;
        _options.PageBackgroundImage = image;
        return this;
    }

    /// <summary>Adds a document-wide default page background shape rendered behind page content.</summary>
    public PdfDoc BackgroundShape(PdfPageBackgroundShape shape) {
        _options.AddPageBackgroundShape(shape);
        return this;
    }

    /// <summary>Replaces or clears document-wide default page background shapes.</summary>
    public PdfDoc BackgroundShapes(System.Collections.Generic.IEnumerable<PdfPageBackgroundShape>? shapes) {
        _options.PageBackgroundShapes = shapes?.ToList();
        return this;
    }

    /// <summary>Clears document-wide default page background shapes.</summary>
    public PdfDoc ClearBackgroundShapes() {
        _options.ClearPageBackgroundShapes();
        return this;
    }

    /// <summary>Adds a document-wide default rectangle page background shape.</summary>
    public PdfDoc BackgroundRectangle(double x, double y, double width, double height, PdfColor? fill = null, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeLinearGradient? fillGradient = null) =>
        BackgroundShape(PdfPageBackgroundShape.Rectangle(x, y, width, height, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient));

    /// <summary>Adds a document-wide default rounded rectangle page background shape.</summary>
    public PdfDoc BackgroundRoundedRectangle(double x, double y, double width, double height, double cornerRadius, PdfColor? fill = null, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeLinearGradient? fillGradient = null) =>
        BackgroundShape(PdfPageBackgroundShape.RoundedRectangle(x, y, width, height, cornerRadius, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient));

    /// <summary>Adds a document-wide default ellipse page background shape.</summary>
    public PdfDoc BackgroundEllipse(double x, double y, double width, double height, PdfColor? fill = null, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeLinearGradient? fillGradient = null) =>
        BackgroundShape(PdfPageBackgroundShape.Ellipse(x, y, width, height, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient));

    /// <summary>Adds a document-wide default top page background band using the current page size.</summary>
    public PdfDoc BackgroundTopBand(double height, PdfColor? fill = null, double insetX = 0D, double offsetY = 0D, double cornerRadius = 0D, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeLinearGradient? fillGradient = null) =>
        BackgroundShape(PdfPageBackgroundShape.TopBand(_options.PageWidth, _options.PageHeight, height, fill, insetX, offsetY, cornerRadius, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient));

    /// <summary>Adds a document-wide default bottom page background band using the current page size.</summary>
    public PdfDoc BackgroundBottomBand(double height, PdfColor? fill = null, double insetX = 0D, double offsetY = 0D, double cornerRadius = 0D, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeLinearGradient? fillGradient = null) =>
        BackgroundShape(PdfPageBackgroundShape.BottomBand(_options.PageWidth, _options.PageHeight, height, fill, insetX, offsetY, cornerRadius, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient));

    /// <summary>Adds a document-wide default left page background band using the current page size.</summary>
    public PdfDoc BackgroundLeftBand(double width, PdfColor? fill = null, double insetY = 0D, double offsetX = 0D, double cornerRadius = 0D, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeLinearGradient? fillGradient = null) =>
        BackgroundShape(PdfPageBackgroundShape.LeftBand(_options.PageWidth, _options.PageHeight, width, fill, insetY, offsetX, cornerRadius, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient));

    /// <summary>Adds a document-wide default right page background band using the current page size.</summary>
    public PdfDoc BackgroundRightBand(double width, PdfColor? fill = null, double insetY = 0D, double offsetX = 0D, double cornerRadius = 0D, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeLinearGradient? fillGradient = null) =>
        BackgroundShape(PdfPageBackgroundShape.RightBand(_options.PageWidth, _options.PageHeight, width, fill, insetY, offsetX, cornerRadius, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient));

    /// <summary>Sets the document-wide default page orientation to portrait.</summary>
    public PdfDoc Portrait() => Orientation(PdfPageOrientation.Portrait);

    /// <summary>Sets the document-wide default page orientation to landscape.</summary>
    public PdfDoc Landscape() => Orientation(PdfPageOrientation.Landscape);

    /// <summary>Sets uniform document-wide default page margins in points.</summary>
    public PdfDoc Margin(double all) {
        _options.Margins = PageMargins.Uniform(all);
        return this;
    }

    /// <summary>Sets document-wide default page margins from a reusable margin value.</summary>
    public PdfDoc Margin(PageMargins margins) {
        _options.Margins = margins;
        return this;
    }

    /// <summary>Sets document-wide default page margins in points.</summary>
    public PdfDoc Margin(double left, double top, double right, double bottom) {
        _options.Margins = new PageMargins(left, top, right, bottom);
        return this;
    }

    /// <summary>Sets the first visible page number for the document-wide flow.</summary>
    public PdfDoc PageNumberStart(int start) {
        _options.PageNumberStart = start;
        return this;
    }

    /// <summary>Sets the document-wide visible page-number style for header/footer tokens.</summary>
    public PdfDoc PageNumberStyle(PdfPageNumberStyle style) {
        _options.PageNumberStyle = style;
        return this;
    }

    /// <summary>Defines the document-wide default header layout and content.</summary>
    public PdfDoc Header(System.Action<PdfHeaderCompose> build) {
        Guard.NotNull(build, nameof(build));
        var header = new PdfHeaderCompose(_options);
        build(header);
        return this;
    }

    /// <summary>Defines the document-wide default footer layout and content.</summary>
    public PdfDoc Footer(System.Action<PdfFooterCompose> build) {
        Guard.NotNull(build, nameof(build));
        var footer = new PdfFooterCompose(_options);
        build(footer);
        return this;
    }

    /// <summary>Applies reusable document-wide default styles.</summary>
    public PdfDoc Theme(PdfTheme theme) {
        Guard.NotNull(theme, nameof(theme));
        theme.Clone().ApplyTo(_options);
        return this;
    }

    /// <summary>Requests a generated-PDF compliance profile for this document.</summary>
    public PdfDoc Compliance(PdfComplianceProfile profile) {
        _options.RequireCompliance(profile);
        return this;
    }

    /// <summary>Sets the generated PDF file header version.</summary>
    public PdfDoc FileVersion(PdfFileVersion version) {
        _options.SetFileVersion(version);
        return this;
    }

    /// <summary>Sets PDF/A XMP identification metadata. This does not by itself certify PDF/A conformance.</summary>
    public PdfDoc PdfAIdentification(PdfAIdentification? identification) {
        _options.SetPdfAIdentification(identification);
        return this;
    }

    /// <summary>Sets PDF/A XMP identification metadata. This does not by itself certify PDF/A conformance.</summary>
    public PdfDoc PdfAIdentification(int part, string conformance) {
        _options.SetPdfAIdentification(part, conformance);
        return this;
    }

    /// <summary>Sets PDF/UA XMP identification metadata. This does not by itself certify PDF/UA conformance.</summary>
    public PdfDoc PdfUaIdentification(PdfUaIdentification? identification) {
        _options.SetPdfUaIdentification(identification);
        return this;
    }

    /// <summary>Sets PDF/UA XMP identification metadata. This does not by itself certify PDF/UA conformance.</summary>
    public PdfDoc PdfUaIdentification(int part = 1) {
        _options.SetPdfUaIdentification(part);
        return this;
    }

    /// <summary>Sets Factur-X/ZUGFeRD XMP extension metadata. This does not by itself certify e-invoice conformance.</summary>
    public PdfDoc ElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata? metadata) {
        _options.SetElectronicInvoiceMetadata(metadata);
        return this;
    }

    /// <summary>Sets Factur-X/ZUGFeRD XMP extension metadata. This does not by itself certify e-invoice conformance.</summary>
    public PdfDoc ElectronicInvoiceMetadata(string conformanceLevel, string version = "1.0") {
        _options.SetElectronicInvoiceMetadata(conformanceLevel, version);
        return this;
    }

    /// <summary>Sets or clears the generated catalog output intent.</summary>
    public PdfDoc OutputIntent(PdfOutputIntent? outputIntent) {
        _options.SetOutputIntent(outputIntent);
        return this;
    }

    /// <summary>Sets a generated catalog output intent from ICC profile bytes using the default sRGB output condition identifier.</summary>
    public PdfDoc OutputIntent(byte[] iccProfile) {
        _options.SetOutputIntent(iccProfile);
        return this;
    }

    /// <summary>Sets a generated catalog output intent from ICC profile bytes.</summary>
    public PdfDoc OutputIntent(byte[] iccProfile, string outputConditionIdentifier) {
        _options.SetOutputIntent(iccProfile, outputConditionIdentifier);
        return this;
    }

    /// <summary>Sets a generated catalog output intent from ICC profile bytes using the default sRGB output condition identifier.</summary>
    public PdfDoc OutputIntent(byte[] iccProfile, PdfOutputIntentPolicy policy) {
        _options.SetOutputIntent(iccProfile, policy);
        return this;
    }

    /// <summary>Sets a generated catalog output intent from ICC profile bytes.</summary>
    public PdfDoc OutputIntent(byte[] iccProfile, string outputConditionIdentifier, PdfOutputIntentPolicy policy) {
        _options.SetOutputIntent(iccProfile, outputConditionIdentifier, policy);
        return this;
    }

    /// <summary>Sets the generated catalog output intent to OfficeIMO's built-in sRGB IEC61966-2.1 ICC profile.</summary>
    public PdfDoc SrgbOutputIntent() {
        _options.SetSrgbOutputIntent();
        return this;
    }

    /// <summary>Sets the generated tagged-PDF groundwork mode.</summary>
    public PdfDoc TaggedStructure(PdfTaggedStructureMode mode) {
        _options.SetTaggedStructureMode(mode);
        return this;
    }

    /// <summary>Emits catalog-level tagged-PDF markers without claiming full tagged-content generation.</summary>
    public PdfDoc TaggedPdfCatalogMarkers() {
        _options.EnableTaggedPdfCatalogMarkers();
        return this;
    }

    /// <summary>Sets or clears the generated catalog document language, for example "en-US".</summary>
    public PdfDoc Language(string? language) {
        _options.SetLanguage(language);
        return this;
    }

    /// <summary>Enables generated catalog page labels that match the configured page-number style and start number.</summary>
    public PdfDoc PageLabels(string? prefix = null) {
        _options.SetPageLabels(true, prefix);
        return this;
    }

    /// <summary>Enables or disables generated catalog page labels.</summary>
    public PdfDoc PageLabels(bool include, string? prefix = null) {
        _options.SetPageLabels(include, prefix);
        return this;
    }

    /// <summary>Sets or clears generated catalog viewer preferences.</summary>
    public PdfDoc ViewerPreferences(PdfViewerPreferencesOptions? preferences) {
        _options.SetViewerPreferences(preferences);
        return this;
    }

    /// <summary>Configures generated catalog viewer preferences.</summary>
    public PdfDoc ViewerPreferences(System.Action<PdfViewerPreferencesOptions> configure) {
        _options.ConfigureViewerPreferences(configure);
        return this;
    }

    /// <summary>Configures common PDF/UA-1 groundwork without enabling formal compliance profile generation.</summary>
    public PdfDoc ConfigurePdfUaGroundwork(string language = "en-US") {
        _options.ConfigurePdfUaGroundwork(language);
        return this;
    }

    /// <summary>Configures common PDF/A-2 or PDF/A-3 groundwork without enabling formal compliance profile generation.</summary>
    public PdfDoc ConfigurePdfAGroundwork(PdfComplianceProfile profile, string language = "en-US") {
        _options.ConfigurePdfAGroundwork(profile, language);
        return this;
    }

    /// <summary>Adds an embedded file associated with the generated PDF catalog.</summary>
    public PdfDoc AttachFile(PdfEmbeddedFile file) {
        _options.AddEmbeddedFile(file);
        return this;
    }

    /// <summary>Adds an embedded file associated with the generated PDF catalog.</summary>
    public PdfDoc AttachFile(
        string fileName,
        byte[] data,
        string? mimeType = null,
        PdfAssociatedFileRelationship relationship = PdfAssociatedFileRelationship.Unspecified,
        string? description = null) {
        _options.AddEmbeddedFile(fileName, data, mimeType, relationship, description);
        return this;
    }

    /// <summary>Adds the canonical Factur-X/ZUGFeRD CrossIndustryInvoice XML payload and matching XMP extension metadata.</summary>
    public PdfDoc AttachFacturXInvoiceXml(
        byte[] ciiXml,
        string conformanceLevel = "EN 16931",
        string version = "1.0",
        PdfAssociatedFileRelationship relationship = PdfAssociatedFileRelationship.Data,
        string? description = "Factur-X/ZUGFeRD invoice XML") {
        _options.AddFacturXInvoiceXml(ciiXml, conformanceLevel, version, relationship, description);
        return this;
    }

    /// <summary>Configures common PDF/A-3 Factur-X/ZUGFeRD groundwork without enabling formal compliance profile generation.</summary>
    public PdfDoc ConfigureFacturXGroundwork(
        byte[] ciiXml,
        string conformanceLevel = "EN 16931",
        string version = "1.0",
        PdfAssociatedFileRelationship relationship = PdfAssociatedFileRelationship.Data,
        string? description = "Factur-X/ZUGFeRD invoice XML") {
        _options.ConfigureFacturXGroundwork(ciiXml, conformanceLevel, version, relationship, description);
        return this;
    }

    /// <summary>Sets document-wide default text styling used by following page-flow content.</summary>
    public PdfDoc DefaultTextStyle(System.Action<PdfTextStyleCompose> style) {
        Guard.NotNull(style, nameof(style));
        var compose = new PdfTextStyleCompose(_options);
        style(compose);
        return this;
    }

    /// <summary>Sets document-wide default text styling from a reusable text style object.</summary>
    public PdfDoc DefaultTextStyle(PdfTextStyle style) {
        Guard.NotNull(style, nameof(style));
        style.Clone().ApplyTo(_options);
        return this;
    }

    /// <summary>Embeds a TrueType font file for a generated standard-font slot in this document.</summary>
    public PdfDoc EmbedStandardFont(PdfStandardFont font, byte[] data, string? fontName = null) {
        _options.EmbedStandardFont(font, data, fontName);
        return this;
    }

    /// <summary>Embeds a TrueType font file from disk for a generated standard-font slot in this document.</summary>
    public PdfDoc EmbedStandardFont(PdfStandardFont font, string path, string? fontName = null) {
        _options.EmbedStandardFont(font, path, fontName);
        return this;
    }

    /// <summary>Uses a caller-supplied TrueType font family as the generated document's default font family.</summary>
    public PdfDoc UseFontFamily(
        string familyName,
        byte[] regular,
        byte[]? bold = null,
        byte[]? italic = null,
        byte[]? boldItalic = null) {
        _options.UseFontFamily(familyName, regular, bold, italic, boldItalic);
        return this;
    }

    /// <summary>Uses a reusable caller-supplied TrueType font family as the generated document's default font family.</summary>
    public PdfDoc UseFontFamily(PdfEmbeddedFontFamily fontFamily) {
        _options.UseFontFamily(fontFamily);
        return this;
    }

    /// <summary>Uses caller-supplied TrueType font files as the generated document's default font family.</summary>
    public PdfDoc UseFontFamily(
        string familyName,
        string regularPath,
        string? boldPath = null,
        string? italicPath = null,
        string? boldItalicPath = null) {
        _options.UseFontFamily(familyName, regularPath, boldPath, italicPath, boldItalicPath);
        return this;
    }

    /// <summary>Sets the document-wide default paragraph style used by paragraphs that do not provide an explicit style.</summary>
    public PdfDoc DefaultParagraphStyle(PdfParagraphStyle style) {
        Guard.NotNull(style, nameof(style));
        _options.DefaultParagraphStyle = style;
        return this;
    }

    /// <summary>Sets the document-wide default table style used by tables that do not provide an explicit style.</summary>
    public PdfDoc DefaultTableStyle(PdfTableStyle style) {
        Guard.NotNull(style, nameof(style));
        _options.DefaultTableStyle = style;
        return this;
    }

    /// <summary>Sets the document-wide default table style from a supported Word table style name.</summary>
    public PdfDoc DefaultTableStyle(string wordTableStyleName) {
        _options.DefaultTableStyle = TableStyles.FromWordTableStyle(wordTableStyleName);
        return this;
    }

    /// <summary>Sets the document-wide default style for a built-in heading level.</summary>
    public PdfDoc DefaultHeadingStyle(int level, PdfHeadingStyle style) {
        Guard.NotNull(style, nameof(style));
        _options.SetDefaultHeadingStyle(level, style);
        return this;
    }
}
