namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    /// <summary>Creates a deep copy of the options.</summary>
    public PdfOptions Clone() {
        var clone = new PdfOptions {
            PageWidth = PageWidth,
            PageHeight = PageHeight,
            BackgroundColor = BackgroundColor,
            TextWatermark = _textWatermark?.Clone(),
            FirstPageTextWatermark = _firstPageTextWatermark?.Clone(),
            EvenPageTextWatermark = _evenPageTextWatermark?.Clone(),
            _suppressFirstPageTextWatermark = _suppressFirstPageTextWatermark,
            _suppressEvenPageTextWatermark = _suppressEvenPageTextWatermark,
            ImageWatermark = _imageWatermark?.Clone(),
            FirstPageImageWatermark = _firstPageImageWatermark?.Clone(),
            EvenPageImageWatermark = _evenPageImageWatermark?.Clone(),
            _suppressFirstPageImageWatermark = _suppressFirstPageImageWatermark,
            _suppressEvenPageImageWatermark = _suppressEvenPageImageWatermark,
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
            ObjectBufferMemoryLimitBytes = ObjectBufferMemoryLimitBytes,
            PageContentMemoryLimitBytes = PageContentMemoryLimitBytes,
            IncludeStandardFontToUnicodeMaps = IncludeStandardFontToUnicodeMaps,
            CompressEmbeddedFonts = CompressEmbeddedFonts,
            IncludeXmpMetadata = IncludeXmpMetadata,
            IncludePageLabels = IncludePageLabels,
            PageLabelPrefix = PageLabelPrefix,
            _pageLabelRanges = ClonePageLabelRanges(_pageLabelRanges),
            FlattenVisualAnnotations = FlattenVisualAnnotations,
            FileVersion = FileVersion,
            ComplianceProfile = ComplianceProfile,
            PdfAIdentification = _pdfAIdentification?.Clone(),
            PdfUaIdentification = _pdfUaIdentification?.Clone(),
            ElectronicInvoiceMetadata = _electronicInvoiceMetadata?.Clone(),
            OutputIntent = _outputIntent?.Clone(),
            TaggedStructureMode = TaggedStructureMode,
            Language = Language,
            CatalogPageMode = CatalogPageMode,
            CatalogPageLayout = CatalogPageLayout,
            OpenAction = _openAction?.Clone(),
            ViewerPreferences = _viewerPreferences?.Clone(),
            CatalogUriBase = CatalogUriBase,
            Encryption = _encryption?.Clone(),
            AcroFormDefaultTextAlignment = AcroFormDefaultTextAlignment,
            _embeddedFontFallbacks = _embeddedFontFallbacks?.Clone(),
            TextLineBreakCallback = _textLineBreakCallback,
            TextHyphenationCallback = _textHyphenationCallback,
            TextShapingMode = TextShapingMode,
            TextShapingProvider = TextShapingProvider,
            _diagnosticsReport = _diagnosticsReport,
            _diagnosticsConverter = _diagnosticsConverter,
            _embeddedFonts = CloneEmbeddedFonts(_embeddedFonts),
            _namedFontFamilies = CloneNamedFontFamilies(_namedFontFamilies),
            _embeddedFiles = CloneEmbeddedFiles(_embeddedFiles),
            Portfolio = _portfolio?.Clone(),
            ShowHeader = ShowHeader,
            HeaderFormat = HeaderFormat,
            DifferentFirstPageHeaderFooter = DifferentFirstPageHeaderFooter,
            FirstPageHeaderFormat = FirstPageHeaderFormat,
            DifferentOddAndEvenPagesHeaderFooter = DifferentOddAndEvenPagesHeaderFooter,
            EvenPageHeaderFormat = EvenPageHeaderFormat,
            HeaderFont = HeaderFont,
            HeaderFontFamily = HeaderFontFamily,
            HeaderFontSize = HeaderFontSize,
            HeaderTextColor = HeaderTextColor,
            HeaderAlign = HeaderAlign,
            HeaderOffsetY = HeaderOffsetY,
            ShowPageNumbers = ShowPageNumbers,
            FooterFormat = FooterFormat,
            FirstPageFooterFormat = FirstPageFooterFormat,
            EvenPageFooterFormat = EvenPageFooterFormat,
            FooterFont = FooterFont,
            FooterFontFamily = FooterFontFamily,
            FooterFontSize = FooterFontSize,
            FooterTextColor = FooterTextColor,
            FooterAlign = FooterAlign,
            FooterOffsetY = FooterOffsetY,
            PageNumberStyle = PageNumberStyle,
            DefaultTextColor = DefaultTextColor,
            DefaultParagraphStyle = _defaultParagraphStyle?.Clone(),
            DefaultHeadingStyles = _defaultHeadingStyles?.Clone(),
            DefaultListStyle = _defaultListStyle?.Clone(),
            DefaultPanelStyle = _defaultPanelStyle?.Clone(),
            DefaultHorizontalRuleStyle = _defaultHorizontalRuleStyle?.Clone(),
            DefaultImageStyle = _defaultImageStyle?.Clone(),
            ImageOptimization = _imageOptimization?.Clone(),
            DefaultDrawingStyle = _defaultDrawingStyle?.Clone(),
            DefaultRowStyle = _defaultRowStyle?.Clone(),
            CreateOutlineFromHeadings = CreateOutlineFromHeadings,
            OutlineExpansionLevel = OutlineExpansionLevel,
            Debug = Debug is null ? null : new PdfDebugOptions {
                ShowContentArea = Debug.ShowContentArea,
                ShowFlowObjectBoxes = Debug.ShowFlowObjectBoxes,
                ShowCanvasItemBoxes = Debug.ShowCanvasItemBoxes,
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
        clone._defaultTableStyle = _defaultTableStyle?.Clone();
        clone._hasExplicitDefaultTableStyle = _hasExplicitDefaultTableStyle;
        clone._hasExplicitDefaultFont = _hasExplicitDefaultFont;
        clone._hasExplicitHeaderFont = _hasExplicitHeaderFont;
        clone._hasExplicitFooterFont = _hasExplicitFooterFont;
        clone._pageNumberStart = _pageNumberStart;
        clone._hasExplicitPageNumberStart = _hasExplicitPageNumberStart;
        return clone;
    }

    private static System.Collections.Generic.List<PdfPageLabelRange>? ClonePageLabelRanges(System.Collections.Generic.IEnumerable<PdfPageLabelRange>? ranges) {
        if (ranges == null) {
            return null;
        }

        var clone = new System.Collections.Generic.List<PdfPageLabelRange>();
        foreach (PdfPageLabelRange range in ranges) {
            clone.Add(new PdfPageLabelRange(range.StartPageNumber, range.Style, range.StartNumber, range.Prefix));
        }

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

    private static System.Collections.Generic.Dictionary<string, PdfEmbeddedFontFamily>? CloneNamedFontFamilies(System.Collections.Generic.Dictionary<string, PdfEmbeddedFontFamily>? fonts) {
        if (fonts == null) {
            return null;
        }

        var clone = new System.Collections.Generic.Dictionary<string, PdfEmbeddedFontFamily>(System.StringComparer.Ordinal);
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

}
