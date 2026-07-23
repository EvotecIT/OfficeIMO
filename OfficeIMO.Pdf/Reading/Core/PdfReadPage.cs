using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Represents a single page parsed from the PDF.
/// Provides access to plain text and basic text spans based on content stream operators.
/// </summary>
public sealed partial class PdfReadPage {
    private readonly PdfDictionary _pageDict;
    private readonly Dictionary<int, PdfIndirectObject> _objects;
    private readonly int _maxDecodedStreamBytes;
    private readonly PdfReadLimits _limits;
    private readonly Action? _demandTextExtraction;
    private readonly Action<string>? _demandContentExtraction;

    internal PdfReadPage(int objectNumber, PdfDictionary pageDict, Dictionary<int, PdfIndirectObject> objects)
        : this(objectNumber, pageDict, objects, new PdfReadLimits()) { }

    internal PdfReadPage(
        int objectNumber,
        PdfDictionary pageDict,
        Dictionary<int, PdfIndirectObject> objects,
        PdfReadLimits limits,
        Action? demandTextExtraction = null,
        Action<string>? demandContentExtraction = null) {
        ObjectNumber = objectNumber;
        _pageDict = pageDict;
        _objects = objects;
        _limits = limits;
        _maxDecodedStreamBytes = limits.MaxDecodedStreamBytes;
        _demandTextExtraction = demandTextExtraction;
        _demandContentExtraction = demandContentExtraction;
    }

    /// <summary>Underlying object number for the page.</summary>
    public int ObjectNumber { get; }

    /// <summary>Extracts plain text from this page without column reordering.</summary>
    public string ExtractText() {
        var spans = GetTextSpans();
        var opts = new TextLayoutEngine.Options { ForceSingleColumn = true };
        var lines = TextLayoutEngine.BuildLines(spans, opts);
        return TextLayoutEngine.EmitText(lines, TextLayoutEngine.DetectColumns(lines, GetPageSize().Width, opts), null);
    }

    /// <summary>
    /// Attempts to read page size from CropBox (or MediaBox) and returns width/height in points.
    /// Falls back to 612x792 (US Letter) when not present or malformed.
    /// </summary>
    public (double Width, double Height) GetPageSize() {
        PdfPageBox box = GetPageBoundaryBox();
        return (box.Width, box.Height);
    }

    private (double Width, double Height) GetVisualPageSize() {
        (double Width, double Height) pageSize = GetPageSize();
        int rotation = GetRotationDegrees();
        return rotation == 90 || rotation == 270
            ? (pageSize.Height, pageSize.Width)
            : pageSize;
    }

    private Matrix2D GetVisualPageTransform() {
        PdfPageBox pageBox = GetPageBoundaryBox();
        Matrix2D cropOrigin = Matrix2D.Translation(-pageBox.Left, -pageBox.Bottom);
        Matrix2D rotation = GetRotationDegrees() switch {
            90 => new Matrix2D(0D, 1D, -1D, 0D, pageBox.Height, 0D),
            180 => new Matrix2D(-1D, 0D, 0D, -1D, pageBox.Width, pageBox.Height),
            270 => new Matrix2D(0D, -1D, 1D, 0D, 0D, pageBox.Width),
            _ => Matrix2D.Identity
        };

        return Matrix2D.Multiply(rotation, cropOrigin);
    }

    internal (double Width, double Height) GetInteractionPageSize() => GetVisualPageSize();

    internal (double X, double Y) TransformPointToVisual(double x, double y) => GetVisualPageTransform().Transform(x, y);

    internal IReadOnlyList<PdfTextSpan> GetInteractionTextSpans() {
        _demandTextExtraction?.Invoke();
        (double Width, double Height) size = GetVisualPageSize();
        return GetVisualTextSpans(size.Height, GetVisualPageTransform());
    }

    private PdfPageBox GetPageBoundaryBox() {
        if (TryReadPageBox("CropBox", out PdfPageBox? cropBox) && cropBox != null) {
            return cropBox;
        }

        if (TryReadPageBox("MediaBox", out PdfPageBox? mediaBox) && mediaBox != null) {
            return mediaBox;
        }

        return new PdfPageBox("MediaBox", 0D, 0D, 612D, 792D);
    }
    /// <summary>Gets inherited page rotation in degrees normalized to 0, 90, 180, or 270.</summary>
    public int GetRotationDegrees() {
        var rotate = GetInheritedValue("Rotate");
        if (rotate is PdfNumber number) {
            int degrees = (int)Math.Round(number.Value);
            degrees %= 360;
            if (degrees < 0) {
                degrees += 360;
            }

            return degrees;
        }

        return 0;
    }

    /// <summary>Gets text spans (text with position and font info) from this page.</summary>
    public IReadOnlyList<PdfTextSpan> GetTextSpans() {
        _demandTextExtraction?.Invoke();
        var spans = new List<PdfTextSpan>();
        var pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        var pageDecoders = ResourceResolver.GetBudgetedFontDecoders(_pageDict, _objects);
        var pageWidthProviders = ResourceResolver.GetFontWidthProviders(_pageDict, _objects);
        var pageFonts = ResourceResolver.GetFontsForResources(pageResources, _objects);
        var activeForms = new HashSet<PdfStream>();
        double pageHeight = GetPageSize().Height;
        var pageContentBudget = new PageContentBudget(this);

        string content = GetContentStreamContent(pageContentBudget);
        if (content.Length > 0) {
            CollectTextAndForms(
                content,
                pageResources,
                pageDecoders,
                pageWidthProviders,
                pageFonts,
                spans,
                activeForms,
                pageHeight,
                pageContentBudget: pageContentBudget);
        }

        return spans;
    }

    /// <summary>Reads simple URI, named-destination, direct-destination, named-action, and remote GoTo link annotations from this page.</summary>
    public IReadOnlyList<PdfLinkAnnotation> GetLinkAnnotations() {
        _demandContentExtraction?.Invoke("link annotation");
        return GetLinkAnnotationsUnchecked();
    }

    internal IReadOnlyList<PdfLinkAnnotation> GetLinkAnnotationsUnchecked() {
        if (!_pageDict.Items.TryGetValue("Annots", out var annotsObject)) {
            return Array.Empty<PdfLinkAnnotation>();
        }

        var annotations = ResolveArray(annotsObject);
        if (annotations is null) {
            return Array.Empty<PdfLinkAnnotation>();
        }
        EnsureAnnotationBudget(annotations);

        var result = new List<PdfLinkAnnotation>();
        foreach (var item in annotations.Items) {
            var annotation = ResolveDictionary(item);
            if (annotation is null ||
                annotation.Get<PdfName>("Subtype")?.Name != "Link" ||
                !TryReadRectangle(annotation.Items.TryGetValue("Rect", out var rectObject) ? rectObject : null, out var rect)) {
                continue;
            }

            var action = ResolveDictionary(annotation.Items.TryGetValue("A", out var actionObject) ? actionObject : null);
            TryGetString(annotation.Items.TryGetValue("Contents", out var contentsObject) ? contentsObject : null, out string? contents);

            if (action != null &&
                action.Get<PdfName>("S")?.Name == "URI" &&
                TryGetString(action.Items.TryGetValue("URI", out var uriObject) ? uriObject : null, out string? uri) &&
                Guard.IsUriAction(uri)) {
                result.Add(new PdfLinkAnnotation(uri!, contents, rect.X1, rect.Y1, rect.X2, rect.Y2));
                continue;
            }

            if (action != null &&
                action.Get<PdfName>("S")?.Name == "GoTo" &&
                TryReadLinkDestination(action.Items.TryGetValue("D", out var actionDestination) ? actionDestination : null, out string? actionDestinationName, out int? actionDestinationPageObjectNumber, out double? actionDestinationTop, out PdfOpenActionDestinationMode? actionDestinationMode, out double? actionDestinationLeft, out double? actionDestinationBottom, out double? actionDestinationRight)) {
                result.Add(new PdfLinkAnnotation(null, actionDestinationName, contents, rect.X1, rect.Y1, rect.X2, rect.Y2, destinationPageObjectNumber: actionDestinationPageObjectNumber, destinationTop: actionDestinationTop, destinationMode: actionDestinationMode, destinationLeft: actionDestinationLeft, destinationBottom: actionDestinationBottom, destinationRight: actionDestinationRight));
                continue;
            }

            if (action != null &&
                action.Get<PdfName>("S")?.Name == "Named" &&
                TryGetNameOrString(action.Items.TryGetValue("N", out var namedActionObject) ? namedActionObject : null, out string? namedAction)) {
                result.Add(new PdfLinkAnnotation(null, null, contents, rect.X1, rect.Y1, rect.X2, rect.Y2, namedAction: namedAction));
                continue;
            }

            if (action != null &&
                action.Get<PdfName>("S")?.Name == "GoToR" &&
                TryReadFileSpecification(action.Items.TryGetValue("F", out var remoteFileObject) ? remoteFileObject : null, out string? remoteFile)) {
                TryReadRemoteDestination(action.Items.TryGetValue("D", out var remoteDestinationObject) ? remoteDestinationObject : null, out string? remoteDestinationName, out int? remoteDestinationPageNumber, out double? remoteDestinationTop, out PdfOpenActionDestinationMode? remoteDestinationMode, out double? remoteDestinationLeft, out double? remoteDestinationBottom, out double? remoteDestinationRight);
                result.Add(new PdfLinkAnnotation(null, null, contents, rect.X1, rect.Y1, rect.X2, rect.Y2, remoteFile: remoteFile, remoteDestinationName: remoteDestinationName, remoteDestinationPageNumber: remoteDestinationPageNumber, remoteDestinationTop: remoteDestinationTop, remoteDestinationMode: remoteDestinationMode, remoteDestinationLeft: remoteDestinationLeft, remoteDestinationBottom: remoteDestinationBottom, remoteDestinationRight: remoteDestinationRight));
                continue;
            }

            if (TryReadLinkDestination(annotation.Items.TryGetValue("Dest", out var directDestination) ? directDestination : null, out string? directDestinationName, out int? directDestinationPageObjectNumber, out double? directDestinationTop, out PdfOpenActionDestinationMode? directDestinationMode, out double? directDestinationLeft, out double? directDestinationBottom, out double? directDestinationRight)) {
                result.Add(new PdfLinkAnnotation(null, directDestinationName, contents, rect.X1, rect.Y1, rect.X2, rect.Y2, destinationPageObjectNumber: directDestinationPageObjectNumber, destinationTop: directDestinationTop, destinationMode: directDestinationMode, destinationLeft: directDestinationLeft, destinationBottom: directDestinationBottom, destinationRight: directDestinationRight));
            }
        }

        return result.AsReadOnly();
    }

    /// <summary>Reads generic annotation metadata from this page.</summary>
    public IReadOnlyList<PdfAnnotation> GetAnnotations() {
        _demandContentExtraction?.Invoke("annotation");
        return GetAnnotationsUnchecked();
    }

    internal IReadOnlyList<PdfAnnotation> GetAnnotationsUnchecked() {
        if (!_pageDict.Items.TryGetValue("Annots", out var annotsObject)) {
            return Array.Empty<PdfAnnotation>();
        }

        var annotations = ResolveArray(annotsObject);
        if (annotations is null) {
            return Array.Empty<PdfAnnotation>();
        }
        EnsureAnnotationBudget(annotations);

        var result = new List<PdfAnnotation>();
        foreach (var item in annotations.Items) {
            int? objectNumber = item is PdfReference reference ? reference.ObjectNumber : null;
            var annotation = ResolveDictionary(item);
            string? subtype = annotation?.Get<PdfName>("Subtype")?.Name;
            if (annotation is null ||
                string.IsNullOrWhiteSpace(subtype) ||
                !TryReadRectangle(annotation.Items.TryGetValue("Rect", out var rectObject) ? rectObject : null, out var rect)) {
                continue;
            }

            TryGetString(annotation.Items.TryGetValue("Contents", out var contentsObject) ? contentsObject : null, out string? contents);
            bool hasNormalAppearance = HasNormalAppearance(annotation);
            annotation.Items.TryGetValue("A", out var actionObject);
            annotation.Items.TryGetValue("AA", out var additionalActionsObject);
            string? actionType = TryReadActionType(actionObject);
            IReadOnlyList<PdfAnnotationAdditionalAction> additionalActions = ReadAdditionalActions(additionalActionsObject);
            IReadOnlyList<PdfAnnotationChainedAction> chainedActions = ReadAnnotationChainedActions(actionObject, additionalActionsObject);
            int? flags = TryReadInteger(annotation.Items.TryGetValue("F", out var flagsObject) ? flagsObject : null);
            TryGetString(annotation.Items.TryGetValue("NM", out var nameObject) ? nameObject : null, out string? name);
            TryGetString(annotation.Items.TryGetValue("T", out var titleObject) ? titleObject : null, out string? title);
            TryGetString(annotation.Items.TryGetValue("M", out var modifiedObject) ? modifiedObject : null, out string? modified);
            IReadOnlyList<double> color = ReadNumberArray(annotation.Items.TryGetValue("C", out var colorObject) ? colorObject : null);
            ReadFreeTextAppearanceMetadata(
                annotation,
                subtype!,
                out string? defaultAppearance,
                out string? defaultStyle,
                out string? richContents,
                out string? richContentsPlainText,
                out double? effectiveFontSize,
                out PdfColor? effectiveTextColor,
                out PdfAlign? effectiveTextAlign);
            ReadAnnotationVisualStyleMetadata(
                annotation,
                subtype!,
                rect.X2 - rect.X1,
                rect.Y2 - rect.Y1,
                out IReadOnlyList<double> interiorColor,
                out double? opacity,
                out double? borderWidth,
                out string? borderStyle,
                out IReadOnlyList<double> borderDashPattern,
                out string? borderEffectStyle,
                out double? borderEffectIntensity,
                out IReadOnlyList<double> rectangleDifferences,
                out IReadOnlyList<double> calloutLine,
                out string? calloutLineEnding,
                out string? lineStartEnding,
                out string? lineEndEnding);
            ReadAnnotationPathGeometryMetadata(
                annotation,
                out IReadOnlyList<double> quadPoints,
                out IReadOnlyList<double> lineCoordinates,
                out IReadOnlyList<double> vertices,
                out IReadOnlyList<IReadOnlyList<double>> inkList);
            result.Add(new PdfAnnotation(objectNumber, null, subtype!, contents, rect.X1, rect.Y1, rect.X2, rect.Y2, hasNormalAppearance, actionType, additionalActions, chainedActions, flags, name, title, modified, color, defaultAppearance, defaultStyle, richContents, richContentsPlainText, effectiveFontSize, effectiveTextColor, effectiveTextAlign, interiorColor, opacity, borderWidth, borderStyle, borderDashPattern, borderEffectStyle, borderEffectIntensity, rectangleDifferences, calloutLine, calloutLineEnding, lineStartEnding, lineEndEnding, quadPoints, lineCoordinates, vertices, inkList));
        }

        return result.Count == 0 ? Array.Empty<PdfAnnotation>() : result.AsReadOnly();
    }

    internal IReadOnlyList<int> GetAnnotationObjectNumbers(string subtypeName) {
        if (!_pageDict.Items.TryGetValue("Annots", out var annotsObject)) {
            return Array.Empty<int>();
        }

        var annotations = ResolveArray(annotsObject);
        if (annotations is null) {
            return Array.Empty<int>();
        }
        EnsureAnnotationBudget(annotations);

        var result = new List<int>();
        foreach (var item in annotations.Items) {
            if (item is not PdfReference reference) {
                continue;
            }

            var annotation = ResolveDictionary(reference);
            if (annotation?.Get<PdfName>("Subtype")?.Name == subtypeName) {
                result.Add(reference.ObjectNumber);
            }
        }

        return result.Count == 0 ? Array.Empty<int>() : result.AsReadOnly();
    }

    private void EnsureAnnotationBudget(PdfArray annotations) {
        if (annotations.Items.Count > _limits.MaxAnnotationsPerPage) {
            throw PdfReadLimitException.Create(
                PdfReadLimitKind.AnnotationsPerPage,
                _limits.MaxAnnotationsPerPage,
                annotations.Items.Count);
        }
    }

    /// <summary>Extracts image XObjects referenced by this page.</summary>
    public IReadOnlyList<PdfExtractedImage> GetImages() {
        _demandContentExtraction?.Invoke("image");
        return GetImages(0);
    }

    internal IReadOnlyList<PdfExtractedImage> GetImages(int pageNumber) {
        return GetImages(pageNumber, GetImagePlacements(pageNumber));
    }

    internal IReadOnlyList<PdfExtractedImage> GetImages(int pageNumber, IReadOnlyList<PdfImagePlacement>? imagePlacements) {
        return GetImages(pageNumber, imagePlacements, colorizeImageMasks: false);
    }

    internal IReadOnlyList<PdfExtractedImage> GetImages(int pageNumber, IReadOnlyList<PdfImagePlacement>? imagePlacements, bool colorizeImageMasks) {
        return GetImagesForResources(ResolveDictionary(GetInheritedValue("Resources")), pageNumber, imagePlacements, colorizeImageMasks);
    }

    private IReadOnlyList<PdfExtractedImage> GetImagesForResources(PdfDictionary? resources, int pageNumber, IReadOnlyList<PdfImagePlacement>? imagePlacements, bool colorizeImageMasks = false) {
        var images = resources == null
            ? new List<PdfExtractedImage>()
            : new List<PdfExtractedImage>(ResourceResolver.GetImageXObjectsForResources(resources, _objects, pageNumber, imagePlacements, colorizeImageMasks, _limits));
        if (imagePlacements is not null) {
            for (int i = 0; i < imagePlacements.Count; i++) {
                PdfImagePlacement placement = imagePlacements[i];
                if (placement.InlineImageStream == null) {
                    continue;
                }

                images.Add(ResourceResolver.BuildExtractedImage(
                    pageNumber,
                    placement.ResourceName,
                    placement.ObjectNumber,
                    placement.DirectStreamIdentity,
                    placement.InlineImageStream,
                    _objects,
                    placement.ImageMaskColor,
                    placement.InlineImageResources ?? resources,
                    colorizeImageMasks));
            }
        }

        return images.Count == 0 ? Array.Empty<PdfExtractedImage>() : images.AsReadOnly();
    }

    /// <summary>Extracts image XObject placement invocations from this page.</summary>
    public IReadOnlyList<PdfImagePlacement> GetImagePlacements() {
        _demandContentExtraction?.Invoke("image placement");
        return GetImagePlacements(0);
    }

    internal IReadOnlyList<PdfImagePlacement> GetImagePlacements(int pageNumber) {
        var placements = new List<PdfImagePlacement>();
        var pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        var activeForms = new HashSet<PdfStream>();
        double pageHeight = GetPageSize().Height;
        var pageContentBudget = new PageContentBudget(this);

        string content = GetContentStreamContent(pageContentBudget);
        if (content.Length > 0) {
            CollectImagePlacementsAndForms(
                content,
                pageResources,
                pageNumber,
                Matrix2D.Identity,
                pageHeight,
                placements,
                activeForms,
                pageContentBudget: pageContentBudget);
        }

        return placements.Count == 0 ? Array.Empty<PdfImagePlacement>() : placements.AsReadOnly();
    }

    internal List<string> GetUnsupportedContentStreamFilters() {
        var unsupported = new List<string>();
        var pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        var activeForms = new HashSet<PdfStream>();
        var pageContentBudget = new PageContentBudget(this);
        var content = new System.Text.StringBuilder();
        bool canInspectFormInvocations = true;
        foreach (var stream in GetContentStreamObjects()) {
            AddUnsupportedFilters(stream, unsupported);
            if (Filters.StreamDecoder.GetUnsupportedFilters(stream.Dictionary, _objects).Count != 0) {
                canInspectFormInvocations = false;
                continue;
            }

            content.Append(PdfEncoding.Latin1GetString(pageContentBudget.Decode(stream)));
        }

        if (canInspectFormInvocations && content.Length > 0) {
            CollectUnsupportedFormFilters(content.ToString(), pageResources, unsupported, activeForms, pageContentBudget);
        }

        return unsupported;
    }

    private void CollectUnsupportedFormFilters(
        string content,
        PdfDictionary? resources,
        List<string> unsupported,
        HashSet<PdfStream> activeForms,
        PageContentBudget pageContentBudget,
        int contentNestingDepth = 0) {
        EnsureContentNestingBudget(contentNestingDepth);
        foreach (var invocation in TextContentParser.ExtractFormInvocations(
                     content,
                     maxOperations: _limits.MaxContentOperations,
                     maxNestingDepth: _limits.MaxContentNestingDepth,
                     maxOperands: _limits.MaxContentOperands)) {
            if (!TryGetFormStream(resources, invocation.Name, out var formStream)) {
                continue;
            }

            if (!activeForms.Add(formStream)) {
                continue;
            }

            try {
                AddUnsupportedFilters(formStream, unsupported);
                if (Filters.StreamDecoder.GetUnsupportedFilters(formStream.Dictionary, _objects).Count != 0) {
                    continue;
                }

                var formResources = ResolveDictionary(formStream.Dictionary.Items.TryGetValue("Resources", out var resObj) ? resObj : null) ?? resources;
                CollectUnsupportedFormFilters(PdfEncoding.Latin1GetString(pageContentBudget.Decode(formStream)), formResources, unsupported, activeForms, pageContentBudget, contentNestingDepth + 1);
            } finally {
                activeForms.Remove(formStream);
            }
        }
    }

    private void AddUnsupportedFilters(PdfStream stream, List<string> unsupported) {
        foreach (string filterName in Filters.StreamDecoder.GetUnsupportedFilters(stream.Dictionary, _objects)) {
            if (!ContainsFilter(unsupported, filterName)) {
                unsupported.Add(filterName);
            }
        }
    }

    private void CollectTextAndForms(
        string content,
        PdfDictionary? resources,
        Dictionary<string, Func<byte[], int, string>> decoders,
        Dictionary<string, Func<byte[], double>> widthProviders,
        Dictionary<string, PdfFontResource> fonts,
        List<PdfTextSpan> spans,
        HashSet<PdfStream> activeForms,
        double pageHeight,
        double paintOrderBase = 0D,
        double paintOrderScale = 1D,
        double paintOrderOffset = 0D,
        OfficeColor? initialFillColor = null,
        PdfPageColorSpace initialFillColorSpace = default,
        OfficeColor? initialStrokeColor = null,
        PdfPageColorSpace initialStrokeColorSpace = default,
        double? initialFillOpacity = null,
        double? initialStrokeOpacity = null,
        int initialTextRenderingMode = 0,
        PdfPageClipPath? initialClipPath = null,
        bool useLogicalTextFilters = true,
        int contentNestingDepth = 0,
        TextContentParser.TextOutputBudget? textOutputBudget = null,
        PageContentBudget? pageContentBudget = null) {
        EnsureContentNestingBudget(contentNestingDepth);
        pageContentBudget ??= new PageContentBudget(this);
        textOutputBudget ??= new TextContentParser.TextOutputBudget(
            _limits.MaxActualTextCharacters,
            _limits.MaxDecodedTextCharacters);
        string DecodeWithFontWithinLimit(string fontRes, byte[] bytes, int maximumCharacters) =>
            decoders.TryGetValue(fontRes, out var dec)
                ? dec(bytes, maximumCharacters)
                : PdfWinAnsiEncoding.Decode(bytes, maximumCharacters);
        string DecodeWithFont(string fontRes, byte[] bytes) =>
            DecodeWithFontWithinLimit(fontRes, bytes, _limits.MaxDecodedTextCharacters);
        double SumWidth1000(string fontRes, byte[] bytes) =>
            widthProviders.TryGetValue(fontRes, out var wp) ? wp(bytes) : (bytes?.Length ?? 0) * 500.0;
        string? ResolveBaseFont(string fontRes) =>
            fonts.TryGetValue(fontRes, out PdfFontResource? font) ? font.BaseFont : null;
        string? ResolveDrawingFontFamily(string fontRes) =>
            fonts.TryGetValue(fontRes, out PdfFontResource? font) ? font.DrawingFontFamily : null;
        byte[]? ResolveActualTextProperty(string propertyName) =>
            GetMarkedContentActualTextBytes(resources, propertyName);

        spans.AddRange(TextContentParser.Parse(
            content,
            DecodeWithFont,
            SumWidth1000,
            actualTextForProperty: ResolveActualTextProperty,
            graphicsStates: GetGraphicsStateResources(resources),
            colorSpaces: GetColorSpaceResources(resources),
            baseFontForResource: ResolveBaseFont,
            drawingFontFamilyForResource: ResolveDrawingFontFamily,
            optionalContentVisibility: GetOptionalContentVisibility(resources),
            pageHeight: pageHeight,
            paintOrderBase: paintOrderBase,
            paintOrderScale: paintOrderScale,
            paintOrderOffset: paintOrderOffset,
            initialFillColor: initialFillColor,
            initialFillColorSpace: initialFillColorSpace,
            initialStrokeColor: initialStrokeColor,
            initialStrokeColorSpace: initialStrokeColorSpace,
            initialFillOpacity: initialFillOpacity,
            initialStrokeOpacity: initialStrokeOpacity,
            initialTextRenderingMode: initialTextRenderingMode,
            initialClipPath: initialClipPath,
            useLogicalTextFilters: useLogicalTextFilters,
            maxOperations: _limits.MaxContentOperations,
            maxNestingDepth: _limits.MaxContentNestingDepth,
            maxOperands: _limits.MaxContentOperands,
            maxActualTextCharacters: _limits.MaxActualTextCharacters,
            maxDecodedTextCharacters: _limits.MaxDecodedTextCharacters,
            textOutputBudget: textOutputBudget,
            decodeWithFontWithinLimit: DecodeWithFontWithinLimit));

        foreach (var invocation in TextContentParser.ExtractFormInvocations(
                     content,
                     GetOptionalContentVisibility(resources),
                     paintOrderBase,
                     paintOrderScale,
                     paintOrderOffset,
                     GetGraphicsStateResources(resources),
                     GetColorSpaceResources(resources),
                     pageHeight,
                     initialFillColor,
                     initialFillColorSpace,
                     initialStrokeColor,
                     initialStrokeColorSpace,
                     initialFillOpacity,
                     initialStrokeOpacity,
                     initialTextRenderingMode,
                     initialClipPath,
                     maxOperations: _limits.MaxContentOperations,
                     maxNestingDepth: _limits.MaxContentNestingDepth,
                     maxOperands: _limits.MaxContentOperands)) {
            if (!TryGetFormStream(resources, invocation.Name, out var formStream)) {
                continue;
            }

            if (!activeForms.Add(formStream)) {
                continue;
            }

            try {
                var formDict = formStream.Dictionary;
                var formResources = ResolveDictionary(formDict.Items.TryGetValue("Resources", out var resObj) ? resObj : null) ?? resources;
                var formDecoders = MergeDecoders(decoders, ResourceResolver.GetBudgetedFontDecodersForForm(formDict, _objects));
                var formWidths = MergeWidthProviders(widthProviders, ResourceResolver.GetFontWidthProviders(formDict, _objects));
                var formFonts = MergeFonts(fonts, ResourceResolver.GetFontsForResources(formResources, _objects));
                var combinedTransform = ApplyFormMatrix(invocation.Transform, formDict);
                var formContent = WrapContentWithTransform(WrapFormContentWithBoundingBoxClip(PdfEncoding.Latin1GetString(pageContentBudget.Decode(formStream)), formDict), combinedTransform, out int formContentOffset);

                CollectTextAndForms(
                    formContent,
                    formResources,
                    formDecoders,
                    formWidths,
                    formFonts,
                    spans,
                    activeForms,
                    pageHeight,
                    invocation.PaintOrder,
                    paintOrderScale * 0.000000001D,
                    -formContentOffset,
                    invocation.FillColor,
                    invocation.FillColorSpace,
                    invocation.StrokeColor,
                    invocation.StrokeColorSpace,
                    invocation.FillOpacity,
                    invocation.StrokeOpacity,
                    invocation.TextRenderingMode,
                    invocation.ClipPath,
                    useLogicalTextFilters,
                    contentNestingDepth + 1,
                    textOutputBudget,
                    pageContentBudget);
            } finally {
                activeForms.Remove(formStream);
            }
        }
    }

    private void CollectImagePlacementsAndForms(
        string content,
        PdfDictionary? resources,
        int pageNumber,
        Matrix2D baseTransform,
        double pageHeight,
        List<PdfImagePlacement> placements,
        HashSet<PdfStream> activeForms,
        OfficeColor? initialFillColor = null,
        PdfPageColorSpace initialFillColorSpace = default,
        double? initialFillOpacity = null,
        double paintOrderBase = 0D,
        double paintOrderScale = 1D,
        double paintOrderOffset = 0D,
        PdfPageClipPath? initialClipPath = null,
        int contentNestingDepth = 0,
        PageContentBudget? pageContentBudget = null) {
        EnsureContentNestingBudget(contentNestingDepth);
        pageContentBudget ??= new PageContentBudget(this);
        foreach (var invocation in PdfPageXObjectInvocationParser.Parse(
                     content,
                     baseTransform,
                     pageHeight,
                     GetGraphicsStateResources(resources),
                     GetColorSpaceResources(resources),
                     GetOptionalContentVisibility(resources),
                     initialFillColor,
                     initialFillColorSpace,
                      initialFillOpacity,
                      paintOrderBase,
                      paintOrderScale,
                      paintOrderOffset,
                      initialClipPath,
                      maxOperations: _limits.MaxContentOperations,
                      maxNestingDepth: _limits.MaxContentNestingDepth,
                      maxOperands: _limits.MaxContentOperands)) {
            Matrix2D invocationTransform = invocation.Transform;
            if (invocation.InlineImage != null) {
                placements.Add(BuildImagePlacement(
                    pageNumber,
                    invocation.InlineImage.ResourceName,
                    0,
                    invocation.InlineImage.DirectStreamIdentity,
                    invocationTransform,
                    invocation.ClipPath,
                    invocation.FillColor,
                    invocation.FillOpacity,
                    invocation.InlineImage.Stream,
                    resources,
                    invocation.PaintOrder));
                continue;
            }

            if (TryGetImageXObject(resources, invocation.Name, out int imageObjectNumber, out int directStreamIdentity)) {
                placements.Add(BuildImagePlacement(pageNumber, invocation.Name, imageObjectNumber, directStreamIdentity, invocationTransform, invocation.ClipPath, invocation.FillColor, invocation.FillOpacity, paintOrder: invocation.PaintOrder));
                continue;
            }

            if (!TryGetFormStream(resources, invocation.Name, out var formStream)) {
                continue;
            }

            if (!activeForms.Add(formStream)) {
                continue;
            }

            try {
                var formDict = formStream.Dictionary;
                var formResources = ResolveDictionary(formDict.Items.TryGetValue("Resources", out var resObj) ? resObj : null) ?? resources;
                Matrix2D formTransform = ApplyFormMatrix(invocationTransform, formDict);
                string formContent = WrapFormContentWithBoundingBoxClip(PdfEncoding.Latin1GetString(pageContentBudget.Decode(formStream)), formDict);
                CollectImagePlacementsAndForms(
                    formContent,
                    formResources,
                    pageNumber,
                    formTransform,
                    pageHeight,
                    placements,
                    activeForms,
                    invocation.FillColor,
                    invocation.FillColorSpace,
                    invocation.FillOpacity,
                    invocation.PaintOrder,
                    paintOrderScale * 0.000000001D,
                    initialClipPath: invocation.ClipPath,
                    contentNestingDepth: contentNestingDepth + 1,
                    pageContentBudget: pageContentBudget);
            } finally {
                activeForms.Remove(formStream);
            }
        }
    }

    private void EnsureContentNestingBudget(int contentNestingDepth) {
        if (contentNestingDepth > _limits.MaxContentNestingDepth) {
            throw PdfReadLimitException.Create(
                PdfReadLimitKind.ContentNestingDepth,
                _limits.MaxContentNestingDepth,
                contentNestingDepth);
        }
    }

    private bool TryGetFormStream(PdfDictionary? resources, string name, out PdfStream formStream) {
        if (resources is null || !resources.Items.TryGetValue("XObject", out var xoObj)) {
            formStream = null!;
            return false;
        }

        var xoDict = ResolveDictionary(xoObj);
        if (xoDict is null || !xoDict.Items.TryGetValue(name, out var formObj)) {
            formStream = null!;
            return false;
        }

        if (formObj is PdfReference formRef &&
            PdfObjectLookup.TryGet(_objects, formRef, out var indirectForm) &&
            indirectForm.Value is PdfStream stream &&
            string.Equals(stream.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal)) {
            formStream = stream;
            return true;
        }

        if (formObj is PdfStream directStream &&
            string.Equals(directStream.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal)) {
            formStream = directStream;
            return true;
        }

        formStream = null!;
        return false;
    }

    private bool TryGetImageXObject(PdfDictionary? resources, string name, out int objectNumber, out int directStreamIdentity) {
        objectNumber = 0;
        directStreamIdentity = 0;
        if (resources is null || !resources.Items.TryGetValue("XObject", out var xoObj)) {
            return false;
        }

        var xoDict = ResolveDictionary(xoObj);
        if (xoDict is null || !xoDict.Items.TryGetValue(name, out var imageObj)) {
            return false;
        }

        PdfStream? stream = null;
        if (imageObj is PdfReference imageRef &&
            PdfObjectLookup.TryGet(_objects, imageRef, out var indirectImage) &&
            indirectImage.Value is PdfStream referencedStream) {
            objectNumber = imageRef.ObjectNumber;
            stream = referencedStream;
        } else if (imageObj is PdfStream directStream) {
            stream = directStream;
            directStreamIdentity = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(directStream);
        }

        return stream is not null &&
            string.Equals(stream.Dictionary.Get<PdfName>("Subtype")?.Name, "Image", StringComparison.Ordinal);
    }

    private static PdfImagePlacement BuildImagePlacement(
        int pageNumber,
        string resourceName,
        int objectNumber,
        int directStreamIdentity,
        Matrix2D transform,
        PdfPageClipPath? clipPath,
        OfficeColor imageMaskColor,
        double? imageOpacity,
        PdfStream? inlineImageStream = null,
        PdfDictionary? inlineImageResources = null,
        double paintOrder = 0D) {
        var p0 = transform.Transform(0D, 0D);
        var p1 = transform.Transform(1D, 0D);
        var p2 = transform.Transform(0D, 1D);
        var p3 = transform.Transform(1D, 1D);
        double left = Math.Min(Math.Min(p0.X, p1.X), Math.Min(p2.X, p3.X));
        double right = Math.Max(Math.Max(p0.X, p1.X), Math.Max(p2.X, p3.X));
        double bottom = Math.Min(Math.Min(p0.Y, p1.Y), Math.Min(p2.Y, p3.Y));
        double top = Math.Max(Math.Max(p0.Y, p1.Y), Math.Max(p2.Y, p3.Y));

        return new PdfImagePlacement(
            pageNumber,
            resourceName,
            objectNumber,
            directStreamIdentity,
            transform.A,
            transform.B,
            transform.C,
            transform.D,
            transform.E,
            transform.F,
            left,
            bottom,
            Math.Max(0D, right - left),
            Math.Max(0D, top - bottom),
            clipPath,
            imageMaskColor,
            imageOpacity,
            inlineImageStream,
            inlineImageResources,
            paintOrder);
    }

    private byte[]? GetMarkedContentActualTextBytes(PdfDictionary? resources, string propertyName) {
        if (resources is null ||
            !resources.Items.TryGetValue("Properties", out var propertiesObj)) {
            return null;
        }

        var properties = ResolveDictionary(propertiesObj);
        if (properties is null ||
            !properties.Items.TryGetValue(propertyName, out var propertyObj)) {
            return null;
        }

        var propertyDictionary = ResolveDictionary(propertyObj);
        if (propertyDictionary is null ||
            !propertyDictionary.Items.TryGetValue("ActualText", out var actualTextObj) ||
            ResolveObject(actualTextObj) is not PdfStringObj actualText) {
            return null;
        }

        return actualText.RawBytes;
    }

    private PdfPageOptionalContentVisibility? GetOptionalContentVisibility(PdfDictionary? resources) =>
        PdfPageOptionalContentVisibility.Create(resources, _objects);

    private static Dictionary<string, Func<byte[], int, string>> MergeDecoders(
        Dictionary<string, Func<byte[], int, string>> parent,
        Dictionary<string, Func<byte[], int, string>> local) {
        var merged = new Dictionary<string, Func<byte[], int, string>>(parent, StringComparer.Ordinal);
        foreach (var entry in local) {
            merged[entry.Key] = entry.Value;
        }

        return merged;
    }

    private static Dictionary<string, Func<byte[], string>> MergeDecoders(
        Dictionary<string, Func<byte[], string>> parent,
        Dictionary<string, Func<byte[], string>> local) {
        var merged = new Dictionary<string, Func<byte[], string>>(parent, StringComparer.Ordinal);
        foreach (var entry in local) {
            merged[entry.Key] = entry.Value;
        }

        return merged;
    }

    private static Dictionary<string, Func<byte[], double>> MergeWidthProviders(
        Dictionary<string, Func<byte[], double>> parent,
        Dictionary<string, Func<byte[], double>> local) {
        var merged = new Dictionary<string, Func<byte[], double>>(parent, StringComparer.Ordinal);
        foreach (var entry in local) {
            merged[entry.Key] = entry.Value;
        }

        return merged;
    }

    private static Dictionary<string, PdfFontResource> MergeFonts(
        Dictionary<string, PdfFontResource> parent,
        Dictionary<string, PdfFontResource> local) {
        var merged = new Dictionary<string, PdfFontResource>(parent, StringComparer.Ordinal);
        foreach (var entry in local) {
            merged[entry.Key] = entry.Value;
        }

        return merged;
    }

    private static string WrapContentWithTransform(string content, Matrix2D transform) => WrapContentWithTransform(content, transform, out _);

    private static string WrapContentWithTransform(string content, Matrix2D transform, out int contentOffset) {
        string prefix = string.Format(
            System.Globalization.CultureInfo.InvariantCulture,
            "q {0} {1} {2} {3} {4} {5} cm ",
            transform.A,
            transform.B,
            transform.C,
            transform.D,
            transform.E,
            transform.F);
        contentOffset = prefix.Length;
        return prefix + content + " Q";
    }

    private string WrapFormContentWithBoundingBoxClip(string content, PdfDictionary? formDict) {
        if (formDict is null ||
            !TryReadBox(formDict.Items.TryGetValue("BBox", out PdfObject? bboxObject) ? bboxObject : null, out (double X1, double Y1, double X2, double Y2) bbox)) {
            return content;
        }

        double width = bbox.X2 - bbox.X1;
        double height = bbox.Y2 - bbox.Y1;
        if (width <= 0D || height <= 0D) {
            return content;
        }

        string prefix = string.Format(
            System.Globalization.CultureInfo.InvariantCulture,
            "q {0} {1} {2} {3} re W n ",
            bbox.X1,
            bbox.Y1,
            width,
            height);
        return prefix + content + " Q";
    }

    private static Matrix2D ApplyFormMatrix(Matrix2D invocationTransform, PdfDictionary? formDict) {
        if (formDict is null ||
            !formDict.Items.TryGetValue("Matrix", out var matrixObj) ||
            matrixObj is not PdfArray arr ||
            arr.Items.Count < 6) {
            return invocationTransform;
        }

        var formMatrix = new Matrix2D(
            (arr.Items[0] as PdfNumber)?.Value ?? 1,
            (arr.Items[1] as PdfNumber)?.Value ?? 0,
            (arr.Items[2] as PdfNumber)?.Value ?? 0,
            (arr.Items[3] as PdfNumber)?.Value ?? 1,
            (arr.Items[4] as PdfNumber)?.Value ?? 0,
            (arr.Items[5] as PdfNumber)?.Value ?? 0);

        return Matrix2D.Multiply(invocationTransform, formMatrix);
    }

    private PdfObject? GetInheritedValue(string key) {
        PdfDictionary? current = _pageDict;
        int guard = 0;
        while (current is not null && guard++ < 100) {
            if (current.Items.TryGetValue(key, out var value)) {
                return value;
            }

            if (!current.Items.TryGetValue("Parent", out var parentObj) ||
                parentObj is not PdfReference parentRef ||
                !PdfObjectLookup.TryGet(_objects, parentRef, out var parentIndirect) ||
                parentIndirect.Value is not PdfDictionary parentDict) {
                break;
            }

            current = parentDict;
        }

        return null;
    }

    private PdfDictionary? ResolveDictionary(PdfObject? obj) {
        if (obj is PdfDictionary dictionary) {
            return dictionary;
        }

        if (obj is PdfReference reference &&
            PdfObjectLookup.TryGet(_objects, reference, out var indirect) &&
            indirect.Value is PdfDictionary referencedDictionary) {
            return referencedDictionary;
        }

        return null;
    }

    private PdfObject? ResolveObject(PdfObject? obj) {
        return PdfObjectLookup.Resolve(_objects, obj);
    }

    private PdfArray? ResolveArray(PdfObject? obj) {
        if (obj is PdfArray array) {
            return array;
        }

        if (obj is PdfReference reference &&
            PdfObjectLookup.TryGet(_objects, reference, out var indirect) &&
            indirect.Value is PdfArray referencedArray) {
            return referencedArray;
        }

        return null;
    }

    private bool TryGetString(PdfObject? obj, out string? value) {
        if (ResolveObject(obj) is PdfStringObj text) {
            value = text.Value;
            return true;
        }

        value = null;
        return false;
    }

    private bool TryGetDestinationName(PdfObject? obj, out string? value) {
        return TryGetNameOrString(obj, out value);
    }

    private bool TryGetNameOrString(PdfObject? obj, out string? value) {
        switch (ResolveObject(obj)) {
            case PdfStringObj text when !string.IsNullOrEmpty(text.Value):
                value = text.Value;
                return true;
            case PdfName name when !string.IsNullOrEmpty(name.Name):
                value = name.Name;
                return true;
            default:
                value = null;
                return false;
        }
    }

    private string? TryReadActionType(PdfObject? obj) {
        var action = ResolveDictionary(obj);
        string? actionType = action?.Get<PdfName>("S")?.Name;
        return string.IsNullOrEmpty(actionType) ? null : actionType;
    }

    private int? TryReadInteger(PdfObject? obj) {
        if (ResolveObject(obj) is PdfNumber number &&
            number.Value >= int.MinValue &&
            number.Value <= int.MaxValue &&
            Math.Abs(number.Value - Math.Truncate(number.Value)) < double.Epsilon) {
            return (int)number.Value;
        }

        return null;
    }

    private void ReadFreeTextAppearanceMetadata(
        PdfDictionary annotation,
        string subtype,
        out string? defaultAppearance,
        out string? defaultStyle,
        out string? richContents,
        out string? richContentsPlainText,
        out double? effectiveFontSize,
        out PdfColor? effectiveTextColor,
        out PdfAlign? effectiveTextAlign) {
        defaultAppearance = null;
        defaultStyle = null;
        richContents = null;
        richContentsPlainText = null;
        effectiveFontSize = null;
        effectiveTextColor = null;
        effectiveTextAlign = null;
        if (!string.Equals(subtype, "FreeText", StringComparison.Ordinal)) {
            return;
        }

        TryGetString(annotation.Items.TryGetValue("DA", out PdfObject? defaultAppearanceObject) ? defaultAppearanceObject : null, out defaultAppearance);
        TryGetString(annotation.Items.TryGetValue("DS", out PdfObject? defaultStyleObject) ? defaultStyleObject : null, out defaultStyle);
        TryGetString(annotation.Items.TryGetValue("RC", out PdfObject? richContentsObject) ? richContentsObject : null, out richContents);
        richContentsPlainText = PdfFreeTextStyleParser.ExtractPlainText(richContents);
        PdfFreeTextDefaultStyle parsedDefaultStyle = PdfFreeTextStyleParser.ParseDefaultStyle(defaultStyle);
        effectiveFontSize = PdfDefaultAppearanceParser.TryReadFontSize(defaultAppearance, out double defaultAppearanceFontSize)
            ? defaultAppearanceFontSize
            : parsedDefaultStyle.FontSize;
        effectiveTextColor = PdfDefaultAppearanceParser.TryReadTextColor(defaultAppearance, out PdfColor defaultAppearanceTextColor)
            ? defaultAppearanceTextColor
            : parsedDefaultStyle.TextColor;
        effectiveTextAlign = TryReadFreeTextAlignment(annotation, parsedDefaultStyle.TextAlign);
    }

    private PdfAlign? TryReadFreeTextAlignment(PdfDictionary annotation, PdfAlign? defaultAlignment) {
        int? alignment = TryReadInteger(annotation.Items.TryGetValue("Q", out PdfObject? alignmentObject) ? alignmentObject : null);
        if (!alignment.HasValue) {
            return defaultAlignment;
        }

        return alignment.Value == 1
            ? PdfAlign.Center
            : alignment.Value == 2
                ? PdfAlign.Right
                : PdfAlign.Left;
    }

    private IReadOnlyList<double> ReadNumberArray(PdfObject? obj) {
        PdfArray? array = ResolveArray(obj);
        if (array is null || array.Items.Count == 0) {
            return Array.Empty<double>();
        }

        var values = new List<double>();
        for (int i = 0; i < array.Items.Count; i++) {
            if (ResolveObject(array.Items[i]) is PdfNumber number) {
                values.Add(number.Value);
            }
        }

        return values.Count == 0 ? Array.Empty<double>() : values.AsReadOnly();
    }

    private IReadOnlyList<PdfAnnotationAdditionalAction> ReadAdditionalActions(PdfObject? obj) {
        var additionalActions = ResolveDictionary(obj);
        if (additionalActions is null || additionalActions.Items.Count == 0) {
            return Array.Empty<PdfAnnotationAdditionalAction>();
        }

        var actions = new List<PdfAnnotationAdditionalAction>();
        foreach (var item in additionalActions.Items) {
            if (string.IsNullOrEmpty(item.Key)) {
                continue;
            }

            string? actionType = TryReadActionType(item.Value);
            if (!string.IsNullOrEmpty(actionType)) {
                actions.Add(new PdfAnnotationAdditionalAction(item.Key, actionType!));
            }
        }

        return actions.Count == 0 ? Array.Empty<PdfAnnotationAdditionalAction>() : actions.AsReadOnly();
    }

    private bool TryReadFileSpecification(PdfObject? obj, out string? file) {
        PdfObject? resolved = ResolveObject(obj);
        if (resolved is PdfStringObj text && !string.IsNullOrEmpty(text.Value)) {
            file = text.Value;
            return true;
        }

        if (resolved is PdfDictionary dictionary) {
            if (TryGetString(dictionary.Items.TryGetValue("UF", out var unicodeFileObject) ? unicodeFileObject : null, out string? unicodeFile) &&
                !string.IsNullOrEmpty(unicodeFile)) {
                file = unicodeFile;
                return true;
            }

            if (TryGetString(dictionary.Items.TryGetValue("F", out var fileObject) ? fileObject : null, out string? fallbackFile) &&
                !string.IsNullOrEmpty(fallbackFile)) {
                file = fallbackFile;
                return true;
            }
        }

        file = null;
        return false;
    }

    private bool TryReadRemoteDestination(
        PdfObject? obj,
        out string? destinationName,
        out int? destinationPageNumber,
        out double? destinationTop,
        out PdfOpenActionDestinationMode? destinationMode,
        out double? destinationLeft,
        out double? destinationBottom,
        out double? destinationRight) {
        if (TryGetDestinationName(obj, out destinationName)) {
            destinationPageNumber = null;
            destinationTop = null;
            destinationMode = null;
            destinationLeft = null;
            destinationBottom = null;
            destinationRight = null;
            return true;
        }

        destinationName = null;
        destinationPageNumber = null;
        destinationTop = null;
        destinationMode = null;
        destinationLeft = null;
        destinationBottom = null;
        destinationRight = null;

        PdfObject? resolved = ResolveObject(obj);
        if (resolved is PdfDictionary dictionary &&
            dictionary.Items.TryGetValue("D", out var explicitDestination)) {
            resolved = ResolveObject(explicitDestination);
        }

        if (resolved is not PdfArray destination || destination.Items.Count < 2) {
            return false;
        }

        if (ResolveObject(destination.Items[0]) is PdfNumber pageIndex &&
            pageIndex.Value >= 0 &&
            pageIndex.Value < int.MaxValue &&
            Math.Abs(pageIndex.Value - Math.Truncate(pageIndex.Value)) < double.Epsilon) {
            destinationPageNumber = (int)pageIndex.Value + 1;
        }

        ReadDestinationCoordinates(destination, out destinationTop, out destinationMode, out destinationLeft, out destinationBottom, out destinationRight);
        return destinationPageNumber.HasValue || destinationTop.HasValue || destinationMode.HasValue || destinationLeft.HasValue || destinationBottom.HasValue || destinationRight.HasValue;
    }

    private bool TryReadLinkDestination(
        PdfObject? obj,
        out string? destinationName,
        out int? destinationPageObjectNumber,
        out double? destinationTop,
        out PdfOpenActionDestinationMode? destinationMode,
        out double? destinationLeft,
        out double? destinationBottom,
        out double? destinationRight) {
        if (TryGetDestinationName(obj, out destinationName)) {
            destinationPageObjectNumber = null;
            destinationTop = null;
            destinationMode = null;
            destinationLeft = null;
            destinationBottom = null;
            destinationRight = null;
            return true;
        }

        destinationPageObjectNumber = null;
        destinationTop = null;
        destinationMode = null;
        destinationLeft = null;
        destinationBottom = null;
        destinationRight = null;

        PdfObject? resolved = ResolveObject(obj);
        if (resolved is PdfDictionary dictionary &&
            dictionary.Items.TryGetValue("D", out var explicitDestination)) {
            resolved = ResolveObject(explicitDestination);
        }

        if (resolved is not PdfArray destination || destination.Items.Count < 2) {
            return false;
        }

        if (destination.Items[0] is PdfReference pageReference) {
            destinationPageObjectNumber = pageReference.ObjectNumber;
        }

        ReadDestinationCoordinates(destination, out destinationTop, out destinationMode, out destinationLeft, out destinationBottom, out destinationRight);
        return destinationPageObjectNumber.HasValue || destinationTop.HasValue || destinationMode.HasValue || destinationLeft.HasValue || destinationBottom.HasValue || destinationRight.HasValue;
    }

    private void ReadDestinationCoordinates(
        PdfArray destination,
        out double? destinationTop,
        out PdfOpenActionDestinationMode? destinationMode,
        out double? destinationLeft,
        out double? destinationBottom,
        out double? destinationRight) {
        destinationTop = null;
        destinationMode = null;
        destinationLeft = null;
        destinationBottom = null;
        destinationRight = null;

        if (ResolveObject(destination.Items[1]) is PdfName fitName) {
            switch (fitName.Name) {
                case "XYZ":
                    destinationMode = PdfOpenActionDestinationMode.Xyz;
                    if (destination.Items.Count > 2 && ResolveObject(destination.Items[2]) is PdfNumber xyzLeft) {
                        destinationLeft = xyzLeft.Value;
                    }

                    if (destination.Items.Count > 3 && ResolveObject(destination.Items[3]) is PdfNumber xyzTop) {
                        destinationTop = xyzTop.Value;
                    }

                    break;
                case "Fit":
                    destinationMode = PdfOpenActionDestinationMode.Fit;
                    break;
                case "FitH":
                    destinationMode = PdfOpenActionDestinationMode.FitHorizontal;
                    if (destination.Items.Count > 2 && ResolveObject(destination.Items[2]) is PdfNumber fitTop) {
                        destinationTop = fitTop.Value;
                    }

                    break;
                case "FitV":
                    destinationMode = PdfOpenActionDestinationMode.FitVertical;
                    if (destination.Items.Count > 2 && ResolveObject(destination.Items[2]) is PdfNumber fitLeft) {
                        destinationLeft = fitLeft.Value;
                    }

                    break;
                case "FitR":
                    destinationMode = PdfOpenActionDestinationMode.FitRectangle;
                    if (destination.Items.Count > 5) {
                        if (ResolveObject(destination.Items[2]) is PdfNumber left) {
                            destinationLeft = left.Value;
                        }

                        if (ResolveObject(destination.Items[3]) is PdfNumber bottom) {
                            destinationBottom = bottom.Value;
                        }

                        if (ResolveObject(destination.Items[4]) is PdfNumber right) {
                            destinationRight = right.Value;
                        }

                        if (ResolveObject(destination.Items[5]) is PdfNumber top) {
                            destinationTop = top.Value;
                        }
                    }

                    break;
                case "FitB":
                    destinationMode = PdfOpenActionDestinationMode.FitBoundingBox;
                    break;
                case "FitBH":
                    destinationMode = PdfOpenActionDestinationMode.FitBoundingBoxHorizontal;
                    if (destination.Items.Count > 2 && ResolveObject(destination.Items[2]) is PdfNumber fitBoundingTop) {
                        destinationTop = fitBoundingTop.Value;
                    }

                    break;
                case "FitBV":
                    destinationMode = PdfOpenActionDestinationMode.FitBoundingBoxVertical;
                    if (destination.Items.Count > 2 && ResolveObject(destination.Items[2]) is PdfNumber fitBoundingLeft) {
                        destinationLeft = fitBoundingLeft.Value;
                    }

                    break;
                default:
                    if (destination.Items.Count > 3 && ResolveObject(destination.Items[3]) is PdfNumber fallbackTop) {
                        destinationTop = fallbackTop.Value;
                    }

                    break;
            }
        }
    }

    private bool HasNormalAppearance(PdfDictionary annotation) {
        var appearance = ResolveDictionary(annotation.Items.TryGetValue("AP", out var appearanceObject) ? appearanceObject : null);
        return appearance != null && appearance.Items.ContainsKey("N");
    }

    private bool TryReadRectangle(PdfObject? obj, out (double X1, double Y1, double X2, double Y2) rect) {
        rect = default;
        var array = ResolveArray(obj);
        if (array is null || array.Items.Count < 4) {
            return false;
        }

        if (ResolveObject(array.Items[0]) is not PdfNumber x1 ||
            ResolveObject(array.Items[1]) is not PdfNumber y1 ||
            ResolveObject(array.Items[2]) is not PdfNumber x2 ||
            ResolveObject(array.Items[3]) is not PdfNumber y2) {
            return false;
        }

        double left = Math.Min(x1.Value, x2.Value);
        double right = Math.Max(x1.Value, x2.Value);
        double bottom = Math.Min(y1.Value, y2.Value);
        double top = Math.Max(y1.Value, y2.Value);
        if (double.IsNaN(left) || double.IsInfinity(left) ||
            double.IsNaN(right) || double.IsInfinity(right) ||
            double.IsNaN(bottom) || double.IsInfinity(bottom) ||
            double.IsNaN(top) || double.IsInfinity(top) ||
            right <= left ||
            top <= bottom) {
            return false;
        }

        rect = (left, bottom, right, top);
        return true;
    }

    private bool TryParseBox(PdfObject? box, out (double Width, double Height) size) {
        var arr = ResolveArray(box);
        if (arr is not null &&
            arr.Items.Count >= 4 &&
            arr.Items[0] is PdfNumber llx &&
            arr.Items[1] is PdfNumber lly &&
            arr.Items[2] is PdfNumber urx &&
            arr.Items[3] is PdfNumber ury) {
            double width = urx.Value - llx.Value;
            double height = ury.Value - lly.Value;
            if (width > 0 && height > 0) {
                size = (width, height);
                return true;
            }
        }

        size = default;
        return false;
    }

    private static double GlyphWidthEmForBase(string baseFont) {
        if (string.IsNullOrEmpty(baseFont)) return 0.55;
        if (ContainsIgnoreCase(baseFont, "courier")) return 0.6;
        if (ContainsIgnoreCase(baseFont, "times")) return 0.5;
        if (ContainsIgnoreCase(baseFont, "helvetica")) return 0.55;
        return 0.55;
    }

    private static bool ContainsIgnoreCase(string source, string value) {
#if NET8_0_OR_GREATER
        return source.Contains(value, System.StringComparison.OrdinalIgnoreCase);
#else
        return source.IndexOf(value, System.StringComparison.OrdinalIgnoreCase) >= 0;
#endif
    }

    /// <summary>
    /// Returns decoded page content with stream arrays concatenated in PDF processing order.
    /// </summary>
    private string GetContentStreamContent(PageContentBudget? pageContentBudget = null) {
        pageContentBudget ??= new PageContentBudget(this);
        var builder = new System.Text.StringBuilder();
        foreach (var stream in GetContentStreamObjects()) {
            builder.Append(PdfEncoding.Latin1GetString(pageContentBudget.Decode(stream)));
        }

        return builder.ToString();
    }

    private List<PdfStream> GetContentStreamObjects() {
        var result = new List<PdfStream>();
        var contents = _pageDict.Items.TryGetValue("Contents", out var obj) ? obj : null;
        if (contents is PdfReference r) {
            if (PdfObjectLookup.TryGet(_objects, r, out var ind) && ind.Value is PdfStream s) {
                result.Add(s);
                return result;
            }
        }

        var contentArray = ResolveArray(contents);
        if (contentArray is null) {
            return result;
        }

        foreach (var item in contentArray.Items) {
            if (item is PdfReference rr &&
                PdfObjectLookup.TryGet(_objects, rr, out var ind2) &&
                ind2.Value is PdfStream s2) {
                result.Add(s2);
            } else if (item is PdfStream directStream) {
                result.Add(directStream);
            }
        }

        return result;
    }

    private static bool ContainsFilter(List<string> filters, string filterName) {
        for (int i = 0; i < filters.Count; i++) {
            if (string.Equals(filters[i], filterName, StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
    }

    private byte[] DecodeIfNeeded(PdfStream s, int maxDecodedBytes) {
        return Filters.StreamDecoder.Decode(s.Dictionary, s.Data, _objects, maxDecodedBytes);
    }

    private sealed class PageContentBudget {
        private readonly PdfReadPage _page;
        private readonly Dictionary<PdfStream, byte[]> _decodedStreams = new();
        private long _decodedBytes;

        internal PageContentBudget(PdfReadPage page) {
            _page = page;
        }

        internal byte[] Decode(PdfStream stream) {
            if (_decodedStreams.TryGetValue(stream, out byte[]? cached)) {
                Charge(cached.LongLength);
                return cached;
            }

            long remainingPageBytes = (long)_page._limits.MaxPageContentBytes - _decodedBytes;
            if (remainingPageBytes <= 0L) {
                throw PdfReadLimitException.Create(
                    PdfReadLimitKind.PageContentBytes,
                    _page._limits.MaxPageContentBytes,
                    (long)_page._limits.MaxPageContentBytes + 1L);
            }

            int streamDecodeLimit = (int)Math.Min(_page._maxDecodedStreamBytes, remainingPageBytes);
            byte[] decoded;
            try {
                decoded = _page.DecodeIfNeeded(stream, streamDecodeLimit);
            } catch (PdfReadLimitException exception) when (
                exception.Kind == PdfReadLimitKind.DecodedStreamBytes &&
                remainingPageBytes <= _page._maxDecodedStreamBytes) {
                throw PdfReadLimitException.Create(
                    PdfReadLimitKind.PageContentBytes,
                    _page._limits.MaxPageContentBytes,
                    (long)_page._limits.MaxPageContentBytes + 1L);
            }

            Charge(decoded.LongLength);
            _decodedStreams[stream] = decoded;
            return decoded;
        }

        private void Charge(long decodedBytes) {
            _decodedBytes += decodedBytes;
            if (_decodedBytes > _page._limits.MaxPageContentBytes) {
                throw PdfReadLimitException.Create(
                    PdfReadLimitKind.PageContentBytes,
                    _page._limits.MaxPageContentBytes,
                    _decodedBytes);
            }
        }
    }
}
