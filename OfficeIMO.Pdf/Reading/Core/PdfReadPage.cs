using OfficeIMO.Drawing;
using System.IO;

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
    private readonly PdfFontResourceCache _fontResourceCache;
    private readonly bool _includeArtifactText;
    private readonly Action? _demandTextExtraction;
    private readonly Action<string>? _demandContentExtraction;
    private readonly PdfOutputIntentColorTransform? _outputIntentColorTransform;
    private readonly Lazy<bool>? _hasOutputIntentCompositionInteraction;

    internal PdfDictionary PageDictionary => _pageDict;

    internal PdfReadPage(int objectNumber, PdfDictionary pageDict, Dictionary<int, PdfIndirectObject> objects)
        : this(objectNumber, pageDict, objects, new PdfReadLimits(), new PdfFontResourceCache()) { }

    internal PdfReadPage(
        int objectNumber,
        PdfDictionary pageDict,
        Dictionary<int, PdfIndirectObject> objects,
        PdfReadLimits limits,
        PdfFontResourceCache fontResourceCache,
        Action? demandTextExtraction = null,
        Action<string>? demandContentExtraction = null,
        bool includeArtifactText = false,
        PdfOutputIntentColorTransform? outputIntentColorTransform = null) {
        ObjectNumber = objectNumber;
        _pageDict = pageDict;
        _objects = objects;
        _limits = limits;
        _fontResourceCache = fontResourceCache;
        _includeArtifactText = includeArtifactText;
        _maxDecodedStreamBytes = limits.MaxDecodedStreamBytes;
        _demandTextExtraction = demandTextExtraction;
        _demandContentExtraction = demandContentExtraction;
        _outputIntentColorTransform = outputIntentColorTransform;
        _hasOutputIntentCompositionInteraction = outputIntentColorTransform == null
            ? null
            : new Lazy<bool>(
                HasOutputIntentCompositionInteraction,
                System.Threading.LazyThreadSafetyMode.ExecutionAndPublication);
    }

    private PdfOutputIntentColorTransform? EffectiveOutputIntentColorTransform =>
        _outputIntentColorTransform != null && _hasOutputIntentCompositionInteraction?.Value != true
            ? _outputIntentColorTransform
            : null;

    internal bool HasEffectiveOutputIntentColorTransform =>
        EffectiveOutputIntentColorTransform?.IsSupported == true;

    /// <summary>Underlying object number for the page.</summary>
    public int ObjectNumber { get; }

    /// <summary>Extracts plain text from this page without column reordering.</summary>
    public string ExtractText() => ExtractText(System.Threading.CancellationToken.None);

    internal string ExtractText(System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var spans = GetTextSpans(cancellationToken);
        var opts = new TextLayoutEngine.Options { ForceSingleColumn = true };
        var lines = TextLayoutEngine.BuildLines(
            spans,
            opts,
            consumeWork: null,
            cancellationCheck: cancellationToken.ThrowIfCancellationRequested);
        cancellationToken.ThrowIfCancellationRequested();
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

    internal (double Width, double Height) GetVisualPageSize() {
        PdfPageBox pageBox = GetPageBoundaryBox();
        return PdfVisualCoordinateMapper.GetVisualSize(pageBox, GetRotationDegrees(), GetEffectiveUserUnit());
    }

    internal Matrix2D GetVisualPageTransform() =>
        PdfVisualCoordinateMapper.CreateTransform(GetPageBoundaryBox(), GetRotationDegrees(), GetEffectiveUserUnit());

    internal PdfVisualBounds TransformBoundsToVisual(double left, double bottom, double right, double top) =>
        PdfVisualCoordinateMapper.TransformBounds(GetPageBoundaryBox(), GetRotationDegrees(), left, bottom, right, top, GetEffectiveUserUnit());

    internal PdfVisualBounds TransformVisualBoundsToUser(double left, double top, double right, double bottom) =>
        PdfVisualCoordinateMapper.TransformVisualBoundsToUser(GetPageBoundaryBox(), GetRotationDegrees(), left, top, right, bottom, GetEffectiveUserUnit());

    private double GetEffectiveUserUnit() => TryReadDirectPositiveNumber("UserUnit") ?? 1D;

    internal (double Width, double Height) GetInteractionPageSize() => GetVisualPageSize();

    internal (double X, double Y) TransformPointToVisual(double x, double y) => GetVisualPageTransform().Transform(x, y);

    internal IReadOnlyList<PdfTextSpan> GetInteractionTextSpans() {
        _demandTextExtraction?.Invoke();
        (double Width, double Height) size = GetVisualPageSize();
        return GetVisualTextSpans(
            size.Height,
            GetVisualPageTransform(),
            useLogicalTextFilters: true,
            includeArtifactText: true);
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
        return GetTextSpans(_includeArtifactText, default);
    }

    internal IReadOnlyList<PdfTextSpan> GetTextSpans(bool includeArtifactText) {
        return GetTextSpans(includeArtifactText, default);
    }

    internal IReadOnlyList<PdfTextSpan> GetTextSpans(System.Threading.CancellationToken cancellationToken) {
        return GetTextSpans(_includeArtifactText, cancellationToken);
    }

    internal IReadOnlyList<PdfTextSpan> GetTextSpans(
        bool includeArtifactText,
        System.Threading.CancellationToken cancellationToken,
        bool includeHiddenOptionalContent = false) {
        cancellationToken.ThrowIfCancellationRequested();
        _demandTextExtraction?.Invoke();
        var spans = new List<PdfTextSpan>();
        var pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        PdfFontResourceSet pageFontResources = _fontResourceCache.GetOrCreate(pageResources, _objects);
        Dictionary<string, Func<byte[], int, string>> pageDecoders = pageFontResources.Decoders;
        Dictionary<string, Func<byte[], double>> pageWidthProviders = pageFontResources.WidthProviders;
        Dictionary<string, PdfFontResource> pageFonts = pageFontResources.Fonts;
        var activeForms = new HashSet<PdfStream>();
        double pageHeight = GetPageSize().Height;
        var pageContentBudget = new PageContentBudget(this);

        PageContentStreamSequence contentSequence = GetContentStreamSequence(pageContentBudget);
        string content = contentSequence.Content;
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
                includeArtifactText: includeArtifactText,
                includeHiddenOptionalContent: includeHiddenOptionalContent,
                pageContentBudget: pageContentBudget,
                contentOrderPrefix: PdfContentOrderKey.Root,
                contentStreamObjectNumberAtOffset: contentSequence.GetObjectNumber,
                cancellationCheck: cancellationToken.CanBeCanceled
                    ? cancellationToken.ThrowIfCancellationRequested
                    : null);
        }

        return spans;
    }

    internal (double X, double Y) GetPageBoundaryOrigin() {
        PdfPageBox box = GetPageBoundaryBox();
        return (box.Left, box.Bottom);
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

    internal IReadOnlyList<PdfAnnotation> GetAnnotationsUnchecked() => GetAnnotationsUnchecked(includeUnreadableRectangles: false);

    internal IReadOnlyList<PdfAnnotation> GetAnnotationsForContentSafety() => GetAnnotationsUnchecked(includeUnreadableRectangles: true);

    private IReadOnlyList<PdfAnnotation> GetAnnotationsUnchecked(bool includeUnreadableRectangles) {
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
                string.IsNullOrWhiteSpace(subtype)) {
                continue;
            }
            bool hasReadableRectangle = TryReadAnnotationRectangle(
                annotation.Items.TryGetValue("Rect", out var rectObject) ? rectObject : null,
                out var rect);
            if (!hasReadableRectangle && !includeUnreadableRectangles) continue;

            TryGetString(annotation.Items.TryGetValue("Contents", out var contentsObject) ? contentsObject : null, out string? contents);
            bool hasNormalAppearance = HasNormalAppearance(annotation);
            PdfDictionary? appearances = ResolveDictionary(annotation.Items.TryGetValue("AP", out PdfObject? appearancesObject) ? appearancesObject : null);
            PdfObject? normalAppearanceObject = appearances != null && appearances.Items.TryGetValue("N", out PdfObject? normalAppearance)
                ? normalAppearance
                : null;
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
            ReadAnnotationAppearanceMetadata(
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
            PdfAnnotationReviewInfo? review = ReadAnnotationReviewInfo(annotation);
            string? appearanceState = annotation.Get<PdfName>("AS")?.Name;
            result.Add(new PdfAnnotation(objectNumber, null, subtype!, contents, rect.X1, rect.Y1, rect.X2, rect.Y2, hasNormalAppearance, actionType, additionalActions, chainedActions, flags, name, title, modified, color, defaultAppearance, defaultStyle, richContents, richContentsPlainText, effectiveFontSize, effectiveTextColor, effectiveTextAlign, interiorColor, opacity, borderWidth, borderStyle, borderDashPattern, borderEffectStyle, borderEffectIntensity, rectangleDifferences, calloutLine, calloutLineEnding, lineStartEnding, lineEndEnding, quadPoints, lineCoordinates, vertices, inkList, review, normalAppearanceObject, appearanceState, annotation, hasReadableRectangle));
        }

        return result.Count == 0 ? Array.Empty<PdfAnnotation>() : result.AsReadOnly();
    }

    private bool TryReadAnnotationRectangle(PdfObject? obj, out (double X1, double Y1, double X2, double Y2) rect) {
        rect = default;
        var array = ResolveArray(obj);
        if (array is null || array.Items.Count < 4 ||
            ResolveObject(array.Items[0]) is not PdfNumber x1 ||
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
            double.IsNaN(top) || double.IsInfinity(top)) {
            return false;
        }

        rect = (left, bottom, right, top);
        return true;
    }

    private PdfAnnotationReviewInfo? ReadAnnotationReviewInfo(PdfDictionary annotation) {
        int? inReplyToObjectNumber = annotation.Items.TryGetValue("IRT", out PdfObject? replyObject) && replyObject is PdfReference replyReference
            ? replyReference.ObjectNumber
            : null;
        string? replyType = annotation.Get<PdfName>("RT")?.Name;
        string? state = annotation.Get<PdfName>("State")?.Name;
        string? stateModel = annotation.Get<PdfName>("StateModel")?.Name;
        TryGetString(annotation.Items.TryGetValue("Subj", out PdfObject? subjectObject) ? subjectObject : null, out string? subject);
        string? intent = annotation.Get<PdfName>("IT")?.Name;
        if (!inReplyToObjectNumber.HasValue && replyType is null && state is null && stateModel is null && subject is null && intent is null) {
            return null;
        }

        return new PdfAnnotationReviewInfo(inReplyToObjectNumber, replyType, state, stateModel, subject, intent);
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
        return GetImages(pageNumber, imagePlacements, colorizeImageMasks, new PageContentBudget(this));
    }

    private IReadOnlyList<PdfExtractedImage> GetImages(
        int pageNumber,
        IReadOnlyList<PdfImagePlacement>? imagePlacements,
        bool colorizeImageMasks,
        PageContentBudget pageContentBudget) {
        return GetImagesForResources(
            ResolveDictionary(GetInheritedValue("Resources")),
            pageNumber,
            imagePlacements,
            colorizeImageMasks,
            pageContentBudget);
    }

    private IReadOnlyList<PdfExtractedImage> GetImagesForResources(
        PdfDictionary? resources,
        int pageNumber,
        IReadOnlyList<PdfImagePlacement>? imagePlacements,
        bool colorizeImageMasks = false,
        PageContentBudget? pageContentBudget = null) {
        var images = resources == null
            ? new List<PdfExtractedImage>()
            : new List<PdfExtractedImage>(ResourceResolver.GetImageXObjectsForResources(
                resources,
                _objects,
                pageNumber,
                imagePlacements,
                colorizeImageMasks,
                _limits,
                EffectiveOutputIntentColorTransform,
                pageContentBudget == null ? null : pageContentBudget.TryConsumeColorFunctionEvaluations,
                pageContentBudget?.ColorFunctionResolutionContext));
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
                    colorizeImageMasks,
                    _limits.MaxDecodedStreamBytes,
                    placement.RenderingIntent,
                    EffectiveOutputIntentColorTransform,
                    pageContentBudget == null ? null : pageContentBudget.TryConsumeColorFunctionEvaluations,
                    pageContentBudget?.ColorFunctionResolutionContext,
                    inheritedHasAuthoredRenderingIntent: placement.HasAuthoredRenderingIntent));
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
        return GetImagePlacements(pageNumber, includeHiddenOptionalContent: false);
    }

    internal IReadOnlyList<PdfImagePlacement> GetImagePlacementsIncludingHiddenOptionalContent(int pageNumber) {
        IReadOnlyList<PdfImagePlacement> visible = GetImagePlacements(pageNumber, includeHiddenOptionalContent: false);
        IReadOnlyList<PdfImagePlacement> all = GetImagePlacements(pageNumber, includeHiddenOptionalContent: true);
        if (all.Count == visible.Count) return all;

        var visibleContentOrderKeys = new HashSet<PdfContentOrderKey>(visible
            .Select(static placement => placement.ContentOrderKey)
            .OfType<PdfContentOrderKey>());
        return all
            .Select(placement => placement.ContentOrderKey is not null && visibleContentOrderKeys.Contains(placement.ContentOrderKey)
                ? placement
                : placement.WithHiddenOptionalContent(true))
            .ToArray();
    }

    private IReadOnlyList<PdfImagePlacement> GetImagePlacements(int pageNumber, bool includeHiddenOptionalContent) {
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
                includeHiddenOptionalContent: includeHiddenOptionalContent,
                pageContentBudget: pageContentBudget,
                contentOrderPrefix: PdfContentOrderKey.Root);
        }

        return placements.Count == 0 ? Array.Empty<PdfImagePlacement>() : placements.AsReadOnly();
    }

    internal List<string> GetUnsupportedContentStreamFilters() {
        var unsupported = new List<string>();
        var pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        var activeForms = new HashSet<PdfStream>();
        var pageContentBudget = new PageContentBudget(this);
        bool mayContainForms = MayContainFormXObjects(pageResources);
        System.Text.StringBuilder? content = mayContainForms ? new System.Text.StringBuilder() : null;
        bool canInspectFormInvocations = mayContainForms;
        foreach (PageContentStreamEntry entry in GetContentStreamObjects()) {
            PdfStream stream = entry.Stream;
            AddUnsupportedFilters(stream, unsupported);
            if (Filters.StreamDecoder.GetUnsupportedFilters(stream.Dictionary, _objects).Count != 0) {
                canInspectFormInvocations = false;
                continue;
            }

            byte[] decoded = pageContentBudget.Decode(stream);
            if (content is not null) {
                content.Append(PdfEncoding.Latin1GetString(decoded));
            }
        }

        if (canInspectFormInvocations && content is { Length: > 0 }) {
            CollectUnsupportedFormFilters(content.ToString(), pageResources, unsupported, activeForms, pageContentBudget);
        }

        return unsupported;
    }

    private bool MayContainFormXObjects(PdfDictionary? resources) {
        if (resources is null || !resources.Items.TryGetValue("XObject", out var xObjectValue)) {
            return false;
        }

        var xObjects = ResolveDictionary(xObjectValue);
        if (xObjects is null) {
            return true;
        }

        foreach (var value in xObjects.Items.Values) {
            PdfStream? stream;
            if (value is PdfReference reference) {
                if (!PdfObjectLookup.TryGet(_objects, reference, out var indirectObject) ||
                    indirectObject.Value is not PdfStream referencedStream) {
                    return true;
                }

                stream = referencedStream;
            } else if (value is PdfStream directStream) {
                stream = directStream;
            } else {
                return true;
            }

            string? subtype = stream.Dictionary.Get<PdfName>("Subtype")?.Name;
            if (subtype is null || string.Equals(subtype, "Form", StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
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
                     maxOperands: _limits.MaxContentOperands,
                     inlineImageComponentCount: name => GetDeclaredColorSpaceComponentCount(resources, name))) {
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
        bool initialUnsupportedTextEffect = false,
        bool useLogicalTextFilters = true,
        bool includeArtifactText = false,
        int contentNestingDepth = 0,
        TextContentParser.TextOutputBudget? textOutputBudget = null,
        PdfTextClippingBudget? textClippingBudget = null,
        PageContentBudget? pageContentBudget = null,
        PdfContentOrderKey? contentOrderPrefix = null,
        int contentOrderOffset = 0,
        OfficeIccRenderingIntent initialRenderingIntent = OfficeIccRenderingIntent.RelativeColorimetric,
        PdfPaintColorSelection? initialFillColorSelection = null,
        PdfPaintColorSelection? initialStrokeColorSelection = null,
        int? contentStreamObjectNumber = null,
        Func<int, int?>? contentStreamObjectNumberAtOffset = null,
        Action? cancellationCheck = null,
        bool includeHiddenOptionalContent = false) {
        cancellationCheck?.Invoke();
        EnsureContentNestingBudget(contentNestingDepth);
        pageContentBudget ??= new PageContentBudget(this);
        textOutputBudget ??= new TextContentParser.TextOutputBudget(
            _limits.MaxActualTextCharacters,
            _limits.MaxDecodedTextCharacters);
        textClippingBudget ??= new PdfTextClippingBudget();
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
        int? ResolveMarkedContentMcid(string propertyName) =>
            GetMarkedContentMcid(resources, propertyName);

        PdfPageOptionalContentVisibility? optionalContentVisibility = GetOptionalContentVisibility(resources);
        PdfPageInvokedResourceNames invokedResources = GetInvokedResourceNames(content, resources);
        spans.AddRange(TextContentParser.Parse(
            content,
            DecodeWithFont,
            SumWidth1000,
            actualTextForProperty: ResolveActualTextProperty,
            mcidForProperty: ResolveMarkedContentMcid,
            graphicsStates: GetGraphicsStateResources(resources),
            colorSpaces: GetColorSpaceResources(resources, invokedResources.ColorSpaces, pageContentBudget),
            baseFontForResource: ResolveBaseFont,
            drawingFontFamilyForResource: ResolveDrawingFontFamily,
            optionalContentVisibility: includeHiddenOptionalContent ? null : optionalContentVisibility,
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
            includeArtifactText: includeArtifactText,
            maxOperations: _limits.MaxContentOperations,
            maxNestingDepth: _limits.MaxContentNestingDepth,
            maxOperands: _limits.MaxContentOperands,
            maxActualTextCharacters: _limits.MaxActualTextCharacters,
            maxDecodedTextCharacters: _limits.MaxDecodedTextCharacters,
            textOutputBudget: textOutputBudget,
            textClippingBudget: textClippingBudget,
            decodeWithFontWithinLimit: DecodeWithFontWithinLimit,
            contentOrderPrefix: contentOrderPrefix,
            contentOrderOffset: contentOrderOffset,
            initialUnsupportedEffect: initialUnsupportedTextEffect,
            initialRenderingIntent: initialRenderingIntent,
            initialFillColorSelection: initialFillColorSelection,
            initialStrokeColorSelection: initialStrokeColorSelection,
            outputIntentColorTransform: EffectiveOutputIntentColorTransform,
            inlineImageComponentCount: name => GetDeclaredColorSpaceComponentCount(resources, name),
            inlineImageArrayComponentCount: array => GetDeclaredColorSpaceComponentCount(array),
            contentStreamObjectNumber: contentStreamObjectNumber,
            contentStreamObjectNumberAtOffset: contentStreamObjectNumberAtOffset,
            cancellationCheck: cancellationCheck));

        foreach (var invocation in TextContentParser.ExtractFormInvocations(
                     content,
                     includeHiddenOptionalContent ? null : optionalContentVisibility,
                     paintOrderBase,
                     paintOrderScale,
                     paintOrderOffset,
                     GetGraphicsStateResources(resources),
                     GetColorSpaceResources(resources, invokedResources.ColorSpaces, pageContentBudget),
                     pageHeight,
                     initialFillColor,
                     initialFillColorSpace,
                     initialStrokeColor,
                     initialStrokeColorSpace,
                     initialFillOpacity,
                     initialStrokeOpacity,
                     initialTextRenderingMode,
                     initialClipPath,
                     initialUnsupportedTextEffect,
                     mcidForProperty: ResolveMarkedContentMcid,
                     maxOperations: _limits.MaxContentOperations,
                     maxNestingDepth: _limits.MaxContentNestingDepth,
                     maxOperands: _limits.MaxContentOperands,
                     textClippingBudget: textClippingBudget,
                     initialRenderingIntent: initialRenderingIntent,
                     initialFillColorSelection: initialFillColorSelection,
                     initialStrokeColorSelection: initialStrokeColorSelection,
                     outputIntentColorTransform: EffectiveOutputIntentColorTransform,
                     inlineImageComponentCount: name => GetDeclaredColorSpaceComponentCount(resources, name),
                     inlineImageArrayComponentCount: array => GetDeclaredColorSpaceComponentCount(array),
                     cancellationCheck: cancellationCheck)) {
            if (!TryGetFormStream(resources, invocation.Name, out int? formObjectNumber, out var formStream)) {
                continue;
            }

            if (!activeForms.Add(formStream)) {
                continue;
            }

            try {
                var formDict = formStream.Dictionary;
                if (!includeHiddenOptionalContent &&
                    optionalContentVisibility is not null &&
                    formDict.Items.TryGetValue("OC", out PdfObject? formOptionalContent) &&
                    optionalContentVisibility.IsHidden(formOptionalContent)) {
                    continue;
                }
                var formResources = ResolveDictionary(formDict.Items.TryGetValue("Resources", out var resObj) ? resObj : null) ?? resources;
                PdfFontResourceSet formFontResources = _fontResourceCache.GetOrCreate(formResources, _objects);
                var formDecoders = MergeDecoders(decoders, formFontResources.Decoders);
                var formWidths = MergeWidthProviders(widthProviders, formFontResources.WidthProviders);
                var formFonts = MergeFonts(fonts, formFontResources.Fonts);
                var combinedTransform = ApplyFormMatrix(invocation.Transform, formDict);
                var formContent = WrapContentWithTransform(WrapFormContentWithBoundingBoxClip(PdfEncoding.Latin1GetString(pageContentBudget.Decode(formStream)), formDict), combinedTransform, out int formContentOffset);
                PdfContentOrderKey? formOrderPrefix = contentOrderPrefix?.Append(invocation.SourceOperatorIndex + contentOrderOffset);

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
                    invocation.HasUnsupportedEffect || !invocation.FillColorResolved || HasTransparencyGroupForTextEditing(formDict),
                    useLogicalTextFilters,
                    includeArtifactText,
                    contentNestingDepth + 1,
                    textOutputBudget,
                    textClippingBudget,
                    pageContentBudget,
                    formOrderPrefix,
                    -formContentOffset,
                    invocation.RenderingIntent,
                    invocation.FillColorSelection,
                    invocation.StrokeColorSelection,
                    formObjectNumber,
                    contentStreamObjectNumberAtOffset: null,
                    cancellationCheck: cancellationCheck,
                    includeHiddenOptionalContent: includeHiddenOptionalContent);
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
        OfficeBlendMode initialBlendMode = OfficeBlendMode.Normal,
        OfficeBlendMode? initialAuthoredBlendMode = null,
        bool initialHasUnsupportedBlendMode = false,
        bool initialHasSoftMask = false,
        bool initialHasAuthoredRenderingIntent = false,
        OfficeIccRenderingIntent initialRenderingIntent = OfficeIccRenderingIntent.RelativeColorimetric,
        PdfPaintColorSelection? initialFillColorSelection = null,
        PdfPaintColorSelection? initialStrokeColorSelection = null,
        int contentNestingDepth = 0,
        PdfTextClippingBudget? textClippingBudget = null,
        PageContentBudget? pageContentBudget = null,
        PdfContentOrderKey? contentOrderPrefix = null,
        bool skipTransparencyGroupForms = false,
        bool includeHiddenOptionalContent = false) {
        EnsureContentNestingBudget(contentNestingDepth);
        pageContentBudget ??= new PageContentBudget(this);
        textClippingBudget ??= new PdfTextClippingBudget();
        PdfPageInvokedResourceNames invokedResources = GetInvokedResourceNames(content, resources);
        PdfPageOptionalContentVisibility? optionalContentVisibility = includeHiddenOptionalContent
            ? null
            : GetOptionalContentVisibility(resources);
        foreach (var invocation in PdfPageXObjectInvocationParser.Parse(
                     content,
                     baseTransform,
                     pageHeight,
                     GetGraphicsStateResources(resources),
                     GetColorSpaceResources(resources, invokedResources.ColorSpaces, pageContentBudget),
                     optionalContentVisibility,
                     initialFillColor,
                     initialFillColorSpace,
                      initialFillOpacity,
                      paintOrderBase,
                      paintOrderScale,
                      paintOrderOffset,
                      initialClipPath,
                     maxOperations: _limits.MaxContentOperations,
                     maxNestingDepth: _limits.MaxContentNestingDepth,
                     maxOperands: _limits.MaxContentOperands,
                     initialBlendMode: initialBlendMode,
                     initialAuthoredBlendMode: initialAuthoredBlendMode,
                     initialHasUnsupportedBlendMode: initialHasUnsupportedBlendMode,
                     initialHasSoftMask: initialHasSoftMask,
                     initialHasAuthoredRenderingIntent: initialHasAuthoredRenderingIntent,
                     initialRenderingIntent: initialRenderingIntent,
                     initialFillColorSelection: initialFillColorSelection,
                     initialStrokeColorSelection: initialStrokeColorSelection,
                     outputIntentColorTransform: EffectiveOutputIntentColorTransform,
                     textClippingBudget: textClippingBudget,
                     inlineImageArrayComponentCount: array => GetDeclaredColorSpaceComponentCount(array))) {
            Matrix2D invocationTransform = invocation.Transform;
            PdfContentOrderKey? invocationOrder = contentOrderPrefix?.Append(invocation.SourceOperatorIndex);
            if (invocation.InlineImage != null) {
                PdfImagePlacement placement = BuildImagePlacement(
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
                    invocation.PaintOrder,
                    fillPattern: invocation.FillPattern,
                    effectiveResources: resources,
                    authoredBlendMode: invocation.AuthoredBlendMode,
                    hasUnsupportedBlendMode: invocation.HasUnsupportedBlendMode,
                    hasSoftMask: invocation.HasSoftMask,
                    hasAuthoredRenderingIntent: invocation.HasAuthoredRenderingIntent,
                    renderingIntent: invocation.RenderingIntent,
                    objects: _objects);
                placements.Add(invocationOrder == null ? placement : placement.WithContentOrderKey(invocationOrder));
                continue;
            }

            if (TryGetImageXObject(
                    resources,
                    invocation.Name,
                    out int imageObjectNumber,
                    out int directStreamIdentity,
                    out PdfStream? imageStream)) {
                if (!includeHiddenOptionalContent &&
                    optionalContentVisibility is not null &&
                    imageStream!.Dictionary.Items.TryGetValue("OC", out PdfObject? imageOptionalContent) &&
                    optionalContentVisibility.IsHidden(imageOptionalContent)) {
                    continue;
                }
                PdfImagePlacement placement = BuildImagePlacement(
                    pageNumber,
                    invocation.Name,
                    imageObjectNumber,
                    directStreamIdentity,
                    invocationTransform,
                    invocation.ClipPath,
                    invocation.FillColor,
                    invocation.FillOpacity,
                    paintOrder: invocation.PaintOrder,
                    fillPattern: invocation.FillPattern,
                    effectiveResources: resources,
                    authoredBlendMode: invocation.AuthoredBlendMode,
                    hasUnsupportedBlendMode: invocation.HasUnsupportedBlendMode,
                    hasSoftMask: invocation.HasSoftMask,
                    hasAuthoredRenderingIntent: invocation.HasAuthoredRenderingIntent,
                    renderingIntent: invocation.RenderingIntent,
                    imageDictionary: imageStream!.Dictionary,
                    objects: _objects);
                placements.Add(invocationOrder == null ? placement : placement.WithContentOrderKey(invocationOrder));
                continue;
            }

            if (!TryGetFormStream(resources, invocation.Name, out var formStream)) {
                continue;
            }

            if (skipTransparencyGroupForms &&
                (!TryClassifyType3TransparencyGroup(formStream.Dictionary, out bool isTransparencyGroup) || isTransparencyGroup)) {
                continue;
            }

            if (!activeForms.Add(formStream)) {
                continue;
            }

            try {
                var formDict = formStream.Dictionary;
                if (!includeHiddenOptionalContent &&
                    optionalContentVisibility is not null &&
                    formDict.Items.TryGetValue("OC", out PdfObject? formOptionalContent) &&
                    optionalContentVisibility.IsHidden(formOptionalContent)) {
                    continue;
                }
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
                    initialBlendMode: invocation.BlendMode,
                    initialAuthoredBlendMode: invocation.AuthoredBlendMode,
                    initialHasUnsupportedBlendMode: invocation.HasUnsupportedBlendMode,
                    initialHasSoftMask: invocation.HasSoftMask,
                    initialHasAuthoredRenderingIntent: invocation.HasAuthoredRenderingIntent,
                    initialRenderingIntent: invocation.RenderingIntent,
                    initialFillColorSelection: invocation.FillColorSelection,
                    initialStrokeColorSelection: invocation.StrokeColorSelection,
                    contentNestingDepth: contentNestingDepth + 1,
                    textClippingBudget: textClippingBudget,
                    pageContentBudget: pageContentBudget,
                    contentOrderPrefix: invocationOrder,
                    skipTransparencyGroupForms: skipTransparencyGroupForms,
                    includeHiddenOptionalContent: includeHiddenOptionalContent);
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
        return TryGetFormStream(resources, name, out _, out formStream);
    }

    private bool TryGetFormStream(PdfDictionary? resources, string name, out int? objectNumber, out PdfStream formStream) {
        objectNumber = null;
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
            objectNumber = formRef.ObjectNumber;
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
        return TryGetImageXObject(resources, name, out objectNumber, out directStreamIdentity, out _);
    }

    private bool TryGetXObjectStream(PdfDictionary? resources, string name, out PdfStream? stream) {
        stream = null;
        if (resources is null ||
            !resources.Items.TryGetValue("XObject", out PdfObject? xObjectsObject) ||
            ResolveDictionary(xObjectsObject) is not PdfDictionary xObjects ||
            !xObjects.Items.TryGetValue(name, out PdfObject? xObject)) return false;
        stream = ResolveObject(xObject) as PdfStream;
        return stream != null;
    }

    private bool TryGetImageXObject(PdfDictionary? resources, string name, out int objectNumber, out int directStreamIdentity, out PdfStream? imageStream) {
        objectNumber = 0;
        directStreamIdentity = 0;
        imageStream = null;
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
            directStreamIdentity = PdfDirectStreamIdentity.Compute(directStream);
        }

        bool isImage = stream is not null &&
            string.Equals(stream.Dictionary.Get<PdfName>("Subtype")?.Name, "Image", StringComparison.Ordinal);
        if (isImage) imageStream = stream;
        return isImage;
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
        double paintOrder = 0D,
        PdfPagePatternSelection? fillPattern = null,
        PdfDictionary? effectiveResources = null,
        OfficeBlendMode? authoredBlendMode = null,
        bool hasUnsupportedBlendMode = false,
        bool hasSoftMask = false,
        bool hasAuthoredRenderingIntent = false,
        OfficeIccRenderingIntent renderingIntent = OfficeIccRenderingIntent.RelativeColorimetric,
        PdfDictionary? imageDictionary = null,
        Dictionary<int, PdfIndirectObject>? objects = null) {
        PdfDictionary? intentOwner = imageDictionary ?? inlineImageStream?.Dictionary;
        if (intentOwner is not null && objects is not null && PdfRenderingIntentResolver.TryRead(
                intentOwner,
                "Intent",
                objects,
                out OfficeIccRenderingIntent authoredImageIntent)) {
            hasAuthoredRenderingIntent = true;
            renderingIntent = authoredImageIntent;
        }
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
            paintOrder,
            fillPattern: fillPattern,
            effectiveResources: effectiveResources,
            blendMode: authoredBlendMode,
            hasUnsupportedBlendMode: hasUnsupportedBlendMode,
            hasSoftMask: hasSoftMask,
            hasAuthoredRenderingIntent: hasAuthoredRenderingIntent,
            renderingIntent: renderingIntent);
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

    private int? GetMarkedContentMcid(PdfDictionary? resources, string propertyName) {
        if (resources is null || !resources.Items.TryGetValue("Properties", out PdfObject? propertiesObject)) return null;
        PdfDictionary? properties = ResolveDictionary(propertiesObject);
        if (properties is null || !properties.Items.TryGetValue(propertyName, out PdfObject? propertyObject)) return null;
        PdfDictionary? property = ResolveDictionary(propertyObject);
        return property is null ? null : TryReadInteger(property, "MCID");
    }

    private PdfPageOptionalContentVisibility? GetOptionalContentVisibility(PdfDictionary? resources) =>
        PdfPageOptionalContentVisibility.Create(resources, _objects, _limits.MaxContentNestingDepth);

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

    private Matrix2D ApplyFormMatrix(Matrix2D invocationTransform, PdfDictionary? formDict) {
        return TryReadFormMatrix(formDict, out Matrix2D formMatrix)
            ? Matrix2D.Multiply(invocationTransform, formMatrix)
            : invocationTransform;
    }

    private bool TryReadFormMatrix(PdfDictionary? formDict, out Matrix2D formMatrix) {
        formMatrix = Matrix2D.Identity;
        if (formDict is null || !formDict.Items.TryGetValue("Matrix", out PdfObject? matrixObject)) return true;
        PdfObject? resolvedMatrix = ResolveEffectObject(matrixObject);
        if (resolvedMatrix is PdfNull) return true;
        if (resolvedMatrix is not PdfArray array || array.Items.Count != 6) return false;
        var values = new double[6];
        for (int index = 0; index < values.Length; index++) {
            if (ResolveEffectObject(array.Items[index]) is not PdfNumber number ||
                double.IsNaN(number.Value) ||
                double.IsInfinity(number.Value)) return false;
            values[index] = number.Value;
        }
        formMatrix = new Matrix2D(values[0], values[1], values[2], values[3], values[4], values[5]);
        return true;
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

    private bool HasTransparencyGroupForTextEditing(PdfDictionary formDictionary) {
        if (!formDictionary.Items.TryGetValue("Group", out PdfObject? groupObject) ||
            ResolveTextEditingObjectChain(groupObject) is not PdfDictionary group ||
            !group.Items.TryGetValue("S", out PdfObject? subtypeObject)) {
            return false;
        }
        return ResolveTextEditingObjectChain(subtypeObject) is PdfName subtype &&
            string.Equals(subtype.Name, "Transparency", StringComparison.Ordinal);
    }

    private PdfObject? ResolveTextEditingObjectChain(PdfObject? value) {
        var visited = new HashSet<(int ObjectNumber, int Generation)>();
        int maximumDepth = Math.Max(1, _limits.MaxObjectNestingDepth);
        for (int depth = 0; depth < maximumDepth && value is PdfReference reference; depth++) {
            if (!visited.Add((reference.ObjectNumber, reference.Generation)) ||
                !PdfObjectLookup.TryGet(_objects, reference, out PdfIndirectObject? indirect)) {
                return null;
            }
            value = indirect.Value;
        }
        return value is PdfReference ? null : value;
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

    private void ReadAnnotationAppearanceMetadata(
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
        TryGetString(annotation.Items.TryGetValue("RC", out PdfObject? richContentsObject) ? richContentsObject : null, out richContents);
        richContentsPlainText = PdfFreeTextStyleParser.ExtractPlainText(richContents);
        if (!string.Equals(subtype, "FreeText", StringComparison.Ordinal)) {
            return;
        }

        TryGetString(annotation.Items.TryGetValue("DA", out PdfObject? defaultAppearanceObject) ? defaultAppearanceObject : null, out defaultAppearance);
        TryGetString(annotation.Items.TryGetValue("DS", out PdfObject? defaultStyleObject) ? defaultStyleObject : null, out defaultStyle);
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
        return TryGetNormalAppearanceStream(annotation, out _);
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
        return GetContentStreamSequence(pageContentBudget).Content;
    }

    private PageContentStreamSequence GetContentStreamSequence(PageContentBudget? pageContentBudget = null) {
        pageContentBudget ??= new PageContentBudget(this);
        var builder = new System.Text.StringBuilder();
        var entries = new List<PageContentStreamOffset>();
        foreach (PageContentStreamEntry entry in GetContentStreamObjects()) {
            if (builder.Length > 0) builder.Append('\n');
            entries.Add(new PageContentStreamOffset(builder.Length, entry.ObjectNumber));
            builder.Append(PdfEncoding.Latin1GetString(pageContentBudget.Decode(entry.Stream)));
        }

        return new PageContentStreamSequence(builder.ToString(), entries.ToArray());
    }

    private List<PageContentStreamEntry> GetContentStreamObjects() {
        var result = new List<PageContentStreamEntry>();
        var contents = _pageDict.Items.TryGetValue("Contents", out var obj) ? obj : null;
        if (contents is PdfReference r) {
            if (PdfObjectLookup.TryGet(_objects, r, out var ind) && ind.Value is PdfStream s) {
                result.Add(new PageContentStreamEntry(s, r.ObjectNumber));
                return result;
            }
        } else if (contents is PdfStream directStream) {
            result.Add(new PageContentStreamEntry(directStream, null));
            return result;
        }

        var contentArray = ResolveArray(contents);
        if (contentArray is null) {
            return result;
        }

        foreach (var item in contentArray.Items) {
            if (item is PdfReference rr &&
                PdfObjectLookup.TryGet(_objects, rr, out var ind2) &&
                ind2.Value is PdfStream s2) {
                result.Add(new PageContentStreamEntry(s2, rr.ObjectNumber));
            } else if (item is PdfStream directStream) {
                result.Add(new PageContentStreamEntry(directStream, null));
            }
        }

        return result;
    }

    internal bool IsPageContentStreamObjectNumber(int? objectNumber) {
        if (!objectNumber.HasValue) return true;
        List<PageContentStreamEntry> streams = GetContentStreamObjects();
        for (int index = 0; index < streams.Count; index++) {
            if (streams[index].ObjectNumber == objectNumber) return true;
        }
        return false;
    }

    private readonly struct PageContentStreamEntry {
        internal PageContentStreamEntry(PdfStream stream, int? objectNumber) {
            Stream = stream;
            ObjectNumber = objectNumber;
        }

        internal PdfStream Stream { get; }
        internal int? ObjectNumber { get; }
    }

    private readonly struct PageContentStreamOffset {
        internal PageContentStreamOffset(int startOffset, int? objectNumber) {
            StartOffset = startOffset;
            ObjectNumber = objectNumber;
        }

        internal int StartOffset { get; }
        internal int? ObjectNumber { get; }
    }

    private sealed class PageContentStreamSequence {
        private readonly PageContentStreamOffset[] _offsets;

        internal PageContentStreamSequence(string content, PageContentStreamOffset[] offsets) {
            Content = content;
            _offsets = offsets;
        }

        internal string Content { get; }

        internal int? GetObjectNumber(int operatorOffset) {
            int low = 0;
            int high = _offsets.Length - 1;
            int match = -1;
            while (low <= high) {
                int middle = low + ((high - low) / 2);
                if (_offsets[middle].StartOffset <= operatorOffset) {
                    match = middle;
                    low = middle + 1;
                } else {
                    high = middle - 1;
                }
            }
            return match >= 0 ? _offsets[match].ObjectNumber : null;
        }
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
        if (s.DecodingFailed) {
            throw new InvalidDataException(
                "PDF page content stream could not be decoded safely" +
                (string.IsNullOrWhiteSpace(s.DecodingError) ? "." : ": " + s.DecodingError));
        }

        return Filters.StreamDecoder.DecodeRequired(s.Dictionary, s.Data, _objects, maxDecodedBytes);
    }

    internal sealed class PageContentBudget {
        private readonly PdfReadPage _page;
        private readonly Dictionary<PdfStream, byte[]> _decodedStreams = new();
        private long _decodedBytes;
        private long _remainingColorFunctionEvaluationWork;

        internal PageContentBudget(PdfReadPage page) {
            _page = page;
            _remainingColorFunctionEvaluationWork = Math.Max(1, page._limits.MaxContentOperations);
            ColorFunctionResolutionContext = new PdfColorFunctionResolutionContext(
                Math.Min(page._limits.MaxDecodedStreamBytes, page._limits.MaxPageContentBytes));
        }

        internal PdfColorFunctionResolutionContext ColorFunctionResolutionContext { get; }

        internal bool TryConsumeColorFunctionEvaluation(int evaluationCost) =>
            TryConsumeColorFunctionEvaluations(evaluationCost, 1L);

        internal bool TryConsumeColorFunctionEvaluations(int evaluationCost, long evaluationCount) {
            if (evaluationCount < 0L) return false;
            long cost;
            try {
                cost = checked(Math.Max(1, evaluationCost) * evaluationCount);
            } catch (OverflowException) {
                return false;
            }
            if (cost > _remainingColorFunctionEvaluationWork) return false;
            _remainingColorFunctionEvaluationWork -= cost;
            return true;
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
