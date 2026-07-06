namespace OfficeIMO.Pdf;

/// <summary>Generic page annotation metadata read from a PDF page.</summary>
public sealed class PdfAnnotation {
    private const int InvisibleFlag = 1;
    private const int HiddenFlag = 2;
    private const int PrintFlag = 4;
    private const int NoZoomFlag = 8;
    private const int NoRotateFlag = 16;
    private const int NoViewFlag = 32;
    private const int ReadOnlyFlag = 64;
    private const int LockedFlag = 128;
    private const int ToggleNoViewFlag = 256;
    private const int LockedContentsFlag = 512;

    internal PdfAnnotation(int? objectNumber, int? pageNumber, string subtype, string? contents, double x1, double y1, double x2, double y2, bool hasNormalAppearance, string? actionType = null, IReadOnlyList<PdfAnnotationAdditionalAction>? additionalActions = null, IReadOnlyList<PdfAnnotationChainedAction>? chainedActions = null, int? flags = null, string? name = null, string? title = null, string? modified = null, IReadOnlyList<double>? color = null, string? defaultAppearance = null, string? defaultStyle = null, string? richContents = null, string? richContentsPlainText = null, double? effectiveFontSize = null, PdfColor? effectiveTextColor = null, PdfAlign? effectiveTextAlign = null, IReadOnlyList<double>? interiorColor = null, double? opacity = null, double? borderWidth = null, string? borderStyle = null, IReadOnlyList<double>? borderDashPattern = null, string? borderEffectStyle = null, double? borderEffectIntensity = null, IReadOnlyList<double>? rectangleDifferences = null, IReadOnlyList<double>? calloutLine = null, string? calloutLineEnding = null, string? lineStartEnding = null, string? lineEndEnding = null, IReadOnlyList<double>? quadPoints = null, IReadOnlyList<double>? lineCoordinates = null, IReadOnlyList<double>? vertices = null, IReadOnlyList<IReadOnlyList<double>>? inkList = null) {
        ObjectNumber = objectNumber;
        PageNumber = pageNumber;
        Subtype = subtype;
        Contents = contents;
        X1 = x1;
        Y1 = y1;
        X2 = x2;
        Y2 = y2;
        HasNormalAppearance = hasNormalAppearance;
        ActionType = actionType;
        AdditionalActions = additionalActions ?? Array.Empty<PdfAnnotationAdditionalAction>();
        ChainedActions = chainedActions ?? Array.Empty<PdfAnnotationChainedAction>();
        Flags = flags;
        Name = name;
        Title = title;
        Modified = modified;
        Color = color ?? Array.Empty<double>();
        DefaultAppearance = defaultAppearance;
        DefaultStyle = defaultStyle;
        RichContents = richContents;
        RichContentsPlainText = richContentsPlainText;
        EffectiveFontSize = effectiveFontSize;
        EffectiveTextColor = effectiveTextColor;
        EffectiveTextAlign = effectiveTextAlign;
        InteriorColor = interiorColor ?? Array.Empty<double>();
        Opacity = opacity;
        BorderWidth = borderWidth;
        BorderStyle = borderStyle;
        BorderDashPattern = borderDashPattern ?? Array.Empty<double>();
        BorderEffectStyle = borderEffectStyle;
        BorderEffectIntensity = borderEffectIntensity;
        RectangleDifferences = rectangleDifferences ?? Array.Empty<double>();
        CalloutLine = calloutLine ?? Array.Empty<double>();
        CalloutLineEnding = calloutLineEnding;
        LineStartEnding = lineStartEnding;
        LineEndEnding = lineEndEnding;
        QuadPoints = quadPoints ?? Array.Empty<double>();
        LineCoordinates = lineCoordinates ?? Array.Empty<double>();
        Vertices = vertices ?? Array.Empty<double>();
        InkList = inkList ?? Array.Empty<IReadOnlyList<double>>();
    }

    /// <summary>Indirect annotation object number, when the annotation is referenced indirectly.</summary>
    public int? ObjectNumber { get; }

    /// <summary>One-based page number when known.</summary>
    public int? PageNumber { get; }

    /// <summary>PDF annotation subtype name, for example Link, Text, FreeText, or Highlight.</summary>
    public string Subtype { get; }

    /// <summary>Optional annotation contents metadata from /Contents.</summary>
    public string? Contents { get; }

    /// <summary>Left edge of the annotation rectangle in PDF points.</summary>
    public double X1 { get; }

    /// <summary>Bottom edge of the annotation rectangle in PDF points.</summary>
    public double Y1 { get; }

    /// <summary>Right edge of the annotation rectangle in PDF points.</summary>
    public double X2 { get; }

    /// <summary>Top edge of the annotation rectangle in PDF points.</summary>
    public double Y2 { get; }

    /// <summary>Rectangle width in PDF points.</summary>
    public double Width => X2 - X1;

    /// <summary>Rectangle height in PDF points.</summary>
    public double Height => Y2 - Y1;

    /// <summary>True when the annotation dictionary exposes a normal appearance stream through /AP /N.</summary>
    public bool HasNormalAppearance { get; }

    /// <summary>Primary annotation action type from /A /S, when present.</summary>
    public string? ActionType { get; }

    /// <summary>Additional annotation actions from the /AA dictionary, when present.</summary>
    public IReadOnlyList<PdfAnnotationAdditionalAction> AdditionalActions { get; }

    /// <summary>Chained annotation actions discovered through /Next entries on /A or /AA action dictionaries.</summary>
    public IReadOnlyList<PdfAnnotationChainedAction> ChainedActions { get; }

    /// <summary>Raw annotation flags from /F, when present.</summary>
    public int? Flags { get; }

    /// <summary>Annotation unique name from /NM, when present.</summary>
    public string? Name { get; }

    /// <summary>Annotation title or author from /T, when present.</summary>
    public string? Title { get; }

    /// <summary>Annotation modification date string from /M, when present.</summary>
    public string? Modified { get; }

    /// <summary>Annotation color components from /C, when readable.</summary>
    public IReadOnlyList<double> Color { get; }

    /// <summary>FreeText default appearance string from /DA, when present.</summary>
    public string? DefaultAppearance { get; }

    /// <summary>FreeText default style string from /DS, when present.</summary>
    public string? DefaultStyle { get; }

    /// <summary>FreeText rich contents string from /RC, when present.</summary>
    public string? RichContents { get; }

    /// <summary>Plain text extracted from FreeText /RC when rich contents are present and readable.</summary>
    public string? RichContentsPlainText { get; }

    /// <summary>Effective FreeText font size resolved from /DA or /DS, when readable.</summary>
    public double? EffectiveFontSize { get; }

    /// <summary>Effective FreeText text color resolved from /DA or /DS, when readable.</summary>
    public PdfColor? EffectiveTextColor { get; }

    /// <summary>Effective FreeText text alignment resolved from /Q or /DS, when readable.</summary>
    public PdfAlign? EffectiveTextAlign { get; }

    /// <summary>Annotation interior color components from /IC, when readable.</summary>
    public IReadOnlyList<double> InteriorColor { get; }

    /// <summary>Annotation opacity from /CA, when present and valid.</summary>
    public double? Opacity { get; }

    /// <summary>Annotation border width from /BS /W or /Border, when present and valid.</summary>
    public double? BorderWidth { get; }

    /// <summary>Annotation border style resolved from /BS /S, when present.</summary>
    public string? BorderStyle { get; }

    /// <summary>Annotation border dash pattern from /BS /D, when present and valid.</summary>
    public IReadOnlyList<double> BorderDashPattern { get; }

    /// <summary>Annotation border effect style from /BE /S, when present.</summary>
    public string? BorderEffectStyle { get; }

    /// <summary>Annotation border effect intensity from /BE /I, when present and valid.</summary>
    public double? BorderEffectIntensity { get; }

    /// <summary>FreeText rectangle differences from /RD as left, top, right, bottom values, when valid.</summary>
    public IReadOnlyList<double> RectangleDifferences { get; }

    /// <summary>FreeText callout line coordinates from /CL, when present and valid.</summary>
    public IReadOnlyList<double> CalloutLine { get; }

    /// <summary>FreeText callout line ending from /LE, when present.</summary>
    public string? CalloutLineEnding { get; }

    /// <summary>Line or PolyLine start ending from /LE, when present.</summary>
    public string? LineStartEnding { get; }

    /// <summary>Line or PolyLine end ending from /LE, when present.</summary>
    public string? LineEndEnding { get; }

    /// <summary>Text-markup quad point coordinates from /QuadPoints, when present and valid.</summary>
    public IReadOnlyList<double> QuadPoints { get; }

    /// <summary>Line annotation coordinates from /L, when present and valid.</summary>
    public IReadOnlyList<double> LineCoordinates { get; }

    /// <summary>Polygon or PolyLine vertices from /Vertices, when present and valid.</summary>
    public IReadOnlyList<double> Vertices { get; }

    /// <summary>Ink annotation paths from /InkList, when present and valid.</summary>
    public IReadOnlyList<IReadOnlyList<double>> InkList { get; }

    /// <summary>True when the annotation exposes FreeText appearance metadata such as /DA, /DS, /RC, or /Q.</summary>
    public bool HasFreeTextAppearanceMetadata =>
        !string.IsNullOrWhiteSpace(DefaultAppearance) ||
        !string.IsNullOrWhiteSpace(DefaultStyle) ||
        !string.IsNullOrWhiteSpace(RichContents) ||
        EffectiveFontSize.HasValue ||
        EffectiveTextColor.HasValue ||
        EffectiveTextAlign.HasValue;

    /// <summary>True when the annotation exposes visual styling metadata beyond the rectangle and subtype.</summary>
    public bool HasVisualStyleMetadata =>
        InteriorColor.Count > 0 ||
        Opacity.HasValue ||
        BorderWidth.HasValue ||
        !string.IsNullOrWhiteSpace(BorderStyle) ||
        BorderDashPattern.Count > 0 ||
        !string.IsNullOrWhiteSpace(BorderEffectStyle) ||
        BorderEffectIntensity.HasValue ||
        RectangleDifferences.Count > 0 ||
        CalloutLine.Count > 0 ||
        !string.IsNullOrWhiteSpace(CalloutLineEnding) ||
        !string.IsNullOrWhiteSpace(LineStartEnding) ||
        !string.IsNullOrWhiteSpace(LineEndEnding);

    /// <summary>True when the annotation exposes path geometry such as /QuadPoints, /L, /Vertices, or /InkList.</summary>
    public bool HasPathGeometryMetadata =>
        QuadPoints.Count > 0 ||
        LineCoordinates.Count > 0 ||
        Vertices.Count > 0 ||
        InkList.Count > 0;

    /// <summary>True when the annotation has a readable color array.</summary>
    public bool HasColor => Color.Count > 0;

    /// <summary>True when the annotation has a primary action dictionary.</summary>
    public bool HasAction => !string.IsNullOrEmpty(ActionType);

    /// <summary>True when the annotation has at least one additional action dictionary.</summary>
    public bool HasAdditionalActions => AdditionalActions.Count > 0;

    /// <summary>True when at least one chained /Next action was discovered for this annotation.</summary>
    public bool HasChainedActions => ChainedActions.Count > 0;

    /// <summary>Number of chained /Next actions discovered for this annotation.</summary>
    public int ChainedActionCount => ChainedActions.Count;

    /// <summary>True when the annotation has the PDF Invisible flag.</summary>
    public bool IsInvisible => HasFlag(InvisibleFlag);

    /// <summary>True when the annotation has the PDF Hidden flag.</summary>
    public bool IsHidden => HasFlag(HiddenFlag);

    /// <summary>True when the annotation has the PDF Print flag.</summary>
    public bool IsPrint => HasFlag(PrintFlag);

    /// <summary>True when the annotation has the PDF NoZoom flag.</summary>
    public bool IsNoZoom => HasFlag(NoZoomFlag);

    /// <summary>True when the annotation has the PDF NoRotate flag.</summary>
    public bool IsNoRotate => HasFlag(NoRotateFlag);

    /// <summary>True when the annotation has the PDF NoView flag.</summary>
    public bool IsNoView => HasFlag(NoViewFlag);

    /// <summary>True when the annotation has the PDF ReadOnly flag.</summary>
    public bool IsReadOnly => HasFlag(ReadOnlyFlag);

    /// <summary>True when the annotation has the PDF Locked flag.</summary>
    public bool IsLocked => HasFlag(LockedFlag);

    /// <summary>True when the annotation has the PDF ToggleNoView flag.</summary>
    public bool IsToggleNoView => HasFlag(ToggleNoViewFlag);

    /// <summary>True when the annotation has the PDF LockedContents flag.</summary>
    public bool IsLockedContents => HasFlag(LockedContentsFlag);

    internal PdfAnnotation WithPageNumber(int pageNumber) =>
        PageNumber == pageNumber
            ? this
            : new PdfAnnotation(ObjectNumber, pageNumber, Subtype, Contents, X1, Y1, X2, Y2, HasNormalAppearance, ActionType, AdditionalActions, ChainedActions, Flags, Name, Title, Modified, Color, DefaultAppearance, DefaultStyle, RichContents, RichContentsPlainText, EffectiveFontSize, EffectiveTextColor, EffectiveTextAlign, InteriorColor, Opacity, BorderWidth, BorderStyle, BorderDashPattern, BorderEffectStyle, BorderEffectIntensity, RectangleDifferences, CalloutLine, CalloutLineEnding, LineStartEnding, LineEndEnding, QuadPoints, LineCoordinates, Vertices, InkList);

    private bool HasFlag(int flag) {
        return Flags.HasValue && (Flags.Value & flag) != 0;
    }
}

/// <summary>Additional annotation action metadata read from an annotation /AA dictionary.</summary>
public sealed class PdfAnnotationAdditionalAction {
    internal PdfAnnotationAdditionalAction(string triggerName, string actionType) {
        TriggerName = triggerName;
        ActionType = actionType;
    }

    /// <summary>PDF additional-action trigger name, for example E, X, D, U, Fo, Bl, PO, PC, PV, or PI.</summary>
    public string TriggerName { get; }

    /// <summary>Action type from the additional action dictionary /S entry.</summary>
    public string ActionType { get; }
}

/// <summary>Chained annotation action metadata read from /Next entries without exposing action payloads.</summary>
public sealed class PdfAnnotationChainedAction {
    internal PdfAnnotationChainedAction(string sourceName, string actionPath, string actionType) {
        SourceName = sourceName;
        ActionPath = actionPath;
        ActionType = actionType;
    }

    /// <summary>Primary action source, A, or page additional-action trigger name such as E, X, D, U, Fo, or Bl.</summary>
    public string SourceName { get; }

    /// <summary>Stable chained-action path, for example A.Next, E.Next, or E.Next.0.</summary>
    public string ActionPath { get; }

    /// <summary>Action type from the chained action dictionary /S entry.</summary>
    public string ActionType { get; }
}
