namespace OfficeIMO.Pdf;

/// <summary>
/// Basic geometry and identity information for a page in a PDF document.
/// </summary>
public sealed class PdfPageInfo {
    internal PdfPageInfo(int pageNumber, double width, double height, int rotationDegrees = 0, PdfPageGeometry? geometry = null, IReadOnlyList<PdfLinkAnnotation>? linkAnnotations = null, IReadOnlyList<PdfFormWidget>? formWidgets = null, IReadOnlyList<PdfAnnotation>? annotations = null, IReadOnlyList<PdfPageAction>? pageActions = null) {
        PageNumber = pageNumber;
        Width = width;
        Height = height;
        RotationDegrees = rotationDegrees;
        Geometry = geometry ?? new PdfPageGeometry(null, null, null, null, null, null, null, null, null, false, null, false);
        LinkAnnotations = linkAnnotations ?? Array.Empty<PdfLinkAnnotation>();
        FormWidgets = formWidgets ?? Array.Empty<PdfFormWidget>();
        Annotations = annotations ?? Array.Empty<PdfAnnotation>();
        PageActions = pageActions ?? Array.Empty<PdfPageAction>();
    }

    /// <summary>One-based page number in document order.</summary>
    public int PageNumber { get; }

    /// <summary>Page width in PDF points.</summary>
    public double Width { get; }

    /// <summary>Page height in PDF points.</summary>
    public double Height { get; }

    /// <summary>Inherited page rotation in degrees.</summary>
    public int RotationDegrees { get; }

    /// <summary>Page boundary boxes and page-level presentation metadata.</summary>
    public PdfPageGeometry Geometry { get; }

    /// <summary>Inherited /MediaBox boundary, when readable.</summary>
    public PdfPageBox? MediaBox => Geometry.MediaBox;

    /// <summary>Inherited /CropBox boundary, when readable.</summary>
    public PdfPageBox? CropBox => Geometry.CropBox;

    /// <summary>Inherited /BleedBox boundary, when readable.</summary>
    public PdfPageBox? BleedBox => Geometry.BleedBox;

    /// <summary>Inherited /TrimBox boundary, when readable.</summary>
    public PdfPageBox? TrimBox => Geometry.TrimBox;

    /// <summary>Inherited /ArtBox boundary, when readable.</summary>
    public PdfPageBox? ArtBox => Geometry.ArtBox;

    /// <summary>Inherited page user-unit scale from /UserUnit, when present and positive.</summary>
    public double? UserUnit => Geometry.UserUnit;

    /// <summary>Page tab order from /Tabs, when present.</summary>
    public string? TabOrder => Geometry.TabOrder;

    /// <summary>Page display duration from /Dur, in seconds, when present.</summary>
    public double? DurationSeconds => Geometry.DurationSeconds;

    /// <summary>Page transition dictionary from /Trans, when present and readable.</summary>
    public PdfPageTransition? Transition => Geometry.Transition;

    /// <summary>True when page-level /Metadata was present.</summary>
    public bool HasPageMetadata => Geometry.HasMetadata;

    /// <summary>True when page-level /PieceInfo was present.</summary>
    public bool HasPieceInfo => Geometry.HasPieceInfo;

    /// <summary>Simple URI, named-destination, direct-destination, named-action, and remote GoTo link annotations on this page.</summary>
    public IReadOnlyList<PdfLinkAnnotation> LinkAnnotations { get; }

    /// <summary>Simple AcroForm widget annotations on this page.</summary>
    public IReadOnlyList<PdfFormWidget> FormWidgets { get; }

    /// <summary>Generic annotation metadata on this page.</summary>
    public IReadOnlyList<PdfAnnotation> Annotations { get; }

    /// <summary>Page-level additional actions from the page dictionary /AA entry.</summary>
    public IReadOnlyList<PdfPageAction> PageActions { get; }

    /// <summary>Number of page-level additional actions read on this page.</summary>
    public int PageActionCount => PageActions.Count;

    /// <summary>True when at least one page-level additional action was read on this page.</summary>
    public bool HasPageActions => PageActionCount > 0;

    /// <summary>True when at least one generic annotation was read on this page.</summary>
    public bool HasAnnotations => Annotations.Count > 0;

    /// <summary>True when at least one simple AcroForm widget annotation was read on this page.</summary>
    public bool HasFormWidgets => FormWidgets.Count > 0;

    /// <summary>Page size in PDF points.</summary>
    public PageSize Size => new PageSize(Width, Height);

    /// <summary>True when the page is wider than it is tall.</summary>
    public bool IsLandscape => Width > Height;
}
