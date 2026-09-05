namespace OfficeIMO.Pdf;

/// <summary>
/// Page boundary boxes and page-level presentation metadata read from a PDF page dictionary.
/// </summary>
public sealed class PdfPageGeometry {
    internal PdfPageGeometry(
        PdfPageBox? mediaBox,
        PdfPageBox? cropBox,
        PdfPageBox? bleedBox,
        PdfPageBox? trimBox,
        PdfPageBox? artBox,
        double? userUnit,
        string? tabOrder,
        double? durationSeconds,
        PdfPageTransition? transition,
        bool hasMetadata,
        int? metadataObjectNumber,
        bool hasPieceInfo) {
        MediaBox = mediaBox;
        CropBox = cropBox;
        BleedBox = bleedBox;
        TrimBox = trimBox;
        ArtBox = artBox;
        UserUnit = userUnit;
        TabOrder = tabOrder;
        DurationSeconds = durationSeconds;
        Transition = transition;
        HasMetadata = hasMetadata;
        MetadataObjectNumber = metadataObjectNumber;
        HasPieceInfo = hasPieceInfo;
    }

    /// <summary>Inherited /MediaBox boundary, when readable.</summary>
    public PdfPageBox? MediaBox { get; }

    /// <summary>Inherited /CropBox boundary, when readable.</summary>
    public PdfPageBox? CropBox { get; }

    /// <summary>Inherited /BleedBox boundary, when readable.</summary>
    public PdfPageBox? BleedBox { get; }

    /// <summary>Inherited /TrimBox boundary, when readable.</summary>
    public PdfPageBox? TrimBox { get; }

    /// <summary>Inherited /ArtBox boundary, when readable.</summary>
    public PdfPageBox? ArtBox { get; }

    /// <summary>
    /// Effective page box used by OfficeIMO.Pdf for page size and coordinates. When both boxes are
    /// readable, this is the intersection of CropBox and MediaBox as required by the PDF page model.
    /// A null result with readable boxes means their intersection is empty and geometry consumers
    /// must fail closed instead of using out-of-page coordinates.
    /// </summary>
    public PdfPageBox? EffectiveBox {
        get {
            if (CropBox is null) return MediaBox;
            if (MediaBox is null) return CropBox;
            double left = Math.Max(CropBox.Left, MediaBox.Left);
            double bottom = Math.Max(CropBox.Bottom, MediaBox.Bottom);
            double right = Math.Min(CropBox.Right, MediaBox.Right);
            double top = Math.Min(CropBox.Top, MediaBox.Top);
            return right > left && top > bottom
                ? new PdfPageBox("EffectiveBox", left, bottom, right, top)
                : null;
        }
    }

    /// <summary>True when readable MediaBox and CropBox values do not overlap.</summary>
    public bool HasEmptyEffectiveBoxIntersection => MediaBox is not null && CropBox is not null && EffectiveBox is null;

    /// <summary>True when the readable CropBox was clamped to the MediaBox.</summary>
    public bool IsCropBoxClamped => MediaBox is not null && CropBox is not null && EffectiveBox is PdfPageBox effective &&
        (effective.Left != CropBox.Left || effective.Bottom != CropBox.Bottom || effective.Right != CropBox.Right || effective.Top != CropBox.Top);

    /// <summary>Inherited page user-unit scale from /UserUnit, when present and positive.</summary>
    public double? UserUnit { get; }

    /// <summary>Page tab order from /Tabs, when present.</summary>
    public string? TabOrder { get; }

    /// <summary>Page display duration from /Dur, in seconds, when present.</summary>
    public double? DurationSeconds { get; }

    /// <summary>Page transition dictionary from /Trans, when present and readable.</summary>
    public PdfPageTransition? Transition { get; }

    /// <summary>True when the page has a /Trans transition dictionary.</summary>
    public bool HasTransition => Transition is not null;

    /// <summary>True when the page has page-level /Metadata.</summary>
    public bool HasMetadata { get; }

    /// <summary>Object number of page-level /Metadata when it is an indirect reference.</summary>
    public int? MetadataObjectNumber { get; }

    /// <summary>True when the page has a /PieceInfo dictionary.</summary>
    public bool HasPieceInfo { get; }

    /// <summary>True when at least one non-default boundary box was readable.</summary>
    public bool HasNonDefaultBoundaryBoxes => CropBox is not null || BleedBox is not null || TrimBox is not null || ArtBox is not null;

    /// <summary>True when TrimBox, BleedBox, or ArtBox information was readable for production workflows.</summary>
    public bool HasProductionBoundaryBoxes => TrimBox is not null || BleedBox is not null || ArtBox is not null;

    /// <summary>True when both TrimBox and BleedBox are readable, a common print-production preflight pair.</summary>
    public bool HasPrintProductionBoxes => TrimBox is not null && BleedBox is not null;
}
