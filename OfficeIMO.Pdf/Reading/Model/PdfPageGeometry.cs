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

    /// <summary>Effective page box used by OfficeIMO.Pdf for page size, preferring CropBox over MediaBox.</summary>
    public PdfPageBox? EffectiveBox => CropBox ?? MediaBox;

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
