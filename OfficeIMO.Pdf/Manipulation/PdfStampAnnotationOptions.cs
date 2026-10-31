namespace OfficeIMO.Pdf;

/// <summary>Placement, appearance, and mutation policy for a stamp annotation added to an existing PDF page.</summary>
/// <remarks>Native page coordinates use default user-space units: one unit is /UserUnit physical points (1/72 inch each), defaulting to 1. Local text/font and matrix values retain their native coordinate spaces. Explicit visual projections use their documented coordinates.</remarks>
public sealed class PdfStampAnnotationOptions {
    /// <summary>Cancels input reading, PDF processing, and output publication for this operation.</summary>
    public System.Threading.CancellationToken CancellationToken { get; set; }

    private byte[]? _imageBytes;
    /// <summary>One-based target page number.</summary>
    public int PageNumber { get; set; } = 1;

    /// <summary>Left edge in page user-space units.</summary>
    public double X { get; set; } = 36D;

    /// <summary>Bottom edge in page user-space units.</summary>
    public double Y { get; set; } = 36D;

    /// <summary>Annotation width in page user-space units.</summary>
    public double Width { get; set; } = 144D;

    /// <summary>Annotation height in page user-space units.</summary>
    public double Height { get; set; } = 48D;

    /// <summary>Standard or custom PDF stamp name, for example Approved, Draft, or TopSecret.</summary>
    public string StampName { get; set; } = "Approved";

    /// <summary>Optional human-readable annotation contents.</summary>
    public string? Contents { get; set; }

    /// <summary>Optional annotation author/title.</summary>
    public string? Title { get; set; }

    /// <summary>Optional stable annotation name stored in /NM.</summary>
    public string? Name { get; set; }

    /// <summary>Annotation flags. The default value 4 enables printing.</summary>
    public int Flags { get; set; } = 4;

    /// <summary>Border and label color. The default is dark red.</summary>
    public global::OfficeIMO.OfficeColor StrokeColor { get; set; } = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.7D, 0.05D, 0.05D);

    /// <summary>Optional appearance fill color.</summary>
    public global::OfficeIMO.OfficeColor? FillColor { get; set; }

    /// <summary>Border width in page user-space units.</summary>
    public double BorderWidth { get; set; } = 2D;

    /// <summary>Optional PNG or JPEG appearance. When set, the stamp displays this image instead of its text label. This is a visual annotation, not a cryptographic signature.</summary>
    public byte[]? ImageBytes {
        get => _imageBytes is null ? null : (byte[])_imageBytes.Clone();
        set => _imageBytes = value is null ? null : (byte[])value.Clone();
    }

    /// <summary>Maximum encoded image bytes accepted for an image-backed stamp. Defaults to 128 MiB.</summary>
    public long MaximumEncodedImageBytes { get; set; } = PdfImageInput.DefaultMaximumEncodedBytes;

    internal byte[]? GetImageBytes() => _imageBytes;

    /// <summary>Preferred mutation mode. Automatic uses append-only when signature policy requires it.</summary>
    public PdfMutationExecutionPreference ExecutionPreference { get; set; } = PdfMutationExecutionPreference.Automatic;
}
