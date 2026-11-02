namespace OfficeIMO.Pdf;

/// <summary>Visual text or image watermark placed in a rectangle on existing PDF pages.</summary>
public sealed class PdfWatermarkOptions {
    /// <summary>Cancels input reading, PDF processing, and output publication for this operation.</summary>
    public System.Threading.CancellationToken CancellationToken { get; set; }

    /// <summary>Stable identifier for this watermark. Reusing it revises the existing watermark; use a new identifier to add another.</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    /// <summary>Text to draw when <see cref="ImageBytes"/> is null.</summary>
    public string Text { get; set; } = "CONFIDENTIAL";

    /// <summary>Optional encoded image. When supplied, it replaces the text appearance.</summary>
    public byte[]? ImageBytes { get; set; }

    /// <summary>Target pages; null selects every page.</summary>
    public PdfPageSelector? TargetPages { get; set; }

    /// <summary>Nonnegative physical left offset on the visual page; null centers the rectangle horizontally.</summary>
    public global::OfficeIMO.OfficeLength? X { get; set; }

    /// <summary>Nonnegative physical top offset on the visual page; null centers the rectangle vertically.</summary>
    public global::OfficeIMO.OfficeLength? Y { get; set; }

    /// <summary>Physical unrotated rectangle width. Defaults to 360 points.</summary>
    public global::OfficeIMO.OfficeLength Width { get; set; } = global::OfficeIMO.OfficeLength.Points(360D);

    /// <summary>Physical unrotated rectangle height. Defaults to 100 points.</summary>
    public global::OfficeIMO.OfficeLength Height { get; set; } = global::OfficeIMO.OfficeLength.Points(100D);

    /// <summary>Clockwise rotation around the rectangle center, in degrees.</summary>
    public double RotationDegrees { get; set; } = -35D;

    /// <summary>Opacity between zero and one, for either text or image content.</summary>
    public double Opacity { get; set; } = 0.35D;

    /// <summary>Draws the watermark beneath existing page content when true.</summary>
    public bool BehindContent { get; set; }

    /// <summary>Physical text size. Defaults to 42 points.</summary>
    public global::OfficeIMO.OfficeLength FontSize { get; set; } = global::OfficeIMO.OfficeLength.Points(42D);

    /// <summary>Standard text font.</summary>
    public PdfStandardFont Font { get; set; } = PdfStandardFont.HelveticaBold;

    /// <summary>Text color. Image pixels retain their own colors.</summary>
    public global::OfficeIMO.OfficeColor Color { get; set; } = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.5D, 0.5D, 0.5D);

    /// <summary>Creates an independent copy, including any encoded image bytes.</summary>
    public PdfWatermarkOptions Clone() {
        var copy = (PdfWatermarkOptions)MemberwiseClone();
        copy.ImageBytes = ImageBytes is null ? null : (byte[])ImageBytes.Clone();
        return copy;
    }

    internal void Validate() {
        if (!Guid.TryParseExact(Id, "N", out _))
            throw new ArgumentException("Watermark identifier must be a GUID in N format.", nameof(Id));
        Positive(Width, nameof(Width));
        Positive(Height, nameof(Height));
        Positive(FontSize, nameof(FontSize));
        Finite(RotationDegrees, nameof(RotationDegrees));
        _ = new PdfCanvasStampOptions { Opacity = Opacity };
        _ = new PdfCanvasTextBoxStyle { Font = Font };
        if (ImageBytes is null && string.IsNullOrWhiteSpace(Text))
            throw new ArgumentException("Watermark text is required.", nameof(Text));
        if (ImageBytes is { Length: 0 }) throw new ArgumentException("Watermark image is empty.", nameof(ImageBytes));
    }

    private static void Positive(global::OfficeIMO.OfficeLength value, string name) {
        if (value <= global::OfficeIMO.OfficeLength.Zero) throw new ArgumentOutOfRangeException(name, "Watermark dimensions and font size must be positive.");
    }

    private static void Finite(double value, string name) {
        if (double.IsNaN(value) || double.IsInfinity(value))
            throw new ArgumentOutOfRangeException(name, "Watermark geometry must be finite.");
    }
}
