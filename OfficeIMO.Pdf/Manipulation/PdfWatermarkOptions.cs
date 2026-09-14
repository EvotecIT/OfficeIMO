namespace OfficeIMO.Pdf;

/// <summary>Visual text or image watermark placed in a rectangle on existing PDF pages.</summary>
public sealed class PdfWatermarkOptions {
    /// <summary>Stable identifier for this watermark. Reusing it revises the existing watermark; use a new identifier to add another.</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    /// <summary>Text to draw when <see cref="ImageBytes"/> is null.</summary>
    public string Text { get; set; } = "CONFIDENTIAL";

    /// <summary>Optional encoded image. When supplied, it replaces the text appearance.</summary>
    public byte[]? ImageBytes { get; set; }

    /// <summary>Target pages; null selects every page.</summary>
    public PdfPageSelector? TargetPages { get; set; }

    /// <summary>Left edge in visual page points; null centers the rectangle horizontally.</summary>
    public double? X { get; set; }

    /// <summary>Top edge in visual page points; null centers the rectangle vertically.</summary>
    public double? Y { get; set; }

    /// <summary>Unrotated rectangle width in points.</summary>
    public double Width { get; set; } = 360D;

    /// <summary>Unrotated rectangle height in points.</summary>
    public double Height { get; set; } = 100D;

    /// <summary>Clockwise rotation around the rectangle center, in degrees.</summary>
    public double RotationDegrees { get; set; } = -35D;

    /// <summary>Opacity between zero and one, for either text or image content.</summary>
    public double Opacity { get; set; } = 0.35D;

    /// <summary>Draws the watermark beneath existing page content when true.</summary>
    public bool BehindContent { get; set; }

    /// <summary>Text size in points.</summary>
    public double FontSize { get; set; } = 42D;

    /// <summary>Standard text font.</summary>
    public PdfStandardFont Font { get; set; } = PdfStandardFont.HelveticaBold;

    /// <summary>Text color. Image pixels retain their own colors.</summary>
    public PdfColor Color { get; set; } = PdfColor.Gray;

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
        if (X.HasValue) Finite(X.Value, nameof(X));
        if (Y.HasValue) Finite(Y.Value, nameof(Y));
        _ = new PdfCanvasStampOptions { Opacity = Opacity };
        _ = new PdfCanvasTextBoxStyle { Font = Font };
        if (ImageBytes is null && string.IsNullOrWhiteSpace(Text))
            throw new ArgumentException("Watermark text is required.", nameof(Text));
        if (ImageBytes is { Length: 0 }) throw new ArgumentException("Watermark image is empty.", nameof(ImageBytes));
    }

    private static void Positive(double value, string name) {
        Finite(value, name);
        if (value <= 0D) throw new ArgumentOutOfRangeException(name, "Watermark dimensions and font size must be positive.");
    }

    private static void Finite(double value, string name) {
        if (double.IsNaN(value) || double.IsInfinity(value))
            throw new ArgumentOutOfRangeException(name, "Watermark geometry must be finite.");
    }
}
