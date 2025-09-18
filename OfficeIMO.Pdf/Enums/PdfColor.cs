namespace OfficeIMO.Pdf;

/// <summary>
/// Represents an RGB color where each component is expressed in the 0..1 range.
/// Use with writer APIs to color text and vector shapes.
/// </summary>
public readonly struct PdfColor {
    /// <summary>
    /// Red component in the 0..1 range.
    /// </summary>
    public double R { get; }

    /// <summary>
    /// Green component in the 0..1 range.
    /// </summary>
    public double G { get; }

    /// <summary>
    /// Blue component in the 0..1 range.
    /// </summary>
    public double B { get; }

    /// <summary>
    /// Initializes a new <see cref="PdfColor"/>.
    /// </summary>
    /// <param name="r">Red component (0..1).</param>
    /// <param name="g">Green component (0..1).</param>
    /// <param name="b">Blue component (0..1).</param>
    public PdfColor(double r, double g, double b) { R = r; G = g; B = b; }

    /// <summary>
    /// Creates a color from 0..255 byte RGB components.
    /// </summary>
    public static PdfColor FromRgb(byte r, byte g, byte b) => new PdfColor(r / 255.0, g / 255.0, b / 255.0);

    /// <summary>
    /// Pure black color (#000000).
    /// </summary>
    public static PdfColor Black => new PdfColor(0, 0, 0);

    /// <summary>
    /// Pure white color (#FFFFFF).
    /// </summary>
    public static PdfColor White => new PdfColor(1, 1, 1);

    /// <summary>
    /// Light gray color useful for fills and separators.
    /// </summary>
    public static PdfColor LightGray => new PdfColor(0.9, 0.9, 0.9);

    /// <summary>
    /// Mid gray color useful for text and borders.
    /// </summary>
    public static PdfColor Gray => new PdfColor(0.5, 0.5, 0.5);
}
