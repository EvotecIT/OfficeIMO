using OfficeIMO.Drawing;

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
    public PdfColor(double r, double g, double b) {
        ValidateComponent(r, nameof(r));
        ValidateComponent(g, nameof(g));
        ValidateComponent(b, nameof(b));
        R = r; G = g; B = b;
    }

    /// <summary>
    /// Creates a color from 0..255 byte RGB components.
    /// </summary>
    public static PdfColor FromRgb(byte r, byte g, byte b) => new PdfColor(r / 255.0, g / 255.0, b / 255.0);

    /// <summary>
    /// Creates a PDF RGB color from a shared OfficeIMO color. Alpha is ignored because
    /// the current PDF writer does not yet support transparency.
    /// </summary>
    public static PdfColor FromOfficeColor(OfficeColor color) => FromRgb(color.R, color.G, color.B);

    /// <summary>
    /// Creates a PDF RGB color from a shared OfficeIMO color, returning null for fully transparent colors.
    /// Partially transparent colors are converted to their RGB components until PDF transparency is supported.
    /// </summary>
    public static PdfColor? FromOfficeColorOrNull(OfficeColor color) => color.A == 0 ? (PdfColor?)null : FromOfficeColor(color);

    /// <summary>
    /// Converts this color to the shared OfficeIMO color type.
    /// </summary>
    public OfficeColor ToOfficeColor(byte alpha = 255) => OfficeColor.FromRgba(ToByte(R), ToByte(G), ToByte(B), alpha);

    /// <summary>
    /// Converts a shared OfficeIMO color to a PDF RGB color.
    /// </summary>
    public static implicit operator PdfColor(OfficeColor color) => FromOfficeColor(color);

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

    private static void ValidateComponent(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0 || value > 1) {
            throw new ArgumentOutOfRangeException(paramName, "PDF color components must be finite values between 0 and 1.");
        }
    }

    private static byte ToByte(double value) {
        if (value <= 0) {
            return 0;
        }

        if (value >= 1) {
            return 255;
        }

        return (byte)Math.Round(value * 255.0);
    }
}
