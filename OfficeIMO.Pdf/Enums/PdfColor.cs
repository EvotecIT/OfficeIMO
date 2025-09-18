namespace OfficeIMO.Pdf;

/// <summary>
/// RGB color (0..1 range) for text and shapes.
/// </summary>
public readonly struct PdfColor {
    public double R { get; }
    public double G { get; }
    public double B { get; }
    public PdfColor(double r, double g, double b) { R = r; G = g; B = b; }
    public static PdfColor FromRgb(byte r, byte g, byte b) => new PdfColor(r / 255.0, g / 255.0, b / 255.0);
    public static PdfColor Black => new PdfColor(0, 0, 0);
    public static PdfColor White => new PdfColor(1, 1, 1);
    public static PdfColor LightGray => new PdfColor(0.9, 0.9, 0.9);
    public static PdfColor Gray => new PdfColor(0.5, 0.5, 0.5);
}
