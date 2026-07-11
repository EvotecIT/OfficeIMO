namespace OfficeIMO.OpenDocument;

/// <summary>Position and size expressed as ODF lengths.</summary>
public readonly struct OdfRect {
    /// <summary>Creates a rectangle.</summary>
    public OdfRect(OdfLength x, OdfLength y, OdfLength width, OdfLength height) {
        X = x;
        Y = y;
        Width = width;
        Height = height;
    }

    /// <summary>Horizontal position.</summary>
    public OdfLength X { get; }
    /// <summary>Vertical position.</summary>
    public OdfLength Y { get; }
    /// <summary>Width.</summary>
    public OdfLength Width { get; }
    /// <summary>Height.</summary>
    public OdfLength Height { get; }

    /// <summary>Creates a rectangle in centimeters.</summary>
    public static OdfRect FromCentimeters(double x, double y, double width, double height) =>
        new OdfRect(OdfLength.Centimeters(x), OdfLength.Centimeters(y), OdfLength.Centimeters(width), OdfLength.Centimeters(height));
}
