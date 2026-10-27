namespace OfficeIMO.Pdf;

/// <summary>Represents a page size in points.</summary>
public readonly struct PdfPageSize {
    /// <summary>Width in points.</summary>
    public double Width { get; }
    /// <summary>Height in points.</summary>
    public double Height { get; }
    /// <summary>Creates a new page size.</summary>
    public PdfPageSize(double width, double height) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Width = width;
        Height = height;
    }

    /// <summary>Creates a page size from dimensions in inches.</summary>
    public static PdfPageSize FromInches(double width, double height) =>
        new PdfPageSize(ToPointsFromInches(width, nameof(width)), ToPointsFromInches(height, nameof(height)));

    /// <summary>Creates a page size from dimensions in centimeters.</summary>
    public static PdfPageSize FromCentimeters(double width, double height) =>
        new PdfPageSize(ToPointsFromCentimeters(width, nameof(width)), ToPointsFromCentimeters(height, nameof(height)));

    /// <summary>Returns this size in portrait orientation.</summary>
    public PdfPageSize Portrait() => Width <= Height ? this : new PdfPageSize(Height, Width);

    /// <summary>Returns this size in landscape orientation.</summary>
    public PdfPageSize Landscape() => Width >= Height ? this : new PdfPageSize(Height, Width);

    /// <summary>Returns this size in the requested orientation.</summary>
    public PdfPageSize WithOrientation(OfficePageOrientation orientation) {
        Guard.PageOrientation(orientation, nameof(orientation));
        return orientation == OfficePageOrientation.Landscape ? Landscape() : Portrait();
    }

    private static double ToPointsFromInches(double inches, string paramName) {
        Guard.Positive(inches, paramName);
        return inches * 72D;
    }

    private static double ToPointsFromCentimeters(double centimeters, string paramName) {
        Guard.Positive(centimeters, paramName);
        return centimeters * 72D / 2.54D;
    }
}
