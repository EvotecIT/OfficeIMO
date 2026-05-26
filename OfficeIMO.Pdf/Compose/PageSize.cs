namespace OfficeIMO.Pdf;

/// <summary>Represents a page size in points.</summary>
public readonly struct PageSize {
    /// <summary>Width in points.</summary>
    public double Width { get; }
    /// <summary>Height in points.</summary>
    public double Height { get; }
    /// <summary>Creates a new page size.</summary>
    public PageSize(double width, double height) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Width = width;
        Height = height;
    }

    /// <summary>Creates a page size from dimensions in inches.</summary>
    public static PageSize FromInches(double width, double height) =>
        new PageSize(ToPointsFromInches(width, nameof(width)), ToPointsFromInches(height, nameof(height)));

    /// <summary>Creates a page size from dimensions in centimeters.</summary>
    public static PageSize FromCentimeters(double width, double height) =>
        new PageSize(ToPointsFromCentimeters(width, nameof(width)), ToPointsFromCentimeters(height, nameof(height)));

    /// <summary>Returns this size in portrait orientation.</summary>
    public PageSize Portrait() => Width <= Height ? this : new PageSize(Height, Width);

    /// <summary>Returns this size in landscape orientation.</summary>
    public PageSize Landscape() => Width >= Height ? this : new PageSize(Height, Width);

    /// <summary>Returns this size in the requested orientation.</summary>
    public PageSize WithOrientation(PdfPageOrientation orientation) {
        Guard.PageOrientation(orientation, nameof(orientation));
        return orientation == PdfPageOrientation.Landscape ? Landscape() : Portrait();
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
