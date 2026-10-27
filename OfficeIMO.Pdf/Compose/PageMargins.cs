namespace OfficeIMO.Pdf;

/// <summary>Represents page margins in points.</summary>
public readonly struct PdfPageMargins {
    /// <summary>Left margin in points.</summary>
    public double Left { get; }
    /// <summary>Top margin in points.</summary>
    public double Top { get; }
    /// <summary>Right margin in points.</summary>
    public double Right { get; }
    /// <summary>Bottom margin in points.</summary>
    public double Bottom { get; }

    /// <summary>Creates page margins in points.</summary>
    public PdfPageMargins(double left, double top, double right, double bottom) {
        Guard.NonNegative(left, nameof(left));
        Guard.NonNegative(top, nameof(top));
        Guard.NonNegative(right, nameof(right));
        Guard.NonNegative(bottom, nameof(bottom));
        Left = left;
        Top = top;
        Right = right;
        Bottom = bottom;
    }

    /// <summary>Creates uniform page margins in points.</summary>
    public static PdfPageMargins Uniform(double points) => new PdfPageMargins(points, points, points, points);

    /// <summary>Creates uniform page margins in inches.</summary>
    public static PdfPageMargins UniformInches(double inches) => Uniform(ToPointsFromInches(inches, nameof(inches)));

    /// <summary>Creates page margins in inches.</summary>
    public static PdfPageMargins FromInches(double left, double top, double right, double bottom) =>
        new PdfPageMargins(
            ToPointsFromInches(left, nameof(left)),
            ToPointsFromInches(top, nameof(top)),
            ToPointsFromInches(right, nameof(right)),
            ToPointsFromInches(bottom, nameof(bottom)));

    /// <summary>Creates uniform page margins in centimeters.</summary>
    public static PdfPageMargins UniformCentimeters(double centimeters) => Uniform(ToPointsFromCentimeters(centimeters, nameof(centimeters)));

    /// <summary>Creates page margins in centimeters.</summary>
    public static PdfPageMargins FromCentimeters(double left, double top, double right, double bottom) =>
        new PdfPageMargins(
            ToPointsFromCentimeters(left, nameof(left)),
            ToPointsFromCentimeters(top, nameof(top)),
            ToPointsFromCentimeters(right, nameof(right)),
            ToPointsFromCentimeters(bottom, nameof(bottom)));

    /// <summary>Normal Word-compatible one-inch margins.</summary>
    public static PdfPageMargins Normal => Uniform(72);

    /// <summary>Narrow Word-compatible half-inch margins.</summary>
    public static PdfPageMargins Narrow => Uniform(36);

    /// <summary>Moderate Word-compatible margins.</summary>
    public static PdfPageMargins Moderate => new PdfPageMargins(54, 72, 54, 72);

    /// <summary>Wide Word-compatible margins.</summary>
    public static PdfPageMargins Wide => new PdfPageMargins(144, 72, 144, 72);

    /// <summary>Mirrored Word-compatible margin preset values.</summary>
    public static PdfPageMargins Mirrored => new PdfPageMargins(90, 72, 72, 72);

    /// <summary>Word 2003 compatible default margin preset values.</summary>
    public static PdfPageMargins Office2003Default => new PdfPageMargins(90, 72, 90, 72);

    private static double ToPointsFromInches(double inches, string paramName) {
        Guard.NonNegative(inches, paramName);
        return inches * 72D;
    }

    private static double ToPointsFromCentimeters(double centimeters, string paramName) {
        Guard.NonNegative(centimeters, paramName);
        return centimeters * 72D / 2.54D;
    }
}
