namespace OfficeIMO.Pdf;

/// <summary>Represents page margins in points.</summary>
public readonly struct PageMargins {
    /// <summary>Left margin in points.</summary>
    public double Left { get; }
    /// <summary>Top margin in points.</summary>
    public double Top { get; }
    /// <summary>Right margin in points.</summary>
    public double Right { get; }
    /// <summary>Bottom margin in points.</summary>
    public double Bottom { get; }

    /// <summary>Creates page margins in points.</summary>
    public PageMargins(double left, double top, double right, double bottom) {
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
    public static PageMargins Uniform(double points) => new PageMargins(points, points, points, points);

    /// <summary>Creates uniform page margins in inches.</summary>
    public static PageMargins UniformInches(double inches) => Uniform(ToPointsFromInches(inches, nameof(inches)));

    /// <summary>Creates page margins in inches.</summary>
    public static PageMargins FromInches(double left, double top, double right, double bottom) =>
        new PageMargins(
            ToPointsFromInches(left, nameof(left)),
            ToPointsFromInches(top, nameof(top)),
            ToPointsFromInches(right, nameof(right)),
            ToPointsFromInches(bottom, nameof(bottom)));

    /// <summary>Creates uniform page margins in centimeters.</summary>
    public static PageMargins UniformCentimeters(double centimeters) => Uniform(ToPointsFromCentimeters(centimeters, nameof(centimeters)));

    /// <summary>Creates page margins in centimeters.</summary>
    public static PageMargins FromCentimeters(double left, double top, double right, double bottom) =>
        new PageMargins(
            ToPointsFromCentimeters(left, nameof(left)),
            ToPointsFromCentimeters(top, nameof(top)),
            ToPointsFromCentimeters(right, nameof(right)),
            ToPointsFromCentimeters(bottom, nameof(bottom)));

    /// <summary>Normal Word-compatible one-inch margins.</summary>
    public static PageMargins Normal => Uniform(72);

    /// <summary>Narrow Word-compatible half-inch margins.</summary>
    public static PageMargins Narrow => Uniform(36);

    /// <summary>Moderate Word-compatible margins.</summary>
    public static PageMargins Moderate => new PageMargins(54, 72, 54, 72);

    /// <summary>Wide Word-compatible margins.</summary>
    public static PageMargins Wide => new PageMargins(144, 72, 144, 72);

    /// <summary>Mirrored Word-compatible margin preset values.</summary>
    public static PageMargins Mirrored => new PageMargins(90, 72, 72, 72);

    /// <summary>Word 2003 compatible default margin preset values.</summary>
    public static PageMargins Office2003Default => new PageMargins(90, 72, 90, 72);

    private static double ToPointsFromInches(double inches, string paramName) {
        Guard.NonNegative(inches, paramName);
        return inches * 72D;
    }

    private static double ToPointsFromCentimeters(double centimeters, string paramName) {
        Guard.NonNegative(centimeters, paramName);
        return centimeters * 72D / 2.54D;
    }
}
