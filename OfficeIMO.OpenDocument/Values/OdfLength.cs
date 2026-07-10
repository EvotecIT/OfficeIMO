namespace OfficeIMO.OpenDocument;

/// <summary>Preserves an ODF length's invariant lexical representation.</summary>
public readonly struct OdfLength : IEquatable<OdfLength> {
    private readonly string? _value;

    private OdfLength(string value) {
        _value = value;
    }

    /// <summary>Creates a length from an ODF lexical value such as `2.54cm` or `12pt`.</summary>
    public static OdfLength Parse(string value) {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Length cannot be empty.", nameof(value));
        return new OdfLength(value.Trim());
    }

    /// <summary>Creates a centimeter length.</summary>
    public static OdfLength Centimeters(double value) => new OdfLength(value.ToString("0.###", CultureInfo.InvariantCulture) + "cm");
    /// <summary>Creates an inch length.</summary>
    public static OdfLength Inches(double value) => new OdfLength(value.ToString("0.###", CultureInfo.InvariantCulture) + "in");
    /// <summary>Creates a point length.</summary>
    public static OdfLength Points(double value) => new OdfLength(value.ToString("0.###", CultureInfo.InvariantCulture) + "pt");

    /// <inheritdoc />
    public override string ToString() => _value ?? "0cm";
    /// <inheritdoc />
    public bool Equals(OdfLength other) => string.Equals(ToString(), other.ToString(), StringComparison.Ordinal);
    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OdfLength other && Equals(other);
    /// <inheritdoc />
    public override int GetHashCode() => StringComparer.Ordinal.GetHashCode(ToString());
}
