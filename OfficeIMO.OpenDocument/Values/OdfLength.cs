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

    /// <summary>Converts this absolute ODF length to points.</summary>
    public double ToPoints() {
        string lexical = ToString();
        int split = 0;
        while (split < lexical.Length && (char.IsDigit(lexical[split]) || lexical[split] == '-' || lexical[split] == '+' || lexical[split] == '.')) split++;
        if (split == 0 || !double.TryParse(lexical.Substring(0, split), NumberStyles.Float, CultureInfo.InvariantCulture, out double value)) {
            throw new FormatException($"ODF length '{lexical}' is not an absolute numeric length.");
        }
        string unit = lexical.Substring(split).ToLowerInvariant();
        switch (unit) {
            case "pt": return value;
            case "in": return value * 72D;
            case "cm": return value * 72D / 2.54D;
            case "mm": return value * 72D / 25.4D;
            case "pc": return value * 12D;
            default: throw new NotSupportedException($"ODF length unit '{unit}' cannot be converted to points.");
        }
    }

    /// <summary>Converts this absolute ODF length to centimeters.</summary>
    public double ToCentimeters() => ToPoints() * 2.54D / 72D;

    /// <inheritdoc />
    public override string ToString() => _value ?? "0cm";
    /// <inheritdoc />
    public bool Equals(OdfLength other) => string.Equals(ToString(), other.ToString(), StringComparison.Ordinal);
    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OdfLength other && Equals(other);
    /// <inheritdoc />
    public override int GetHashCode() => StringComparer.Ordinal.GetHashCode(ToString());
}
